$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$script:PidFile = Join-Path $PSScriptRoot ".litera-server.pid"
$script:ProjectRoot = $PSScriptRoot

function Test-CommandExists {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name
  )
  return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Write-Header {
  Write-Host ""
  Write-Host "==============================================" -ForegroundColor DarkCyan
  Write-Host " Litera One Control Center" -ForegroundColor Cyan
  Write-Host " Project: $script:ProjectRoot"
  Write-Host "==============================================" -ForegroundColor DarkCyan
  Write-Host ""
}

function Get-EnvTemplatePath {
  $candidates = @(".env.example", ".exmaple.env", ".example.env")
  foreach ($name in $candidates) {
    $path = Join-Path $script:ProjectRoot $name
    if (Test-Path $path) { return $path }
  }
  return $null
}

function Ensure-EnvFile {
  $envPath = Join-Path $script:ProjectRoot ".env"
  if (Test-Path $envPath) {
    Write-Host ".env already exists." -ForegroundColor Green
    return $true
  }

  $templatePath = Get-EnvTemplatePath
  if (-not $templatePath) {
    Write-Host "Missing .env and no template file found (.env.example/.exmaple.env/.example.env)." -ForegroundColor Red
    return $false
  }

  Copy-Item $templatePath $envPath -Force
  Write-Host "Created .env from template: $(Split-Path $templatePath -Leaf)" -ForegroundColor Yellow
  return $true
}

function Get-EnvVarValue {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $envPath = Join-Path $script:ProjectRoot ".env"
  if (-not (Test-Path $envPath)) { return "" }

  $line = Get-Content $envPath | Where-Object { $_ -match "^\s*$Name\s*=" } | Select-Object -First 1
  if (-not $line) { return "" }

  $value = ($line -split "=", 2)[1].Trim()
  return $value.Trim('"').Trim("'").Trim()
}

function Test-EnvCredentials {
  $placeholders = @(
    "",
    "your-client-id",
    "your-client-secret",
    "<your-client-id>",
    "<your-client-secret>",
    "changeme"
  )

  $clientId = Get-EnvVarValue -Name "ENTRA_CLIENT_ID"
  $clientSecret = Get-EnvVarValue -Name "ENTRA_CLIENT_SECRET"

  $okClientId = -not ($placeholders -contains $clientId.ToLowerInvariant())
  $okClientSecret = -not ($placeholders -contains $clientSecret.ToLowerInvariant())

  if ($okClientId -and $okClientSecret) {
    Write-Host ".env validation passed: ENTRA_CLIENT_ID and ENTRA_CLIENT_SECRET are set." -ForegroundColor Green
    return $true
  }

  Write-Host ".env validation failed." -ForegroundColor Red
  if (-not $okClientId) { Write-Host " - ENTRA_CLIENT_ID is missing or placeholder." -ForegroundColor Yellow }
  if (-not $okClientSecret) { Write-Host " - ENTRA_CLIENT_SECRET is missing or placeholder." -ForegroundColor Yellow }
  return $false
}

function Get-PackageMap {
  $packagePath = Join-Path $script:ProjectRoot "package.json"
  if (-not (Test-Path $packagePath)) { throw "package.json not found." }

  $pkg = Get-Content $packagePath -Raw | ConvertFrom-Json
  $map = @{}

  if ($pkg.dependencies) {
    foreach ($p in $pkg.dependencies.PSObject.Properties) { $map[$p.Name] = $p.Value }
  }
  if ($pkg.devDependencies) {
    foreach ($p in $pkg.devDependencies.PSObject.Properties) { $map[$p.Name] = $p.Value }
  }

  return $map
}

function Get-PackageFolder {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PackageName
  )

  $segments = $PackageName -split "/"
  $path = Join-Path $script:ProjectRoot "node_modules"
  foreach ($segment in $segments) {
    $path = Join-Path $path $segment
  }
  return $path
}

function Get-MissingDependencies {
  $deps = Get-PackageMap
  $missing = @()

  foreach ($dep in $deps.Keys) {
    $depFolder = Get-PackageFolder -PackageName $dep
    $depPackageJson = Join-Path $depFolder "package.json"
    if (-not (Test-Path $depPackageJson)) {
      $missing += $dep
    }
  }

  return @($missing | Sort-Object)
}

function Check-Dependencies {
  if (-not (Test-CommandExists -Name "node")) {
    Write-Host "Node.js is not installed or not in PATH." -ForegroundColor Red
    return $false
  }
  if (-not (Test-CommandExists -Name "npm")) {
    Write-Host "npm is not installed or not in PATH." -ForegroundColor Red
    return $false
  }

  $missing = @(Get-MissingDependencies)
  if (@($missing).Count -eq 0) {
    Write-Host "Dependencies check passed. No missing dependencies." -ForegroundColor Green
    return $true
  }

  Write-Host "Dependencies check failed. Missing $(@($missing).Count) package(s)." -ForegroundColor Yellow
  return $false
}

function List-MissingDependencies {
  try {
    $missing = @(Get-MissingDependencies)
    if (@($missing).Count -eq 0) {
      Write-Host "No missing dependencies." -ForegroundColor Green
      return
    }

    Write-Host "Missing dependencies:" -ForegroundColor Yellow
    foreach ($dep in $missing) {
      Write-Host " - $dep"
    }
  } catch {
    Write-Host "Failed to list dependencies: $($_.Exception.Message)" -ForegroundColor Red
  }
}

function Install-Dependencies {
  if (-not (Test-CommandExists -Name "npm")) {
    Write-Host "npm is not installed or not in PATH." -ForegroundColor Red
    return
  }

  Write-Host "Installing dependencies (npm install)..." -ForegroundColor Cyan
  Push-Location $script:ProjectRoot
  try {
    npm install
    Write-Host "Dependency installation completed." -ForegroundColor Green
  } finally {
    Pop-Location
  }
}

function Get-ServerProcess {
  if (-not (Test-Path $script:PidFile)) { return $null }

  $raw = (Get-Content $script:PidFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
  if (-not $raw) {
    Remove-Item $script:PidFile -ErrorAction SilentlyContinue
    return $null
  }

  $processId = 0
  if (-not [int]::TryParse($raw, [ref]$processId)) {
    Remove-Item $script:PidFile -ErrorAction SilentlyContinue
    return $null
  }

  $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
  if (-not $proc) {
    Remove-Item $script:PidFile -ErrorAction SilentlyContinue
    return $null
  }

  return $proc
}

function Start-Server {
  if (-not (Test-CommandExists -Name "node")) {
    Write-Host "Node.js is not installed or not in PATH." -ForegroundColor Red
    return
  }

  $envPort = Get-EnvVarValue -Name "PORT"
  $port = if ([string]::IsNullOrWhiteSpace($envPort)) { 3000 } else { [int]$envPort }

  $existing = Get-ServerProcess
  if ($existing) {
    Write-Host "Server already running (PID $($existing.Id))." -ForegroundColor Yellow
    Write-Host "URL: http://localhost:$port"
    return
  }

  if (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue) {
    $portInUse = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
    if ($portInUse) {
      Write-Host "Port $port is already in use." -ForegroundColor Red
      return
    }
  }

  if (-not (Ensure-EnvFile)) { return }
  if (-not (Test-EnvCredentials)) { return }
  if (-not (Check-Dependencies)) {
    Write-Host "Install dependencies first (menu option 3)." -ForegroundColor Yellow
    return
  }

  $proc = Start-Process -FilePath "node" -ArgumentList "--experimental-strip-types src/server.ts" -WorkingDirectory $script:ProjectRoot -PassThru -NoNewWindow
  Set-Content -Path $script:PidFile -Value $proc.Id -NoNewline

  Start-Sleep -Seconds 1
  $running = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
  if (-not $running) {
    Remove-Item $script:PidFile -ErrorAction SilentlyContinue
    Write-Host "Server failed to start. Run 'npm start' manually for logs." -ForegroundColor Red
    return
  }

  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$ts] Server started (PID $($proc.Id))." -ForegroundColor Green
  Write-Host "Server running at: http://localhost:$port" -ForegroundColor Cyan
  Write-Host "Press Ctrl+C to stop the server..." -ForegroundColor Yellow

  try {
    Wait-Process -Id $proc.Id
  } finally {
    Remove-Item $script:PidFile -ErrorAction SilentlyContinue
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Server stopped." -ForegroundColor Green
  }
}

function Stop-Server {
  $proc = Get-ServerProcess
  if (-not $proc) {
    Write-Host "Server is not running." -ForegroundColor Yellow
    return
  }

  Stop-Process -Id $proc.Id -ErrorAction SilentlyContinue
  Start-Sleep -Milliseconds 600
  Remove-Item $script:PidFile -ErrorAction SilentlyContinue
  Write-Host "Server stopped (PID $($proc.Id))." -ForegroundColor Green
}

function Restart-Server {
  Stop-Server
  Start-Server
}

function Show-Menu {
  Write-Host ""
  Write-Host "Choose an option:" -ForegroundColor Cyan
  Write-Host " 1) Check dependencies"
  Write-Host " 2) List missing dependencies"
  Write-Host " 3) Install dependencies"
  Write-Host " 4) Create .env from template if missing"
  Write-Host " 5) Validate .env client id & secret"
  Write-Host " 6) Start server"
  Write-Host " 7) Stop server"
  Write-Host " 8) Restart server"
  Write-Host " 9) Exit"
  Write-Host " c) Clear Screen"
}

Push-Location $script:ProjectRoot
try {
  Write-Header
  Show-Menu

  while ($true) {
    Write-Host ""
    $choice = Read-Host "Enter option number (or 'm' for menu)"

    switch ($choice) {
      "m" { Show-Menu }
      "c" { Clear-Host; Write-Header; Show-Menu }
      "1" { [void](Check-Dependencies) }
      "2" { List-MissingDependencies }
      "3" { Install-Dependencies }
      "4" { [void](Ensure-EnvFile) }
      "5" {
        if (Ensure-EnvFile) { [void](Test-EnvCredentials) }
      }
      "6" { Start-Server }
      "7" {
        $confirm = Read-Host "Are you sure you want to stop the server? (y/N)"
        if ($confirm -match "^[yY]") { Stop-Server }
      }
      "8" { Restart-Server }
      "9" {
        $confirm = Read-Host "Are you sure you want to exit? (y/N)"
        if ($confirm -match "^[yY]") {
          Write-Host "Exiting Control Center." -ForegroundColor Cyan
          break
        }
      }
      default { Write-Host "Invalid option. Choose 1-9." -ForegroundColor Yellow }
    }
  }
}
finally {
  Pop-Location
}
