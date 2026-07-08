# Litera One Manager Portal

## 📖 What It Does

The **Litera One Manager Portal** is a full-featured Microsoft Entra (Azure AD) administration web portal designed specifically for managing Litera One enterprise applications. 
It simplifies complex IT administrative tasks by providing an easy-to-use graphical interface for:
- **User & Group Provisioning**: Configuring and managing SCIM sync jobs.
- **Enterprise App Management**: Creating, deleting, and assigning users to Litera apps.
- **Add-in Deployments**: Seamlessly deploying Outlook or Word add-ins across the organization.
- **Enablement Resources**: Browsing and exporting training materials for users.

Built with **React + Vite**. Authenticates securely against the Microsoft Graph API using the **OAuth 2.0 Client Credentials** flow (Client ID + Client Secret).

---

## 🚀 How to Setup (Quick Start)

### Prerequisites
- Node.js 22+ (required for backend TypeScript runtime via `--experimental-strip-types`)
- An Azure AD App Registration with the required API permissions (see below)

### Install & Run

```bash
npm install
npm run dev
```

Open http://localhost:3000

### Build for Production
```bash
npm run build
npm run preview
```

### Run with Docker

For comprehensive details on building and deploying the image, please refer to the [Docker Build Instructions](./DOCKER_BUILD_INSTRUCTIONS.md).

Set a strong `SESSION_SECRET` in `.env` first (required in production mode).
For local HTTP Docker runs, set `SESSION_COOKIE_SECURE=false` (already set in `docker-compose.yml`).

```bash
docker compose up --build -d
```

If you run the image directly (without `docker compose`), pass a session secret explicitly:

```bash
docker run --rm -p 3000:3000 -e SESSION_SECRET=your-strong-secret litera-one-portal
```

Portal: `http://localhost:3000`  
Health: `http://localhost:3000/health`

If sign-in shows `fetch failed`, the container cannot reach Microsoft login endpoints.
Set `HTTPS_PROXY` / `HTTP_PROXY` in `.env` when your network requires a proxy.

Connectivity check from Docker:
```bash
docker compose run --rm ms-connectivity-check
```

---

## 🔐 Authentication Setup

This portal uses the **Client Credentials** OAuth 2.0 flow — no user interaction required after setup.

### 1. Register an Azure AD Application

1. Go to **Azure Portal → Azure Active Directory → App registrations → New registration**
2. Name it (e.g. `Litera One Manager Portal`)
3. Set **Supported account types** to "Accounts in this organizational directory only"
4. Click **Register**

### 2. Add API Permissions

In your app registration, go to **API Permissions → Add permission → Microsoft Graph → Application permissions** and add:

| Permission | Purpose |
|---|---|
| `Directory.Read.All` | Read tenant and service principal details |
| `Application.ReadWrite.All` | Create and manage enterprise applications |
| `AppRoleAssignment.ReadWrite.All` | Assign users/groups to apps |
| `Group.Read.All` | Search and read group members |
| `Synchronization.ReadWrite.All` | Manage SCIM provisioning jobs |
| `DelegatedPermissionGrant.ReadWrite.All` | Grant admin consent |

Then click **Grant admin consent for [your tenant]**.

### 3. Create a Client Secret

1. Go to **Certificates & secrets → New client secret**
2. Set an expiry and click **Add**
3. Copy the **Value** (not the ID) — you only see it once

### 4. Sign In to the Portal

In the portal sidebar, click **Sign In** and enter:
- **Tenant ID**: Your Azure AD Directory ID (from app registration Overview)
- **Client ID**: Your Application (client) ID
- **Client Secret**: The secret value you copied

Check "Remember credentials" to persist them in localStorage.

---

## 📂 Project Structure

```
litera-one-portal/
├── src/
│   ├── main.tsx                   # React entry point
│   ├── App.tsx                    # Root component, auth state, page routing
│   ├── index.css                  # Global styles (dark theme)
│   │
│   ├── services/
│   │   ├── graphClient.js         # Microsoft Graph API client (all API calls)
│   │   ├── graphClient.ts         # Typed wrapper around Graph client
│   │   └── credentialStore.ts     # localStorage credential persistence
│   │
│   ├── hooks/
│   │   ├── useToast.ts            # Toast notification hook
│   │   └── useLogger.ts           # Console log hook
│   │
│   ├── utils/
│   │   ├── enablementData.ts      # Sample enablement resources
│   │   └── exportHelpers.ts       # CSV, text, HTML export utilities
│   │
│   ├── components/
│   │   ├── Sidebar.tsx            # Navigation sidebar + user profile
│   │   ├── AuthModal.tsx          # Client Credentials sign-in modal
│   │   ├── AssignModal.tsx        # User/Group search and selection picker
│   │   ├── ScimModal.tsx          # 3-step SCIM app creation wizard
│   │   ├── GroupMembersModal.tsx  # Group member viewer
│   │   ├── UserAppsModal.tsx      # User's Litera app assignments viewer
│   │   ├── ConsolePanel.tsx       # Slide-up activity console
│   │   ├── Notifications.tsx      # Bell notification panel
│   │   └── Toast.tsx              # Toast notification container
│   │
│   └── pages/
│       ├── Dashboard.tsx          # Welcome + system status
│       ├── EntraApps.tsx          # Enterprise app management (main feature)
│       ├── Deployments.tsx        # Outlook/Word add-in deployment
│       ├── EnablementHub.tsx      # Training resources browser
│       └── LiteraUsersManagement.tsx # Multi-app assignment/removal workflow
│
├── index.html
├── vite.config.js
├── package.json
└── README.md
```

---

## ✨ Features

### Dashboard
- System status with Graph API connection indicator
- Quick navigation tiles to all sections
- Organization info and last directory sync time

### Entra Apps
- Search Microsoft Entra service principals by name
- View API permissions and consent status
- Grant admin consent for all required permissions
- Toggle app visibility in MyApps portal (HideApp tag)
- Manage user and group assignments with search and bulk removal
- Export assignments to CSV
- Check and restart SCIM provisioning jobs
- View group members and user's app assignments
- Delete enterprise applications

### SCIM App Creation Wizard
- 3-step wizard: App Details → Assignments → Create
- Creates app registration + service principal
- Configures SCIM synchronization job with your endpoint URL and secret token
- Assigns initial users/groups during creation
- Auto-hides from MyApps (recommended for SCIM apps)

### Add-in Deployment
- Deploy Outlook or Word add-ins
- Support for AppSource URL or Manifest file/URL
- Target selection: add users/groups or copy from existing Litera app
- Method: prepare via portal or deploy via M365 Admin Center (with CSV export)
- Load and display current deployment status from Graph API

### Enablement Hub
- Browse and filter training resources (Video, Guide, Doc, Template)
- Search by title, description, or tags
- Card view (1–5 per row) or list view
- Select resources for targeted export
- Export to HTML page or CSV

### Console & Notifications
- Slide-up activity console with timestamped log entries
- Export logs to text file
- Bell notification panel with unread indicator

---

## 🔧 Configuration

### Custom Enablement Resources

To use your own enablement resources instead of the built-in samples, create a JSON file matching this schema:

```json
{
  "resources": [
    {
      "id": 1,
      "title": "Getting Started",
      "type": "video",
      "description": "Introduction to Litera One.",
      "url": "https://your-link.com",
      "icon": "🎬",
      "tags": ["onboarding"]
    }
  ]
}
```

Valid `type` values: `video`, `guide`, `doc`, `template`

Store it via `saveEnablementData(json)` from `src/services/credentialStore.ts`, or load it in the browser console:
```js
localStorage.setItem('lo_enablement', JSON.stringify(yourJson))
```

---

## 🛡 Security Notes

- Credentials are only stored in `localStorage` if the user explicitly checks "Remember credentials"
- Client secrets are stored in plaintext in localStorage — suitable for internal/admin tools only
- For production use, consider a server-side token broker instead of storing the secret client-side
- All API calls go directly from the browser to Microsoft's endpoints (`login.microsoftonline.com` and `graph.microsoft.com`)
- No backend server required

---

## 📦 Dependencies

| Package | Purpose |
|---|---|
| `react` `react-dom` | UI framework |
| `react-router-dom` | Routing (available for future use) |
| `vite` `@vitejs/plugin-react` | Build tooling |

No additional UI libraries — all styles are hand-crafted in `index.css`.
