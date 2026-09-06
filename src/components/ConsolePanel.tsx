import { useEffect, useRef } from 'react'

interface ConsolePanelProps {
  logs: string[]
  open: boolean
  onClear: () => void
  onExport: () => void
}

export default function ConsolePanel ({ logs, open, onClear, onExport }: ConsolePanelProps) {
  const logRef = useRef<HTMLDivElement | null>(null)
  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight
  }, [logs])

  return (
    <div className={`console-panel ${open ? 'open' : ''}`}>
      <div className="console-header">
        <span>▶ Activity Console ({logs.length} entries)</span>
        <div className="flex gap-2">
          <button className="btn btn-danger btn-xs" onClick={onClear}>🗑 Clear</button>
          <button className="btn btn-secondary btn-xs" onClick={onExport}>💾 Export</button>
        </div>
      </div>
      <div className="console-log" ref={logRef}>
        {logs.length === 0
          ? '— No activity yet —'
          : logs.join('\n')}
      </div>
    </div>
  )
}
