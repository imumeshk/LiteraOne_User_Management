import { useState, useEffect, useRef } from 'react'

interface NotificationsProps {
  logs: string[]
}

export default function Notifications ({ logs }: NotificationsProps) {
  const [open, setOpen] = useState(false)
  const [seen, setSeen] = useState(0)
  const panelRef = useRef<HTMLDivElement | null>(null)

  const unseen = logs.length - seen

  const toggle = () => {
    setOpen(o => !o)
    setSeen(logs.length)
  }

  // Close on outside click
  useEffect(() => {
    if (!open) return
    const handler = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null
      if (!target) return
      if (!panelRef.current?.contains(target) && !target.closest('.notif-trigger')) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [open])

  return (
    <>
      <button className="notif-trigger" onClick={toggle} title="Activity log">
        🔔
        {unseen > 0 && <span className="notif-badge" />}
      </button>
      {open && (
        <div className="notif-panel" ref={panelRef}>
          <div className="notif-head">
            <span>Recent Activity</span>
            <button
              onClick={() => setSeen(logs.length)}
              style={{ background: 'none', border: 'none', fontSize: 11, color: 'var(--text3)', cursor: 'pointer' }}
            >
              Mark all read
            </button>
          </div>
          {logs.length === 0 && (
            <div style={{ textAlign: 'center', padding: '24px 16px', color: 'var(--text3)', fontSize: 12 }}>
              No activity yet
            </div>
          )}
          {[...logs].reverse().map((line, i) => {
            const [ts, ...rest] = line.replace('[', '').split('] ')
            return (
              <div key={i} className="notif-item">
                <div className="notif-time">{ts}</div>
                <div className="notif-msg">{rest.join('] ')}</div>
              </div>
            )
          })}
        </div>
      )}
    </>
  )
}
