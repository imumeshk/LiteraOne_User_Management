interface ToastItem {
  key: number
  msg: string
  type: string
}

interface ToastContainerProps {
  toasts: ToastItem[]
}

export default function ToastContainer ({ toasts }: ToastContainerProps) {
  return (
    <div className="toast-container">
      {toasts.map(t => (
        <div key={t.key} className={`toast ${t.type}`}>{t.msg}</div>
      ))}
    </div>
  )
}
