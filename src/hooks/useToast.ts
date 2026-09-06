import { useState, useCallback } from 'react'

let id = 0

interface ToastItem {
  key: number
  msg: string
  type: string
}

export function useToast () {
  const [toasts, setToasts] = useState<ToastItem[]>([])

  const addToast = useCallback((msg: string, type = '') => {
    const key = ++id
    setToasts(prev => [...prev, { key, msg, type }])
    setTimeout(() => setToasts(prev => prev.filter(t => t.key !== key)), 3500)
  }, [])

  return { toasts, toast: addToast }
}
