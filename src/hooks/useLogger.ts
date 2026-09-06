import { useState, useCallback } from 'react'

export function useLogger () {
  const [logs, setLogs] = useState<string[]>([])
  const [entries, setEntries] = useState<string[]>([]) // raw strings for export

  const log = useCallback((msg: string) => {
    const ts = new Date().toLocaleTimeString()
    const line = `[${ts}] ${msg}`
    setEntries(prev => [...prev, line])
    setLogs(prev => [...prev, line])
  }, [])

  const clearLogs = useCallback(() => {
    setLogs([])
    setEntries([])
  }, [])

  const exportLogs = useCallback(() => {
    const blob = new Blob([entries.join('\n')], { type: 'text/plain' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = 'litera_one_logs.txt'
    a.click()
  }, [entries])

  return { logs, log, clearLogs, exportLogs }
}
