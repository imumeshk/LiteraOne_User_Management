import { useState, useRef, useCallback, useEffect } from 'react'
import graphClient from '../services/graphClient'
import type { GraphPrincipal } from '../types/graph'

export interface AssignSelection {
  id: string
  name: string
  email: string
  emailOrId: string
  type: 'User' | 'Group'
}

interface AssignModalProps {
  title?: string
  onConfirm: (selected: AssignSelection[]) => void
  onClose: () => void
}

export default function AssignModal ({ title, onConfirm, onClose }: AssignModalProps) {
  const [tab, setTab] = useState<'users' | 'groups'>('users')
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<GraphPrincipal[]>([])
  const [loading, setLoading] = useState(false)
  const [selected, setSelected] = useState<AssignSelection[]>([])
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const selectAllRef = useRef<HTMLInputElement | null>(null)

  const search = useCallback(async (q: string, t: 'users' | 'groups') => {
    if (!q.trim()) { setResults([]); return }
    setLoading(true)
    try {
      const res = await graphClient.searchPrincipals(t, q)
      setResults(res?.value || [])
    } catch (e) {
      setResults([])
    } finally {
      setLoading(false)
    }
  }, [])

  const handleInput = (q: string) => {
    setQuery(q)
    clearTimeout(timer.current)
    timer.current = setTimeout(() => search(q, tab), 300)
  }

  const handleTabChange = (t: 'users' | 'groups') => {
    setTab(t)
    if (query) search(query, t)
  }

  const toggle = (item: GraphPrincipal) => {
    const type = tab === 'groups' ? 'Group' : 'User'
    const sub = item.userPrincipalName || item.mail || item.id
    setSelected(prev => {
      const exists = prev.find(s => s.id === item.id)
      if (exists) return prev.filter(s => s.id !== item.id)
      return [...prev, { id: item.id, name: item.displayName, email: sub, type, emailOrId: sub }]
    })
  }

  const isSelected = (id: string) => selected.some(s => s.id === id)
  const selectedResultCount = results.filter(r => isSelected(r.id)).length
  const allResultsSelected = results.length > 0 && selectedResultCount === results.length
  const someResultsSelected = selectedResultCount > 0 && !allResultsSelected

  useEffect(() => {
    if (!selectAllRef.current) return
    selectAllRef.current.indeterminate = someResultsSelected
  }, [someResultsSelected])

  const toggleSelectAllResults = (checked: boolean) => {
    if (!results.length) return
    setSelected(prev => {
      if (!checked) {
        const ids = new Set(results.map(r => r.id))
        return prev.filter(s => !ids.has(s.id))
      }
      const next = [...prev]
      const existing = new Set(prev.map(s => s.id))
      for (const item of results) {
        if (existing.has(item.id)) continue
        const type = tab === 'groups' ? 'Group' : 'User'
        const sub = item.userPrincipalName || item.mail || item.id
        next.push({ id: item.id, name: item.displayName, email: sub, type, emailOrId: sub })
      }
      return next
    })
  }

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="modal-title">{title || '＋ Assign Users / Groups'}</div>

        <div className="tab-bar">
          {['users', 'groups'].map(t => (
            <button
              key={t}
              className={`tab-btn ${tab === t ? 'active' : ''}`}
              onClick={() => handleTabChange(t)}
            >
              {t === 'users' ? '👤 Users' : '👥 Groups'}
            </button>
          ))}
        </div>

        <div className="inp-group mb-2">
          <span className="inp-icon">🔍</span>
          <input
            className="inp"
            value={query}
            onChange={e => handleInput(e.target.value)}
            placeholder={`Search ${tab}...`}
            autoFocus
          />
        </div>

        <label className="chk-label mb-2">
          <input
            ref={selectAllRef}
            type="checkbox"
            checked={allResultsSelected}
            onChange={e => toggleSelectAllResults(e.target.checked)}
          />
          Select all results shown
        </label>

        <div className="search-result-list">
          {loading && (
            <div style={{ textAlign: 'center', padding: 16 }}>
              <span className="spinner spinner-dark" />
            </div>
          )}
          {!loading && !query && (
            <div style={{ textAlign: 'center', padding: '20px 16px', color: 'var(--text3)', fontSize: 13 }}>
              Type to search {tab}…
            </div>
          )}
          {!loading && query && results.length === 0 && (
            <div style={{ textAlign: 'center', padding: '20px 16px', color: 'var(--text3)', fontSize: 13 }}>
              No results for "{query}"
            </div>
          )}
          {!loading && results.map(item => {
            const sub = item.userPrincipalName || item.mail || item.id
            const sel = isSelected(item.id)
            return (
              <div
                key={item.id}
                className="search-result-item"
                onClick={() => toggle(item)}
              >
                <div>
                  <div className="sname">{item.displayName}</div>
                  <div className="semail">{sub}</div>
                </div>
                <span style={{ color: sel ? 'var(--accent3)' : 'var(--text3)', fontSize: 13, fontWeight: 600 }}>
                  {sel ? '✓ Added' : '+ Add'}
                </span>
              </div>
            )
          })}
        </div>

        {selected.length > 0 && (
          <div className="mt-2 mb-2">
            <span className="section-label">Selected ({selected.length})</span>
            <div className="flex flex-wrap gap-1 mt-1">
              {selected.map(s => (
                <span key={s.id} className="chip">
                  {s.name}
                  <button className="chip-remove" onClick={() => toggle({ id: s.id, displayName: s.name, mail: s.email })}>×</button>
                </span>
              ))}
            </div>
          </div>
        )}

        <div className="modal-footer">
          <button className="btn btn-secondary" onClick={onClose}>Cancel</button>
          <button
            className="btn btn-primary"
            onClick={() => onConfirm(selected)}
            disabled={selected.length === 0}
          >
            {selected.length > 0 ? `Confirm ${selected.length} selection(s)` : 'Select principals'}
          </button>
        </div>
      </div>
    </div>
  )
}
