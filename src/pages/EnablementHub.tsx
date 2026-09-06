import { useState, useMemo } from 'react'
import { SAMPLE_ENABLEMENT } from '../utils/enablementData'
import { downloadHtml, downloadCsv, enablementToHtml } from '../utils/exportHelpers'
import { loadEnablementData } from '../services/credentialStore'
import type { EnablementResource } from '../types/enablement'
import type { LogFn, ToastFn } from '../types/ui'

const FILTERS = ['All', 'video', 'guide', 'doc', 'template']
const TYPE_LABELS = { All: 'All', video: '📹 Videos', guide: '📘 Guides', doc: '📄 Docs', template: '🗂 Templates' }
const TYPE_CLASS = { video: 'type-video', guide: 'type-guide', doc: 'type-doc', template: 'type-template' }

interface EnablementHubProps {
  log: LogFn
  toast: ToastFn
}

export default function EnablementHub ({ log, toast }: EnablementHubProps) {
  const [filter, setFilter] = useState('All')
  const [search, setSearch] = useState('')
  const [perRow, setPerRow] = useState(3)
  const [viewMode, setViewMode] = useState('cards') // 'cards' | 'list'
  const [selected, setSelected] = useState<Set<string | number>>(new Set())
  const [selectAll, setSelectAll] = useState(false)

  // Use custom JSON if stored, else samples
  const allResources = useMemo<EnablementResource[]>(() => {
    const custom = loadEnablementData()
    return custom?.resources || SAMPLE_ENABLEMENT
  }, [])

  const filtered = useMemo(() => {
    return allResources.filter(r => {
      if (filter !== 'All' && r.type !== filter) return false
      if (search) {
        const q = search.toLowerCase()
        return (r.title || '').toLowerCase().includes(q) ||
          (r.description || '').toLowerCase().includes(q) ||
          (r.tags || []).some(t => t.toLowerCase().includes(q))
      }
      return true
    })
  }, [allResources, filter, search])

  const toggleSelect = (id: string | number) => {
    setSelected(prev => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  }

  const handleSelectAll = (checked: boolean) => {
    setSelectAll(checked)
    setSelected(checked ? new Set(filtered.map(r => r.id)) : new Set())
  }

  const getExportItems = () => {
    if (selected.size > 0) return filtered.filter(r => selected.has(r.id))
    return filtered
  }

  const exportHtml = () => {
    const items = getExportItems()
    downloadHtml(enablementToHtml(items), 'litera_enablement.html')
    toast(`Exported ${items.length} resource(s)`, 'success')
    log(`Exported enablement resources to HTML (${items.length} items)`)
  }

  const exportCsv = () => {
    const items = getExportItems()
    downloadCsv(
      [['Title', 'Type', 'Description', 'URL'], ...items.map(r => [r.title, r.type, r.description, r.url])],
      'litera_enablement.csv'
    )
    toast(`Exported ${items.length} resource(s) to CSV`, 'success')
  }

  const copyLink = (url: string) => {
    navigator.clipboard.writeText(url)
      .then(() => toast('Link copied!'))
      .catch(() => toast('Copy failed', 'error'))
  }

  return (
    <div className="page-container page-enter">
      <div className="page-header">
        <div>
          <div className="page-title">Enablement Hub</div>
          <div className="page-subtitle">Training resources, guides, and templates for end-users</div>
        </div>
        <div className="flex gap-2">
          <button className="btn btn-secondary" onClick={exportCsv}>📋 Export CSV</button>
          <button className="btn btn-secondary" onClick={exportHtml}>📤 Export HTML</button>
        </div>
      </div>

      {/* Toolbar */}
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <div className="flex items-center gap-2 flex-wrap">
          {FILTERS.map(f => (
            <button
              key={f}
              className={`btn btn-sm ${filter === f ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setFilter(f)}
            >
              {TYPE_LABELS[f]}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2">
          <label className="chk-label" style={{ fontSize: 12.5 }}>
            <input type="checkbox" checked={selectAll} onChange={e => handleSelectAll(e.target.checked)} />
            Select All
          </label>
          <div className="inp-group" style={{ width: 200 }}>
            <span className="inp-icon">🔍</span>
            <input
              className="inp inp-sm"
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="Search resources…"
            />
          </div>
          <select
            className="inp inp-sm"
            value={perRow}
            onChange={e => setPerRow(Number(e.target.value))}
            style={{ width: 80 }}
          >
            <option value={2}>2/row</option>
            <option value={3}>3/row</option>
            <option value={4}>4/row</option>
            <option value={5}>5/row</option>
          </select>
          <button
            className={`btn btn-sm ${viewMode === 'list' ? 'btn-primary' : 'btn-secondary'}`}
            onClick={() => setViewMode(m => m === 'cards' ? 'list' : 'cards')}
            title="Toggle list/card view"
          >
            {viewMode === 'cards' ? '☰ List' : '⊞ Cards'}
          </button>
        </div>
      </div>

      {/* Results info */}
      <div className="flex items-center gap-2 mb-3">
        <span className="text-xs">{filtered.length} resource{filtered.length !== 1 ? 's' : ''}</span>
        {selected.size > 0 && (
          <span className="badge badge-blue">{selected.size} selected for export</span>
        )}
      </div>

      {filtered.length === 0 && (
        <div className="empty-state">
          <div className="empty-icon">🔍</div>
          <div>No resources match your filter</div>
        </div>
      )}

      {/* Cards view */}
      {viewMode === 'cards' && filtered.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${perRow}, 1fr)`, gap: 14 }}>
          {filtered.map(r => (
            <div key={r.id} className="en-card">
              <div className="flex items-center justify-between mb-2">
                <label className="chk-label" style={{ fontSize: 12 }}>
                  <input
                    type="checkbox"
                    style={{ width: 13, height: 13 }}
                    checked={selected.has(r.id)}
                    onChange={() => toggleSelect(r.id)}
                  />
                </label>
                <span className={`en-type ${TYPE_CLASS[r.type] || 'type-guide'}`}>{r.type}</span>
              </div>
              <div className="en-title">{r.icon} {r.title}</div>
              <div className="en-desc">{r.description}</div>
              {r.tags?.length > 0 && (
                <div className="flex flex-wrap gap-1 mb-3">
                  {r.tags.map(t => (
                    <span key={t} style={{ fontSize: 10, background: 'var(--surface3)', border: '1px solid var(--border)', borderRadius: 20, padding: '1px 7px', color: 'var(--text3)' }}>
                      {t}
                    </span>
                  ))}
                </div>
              )}
              <div className="flex gap-2">
                <a href={r.url} target="_blank" rel="noreferrer" className="btn btn-secondary btn-xs">↗ Open</a>
                <button className="btn btn-secondary btn-xs" onClick={() => copyLink(r.url)}>🔗 Copy</button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* List view */}
      {viewMode === 'list' && filtered.length > 0 && (
        <div className="tbl-wrap">
          <table className="tbl">
            <thead>
              <tr>
                <th className="chk-col">
                  <input type="checkbox" checked={selectAll} onChange={e => handleSelectAll(e.target.checked)} />
                </th>
                <th>Resource</th>
                <th style={{ width: 80 }}>Type</th>
                <th style={{ width: 120 }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(r => (
                <tr key={r.id}>
                  <td className="chk-col">
                    <input type="checkbox" checked={selected.has(r.id)} onChange={() => toggleSelect(r.id)} />
                  </td>
                  <td>
                    <div style={{ color: 'var(--text)', fontWeight: 600, marginBottom: 2 }}>{r.icon} {r.title}</div>
                    <div className="text-xs">{r.description}</div>
                  </td>
                  <td>
                    <span className={`en-type ${TYPE_CLASS[r.type] || 'type-guide'}`}>{r.type}</span>
                  </td>
                  <td>
                    <div className="flex gap-1">
                      <a href={r.url} target="_blank" rel="noreferrer" className="btn btn-secondary btn-xs">↗</a>
                      <button className="btn btn-secondary btn-xs" onClick={() => copyLink(r.url)}>🔗</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
