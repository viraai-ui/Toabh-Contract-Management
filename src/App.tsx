import { useEffect, useMemo, useState, type ReactNode } from 'react'
import { Ban, CheckCircle2, ExternalLink, FilePenLine, FileText, MoreVertical, PauseCircle, RefreshCcw, RefreshCw, RotateCw, Save, Search, Send, StickyNote } from 'lucide-react'
import { fetchDashboardData, postAction } from './api'
import type { ContractRecord, ExpiringFilter, RenewalEditableFields, RenewalRecord, TabKey } from './types'

const tabOptions: Array<{ key: TabKey; label: string; subtitle: string }> = [
  { key: 'all', label: 'All Signed Contracts', subtitle: 'Only rows with Signed PDF URL' },
  { key: 'expiring', label: 'Contracts Expiring', subtitle: 'Renewal watchlist by expiry' },
  { key: 'renewals', label: 'Contract Renewal', subtitle: 'Simple renewal operations' },
]

const expiringOptions: Array<{ value: ExpiringFilter; label: string }> = [
  { value: 'this_month', label: 'Expiring This Month' },
  { value: 'next_month', label: 'Expiring Next Month' },
  { value: 'next_3_months', label: 'Expiring in Next 3 Months' },
  { value: 'next_6_months', label: 'Expiring in Next 6 Months' },
  { value: 'already_expired', label: 'Already Expired' },
  { value: 'custom', label: 'Custom Date Range' },
]

function formatDate(value: string) {
  return value || '—'
}

function formatDays(value: number | null) {
  if (value === null || Number.isNaN(value)) return '—'
  return `${value} days`
}

function toneLabel(tone: ContractRecord['statusTone']) {
  switch (tone) {
    case 'active':
      return 'Active'
    case 'due-soon':
      return 'Due Soon'
    case 'urgent':
      return 'Urgent'
    case 'expired':
      return 'Expired'
    default:
      return 'Unknown'
  }
}

function toneClass(tone: ContractRecord['statusTone']) {
  return `status-pill status-pill-${tone}`
}

function safeOpen(url: string) {
  if (!url) return
  window.open(url, '_blank', 'noopener,noreferrer')
}

function parseDate(value: string) {
  if (!value) return null
  const date = new Date(value)
  return Number.isNaN(date.getTime()) ? null : new Date(date.getFullYear(), date.getMonth(), date.getDate())
}

function monthRange(offsetMonths: number) {
  const today = new Date()
  return {
    start: new Date(today.getFullYear(), today.getMonth() + offsetMonths, 1),
    end: new Date(today.getFullYear(), today.getMonth() + offsetMonths + 1, 0),
  }
}

function addMonthsRange(months: number) {
  const today = new Date()
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate())
  const end = new Date(today.getFullYear(), today.getMonth() + months, 0)
  return { start, end }
}

function isWithinRange(date: Date, start: Date, end: Date) {
  return date >= start && date <= end
}

function ActionButton({ onClick, children, disabled = false }: { onClick: () => void; children: ReactNode; disabled?: boolean }) {
  return (
    <button className="secondary-button small-button" onClick={onClick} type="button" disabled={disabled}>
      {children}
    </button>
  )
}

function DataCell({ children }: { children: ReactNode }) {
  return <td>{children || '—'}</td>
}

function ActionMenu({ children }: { children: ReactNode }) {
  const [open, setOpen] = useState(false)

  useEffect(() => {
    if (!open) return
    function close() {
      setOpen(false)
    }
    window.addEventListener('click', close)
    return () => window.removeEventListener('click', close)
  }, [open])

  return (
    <div className="menu-wrap" onClick={(event) => event.stopPropagation()}>
      <button className={open ? 'menu-button active' : 'menu-button'} onClick={() => setOpen((value) => !value)} type="button">
        <MoreVertical size={16} />
      </button>
      {open ? <div className="menu-panel">{children}</div> : null}
    </div>
  )
}

function MenuItem({ onClick, children, disabled = false, danger = false }: { onClick: () => void; children: ReactNode; disabled?: boolean; danger?: boolean }) {
  return (
    <button className={danger ? 'menu-item danger' : 'menu-item'} onClick={onClick} type="button" disabled={disabled}>
      {children}
    </button>
  )
}

function parseEditableJson(value: string): RenewalEditableFields {
  if (!value) return {}
  try {
    const parsed = JSON.parse(value)
    return parsed && typeof parsed === 'object' ? parsed : {}
  } catch {
    return {}
  }
}

function uniqueValues(values: string[]) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b))
}

export default function App() {
  const [tab, setTab] = useState<TabKey>('all')
  const [contracts, setContracts] = useState<ContractRecord[]>([])
  const [renewals, setRenewals] = useState<RenewalRecord[]>([])
  const [generatedAt, setGeneratedAt] = useState('')
  const [expiringFilter, setExpiringFilter] = useState<ExpiringFilter>('this_month')
  const [customFrom, setCustomFrom] = useState('')
  const [customTo, setCustomTo] = useState('')
  const [search, setSearch] = useState('')
  const [versionFilter, setVersionFilter] = useState('all')
  const [aiFilter, setAiFilter] = useState('all')
  const [renewalStatusFilter, setRenewalStatusFilter] = useState('all')
  const [expiryStatusFilter, setExpiryStatusFilter] = useState('all')
  const [loading, setLoading] = useState(true)
  const [busyRow, setBusyRow] = useState<string>('')
  const [error, setError] = useState('')
  const [editingRenewalRow, setEditingRenewalRow] = useState<number | null>(null)
  const [renewalDraft, setRenewalDraft] = useState<RenewalEditableFields>({})

  async function loadData() {
    setLoading(true)
    setError('')
    try {
      const data = await fetchDashboardData(expiringFilter, customFrom, customTo)
      setContracts(data.contracts || [])
      setRenewals(data.renewals || [])
      setGeneratedAt(data.generatedAt || '')
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Something went wrong.')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    void loadData()
  }, [expiringFilter, customFrom, customTo])

  const versionOptions = useMemo(() => uniqueValues(contracts.map((contract) => contract.version || '')), [contracts])
  const aiOptions = useMemo(() => uniqueValues(contracts.map((contract) => contract.aiScanStatus || 'Pending')), [contracts])
  const renewalOptions = useMemo(() => uniqueValues(contracts.map((contract) => contract.renewalStatus || '')), [contracts])

  const filteredContracts = useMemo(() => {
    const query = search.trim().toLowerCase()
    return contracts.filter((contract) => {
      const matchesSearch = !query || [contract.name, contract.email, contract.phone].some((value) => value.toLowerCase().includes(query))
      const matchesVersion = versionFilter === 'all' || (contract.version || '') === versionFilter
      const matchesAi = aiFilter === 'all' || (contract.aiScanStatus || 'Pending') === aiFilter
      const matchesRenewal = renewalStatusFilter === 'all' || (contract.renewalStatus || '') === renewalStatusFilter
      const matchesExpiry = expiryStatusFilter === 'all' || contract.statusTone === expiryStatusFilter
      return matchesSearch && matchesVersion && matchesAi && matchesRenewal && matchesExpiry
    })
  }, [contracts, search, versionFilter, aiFilter, renewalStatusFilter, expiryStatusFilter])

  const expiringContracts = useMemo(() => {
    const today = new Date()
    const normalizedToday = new Date(today.getFullYear(), today.getMonth(), today.getDate())

    return filteredContracts.filter((contract) => {
      const expiry = parseDate(contract.contractExpiryDate)
      if (!expiry) return false

      switch (expiringFilter) {
        case 'this_month': {
          const range = monthRange(0)
          return isWithinRange(expiry, range.start, range.end)
        }
        case 'next_month': {
          const range = monthRange(1)
          return isWithinRange(expiry, range.start, range.end)
        }
        case 'next_3_months': {
          const range = addMonthsRange(3)
          return isWithinRange(expiry, range.start, range.end)
        }
        case 'next_6_months': {
          const range = addMonthsRange(6)
          return isWithinRange(expiry, range.start, range.end)
        }
        case 'already_expired':
          return expiry < normalizedToday
        case 'custom': {
          const start = parseDate(customFrom)
          const end = parseDate(customTo)
          if (!start || !end) return true
          return isWithinRange(expiry, start, end)
        }
        default:
          return true
      }
    })
  }, [filteredContracts, expiringFilter, customFrom, customTo])

  async function runAction(action: string, rowId: number, payload: Record<string, unknown> = {}) {
    setBusyRow(`${action}-${rowId}`)
    setError('')
    try {
      await postAction(action, { rowId, ...payload })
      await loadData()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Action failed.')
    } finally {
      setBusyRow('')
    }
  }

  async function addNote(rowId: number, scope: 'main' | 'renewal' = 'main') {
    const note = window.prompt(scope === 'renewal' ? 'Add or append renewal note' : 'Add or append contract note')?.trim()
    if (!note) return
    await runAction('addNote', rowId, { note, scope })
  }

  async function setMainRenewalStatus(rowId: number, status: 'On Hold' | 'Not Renewing') {
    const note = window.prompt(`Optional note for ${status}`)?.trim() || ''
    await runAction('setContractRenewalStatus', rowId, { status, note })
  }

  async function setRenewalStatus(rowId: number, action: 'setRenewalOnHold' | 'setRenewalNotRenewing' | 'cancelRenewal') {
    const label = action === 'setRenewalOnHold' ? 'On Hold' : action === 'setRenewalNotRenewing' ? 'Not Renewing' : 'Cancel Renewal'
    const note = window.prompt(`Optional note for ${label}`)?.trim() || ''
    await runAction(action, rowId, { note })
  }

  function startEditingRenewal(renewal: RenewalRecord) {
    setEditingRenewalRow(renewal.rowId)
    setRenewalDraft({
      ...parseEditableJson(renewal.editableDataJson),
      ...renewal.editableFields,
      name: renewal.editableFields.name ?? renewal.name,
      email: renewal.editableFields.email ?? renewal.email,
      phone: renewal.editableFields.phone ?? renewal.phone,
      version: renewal.editableFields.version ?? renewal.newVersion,
      contractStartDate: renewal.editableFields.contractStartDate ?? '',
      contractValidity: renewal.editableFields.contractValidity ?? '',
      contractExpiryDate: renewal.editableFields.contractExpiryDate ?? renewal.newContractExpiryDate ?? '',
    })
  }

  function cancelEditingRenewal() {
    setEditingRenewalRow(null)
    setRenewalDraft({})
  }

  async function saveRenewal(rowId: number) {
    setBusyRow(`updateRenewal-${rowId}`)
    setError('')
    try {
      await postAction('updateRenewal', { rowId, updates: renewalDraft })
      setEditingRenewalRow(null)
      setRenewalDraft({})
      await loadData()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save renewal.')
    } finally {
      setBusyRow('')
    }
  }

  return (
    <div className="page-shell">
      <header className="hero-card">
        <div>
          <p className="eyebrow">TOABH Contracts</p>
          <h1>Simple Contract Management Dashboard</h1>
          <p className="subtle">Only signed contracts appear here. Unsigned, draft, pending, and failed contracts stay out until Signed PDF URL exists.</p>
        </div>
        <div className="hero-actions">
          {generatedAt ? <p className="generated-at">Last synced: {generatedAt}</p> : null}
          <button className="primary-button" onClick={() => void loadData()} type="button">
            <RefreshCcw size={16} /> Refresh
          </button>
        </div>
      </header>

      <section className="tabs-card">
        <div className="tabs-row premium-tabs">
          {tabOptions.map((option) => (
            <button
              key={option.key}
              className={option.key === tab ? 'tab-button active' : 'tab-button'}
              onClick={() => setTab(option.key)}
              type="button"
            >
              <span>{option.label}</span>
              <small>{option.subtitle}</small>
            </button>
          ))}
        </div>
      </section>

      {tab !== 'renewals' && (
        <section className="table-card">
          <div className="section-header">
            <div>
              <p className="eyebrow">Filters</p>
              <h2>{tab === 'all' ? 'All Signed Contracts' : 'Contracts Expiring'}</h2>
              <p>Keep it fast: search, filter, act.</p>
            </div>
          </div>
          <div className="secondary-toolbar">
            <div className="toolbar-group">
              <label className="search-field">
                <Search size={16} />
                <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search name, email, phone" />
              </label>
              <label className="toolbar-select">
                <select value={versionFilter} onChange={(event) => setVersionFilter(event.target.value)}>
                  <option value="all">All Versions</option>
                  {versionOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
              </label>
              <label className="toolbar-select">
                <select value={aiFilter} onChange={(event) => setAiFilter(event.target.value)}>
                  <option value="all">All AI Status</option>
                  {aiOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
              </label>
              <label className="toolbar-select">
                <select value={renewalStatusFilter} onChange={(event) => setRenewalStatusFilter(event.target.value)}>
                  <option value="all">All Renewal Status</option>
                  {renewalOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
              </label>
              <label className="toolbar-select">
                <select value={expiryStatusFilter} onChange={(event) => setExpiryStatusFilter(event.target.value)}>
                  <option value="all">All Expiry Status</option>
                  <option value="active">Active</option>
                  <option value="due-soon">Due Soon</option>
                  <option value="urgent">Urgent</option>
                  <option value="expired">Expired</option>
                </select>
              </label>
              {tab === 'expiring' && (
                <>
                  <label className="toolbar-select">
                    <select value={expiringFilter} onChange={(event) => setExpiringFilter(event.target.value as ExpiringFilter)}>
                      {expiringOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                    </select>
                  </label>
                  {expiringFilter === 'custom' && (
                    <>
                      <input type="date" value={customFrom} onChange={(event) => setCustomFrom(event.target.value)} />
                      <input type="date" value={customTo} onChange={(event) => setCustomTo(event.target.value)} />
                    </>
                  )}
                </>
              )}
            </div>
          </div>
        </section>
      )}

      {error && <div className="error-banner">{error}</div>}
      {loading ? (
        <div className="empty-card">Loading dashboard…</div>
      ) : (
        <>
          {tab === 'all' && (
            <div className="table-card">
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Talent Name</th>
                      <th>Email</th>
                      <th>Phone</th>
                      <th>Original Contract Link</th>
                      <th>Signed Contract Link</th>
                      <th>Contract Signed On</th>
                      <th>Contract Start Date</th>
                      <th>Contract Validity</th>
                      <th>Contract Expiry Date</th>
                      <th>Days Left</th>
                      <th>Version</th>
                      <th>AI Scan Status</th>
                      <th>Renewal Status</th>
                      <th>Notes</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredContracts.length === 0 ? (
                      <tr><td colSpan={15} className="empty-state">No signed contracts found.</td></tr>
                    ) : filteredContracts.map((contract) => (
                      <tr key={contract.rowId}>
                        <td>
                          <div className="primary-cell">
                            <strong>{contract.name}</strong>
                            <span>{toneLabel(contract.statusTone)}</span>
                          </div>
                        </td>
                        <DataCell>{contract.email}</DataCell>
                        <DataCell>{contract.phone}</DataCell>
                        <DataCell><button className="link-chip" onClick={() => safeOpen(contract.contractLink)} type="button">Open <ExternalLink size={14} /></button></DataCell>
                        <DataCell><button className="link-chip" onClick={() => safeOpen(contract.signedPdfUrl)} type="button">Open <ExternalLink size={14} /></button></DataCell>
                        <DataCell>{formatDate(contract.contractSignedOn)}</DataCell>
                        <DataCell>{formatDate(contract.contractStartDate)}</DataCell>
                        <DataCell>{contract.contractValidity || '—'}</DataCell>
                        <DataCell>{formatDate(contract.contractExpiryDate)}</DataCell>
                        <DataCell>
                          <div className="status-stack">
                            <span>{formatDays(contract.daysLeft)}</span>
                            <span className={toneClass(contract.statusTone)}>{toneLabel(contract.statusTone)}</span>
                          </div>
                        </DataCell>
                        <DataCell>{contract.version || '—'}</DataCell>
                        <td>
                          <div className="stacked-text">
                            <strong>{contract.aiScanStatus || 'Pending'}</strong>
                            <span>{contract.aiScanNotes || '—'}</span>
                          </div>
                        </td>
                        <DataCell>{contract.renewalStatus || '—'}</DataCell>
                        <td>
                          <div className="stacked-text">
                            <strong>{contract.notes || '—'}</strong>
                            <span>{contract.lastSyncedAt || '—'}</span>
                          </div>
                        </td>
                        <td>
                          <ActionMenu>
                            <MenuItem onClick={() => safeOpen(contract.contractLink)}><ExternalLink size={14} /> Open Original Contract</MenuItem>
                            <MenuItem onClick={() => safeOpen(contract.signedPdfUrl)}><ExternalLink size={14} /> Open Signed Contract</MenuItem>
                            <MenuItem onClick={() => void runAction('rescanWithAI', contract.rowId)} disabled={busyRow === `rescanWithAI-${contract.rowId}`}><RotateCw size={14} /> Rescan with AI</MenuItem>
                            <MenuItem onClick={() => void runAction('startRenewal', contract.rowId)} disabled={busyRow === `startRenewal-${contract.rowId}`}><FilePenLine size={14} /> Start Renewal</MenuItem>
                            <MenuItem onClick={() => void addNote(contract.rowId)}><StickyNote size={14} /> Add/Edit Note</MenuItem>
                          </ActionMenu>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {tab === 'expiring' && (
            <div className="table-card">
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Talent Name</th>
                      <th>Phone</th>
                      <th>Email</th>
                      <th>Contract Signed On</th>
                      <th>Contract Start Date</th>
                      <th>Contract Expiry Date</th>
                      <th>Days Left</th>
                      <th>Signed Contract Link</th>
                      <th>Renewal Status</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {expiringContracts.length === 0 ? (
                      <tr><td colSpan={10} className="empty-state">No matching contracts found.</td></tr>
                    ) : expiringContracts.map((contract) => (
                      <tr key={contract.rowId}>
                        <td>
                          <div className="primary-cell">
                            <strong>{contract.name}</strong>
                            <span className={toneClass(contract.statusTone)}>{toneLabel(contract.statusTone)}</span>
                          </div>
                        </td>
                        <DataCell>{contract.phone}</DataCell>
                        <DataCell>{contract.email}</DataCell>
                        <DataCell>{formatDate(contract.contractSignedOn)}</DataCell>
                        <DataCell>{formatDate(contract.contractStartDate)}</DataCell>
                        <DataCell>{formatDate(contract.contractExpiryDate)}</DataCell>
                        <DataCell>{formatDays(contract.daysLeft)}</DataCell>
                        <DataCell><button className="link-chip" onClick={() => safeOpen(contract.signedPdfUrl)} type="button">Open <ExternalLink size={14} /></button></DataCell>
                        <DataCell>{contract.renewalStatus || '—'}</DataCell>
                        <td>
                          <ActionMenu>
                            <MenuItem onClick={() => void runAction('startRenewal', contract.rowId)} disabled={busyRow === `startRenewal-${contract.rowId}`}><FilePenLine size={14} /> Start Renewal</MenuItem>
                            <MenuItem onClick={() => safeOpen(contract.signedPdfUrl)}><ExternalLink size={14} /> Open Signed Contract</MenuItem>
                            <MenuItem onClick={() => void addNote(contract.rowId)}><StickyNote size={14} /> Add Note</MenuItem>
                            <MenuItem onClick={() => void setMainRenewalStatus(contract.rowId, 'Not Renewing')} danger><Ban size={14} /> Mark Not Renewing</MenuItem>
                            <MenuItem onClick={() => void setMainRenewalStatus(contract.rowId, 'On Hold')}><PauseCircle size={14} /> Put On Hold</MenuItem>
                          </ActionMenu>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {tab === 'renewals' && (
            <div className="table-card">
              <div className="section-header">
                <div>
                  <p className="eyebrow">Renewals</p>
                  <h2>Contract Renewal</h2>
                  <p>Draft, generate, send, sync, finish.</p>
                </div>
              </div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Talent Name</th>
                      <th>Phone</th>
                      <th>Email</th>
                      <th>Old Version</th>
                      <th>New Version</th>
                      <th>Old Expiry Date</th>
                      <th>Renewal Status</th>
                      <th>New Contract Link</th>
                      <th>New Zoho Status</th>
                      <th>New Signed PDF URL</th>
                      <th>Notes</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {renewals.length === 0 ? (
                      <tr><td colSpan={12} className="empty-state">No renewals yet.</td></tr>
                    ) : renewals.map((renewal) => {
                      const isEditing = editingRenewalRow === renewal.rowId
                      return (
                        <tr key={renewal.rowId}>
                          <td>
                            {isEditing ? (
                              <div className="stacked-text">
                                <input value={String(renewalDraft.name ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, name: event.target.value }))} placeholder="Name" />
                              </div>
                            ) : (
                              <div className="primary-cell">
                                <strong>{renewal.name}</strong>
                                <span>{renewal.renewalId}</span>
                              </div>
                            )}
                          </td>
                          <td>{isEditing ? <input value={String(renewalDraft.phone ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, phone: event.target.value }))} placeholder="Phone" /> : renewal.phone || '—'}</td>
                          <td>{isEditing ? <input value={String(renewalDraft.email ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, email: event.target.value }))} placeholder="Email" /> : renewal.email || '—'}</td>
                          <td>{renewal.oldVersion || '—'}</td>
                          <td>
                            {isEditing ? (
                              <div className="stacked-text">
                                <input value={String(renewalDraft.version ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, version: event.target.value }))} placeholder="New Version" />
                                <input type="date" value={String(renewalDraft.contractStartDate ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, contractStartDate: event.target.value }))} />
                                <input value={String(renewalDraft.contractValidity ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, contractValidity: event.target.value }))} placeholder="Validity e.g. 3 years" />
                                <input type="date" value={String(renewalDraft.contractExpiryDate ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, contractExpiryDate: event.target.value }))} />
                              </div>
                            ) : renewal.newVersion || '—'}
                          </td>
                          <DataCell>{formatDate(renewal.oldExpiryDate)}</DataCell>
                          <td>
                            <div className="status-stack">
                              <strong>{renewal.renewalStatus || '—'}</strong>
                              <span>{renewal.error || renewal.newContractSignedOn || '—'}</span>
                            </div>
                          </td>
                          <DataCell>{renewal.newContractLink ? <button className="link-chip" onClick={() => safeOpen(renewal.newContractLink)} type="button">Open <ExternalLink size={14} /></button> : '—'}</DataCell>
                          <DataCell>{renewal.newZohoStatus || '—'}</DataCell>
                          <DataCell>{renewal.newSignedPdfUrl ? <button className="link-chip" onClick={() => safeOpen(renewal.newSignedPdfUrl)} type="button">Open <ExternalLink size={14} /></button> : '—'}</DataCell>
                          <DataCell>{renewal.notes || '—'}</DataCell>
                          <td>
                            {isEditing ? (
                              <div className="actions-stack">
                                <ActionButton onClick={() => void saveRenewal(renewal.rowId)} disabled={busyRow === `updateRenewal-${renewal.rowId}`}><Save size={14} /> Save</ActionButton>
                                <ActionButton onClick={cancelEditingRenewal}><Ban size={14} /> Cancel</ActionButton>
                              </div>
                            ) : (
                              <ActionMenu>
                                {renewal.oldSignedPdfUrl ? <MenuItem onClick={() => safeOpen(renewal.oldSignedPdfUrl)}><ExternalLink size={14} /> View Old Signed Contract</MenuItem> : null}
                                <MenuItem onClick={() => startEditingRenewal(renewal)} disabled={!renewal.canEdit}><FilePenLine size={14} /> Edit Renewal Details</MenuItem>
                                <MenuItem onClick={() => void runAction('regenerateRenewalContract', renewal.rowId)} disabled={!renewal.canEdit || busyRow === `regenerateRenewalContract-${renewal.rowId}`}><FileText size={14} /> Regenerate Contract</MenuItem>
                                <MenuItem onClick={() => void runAction('sendRenewalForSigning', renewal.rowId)} disabled={!renewal.canEdit || busyRow === `sendRenewalForSigning-${renewal.rowId}`}><Send size={14} /> Send for Signing</MenuItem>
                                <MenuItem onClick={() => void runAction('refreshRenewalZohoStatus', renewal.rowId)} disabled={busyRow === `refreshRenewalZohoStatus-${renewal.rowId}`}><RefreshCw size={14} /> Refresh Zoho Status</MenuItem>
                                <MenuItem onClick={() => void runAction('markRenewalRenewed', renewal.rowId)}><CheckCircle2 size={14} /> Mark Signed/Renewed</MenuItem>
                                <MenuItem onClick={() => void setRenewalStatus(renewal.rowId, 'setRenewalNotRenewing')} danger><Ban size={14} /> Mark Not Renewing</MenuItem>
                                <MenuItem onClick={() => void setRenewalStatus(renewal.rowId, 'setRenewalOnHold')}><PauseCircle size={14} /> Put On Hold</MenuItem>
                                <MenuItem onClick={() => void setRenewalStatus(renewal.rowId, 'cancelRenewal')} danger><Ban size={14} /> Cancel Renewal</MenuItem>
                                <MenuItem onClick={() => void addNote(renewal.rowId, 'renewal')}><StickyNote size={14} /> Add/Edit Notes</MenuItem>
                              </ActionMenu>
                            )}
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  )
}
