import { useEffect, useMemo, useState, type ReactNode } from 'react'
import { ExternalLink, RefreshCcw, RotateCw, FilePenLine, StickyNote, Send, CheckCircle2, PauseCircle, Ban, RefreshCw, Save, FileText } from 'lucide-react'
import { fetchDashboardData, postAction } from './api'
import type { ContractRecord, ExpiringFilter, RenewalEditableFields, RenewalRecord, TabKey } from './types'

const tabOptions: Array<{ key: TabKey; label: string }> = [
  { key: 'all', label: 'All Signed Contracts' },
  { key: 'expiring', label: 'Contracts Expiring' },
  { key: 'renewals', label: 'Contract Renewal' },
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
  if (!value) return '—'
  return value
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
  return `badge badge-${tone}`
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
    <button className="action-button" onClick={onClick} type="button" disabled={disabled}>
      {children}
    </button>
  )
}

function DataCell({ children }: { children: ReactNode }) {
  return <td>{children || '—'}</td>
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

export default function App() {
  const [tab, setTab] = useState<TabKey>('all')
  const [contracts, setContracts] = useState<ContractRecord[]>([])
  const [renewals, setRenewals] = useState<RenewalRecord[]>([])
  const [generatedAt, setGeneratedAt] = useState('')
  const [expiringFilter, setExpiringFilter] = useState<ExpiringFilter>('this_month')
  const [customFrom, setCustomFrom] = useState('')
  const [customTo, setCustomTo] = useState('')
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

  const expiringContracts = useMemo(() => {
    const today = new Date()
    const normalizedToday = new Date(today.getFullYear(), today.getMonth(), today.getDate())

    return contracts.filter((contract) => {
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
  }, [contracts, expiringFilter, customFrom, customTo])

  async function handleAction(action: string, rowId: number, extra: Record<string, unknown> = {}) {
    setBusyRow(`${action}-${rowId}`)
    setError('')
    try {
      if (action === 'addNote') {
        const note = window.prompt('Add note')?.trim()
        if (!note) return
        await postAction(action, { rowId, note, ...extra })
      } else {
        await postAction(action, { rowId, ...extra })
      }
      await loadData()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Action failed.')
    } finally {
      setBusyRow('')
    }
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
          <p className="subtle">Main Contracts only shows final signed contracts. All in-progress renewal work stays in the Renewals sheet until finalized.</p>
        </div>
        <button className="refresh-button" onClick={() => void loadData()} type="button">
          <RefreshCcw size={16} /> Refresh
        </button>
      </header>

      <section className="tabs-card">
        <div className="tabs-row">
          {tabOptions.map((option) => (
            <button
              key={option.key}
              className={option.key === tab ? 'tab-button active' : 'tab-button'}
              onClick={() => setTab(option.key)}
              type="button"
            >
              {option.label}
            </button>
          ))}
        </div>

        {tab === 'expiring' && (
          <div className="filters-row">
            <select value={expiringFilter} onChange={(event) => setExpiringFilter(event.target.value as ExpiringFilter)}>
              {expiringOptions.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
            {expiringFilter === 'custom' && (
              <>
                <input type="date" value={customFrom} onChange={(event) => setCustomFrom(event.target.value)} />
                <input type="date" value={customTo} onChange={(event) => setCustomTo(event.target.value)} />
              </>
            )}
          </div>
        )}
      </section>

      {generatedAt && <p className="generated-at">Last synced: {generatedAt}</p>}
      {error && <div className="error-banner">{error}</div>}
      {loading ? (
        <div className="empty-card">Loading dashboard…</div>
      ) : (
        <>
          {tab === 'all' && (
            <div className="table-card">
              <table>
                <thead>
                  <tr>
                    <th>Talent Name</th>
                    <th>Email</th>
                    <th>Phone</th>
                    <th>Original Contract</th>
                    <th>Signed Contract</th>
                    <th>Contract Signed On</th>
                    <th>Contract Start Date</th>
                    <th>Contract Validity</th>
                    <th>Contract Expiry Date</th>
                    <th>Days Left</th>
                    <th>Version</th>
                    <th>AI Scan Status</th>
                    <th>Renewal Status</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {contracts.length === 0 ? (
                    <tr>
                      <td colSpan={14} className="empty-state">No signed contracts found.</td>
                    </tr>
                  ) : (
                    contracts.map((contract) => (
                      <tr key={contract.rowId}>
                        <DataCell>{contract.name}</DataCell>
                        <DataCell>{contract.email}</DataCell>
                        <DataCell>{contract.phone}</DataCell>
                        <DataCell><button className="link-button" onClick={() => safeOpen(contract.contractLink)} type="button">Open <ExternalLink size={14} /></button></DataCell>
                        <DataCell><button className="link-button" onClick={() => safeOpen(contract.signedPdfUrl)} type="button">Open <ExternalLink size={14} /></button></DataCell>
                        <DataCell>{formatDate(contract.contractSignedOn)}</DataCell>
                        <DataCell>{formatDate(contract.contractStartDate)}</DataCell>
                        <DataCell>{contract.contractValidity || '—'}</DataCell>
                        <DataCell>{formatDate(contract.contractExpiryDate)}</DataCell>
                        <DataCell>{formatDays(contract.daysLeft)}</DataCell>
                        <DataCell>{contract.version || '—'}</DataCell>
                        <DataCell>{contract.aiScanStatus || 'Pending'}</DataCell>
                        <DataCell>{contract.renewalStatus || '—'}</DataCell>
                        <td>
                          <div className="actions-stack">
                            <ActionButton onClick={() => safeOpen(contract.contractLink)}><ExternalLink size={14} /> Original</ActionButton>
                            <ActionButton onClick={() => safeOpen(contract.signedPdfUrl)}><ExternalLink size={14} /> Signed</ActionButton>
                            <ActionButton onClick={() => void handleAction('rescanWithAI', contract.rowId)}><RotateCw size={14} /> {busyRow === `rescanWithAI-${contract.rowId}` ? 'Working…' : 'Rescan with AI'}</ActionButton>
                            <ActionButton onClick={() => void handleAction('startRenewal', contract.rowId)}><FilePenLine size={14} /> {busyRow === `startRenewal-${contract.rowId}` ? 'Starting…' : 'Start Renewal'}</ActionButton>
                            <ActionButton onClick={() => void handleAction('addNote', contract.rowId)}><StickyNote size={14} /> Add Note</ActionButton>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          )}

          {tab === 'expiring' && (
            <div className="table-card">
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
                    <th>Status</th>
                    <th>Signed Contract</th>
                    <th>Renewal Status</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {expiringContracts.length === 0 ? (
                    <tr>
                      <td colSpan={11} className="empty-state">No matching contracts found.</td>
                    </tr>
                  ) : (
                    expiringContracts.map((contract) => (
                      <tr key={contract.rowId}>
                        <DataCell>{contract.name}</DataCell>
                        <DataCell>{contract.phone}</DataCell>
                        <DataCell>{contract.email}</DataCell>
                        <DataCell>{formatDate(contract.contractSignedOn)}</DataCell>
                        <DataCell>{formatDate(contract.contractStartDate)}</DataCell>
                        <DataCell>{formatDate(contract.contractExpiryDate)}</DataCell>
                        <DataCell>{formatDays(contract.daysLeft)}</DataCell>
                        <DataCell><span className={toneClass(contract.statusTone)}>{toneLabel(contract.statusTone)}</span></DataCell>
                        <DataCell><button className="link-button" onClick={() => safeOpen(contract.signedPdfUrl)} type="button">Open <ExternalLink size={14} /></button></DataCell>
                        <DataCell>{contract.renewalStatus || '—'}</DataCell>
                        <td>
                          <ActionButton onClick={() => void handleAction('startRenewal', contract.rowId)}><FilePenLine size={14} /> {busyRow === `startRenewal-${contract.rowId}` ? 'Starting…' : 'Start Renewal'}</ActionButton>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          )}

          {tab === 'renewals' && (
            <div className="table-card">
              <table>
                <thead>
                  <tr>
                    <th>Renewal ID</th>
                    <th>Original Row ID</th>
                    <th>Name / Contact</th>
                    <th>Versions</th>
                    <th>Old Expiry</th>
                    <th>Status</th>
                    <th>Contract</th>
                    <th>Zoho</th>
                    <th>Signed PDF</th>
                    <th>Final Dates</th>
                    <th>Notes</th>
                    <th>Error</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {renewals.length === 0 ? (
                    <tr>
                      <td colSpan={13} className="empty-state">No renewals yet.</td>
                    </tr>
                  ) : (
                    renewals.map((renewal) => {
                      const isEditing = editingRenewalRow === renewal.rowId
                      return (
                        <tr key={renewal.rowId}>
                          <DataCell>{renewal.renewalId}</DataCell>
                          <DataCell>{renewal.originalContractRowId}</DataCell>
                          <td>
                            {isEditing ? (
                              <div className="actions-stack">
                                <input value={String(renewalDraft.name ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, name: event.target.value }))} placeholder="Name" />
                                <input value={String(renewalDraft.email ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, email: event.target.value }))} placeholder="Email" />
                                <input value={String(renewalDraft.phone ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, phone: event.target.value }))} placeholder="Phone" />
                              </div>
                            ) : (
                              <div>
                                <div>{renewal.name}</div>
                                <div className="subtle">{renewal.email || '—'} • {renewal.phone || '—'}</div>
                              </div>
                            )}
                          </td>
                          <td>
                            {isEditing ? (
                              <div className="actions-stack">
                                <input value={String(renewalDraft.version ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, version: event.target.value }))} placeholder="Version" />
                                <input type="date" value={String(renewalDraft.contractExpiryDate ?? '')} onChange={(event) => setRenewalDraft((prev) => ({ ...prev, contractExpiryDate: event.target.value }))} />
                              </div>
                            ) : (
                              <div>
                                <div>{renewal.oldVersion} → {renewal.newVersion}</div>
                                {renewal.editableFields.contractExpiryDate ? <div className="subtle">Draft expiry: {renewal.editableFields.contractExpiryDate}</div> : null}
                              </div>
                            )}
                          </td>
                          <DataCell>{formatDate(renewal.oldExpiryDate)}</DataCell>
                          <td>
                            <div>{renewal.renewalStatus || '—'}</div>
                            {renewal.isReadyForFinalization ? <div className="subtle">Ready to sync final</div> : null}
                          </td>
                          <DataCell>{renewal.newContractLink ? <button className="link-button" onClick={() => safeOpen(renewal.newContractLink)} type="button">Open <ExternalLink size={14} /></button> : '—'}</DataCell>
                          <td>
                            <div>{renewal.newZohoStatus || '—'}</div>
                            {renewal.newZohoRequestId ? <div className="subtle">Req: {renewal.newZohoRequestId}</div> : null}
                          </td>
                          <DataCell>{renewal.newSignedPdfUrl ? <button className="link-button" onClick={() => safeOpen(renewal.newSignedPdfUrl)} type="button">Open <ExternalLink size={14} /></button> : '—'}</DataCell>
                          <td>
                            <div>{formatDate(renewal.newContractSignedOn)}</div>
                            <div className="subtle">{formatDate(renewal.newContractExpiryDate)}</div>
                          </td>
                          <DataCell>{renewal.notes || '—'}</DataCell>
                          <DataCell>{renewal.error || '—'}</DataCell>
                          <td>
                            <div className="actions-stack">
                              {isEditing ? (
                                <>
                                  <ActionButton onClick={() => void saveRenewal(renewal.rowId)} disabled={busyRow === `updateRenewal-${renewal.rowId}`}><Save size={14} /> {busyRow === `updateRenewal-${renewal.rowId}` ? 'Saving…' : 'Save'}</ActionButton>
                                  <ActionButton onClick={cancelEditingRenewal}><Ban size={14} /> Cancel</ActionButton>
                                </>
                              ) : (
                                <>
                                  <ActionButton onClick={() => startEditingRenewal(renewal)} disabled={!renewal.canEdit}><FilePenLine size={14} /> Edit</ActionButton>
                                  <ActionButton onClick={() => void handleAction('regenerateRenewalContract', renewal.rowId)} disabled={!renewal.canEdit}><FileText size={14} /> {busyRow === `regenerateRenewalContract-${renewal.rowId}` ? 'Regenerating…' : 'Regenerate'}</ActionButton>
                                  <ActionButton onClick={() => void handleAction('sendRenewalForSigning', renewal.rowId)} disabled={!renewal.canEdit}><Send size={14} /> {busyRow === `sendRenewalForSigning-${renewal.rowId}` ? 'Sending…' : 'Send'}</ActionButton>
                                  <ActionButton onClick={() => void handleAction('refreshRenewalZohoStatus', renewal.rowId)}><RefreshCw size={14} /> {busyRow === `refreshRenewalZohoStatus-${renewal.rowId}` ? 'Refreshing…' : 'Refresh Zoho'}</ActionButton>
                                  <ActionButton onClick={() => void handleAction('markRenewalSigned', renewal.rowId)}><CheckCircle2 size={14} /> Mark Signed</ActionButton>
                                  <ActionButton onClick={() => void handleAction('markRenewalRenewed', renewal.rowId)}><CheckCircle2 size={14} /> Mark Renewed</ActionButton>
                                  <ActionButton onClick={() => void handleAction('setRenewalOnHold', renewal.rowId)}><PauseCircle size={14} /> On Hold</ActionButton>
                                  <ActionButton onClick={() => void handleAction('cancelRenewal', renewal.rowId)}><Ban size={14} /> Cancel</ActionButton>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      )
                    })
                  )}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  )
}
