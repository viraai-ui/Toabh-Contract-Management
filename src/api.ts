import type { ContractRecord, DashboardPayload, ExpiringFilter, RenewalRecord } from './types'

const BASE_URL = import.meta.env.VITE_APPS_SCRIPT_URL?.trim()

function requireBaseUrl() {
  if (!BASE_URL) {
    throw new Error('Missing VITE_APPS_SCRIPT_URL in your environment.')
  }
  return BASE_URL
}

async function parseResponse(response: Response) {
  if (!response.ok) {
    throw new Error(`Request failed with status ${response.status}`)
  }
  const payload = await response.json()
  if (payload && payload.ok === false) {
    throw new Error(payload.message || payload.error || 'Request failed.')
  }
  if (payload && payload.error) {
    throw new Error(payload.error)
  }
  return payload
}

async function getAction(action: string, params: Record<string, string> = {}) {
  const url = new URL(requireBaseUrl())
  url.searchParams.set('action', action)
  Object.entries(params).forEach(([key, value]) => {
    if (value) url.searchParams.set(key, value)
  })
  return parseResponse(await fetch(url.toString(), { method: 'GET' }))
}

function numberOrNull(value: unknown) {
  if (value === null || value === undefined || value === '') return null
  const num = Number(value)
  return Number.isNaN(num) ? null : num
}

function stringValue(value: unknown) {
  return value === null || value === undefined ? '' : String(value)
}

function deriveStatusTone(daysLeft: number | null): ContractRecord['statusTone'] {
  if (daysLeft === null || Number.isNaN(daysLeft)) return 'unknown'
  if (daysLeft < 0) return 'expired'
  if (daysLeft < 30) return 'urgent'
  if (daysLeft <= 90) return 'due-soon'
  return 'active'
}

function normalizeContract(contract: any): ContractRecord | null {
  const signedPdfUrl = stringValue(contract?.signedPdfUrl)
  if (!signedPdfUrl) return null
  const daysLeft = numberOrNull(contract?.daysLeft)
  return {
    rowId: Number(contract?.rowId || 0),
    name: stringValue(contract?.name),
    email: stringValue(contract?.email),
    phone: stringValue(contract?.phone),
    contractLink: stringValue(contract?.contractLink),
    signedPdfUrl,
    version: stringValue(contract?.version),
    contractSignedOn: stringValue(contract?.contractSignedOn),
    contractStartDate: stringValue(contract?.contractStartDate),
    contractValidity: stringValue(contract?.contractValidity),
    contractExpiryDate: stringValue(contract?.contractExpiryDate),
    daysLeft,
    aiScanStatus: stringValue(contract?.aiScanStatus),
    aiScanNotes: stringValue(contract?.aiScanNotes),
    renewalStatus: stringValue(contract?.renewalStatus),
    renewalSheetRowId: stringValue(contract?.renewalSheetRowId),
    lastSyncedAt: stringValue(contract?.lastSyncedAt),
    notes: stringValue(contract?.notes),
    statusTone: deriveStatusTone(daysLeft),
  }
}

function normalizeRenewal(renewal: any): RenewalRecord {
  return {
    rowId: Number(renewal?.rowId || 0),
    renewalId: stringValue(renewal?.renewalId),
    originalContractRowId: stringValue(renewal?.originalContractRowId),
    name: stringValue(renewal?.name),
    email: stringValue(renewal?.email),
    phone: stringValue(renewal?.phone),
    oldVersion: stringValue(renewal?.oldVersion),
    newVersion: stringValue(renewal?.newVersion),
    oldContractLink: stringValue(renewal?.oldContractLink),
    oldSignedPdfUrl: stringValue(renewal?.oldSignedPdfUrl),
    oldExpiryDate: stringValue(renewal?.oldExpiryDate),
    renewalStatus: stringValue(renewal?.renewalStatus),
    renewalStartedOn: stringValue(renewal?.renewalStartedOn),
    editableDataJson: stringValue(renewal?.editableDataJson),
    editableFields: renewal?.editableFields && typeof renewal.editableFields === 'object' ? renewal.editableFields : {},
    newContractLink: stringValue(renewal?.newContractLink),
    newZohoRequestId: stringValue(renewal?.newZohoRequestId),
    newZohoStatus: stringValue(renewal?.newZohoStatus),
    newSignedPdfUrl: stringValue(renewal?.newSignedPdfUrl),
    newContractSignedOn: stringValue(renewal?.newContractSignedOn),
    newContractExpiryDate: stringValue(renewal?.newContractExpiryDate),
    notes: stringValue(renewal?.notes),
    error: stringValue(renewal?.error),
    canEdit: Boolean(renewal?.canEdit ?? true),
    isReadyForFinalization: Boolean(renewal?.isReadyForFinalization),
  }
}

export async function fetchDashboardData(filter: ExpiringFilter, from?: string, to?: string): Promise<DashboardPayload> {
  const [contractsPayload, renewalsPayload] = await Promise.all([
    getAction('contracts', { signedOnly: '1', expiringFilter: filter, from: from || '', to: to || '' }),
    getAction('renewals'),
  ])

  const contracts = Array.isArray(contractsPayload?.contracts)
    ? contractsPayload.contracts.map(normalizeContract).filter(Boolean) as ContractRecord[]
    : []

  const renewals = Array.isArray(renewalsPayload?.renewals)
    ? renewalsPayload.renewals.map(normalizeRenewal)
    : []

  return {
    ok: true,
    contracts,
    renewals,
    generatedAt: stringValue(renewalsPayload?.generatedAt || contractsPayload?.generatedAt || new Date().toISOString()),
  }
}

const ACTION_ALIASES: Record<string, string> = {
  rescanWithAI: 'rescan-ai',
  startRenewal: 'start-renewal',
  updateRenewal: 'update-renewal',
  regenerateRenewalContract: 'generate-renewal-contract',
  sendRenewalForSigning: 'send-renewal-for-signing',
  refreshRenewalZohoStatus: 'refresh-renewal-status',
  addNote: 'update-note',
}

export async function postAction(action: string, payload: Record<string, unknown>) {
  const routedAction = ACTION_ALIASES[action] || action
  const url = new URL(requireBaseUrl())
  url.searchParams.set('action', routedAction)

  return parseResponse(
    await fetch(url.toString(), {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: routedAction, ...payload }),
    }),
  )
}
