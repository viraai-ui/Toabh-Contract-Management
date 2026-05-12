export type TabKey = 'all' | 'expiring' | 'renewals'

export type ExpiringFilter =
  | 'this_month'
  | 'next_month'
  | 'next_3_months'
  | 'next_6_months'
  | 'already_expired'
  | 'custom'

export interface ContractRecord {
  rowId: number
  name: string
  email: string
  phone: string
  contractLink: string
  signedPdfUrl: string
  version: string
  contractSignedOn: string
  contractStartDate: string
  contractValidity: string
  contractExpiryDate: string
  daysLeft: number | null
  aiScanStatus: string
  aiScanNotes: string
  renewalStatus: string
  renewalSheetRowId: string
  lastSyncedAt: string
  notes: string
  statusTone: 'active' | 'due-soon' | 'urgent' | 'expired' | 'unknown'
}

export interface RenewalEditableFields {
  name?: string
  email?: string
  phone?: string
  version?: string
  contractStartDate?: string
  contractValidity?: string
  contractExpiryDate?: string
  noKycRequired?: boolean
  [key: string]: unknown
}

export interface RenewalRecord {
  rowId: number
  renewalId: string
  originalContractRowId: string
  name: string
  email: string
  phone: string
  oldVersion: string
  newVersion: string
  oldContractLink: string
  oldSignedPdfUrl: string
  oldExpiryDate: string
  renewalStatus: string
  renewalStartedOn: string
  editableDataJson: string
  editableFields: RenewalEditableFields
  newContractLink: string
  newZohoRequestId: string
  newZohoStatus: string
  newSignedPdfUrl: string
  newContractSignedOn: string
  newContractExpiryDate: string
  notes: string
  error: string
  canEdit: boolean
  isReadyForFinalization: boolean
}

export interface DashboardPayload {
  ok: boolean
  contracts: ContractRecord[]
  renewals: RenewalRecord[]
  generatedAt: string
  message?: string
}
