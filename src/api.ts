import type { DashboardPayload, ExpiringFilter } from './types'

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
    throw new Error(payload.message || 'Request failed.')
  }
  return payload
}

export async function fetchDashboardData(filter: ExpiringFilter, from?: string, to?: string): Promise<DashboardPayload> {
  const url = new URL(requireBaseUrl())
  url.searchParams.set('action', 'dashboard')
  if (filter) url.searchParams.set('expiringFilter', filter)
  if (from) url.searchParams.set('from', from)
  if (to) url.searchParams.set('to', to)

  return parseResponse(await fetch(url.toString(), { method: 'GET' }))
}

export async function postAction(action: string, payload: Record<string, unknown>) {
  return parseResponse(
    await fetch(requireBaseUrl(), {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, ...payload }),
    }),
  )
}
