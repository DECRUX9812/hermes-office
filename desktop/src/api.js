// The plugin's own backend namespace (`/api/plugins/hermes-office`), reached only
// through the `ctx.rest` / `ctx.os` doors handed to register(). Components read
// them from here because a route render has no `ctx` in reach.

export const KEY = 'hermes-office'

const bound = { rest: null, os: null }

export function bindContext(ctx) {
  bound.rest = ctx.rest
  bound.os = ctx.os

  return () => {
    bound.rest = null
    bound.os = null
  }
}

export function osDoor() {
  return bound.os
}

function rest(path, opts) {
  if (!bound.rest) {
    return Promise.reject(new Error('Desk is not loaded'))
  }

  return bound.rest(path, opts)
}

function query(params) {
  const search = new URLSearchParams()

  for (const [name, value] of Object.entries(params)) {
    if (value !== undefined && value !== null && value !== '') {
      search.set(name, String(value))
    }
  }

  return search.toString()
}

// The host rejects a failed REST call with `Error("<status>: <body>")`; FastAPI
// bodies carry `{"detail": ...}`, which is the text worth showing inline.
export function errorText(error) {
  const raw = error instanceof Error ? error.message : String(error ?? '')
  const body = raw.replace(/^\d{3}:\s*/, '')

  try {
    const parsed = JSON.parse(body)

    if (parsed && typeof parsed.detail === 'string') {
      return parsed.detail
    }

    if (parsed && parsed.detail !== undefined) {
      return JSON.stringify(parsed.detail)
    }
  } catch {
    // Not JSON: the message itself is the best text available.
  }

  return body || 'Request failed'
}

export const api = {
  health: () => rest('/health'),
  desk: () => rest('/desk'),
  addFolder: path => rest('/folders', { method: 'POST', body: { path } }),
  createWorkstream: body => rest('/workstreams', { method: 'POST', body }),
  patchWorkstream: (id, body) => rest(`/workstreams/${encodeURIComponent(id)}`, { method: 'PATCH', body }),
  workstream: id => rest(`/workstreams/${encodeURIComponent(id)}`),
  conversation: id => rest(`/workstreams/${encodeURIComponent(id)}/conversation`),
  preview: (path, receiptId) => rest(`/preview?${query({ path, receipt_id: receiptId })}`),
  diff: (path, receiptId) => rest(`/diff?${query({ path, receipt_id: receiptId })}`),
  redline: body => rest('/redline', { method: 'POST', body }),
  revert: receiptId => rest(`/receipts/${encodeURIComponent(receiptId)}/revert`, { method: 'POST', body: {} })
}

export const keys = {
  all: [KEY],
  health: [KEY, 'health'],
  desk: [KEY, 'desk'],
  workstream: id => [KEY, 'workstream', id],
  conversation: id => [KEY, 'conversation', id],
  preview: (path, receiptId, version) => [KEY, 'preview', path, receiptId ?? 'current', version],
  diff: (path, receiptId, version) => [KEY, 'diff', path, receiptId, version]
}

export const ACTIVE_STATUSES = new Set(['running', 'queued'])
