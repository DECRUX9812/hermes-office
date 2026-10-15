// Dev-harness stand-in for `@hermes/plugin-sdk`. It implements ONLY the exports
// desktop/plugin.js imports, so the real built plugin renders in a plain
// browser against the real backend. It never fabricates data: everything comes
// from the plugin's own REST calls, and the Hermes-only doors fail honestly.
import { createElement, useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react'

export const ROUTES_AREA = 'routes'
export const SIDEBAR_NAV_AREA = 'sidebar.nav'

// ── nanostores-shaped atom ────────────────────────────────────────────────
export function atom(initial) {
  let value = initial
  const listeners = new Set()

  return {
    get: () => value,
    set(next) {
      if (Object.is(next, value)) {
        return
      }

      value = next

      for (const listener of [...listeners]) {
        listener(value)
      }
    },
    listen(listener) {
      listeners.add(listener)

      return () => listeners.delete(listener)
    },
    subscribe(listener) {
      listeners.add(listener)
      listener(value)

      return () => listeners.delete(listener)
    }
  }
}

export function useValue(store) {
  return useSyncExternalStore(store.listen, store.get, store.get)
}

// ── a tiny React Query: fetch + interval + prefix invalidation ────────────
const cache = new Map()

function hashKey(queryKey) {
  return JSON.stringify(queryKey)
}

function entryFor(queryKey) {
  const hash = hashKey(queryKey)

  if (!cache.has(hash)) {
    cache.set(hash, { queryKey, data: undefined, error: null, status: 'idle', listeners: new Set(), fetcher: null, inflight: null })
  }

  return cache.get(hash)
}

function notify(entry) {
  for (const listener of [...entry.listeners]) {
    listener()
  }
}

function runQuery(entry) {
  if (!entry.fetcher) {
    return Promise.resolve()
  }

  if (entry.inflight) {
    return entry.inflight
  }

  entry.inflight = Promise.resolve()
    .then(() => entry.fetcher())
    .then(
      data => {
        entry.data = data
        entry.error = null
        entry.status = 'success'
      },
      error => {
        entry.error = error
        entry.status = 'error'
      }
    )
    .finally(() => {
      entry.inflight = null
      notify(entry)
    })

  return entry.inflight
}

function isPrefix(prefix, queryKey) {
  return prefix.every((part, index) => hashKey([part]) === hashKey([queryKey[index]]))
}

export const queryClient = {
  invalidateQueries({ queryKey = [] } = {}) {
    const runs = []

    for (const entry of cache.values()) {
      if (isPrefix(queryKey, entry.queryKey) && entry.listeners.size > 0) {
        runs.push(runQuery(entry))
      }
    }

    return Promise.all(runs).then(() => undefined)
  }
}

export function useQueryClient() {
  return queryClient
}

export function useQuery({ queryKey, queryFn, enabled = true, refetchInterval = false }) {
  const entry = entryFor(queryKey)
  const [, force] = useState(0)
  const fnRef = useRef(queryFn)
  fnRef.current = queryFn
  entry.fetcher = () => fnRef.current()

  const hash = hashKey(queryKey)

  useEffect(() => {
    const listener = () => force(n => n + 1)
    entry.listeners.add(listener)

    if (enabled) {
      if (entry.status === 'idle') {
        entry.status = 'loading'
      }

      void runQuery(entry)
    }

    return () => entry.listeners.delete(listener)
  }, [hash, enabled])

  useEffect(() => {
    if (!enabled || typeof refetchInterval !== 'number' || refetchInterval <= 0) {
      return undefined
    }

    const timer = setInterval(() => void runQuery(entry), refetchInterval)

    return () => clearInterval(timer)
  }, [hash, enabled, refetchInterval])

  const refetch = useCallback(() => runQuery(entry), [hash])
  const pending = enabled && (entry.status === 'idle' || entry.status === 'loading')

  return {
    data: entry.data,
    error: entry.error,
    isLoading: pending && entry.data === undefined,
    isError: entry.status === 'error',
    isSuccess: entry.status === 'success',
    refetch
  }
}

// ── host: only the members the plugin touches, failing honestly ───────────
const NO_GATEWAY = 'no Hermes gateway in the dev harness'

export const host = {
  request: () => Promise.reject(new Error(NO_GATEWAY)),
  onEvent: () => () => undefined,
  openSession: () => Promise.reject(new Error('no Hermes desktop in the dev harness: cannot open a chat session'))
}

// ── UI kit: Button only, styled by desktop/dev/theme.css ─────────────────
export function Button({ variant = 'default', size = 'default', className, type = 'button', ...props }) {
  return createElement('button', {
    ...props,
    type,
    className: ['shim-button', className].filter(Boolean).join(' '),
    'data-size': size,
    'data-slot': 'button',
    'data-variant': variant
  })
}
