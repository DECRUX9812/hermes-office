import { Button } from '@hermes/plugin-sdk'

import { errorText } from './api'

const STATUS_LABELS = {
  running: 'Running',
  queued: 'Queued',
  needs_you: 'Needs you',
  done: 'Done',
  error: 'Error'
}

export function StatusPill({ status }) {
  const known = STATUS_LABELS[status] ? status : 'unknown'

  return (
    <span className={`hd-pill hd-pill--${known}`}>
      <span aria-hidden className="hd-pill-dot" />
      {STATUS_LABELS[status] ?? String(status ?? 'Unknown')}
    </span>
  )
}

function toMillis(ts) {
  if (ts === null || ts === undefined || ts === '') {
    return null
  }

  if (typeof ts === 'number') {
    // Backend mtimes are epoch seconds; event timestamps are ISO strings.
    return ts < 1e12 ? ts * 1000 : ts
  }

  const ms = Date.parse(ts)

  return Number.isNaN(ms) ? null : ms
}

export function timeAgo(ts) {
  const ms = toMillis(ts)

  if (ms === null) {
    return ''
  }

  const seconds = Math.max(0, Math.round((Date.now() - ms) / 1000))

  if (seconds < 45) {
    return 'just now'
  }

  const minutes = Math.round(seconds / 60)

  if (minutes < 60) {
    return `${minutes} min ago`
  }

  const hours = Math.round(minutes / 60)

  if (hours < 24) {
    return `${hours} h ago`
  }

  const days = Math.round(hours / 24)

  return days < 30 ? `${days} d ago` : new Date(ms).toLocaleDateString()
}

export function clockTime(ts) {
  const ms = toMillis(ts)

  return ms === null ? '' : new Date(ms).toLocaleString()
}

const KIND_LABELS = { docx: 'DOCX', pptx: 'PPTX', xlsx: 'XLSX' }

export function KindMark({ kind }) {
  return <span className={`hd-kind hd-kind--${KIND_LABELS[kind] ? kind : 'other'}`}>{KIND_LABELS[kind] ?? 'FILE'}</span>
}

export function baseName(path) {
  return String(path ?? '').split(/[\\/]/).filter(Boolean).pop() ?? ''
}

export function Skeleton({ lines = 3 }) {
  return (
    <div aria-busy="true" className="hd-skeleton">
      {Array.from({ length: lines }, (_, index) => (
        <span className="hd-skeleton-line" key={index} />
      ))}
    </div>
  )
}

export function QueryError({ error, onRetry }) {
  return (
    <div className="hd-error" role="alert">
      <span>{errorText(error)}</span>
      {onRetry ? (
        <Button onClick={() => onRetry()} size="xs" variant="outline">
          Retry
        </Button>
      ) : null}
    </div>
  )
}

export function Empty({ children }) {
  return <div className="hd-empty">{children}</div>
}

export function InlineError({ message }) {
  return message ? (
    <div className="hd-inline-error" role="alert">
      {message}
    </div>
  ) : null
}
