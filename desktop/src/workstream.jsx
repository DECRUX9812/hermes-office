import { Button, host, useQuery, useQueryClient } from '@hermes/plugin-sdk'
import { useEffect, useState } from 'react'

import { ACTIVE_STATUSES, api, errorText, keys, osDoor } from './api'
import { changedCells, DiffSummary, PreviewPane } from './preview'
import { $selectedWorkstream } from './state'
import { baseName, clockTime, Empty, InlineError, KindMark, QueryError, Skeleton, StatusPill, timeAgo } from './ui'

const OPEN_LABELS = { docx: 'Open in Word', xlsx: 'Open in Excel', pptx: 'Open in PowerPoint' }
const STATE_LABELS = { working: 'Working', changed: 'Changed', created: 'Created' }

function ActivityIsland({ workstream, now, events }) {
  const last = workstream.last_event ?? events[events.length - 1] ?? null

  return (
    <div className={`hd-island hd-island--${workstream.status}`}>
      <StatusPill status={workstream.status} />
      <div className="hd-island-text">
        {now ? (
          <>
            <span className="hd-island-label">{now.label}</span>
            {now.tool ? <span className="hd-island-tool">{now.tool}</span> : null}
            {now.path ? <span className="hd-island-file">{baseName(now.path)}</span> : null}
            {now.since ? <span className="hd-muted">since {timeAgo(now.since)}</span> : null}
          </>
        ) : last ? (
          <>
            <span className="hd-island-label">{last.summary || last.kind}</span>
            <span className="hd-muted">{timeAgo(last.ts)}</span>
          </>
        ) : (
          <span className="hd-muted">No activity recorded yet.</span>
        )}
      </div>
      {workstream.note ? <div className="hd-island-note">{workstream.note}</div> : null}
    </div>
  )
}

function Conversation({ workstream, active }) {
  const { data, error, isLoading, refetch } = useQuery({
    queryKey: keys.conversation(workstream.id),
    queryFn: () => api.conversation(workstream.id),
    refetchInterval: active ? 4000 : false
  })
  const [openError, setOpenError] = useState('')

  const continueInChat = async () => {
    setOpenError('')

    try {
      await host.openSession(workstream.session_id)
    } catch (err) {
      setOpenError(errorText(err))
    }
  }

  return (
    <section className="hd-col hd-col--chat">
      <div className="hd-col-head">
        <h3 className="hd-col-title">Conversation</h3>
        {workstream.session_id && typeof host.openSession === 'function' ? (
          <Button onClick={() => void continueInChat()} size="xs" variant="outline">
            Continue in chat
          </Button>
        ) : null}
      </div>
      <InlineError message={openError} />
      <div className="hd-chat">
        {isLoading ? <Skeleton lines={5} /> : null}
        {error ? <QueryError error={error} onRetry={refetch} /> : null}
        {data && !data.available ? (
          <Empty>
            {data.reason === 'missing_session_id'
              ? 'No Hermes session is attached to this request yet.'
              : data.reason || 'The conversation is not available.'}
          </Empty>
        ) : null}
        {data?.available
          ? (data.messages ?? []).map((message, index) => (
              <div className={`hd-bubble hd-bubble--${message.role}`} key={index}>
                <div className="hd-bubble-role">
                  {message.role}
                  {message.ts ? <span className="hd-muted"> · {timeAgo(message.ts)}</span> : null}
                </div>
                <div className="hd-bubble-text">{message.text}</div>
              </div>
            ))
          : null}
      </div>
    </section>
  )
}

function FileList({ files, selected, onSelect }) {
  return (
    <section className="hd-col hd-col--files">
      <div className="hd-col-head">
        <h3 className="hd-col-title">Files</h3>
      </div>
      {files.length ? (
        <ul className="hd-files">
          {files.map(file => (
            <li key={file.path}>
              <button
                aria-pressed={file.path === selected}
                className={file.path === selected ? 'hd-file hd-file--selected' : 'hd-file'}
                onClick={() => onSelect(file.path)}
                title={file.path}
                type="button"
              >
                <KindMark kind={file.kind} />
                <span className="hd-file-name">{file.rel || baseName(file.path)}</span>
                <span className={`hd-state hd-state--${file.state}`}>{STATE_LABELS[file.state] ?? file.state}</span>
              </button>
            </li>
          ))}
        </ul>
      ) : (
        <Empty>No files touched yet.</Empty>
      )}
    </section>
  )
}

function RedlineActions({ path, workstreamId, onDone }) {
  const [busy, setBusy] = useState('')
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')

  const run = async action => {
    setBusy(action)
    setError('')
    setMessage('')

    try {
      const result = await api.redline({ path, action, workstream_id: workstreamId })
      setMessage(`${action === 'accept' ? 'Accepted' : 'Rejected'} ${result.resolved ?? 0} change${result.resolved === 1 ? '' : 's'}.`)
      await onDone()
    } catch (err) {
      setError(errorText(err))
    } finally {
      setBusy('')
    }
  }

  return (
    <>
      <Button disabled={Boolean(busy)} onClick={() => void run('accept')} size="sm">
        {busy === 'accept' ? 'Accepting…' : 'Accept changes'}
      </Button>
      <Button disabled={Boolean(busy)} onClick={() => void run('reject')} size="sm" variant="outline">
        {busy === 'reject' ? 'Rejecting…' : 'Reject changes'}
      </Button>
      {message ? <span className="hd-ok">{message}</span> : null}
      <InlineError message={error} />
    </>
  )
}

function OpenExternal({ file }) {
  const os = osDoor()
  const [error, setError] = useState('')

  if (typeof os?.openExternal !== 'function') {
    return null
  }

  const open = async () => {
    setError('')
    const opened = await os.openExternal(file.path)

    if (!opened) {
      setError('This window cannot open files with the system app.')
    }
  }

  return (
    <>
      <Button onClick={() => void open()} size="sm" variant="ghost">
        {OPEN_LABELS[file.kind] ?? 'Open'}
      </Button>
      <InlineError message={error} />
    </>
  )
}

function hasTrackedRuns(preview) {
  return (preview?.docx?.paragraphs ?? []).some(paragraph => (paragraph.runs ?? []).some(run => run.rev))
}

function FileDetail({ file, receipts, workstream, version, onChanged }) {
  const fileReceipts = receipts.filter(receipt => receipt.path === file.path)
  const baseline = fileReceipts.reduce((first, receipt) => (first === null || receipt.id < first.id ? receipt : first), null)
  const baselineId = baseline?.id

  const diffQuery = useQuery({
    queryKey: keys.diff(file.path, baselineId, version),
    queryFn: () => api.diff(file.path, baselineId),
    enabled: baselineId !== undefined
  })
  const afterQuery = useQuery({
    queryKey: keys.preview(file.path, undefined, version),
    queryFn: () => api.preview(file.path)
  })

  const changed = changedCells(diffQuery.data)
  const showRedline =
    file.kind === 'docx' && ((diffQuery.data?.revisions ?? []).length > 0 || hasTrackedRuns(afterQuery.data))

  return (
    <section className="hd-col hd-col--preview">
      <div className="hd-col-head">
        <h3 className="hd-col-title" title={file.path}>
          {file.rel || baseName(file.path)}
        </h3>
        <div className="hd-actions">
          {showRedline ? <RedlineActions onDone={onChanged} path={file.path} workstreamId={workstream.id} /> : null}
          <OpenExternal file={file} />
        </div>
      </div>
      {diffQuery.error ? <QueryError error={diffQuery.error} onRetry={diffQuery.refetch} /> : null}
      {diffQuery.data ? <DiffSummary diff={diffQuery.data} /> : null}
      <div className="hd-panes">
        <PreviewPane
          changed={changed}
          emptyText="No receipt yet, so there is no saved original to compare against."
          enabled={baselineId !== undefined}
          label="Before"
          path={file.path}
          receiptId={baselineId}
          version={version}
        />
        <PreviewPane changed={changed} label="After" path={file.path} version={version} />
      </div>
    </section>
  )
}

function ReceiptRow({ receipt, onChanged }) {
  const [confirming, setConfirming] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const revert = async () => {
    setBusy(true)
    setError('')

    try {
      await api.revert(receipt.id)
      setConfirming(false)
      await onChanged()
    } catch (err) {
      setError(errorText(err))
    } finally {
      setBusy(false)
    }
  }

  return (
    <li className="hd-receipt">
      <div className="hd-row-main">
        <div className="hd-row-title">
          {receipt.tool} · {baseName(receipt.path)}
        </div>
        <div className="hd-row-meta">
          <span>{receipt.op_summary}</span>
          <span className="hd-sep">·</span>
          <span title={clockTime(receipt.created_at)}>{timeAgo(receipt.created_at)}</span>
          {receipt.reverted_at ? (
            <>
              <span className="hd-sep">·</span>
              <span>reverted {timeAgo(receipt.reverted_at)}</span>
            </>
          ) : null}
        </div>
        <InlineError message={error} />
      </div>
      {receipt.reverted_at || !receipt.backup_path ? null : confirming ? (
        <div className="hd-actions">
          <span className="hd-muted">Restore the file from before this change?</span>
          <Button disabled={busy} onClick={() => void revert()} size="xs" variant="destructive">
            {busy ? 'Reverting…' : 'Revert'}
          </Button>
          <Button disabled={busy} onClick={() => setConfirming(false)} size="xs" variant="ghost">
            Cancel
          </Button>
        </div>
      ) : (
        <Button onClick={() => setConfirming(true)} size="xs" variant="outline">
          Revert
        </Button>
      )}
    </li>
  )
}

function Receipts({ receipts, onChanged }) {
  if (!receipts.length) {
    return null
  }

  return (
    <section className="hd-section">
      <h3 className="hd-col-title">Receipts</h3>
      <ul className="hd-list">
        {receipts.map(receipt => (
          <ReceiptRow key={receipt.id} onChanged={onChanged} receipt={receipt} />
        ))}
      </ul>
    </section>
  )
}

function Timeline({ events }) {
  return (
    <details className="hd-timeline">
      <summary className="hd-diff-summary">Events ({events.length})</summary>
      {events.length ? (
        <ul className="hd-events">
          {events.map(event => (
            <li className={event.kind === 'refused' ? 'hd-event hd-event--refused' : 'hd-event'} key={event.id}>
              <span className="hd-event-ts" title={clockTime(event.ts)}>
                {timeAgo(event.ts)}
              </span>
              <span className="hd-event-kind">{event.kind}</span>
              <span className="hd-event-tool">{event.tool ?? ''}</span>
              <span className="hd-event-summary">{event.summary}</span>
            </li>
          ))}
        </ul>
      ) : (
        <Empty>No events yet.</Empty>
      )}
    </details>
  )
}

export function WorkstreamView({ id }) {
  const client = useQueryClient()
  const [activeHint, setActiveHint] = useState(true)
  const { data, error, isLoading, refetch } = useQuery({
    queryKey: keys.workstream(id),
    queryFn: () => api.workstream(id),
    refetchInterval: activeHint ? 4000 : false
  })
  const [selected, setSelected] = useState(null)

  const workstream = data?.workstream
  const active = ACTIVE_STATUSES.has(workstream?.status)
  const files = data?.files ?? []
  const receipts = data?.receipts ?? []
  const events = data?.events ?? []
  const file = files.find(item => item.path === selected) ?? files[0] ?? null

  useEffect(() => {
    if (workstream) {
      setActiveHint(active)
    }
  }, [workstream, active])

  const refreshAll = () => client.invalidateQueries({ queryKey: keys.all })
  const version = `${workstream?.updated_at ?? ''}:${receipts.length}`

  return (
    <div className="hd-page">
      <button className="hd-back" onClick={() => $selectedWorkstream.set(null)} type="button">
        ← Desk
      </button>

      {isLoading ? <Skeleton lines={5} /> : null}
      {error ? <QueryError error={error} onRetry={refetch} /> : null}

      {workstream ? (
        <>
          <header className="hd-ws-header">
            <h1 className="hd-title">{workstream.title}</h1>
            <div className="hd-row-meta">
              <span>Started {timeAgo(workstream.created_at)}</span>
              <span className="hd-sep">·</span>
              <span>{workstream.files_changed ?? 0} files changed</span>
            </div>
          </header>
          <ActivityIsland events={events} now={data.now} workstream={workstream} />

          <div className="hd-ws-grid">
            <Conversation active={active} workstream={workstream} />
            <FileList files={files} onSelect={setSelected} selected={file?.path ?? null} />
            {file ? (
              <FileDetail
                file={file}
                key={file.path}
                onChanged={refreshAll}
                receipts={receipts}
                version={version}
                workstream={workstream}
              />
            ) : (
              <section className="hd-col hd-col--preview">
                <Empty>Select a file to preview it.</Empty>
              </section>
            )}
          </div>

          <Receipts onChanged={refreshAll} receipts={receipts} />
          <Timeline events={events} />
        </>
      ) : null}
    </div>
  )
}
