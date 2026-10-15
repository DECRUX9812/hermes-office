import { Button, host, useQuery, useQueryClient } from '@hermes/plugin-sdk'
import { useState } from 'react'

import { api, errorText, keys } from './api'
import { startWorkstream } from './session'
import { $selectedWorkstream } from './state'
import { Empty, InlineError, QueryError, Skeleton, StatusPill, timeAgo } from './ui'

function HealthLine() {
  const { data, error, isLoading } = useQuery({ queryKey: keys.health, queryFn: api.health, refetchInterval: 60_000 })

  if (isLoading) {
    return <div className="hd-health hd-health--muted">Checking the office engines…</div>
  }

  if (error) {
    return <div className="hd-health hd-health--bad">Backend unreachable: {errorText(error)}</div>
  }

  const engines = data?.engines ?? {}
  const names = ['docx', 'pptx', 'xlsx']
  const missing = names.filter(name => !engines[name])

  return (
    <div className="hd-health">
      <span className={missing.length ? 'hd-health-bad' : 'hd-health-ok'}>
        {missing.length ? `Engines missing: ${missing.join(', ')}` : 'Engines ok (docx, pptx, xlsx)'}
      </span>
      <span className="hd-sep">·</span>
      <span>LibreOffice {data?.libreoffice ? 'yes' : 'no'}</span>
      {data?.version ? (
        <>
          <span className="hd-sep">·</span>
          <span>v{data.version}</span>
        </>
      ) : null}
    </div>
  )
}

function openWorkstream(id) {
  $selectedWorkstream.set(id)
}

function NeedsYou({ items, folderName }) {
  if (!items.length) {
    return null
  }

  return (
    <section className="hd-section hd-needs">
      <h2 className="hd-section-title hd-needs-title">Needs you</h2>
      <ul className="hd-list">
        {items.map(ws => (
          <li className="hd-row hd-row--needs hd-enter" key={ws.id}>
            <div className="hd-row-main">
              <div className="hd-row-title">{ws.title}</div>
              <div className="hd-row-meta">
                <span>{folderName(ws.folder_id)}</span>
                {ws.note ? (
                  <>
                    <span className="hd-sep">·</span>
                    <span className="hd-row-note">{ws.note}</span>
                  </>
                ) : null}
                <span className="hd-sep">·</span>
                <span>{timeAgo(ws.updated_at)}</span>
              </div>
            </div>
            <Button onClick={() => openWorkstream(ws.id)} size="sm" variant="secondary">
              Review
            </Button>
          </li>
        ))}
      </ul>
    </section>
  )
}

function AddFolder() {
  const client = useQueryClient()
  const [path, setPath] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const canPick = typeof host.pickOpenPaths === 'function'

  const add = async candidate => {
    const trimmed = String(candidate ?? '').trim()

    if (!trimmed) {
      setError('Enter the absolute path of a folder.')

      return
    }

    setBusy(true)
    setError('')

    try {
      await api.addFolder(trimmed)
      setPath('')
      await client.invalidateQueries({ queryKey: keys.desk })
    } catch (err) {
      setError(errorText(err))
    } finally {
      setBusy(false)
    }
  }

  const pick = async () => {
    setError('')

    try {
      const picked = await host.pickOpenPaths({ directories: true, title: 'Add a folder to your Desk' })

      if (picked?.[0]) {
        await add(picked[0])
      }
    } catch (err) {
      setError(errorText(err))
    }
  }

  return (
    <form
      className="hd-add"
      onSubmit={event => {
        event.preventDefault()
        void add(path)
      }}
    >
      <input
        aria-label="Folder path"
        className="hd-input"
        disabled={busy}
        onChange={event => setPath(event.target.value)}
        placeholder="/path/to/a/folder"
        spellCheck={false}
        value={path}
      />
      <Button disabled={busy} size="sm" type="submit" variant="secondary">
        Add folder
      </Button>
      {canPick ? (
        <Button disabled={busy} onClick={() => void pick()} size="sm" type="button" variant="ghost">
          Browse…
        </Button>
      ) : null}
      <InlineError message={error} />
    </form>
  )
}

function Composer({ folder, onClose }) {
  const client = useQueryClient()
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const submit = async () => {
    const request = text.trim()

    if (!request) {
      setError('Describe what you want done in this folder.')

      return
    }

    setBusy(true)
    setError('')

    try {
      const result = await startWorkstream(folder, request)
      await client.invalidateQueries({ queryKey: keys.desk })

      if (result.error) {
        setError(`${result.error} The request is saved on the Desk as queued.`)
        setText('')
      } else {
        onClose()
      }
    } catch (err) {
      setError(errorText(err))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="hd-composer hd-enter">
      <textarea
        aria-label={`New request for ${folder.name}`}
        className="hd-textarea"
        disabled={busy}
        onChange={event => setText(event.target.value)}
        onKeyDown={event => {
          if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
            event.preventDefault()
            void submit()
          }
        }}
        placeholder="e.g. Tighten the executive summary in the board memo and track the changes."
        rows={3}
        value={text}
      />
      <div className="hd-composer-actions">
        <InlineError message={error} />
        <Button disabled={busy} onClick={onClose} size="sm" variant="ghost">
          Cancel
        </Button>
        <Button disabled={busy} onClick={() => void submit()} size="sm">
          {busy ? 'Starting…' : 'Start'}
        </Button>
      </div>
    </div>
  )
}

function FolderCard({ folder }) {
  const [composing, setComposing] = useState(false)
  const counts = folder.counts ?? {}

  return (
    <li className="hd-card hd-enter">
      <div className="hd-card-head">
        <div className="hd-card-name">{folder.name}</div>
        {folder.counts?.needs_you ? <span className="hd-count-needs">{counts.needs_you} need you</span> : null}
      </div>
      <div className="hd-card-path" title={folder.path}>
        {folder.path}
      </div>
      <div className="hd-card-counts">
        <span>{counts.files ?? 0} files</span>
        <span className="hd-sep">·</span>
        <span>{counts.workstreams ?? 0} requests</span>
      </div>
      {composing ? (
        <Composer folder={folder} onClose={() => setComposing(false)} />
      ) : (
        <div className="hd-card-actions">
          <Button onClick={() => setComposing(true)} size="sm" variant="secondary">
            New request
          </Button>
        </div>
      )}
    </li>
  )
}

function Recent({ items, folderName }) {
  return (
    <section className="hd-section">
      <h2 className="hd-section-title">Recent</h2>
      {items.length ? (
        <ul className="hd-list">
          {items.map(ws => (
            <li key={ws.id}>
              <button className="hd-row hd-row--button" onClick={() => openWorkstream(ws.id)} type="button">
                <StatusPill status={ws.status} />
                <div className="hd-row-main">
                  <div className="hd-row-title">{ws.title}</div>
                  <div className="hd-row-meta">
                    <span>{folderName(ws.folder_id)}</span>
                    {ws.last_event?.summary ? (
                      <>
                        <span className="hd-sep">·</span>
                        <span className="hd-row-note">{ws.last_event.summary}</span>
                      </>
                    ) : null}
                  </div>
                </div>
                <div className="hd-row-side">
                  {ws.files_changed ? <span>{ws.files_changed} changed</span> : null}
                  <span>{timeAgo(ws.updated_at)}</span>
                </div>
              </button>
            </li>
          ))}
        </ul>
      ) : (
        <Empty>No requests yet. Start one from a folder above.</Empty>
      )}
    </section>
  )
}

export function DeskHome() {
  const { data, error, isLoading, refetch } = useQuery({ queryKey: keys.desk, queryFn: api.desk, refetchInterval: 15_000 })
  const folders = data?.folders ?? []
  const folderName = id => folders.find(folder => folder.id === id)?.name ?? 'Unknown folder'

  return (
    <div className="hd-page">
      <header className="hd-header">
        <h1 className="hd-title">Desk</h1>
        <HealthLine />
      </header>

      {isLoading ? <Skeleton lines={4} /> : null}
      {error ? <QueryError error={error} onRetry={refetch} /> : null}

      {data ? (
        <>
          <NeedsYou folderName={folderName} items={data.needs_you ?? []} />

          <section className="hd-section">
            <h2 className="hd-section-title">On your desk</h2>
            <AddFolder />
            {folders.length ? (
              <ul className="hd-cards">
                {folders.map(folder => (
                  <FolderCard folder={folder} key={folder.id} />
                ))}
              </ul>
            ) : (
              <Empty>No folders yet. Add a folder of Word, Excel or PowerPoint files to start.</Empty>
            )}
          </section>

          <Recent folderName={folderName} items={data.recent ?? []} />
        </>
      ) : null}
    </div>
  )
}
