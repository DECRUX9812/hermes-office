import { useQuery } from '@hermes/plugin-sdk'

import { api, keys } from './api'
import { Empty, QueryError, Skeleton } from './ui'

function DocView({ docx }) {
  const paragraphs = docx?.paragraphs ?? []
  const comments = docx?.comments ?? []

  if (!paragraphs.length) {
    return <Empty>This document has no text paragraphs.</Empty>
  }

  return (
    <div className="hd-doc">
      {paragraphs.map((paragraph, index) => {
        const heading = /^(heading|title)/i.test(paragraph.style ?? '')

        return (
          <p className={heading ? 'hd-doc-p hd-doc-h' : 'hd-doc-p'} key={index} title={paragraph.style || undefined}>
            {(paragraph.runs ?? []).map((run, runIndex) => (
              <span
                className={run.rev === 'ins' ? 'hd-ins' : run.rev === 'del' ? 'hd-del' : undefined}
                key={runIndex}
                title={run.rev ? `${run.rev === 'ins' ? 'Inserted' : 'Deleted'}${run.author ? ` by ${run.author}` : ''}` : undefined}
              >
                {run.text}
              </span>
            ))}
          </p>
        )
      })}
      {comments.length ? (
        <div className="hd-comments">
          <div className="hd-subtitle">Comments</div>
          {comments.map(comment => (
            <div className={comment.resolved ? 'hd-comment hd-comment--resolved' : 'hd-comment'} key={comment.id}>
              <div className="hd-comment-meta">
                {comment.author || 'Unknown'}
                {comment.anchor ? <span className="hd-comment-anchor"> on “{comment.anchor}”</span> : null}
                {comment.resolved ? <span> · resolved</span> : null}
              </div>
              <div>{comment.text}</div>
            </div>
          ))}
        </div>
      ) : null}
    </div>
  )
}

function normalizeAddress(address) {
  return String(address ?? '')
    .replace(/\$/g, '')
    .replace(/^'(.*)'!/, '$1!')
    .toUpperCase()
}

function isChanged(changed, sheet, ref) {
  const cell = normalizeAddress(ref)

  return changed.has(`${String(sheet).toUpperCase()}!${cell}`) || changed.has(cell)
}

function SheetView({ xlsx, changed }) {
  const sheets = xlsx?.sheets ?? []

  if (!sheets.length) {
    return <Empty>This workbook has no sheets.</Empty>
  }

  return (
    <div className="hd-sheets">
      {sheets.map(sheet => (
        <div className="hd-sheet" key={sheet.name}>
          <div className="hd-subtitle">
            {sheet.name}
            {sheet.dims ? <span className="hd-muted"> · {sheet.dims}</span> : null}
          </div>
          <div className="hd-grid-wrap">
            <table className="hd-grid">
              <tbody>
                {(sheet.rows ?? []).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.map((cell, cellIndex) => (
                      <td
                        className={isChanged(changed, sheet.name, cell.ref) ? 'hd-cell hd-cell--changed' : 'hd-cell'}
                        key={cell.ref ?? cellIndex}
                        title={[cell.ref, cell.formula].filter(Boolean).join('  ') || undefined}
                      >
                        {cell.value === null || cell.value === undefined ? '' : String(cell.value)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  )
}

function pct(value, total) {
  return total ? `${(Number(value || 0) / total) * 100}%` : '0%'
}

function DeckView({ pptx }) {
  const slides = pptx?.slides ?? []
  const width = Number(pptx?.slide_w) || 0
  const height = Number(pptx?.slide_h) || 0

  if (!slides.length) {
    return <Empty>This deck has no slides.</Empty>
  }

  return (
    <div className="hd-deck">
      {slides.map(slide => (
        <div className="hd-slide-wrap" key={slide.index}>
          <div className="hd-subtitle">
            Slide {Number(slide.index) + 1}
            {slide.title ? <span className="hd-muted"> · {slide.title}</span> : null}
          </div>
          <div className="hd-slide" style={{ aspectRatio: width && height ? `${width} / ${height}` : '16 / 9' }}>
            {(slide.shapes ?? []).map((shape, index) => (
              <div
                className="hd-shape"
                key={`${shape.name}-${index}`}
                style={{ left: pct(shape.x, width), top: pct(shape.y, height), width: pct(shape.w, width), height: pct(shape.h, height) }}
                title={`${shape.name} (${shape.kind})`}
              >
                {shape.text}
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

export function PreviewBody({ preview, changed }) {
  if (preview.kind === 'docx') {
    return <DocView docx={preview.docx} />
  }

  if (preview.kind === 'xlsx') {
    return <SheetView changed={changed} xlsx={preview.xlsx} />
  }

  if (preview.kind === 'pptx') {
    return <DeckView pptx={preview.pptx} />
  }

  return <Empty>No preview for this file type.</Empty>
}

export function PreviewPane({ label, path, receiptId, version, changed, enabled = true, emptyText }) {
  const { data, error, isLoading, refetch } = useQuery({
    queryKey: keys.preview(path, receiptId, version),
    queryFn: () => api.preview(path, receiptId),
    enabled
  })

  return (
    <div className="hd-pane">
      <div className="hd-pane-label">{label}</div>
      <div className="hd-pane-body">
        {!enabled ? <Empty>{emptyText}</Empty> : null}
        {enabled && isLoading ? <Skeleton lines={6} /> : null}
        {enabled && error ? <QueryError error={error} onRetry={refetch} /> : null}
        {enabled && data ? <PreviewBody changed={changed} preview={data} /> : null}
      </div>
    </div>
  )
}

function texts(list) {
  return (list ?? []).map(part => part?.text ?? '').join(' ')
}

function DiffList({ items }) {
  return items.length ? (
    <ul className="hd-diff-list">
      {items.map((item, index) => (
        <li key={index}>{item}</li>
      ))}
    </ul>
  ) : null
}

function docxDiff(diff) {
  const revisions = diff.revisions ?? []
  const changed = diff.paragraphs_changed ?? []

  return {
    counts: [`${revisions.length} tracked change${revisions.length === 1 ? '' : 's'}`, `${changed.length} paragraph${changed.length === 1 ? '' : 's'} changed`],
    items: [
      ...revisions.map(rev => `${rev.revision_type}${rev.author ? ` (${rev.author})` : ''}: ${rev.text}`),
      ...changed.map(change => `“${change.before}” → “${change.after}”`)
    ]
  }
}

function xlsxDiff(diff) {
  const changed = diff.changed ?? []
  const shifted = diff.shifted ?? []
  const added = diff.added_sheets ?? []
  const removed = diff.removed_sheets ?? []

  return {
    counts: [
      `${changed.length} cell${changed.length === 1 ? '' : 's'} changed`,
      `${shifted.length} shifted`,
      ...(added.length ? [`sheets added: ${added.join(', ')}`] : []),
      ...(removed.length ? [`sheets removed: ${removed.join(', ')}`] : [])
    ],
    items: [
      ...changed.map(cell => `${cell.address}: ${String(cell.before ?? '∅')} → ${String(cell.after ?? '∅')}`),
      ...shifted.map(cell => `${cell.from} → ${cell.to}${cell.value !== undefined ? ` (${String(cell.value)})` : ''}`)
    ]
  }
}

function countOf(value) {
  return Array.isArray(value) ? value.length : Number(value) || 0
}

function pptxDiff(diff) {
  const slides = diff.slide_changes ?? []
  const changes = slides.flatMap(slide => (slide.text_changes ?? []).map(change => ({ slide, change })))

  return {
    counts: [
      `${countOf(diff.slides_added)} slides added`,
      `${countOf(diff.slides_removed)} removed`,
      `${countOf(diff.slides_moved)} moved`,
      `${changes.length} text change${changes.length === 1 ? '' : 's'}`
    ],
    items: changes.map(
      ({ slide, change }) =>
        `Slide ${slide.slide_id} · ${change.shape?.name ?? 'shape'} (${change.kind}): “${texts(change.before)}” → “${texts(change.after)}”`
    )
  }
}

const DIFF_SUMMARIES = { docx: docxDiff, xlsx: xlsxDiff, pptx: pptxDiff }

export function DiffSummary({ diff }) {
  const summarize = DIFF_SUMMARIES[diff?.kind]

  if (!summarize) {
    return null
  }

  const { counts, items } = summarize(diff)

  return (
    <details className="hd-diff">
      <summary className="hd-diff-summary">{counts.join(' · ')}</summary>
      <DiffList items={items} />
    </details>
  )
}

export function changedCells(diff) {
  return new Set(diff?.kind === 'xlsx' ? (diff.changed ?? []).map(cell => normalizeAddress(cell.address)) : [])
}
