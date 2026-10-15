// Every colour, radius and font comes from the app's theme variables so the page
// reskins with the active theme. Radius/easing fallbacks are geometry only and
// match the values in apps/desktop/src/design/tokens.css.
const tint = (token, amount) => `color-mix(in srgb, var(${token}) ${amount}%, transparent)`

export const STYLE_ID = 'hermes-office-desk-style'

export const DESK_CSS = `
.hd-page {
  --hd-radius-control: var(--radius-control, 0.375rem);
  --hd-radius-row: var(--radius-row, 0.4375rem);
  --hd-radius-card: var(--radius-card, 0.625rem);
  --hd-radius-panel: var(--radius-panel, 0.75rem);
  --hd-ease: var(--ease-standard, cubic-bezier(0.2, 0, 0, 1));
  --hd-fast: var(--dur-fast, 120ms);
  --hd-base: var(--dur-base, 180ms);
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  gap: 1.25rem;
  height: 100%;
  overflow: auto;
  padding: 1.25rem 1.5rem 2rem;
  color: var(--ui-text-primary);
  font-size: 0.8125rem;
  line-height: 1.25rem;
}
.hd-page *, .hd-page *::before, .hd-page *::after { box-sizing: border-box; }
.hd-page h1, .hd-page h2, .hd-page h3, .hd-page p, .hd-page ul { margin: 0; padding: 0; }
.hd-page ul { list-style: none; }

.hd-header, .hd-ws-header { display: flex; flex-direction: column; gap: 0.25rem; }
.hd-title { font-size: 1.25rem; line-height: 1.75rem; font-weight: 600; color: var(--ui-text-primary); }
.hd-health { display: flex; flex-wrap: wrap; gap: 0.375rem; color: var(--ui-text-tertiary); font-size: 0.75rem; min-height: 1rem; }
.hd-health-ok { color: var(--ui-text-secondary); }
.hd-health-bad, .hd-health--bad { color: var(--ui-red); }
.hd-health--muted { color: var(--ui-text-quaternary); }
.hd-sep { color: var(--ui-text-quaternary); }
.hd-muted { color: var(--ui-text-tertiary); }
.hd-ok { color: var(--ui-green); font-size: 0.75rem; }

.hd-section { display: flex; flex-direction: column; gap: 0.5rem; }
.hd-section-title, .hd-col-title { font-size: 0.75rem; line-height: 1rem; font-weight: 600; color: var(--ui-text-tertiary); }
.hd-needs { padding: 0.75rem; border-radius: var(--hd-radius-panel); background: ${tint('--ui-yellow', 8)}; box-shadow: inset 0 0 0 1px ${tint('--ui-yellow', 30)}; }
.hd-needs-title { color: var(--ui-yellow); }

.hd-list { display: flex; flex-direction: column; gap: 0.125rem; }
.hd-row {
  display: flex; align-items: center; gap: 0.75rem; width: 100%;
  min-height: 2.5rem; padding: 0.375rem 0.5rem;
  border-radius: var(--hd-radius-row); color: inherit; font: inherit; text-align: left;
  transition: background-color var(--hd-fast) var(--hd-ease);
}
.hd-row--button { border: 0; background: transparent; cursor: pointer; }
.hd-row--button:hover { background: var(--ui-row-hover-background); }
.hd-row--button:focus-visible, .hd-file:focus-visible, .hd-back:focus-visible { outline: 2px solid ${tint('--ui-accent', 55)}; outline-offset: 1px; }
.hd-row--needs { background: var(--ui-widget-surface-background); }
.hd-row-main { display: flex; flex: 1; min-width: 0; flex-direction: column; }
.hd-row-title { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ui-text-primary); font-weight: 500; }
.hd-row-meta { display: flex; flex-wrap: wrap; gap: 0.375rem; color: var(--ui-text-tertiary); font-size: 0.75rem; line-height: 1rem; }
.hd-row-note { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 28rem; }
.hd-row-side { display: flex; flex-direction: column; align-items: flex-end; color: var(--ui-text-tertiary); font-size: 0.6875rem; white-space: nowrap; }

.hd-add { display: flex; flex-wrap: wrap; align-items: center; gap: 0.5rem; }
.hd-input, .hd-textarea {
  flex: 1; min-width: 14rem; padding: 0.375rem 0.5rem;
  border: 0; border-radius: var(--hd-radius-control);
  background: var(--ui-bg-quaternary); color: var(--ui-text-primary); font: inherit;
  box-shadow: inset 0 0 0 1px var(--ui-stroke-tertiary);
  transition: box-shadow var(--hd-fast) var(--hd-ease);
}
.hd-input:focus, .hd-textarea:focus { outline: none; box-shadow: inset 0 0 0 1px ${tint('--ui-accent', 60)}; }
.hd-input::placeholder, .hd-textarea::placeholder { color: var(--ui-text-quaternary); }
.hd-textarea { width: 100%; resize: vertical; min-height: 4.5rem; }

.hd-cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(16rem, 1fr)); gap: 0.75rem; }
.hd-card {
  display: flex; flex-direction: column; gap: 0.375rem; padding: 0.75rem;
  border-radius: var(--hd-radius-card); background: var(--ui-widget-surface-background);
  box-shadow: inset 0 0 0 1px var(--ui-stroke-tertiary);
}
.hd-card-head { display: flex; align-items: baseline; justify-content: space-between; gap: 0.5rem; }
.hd-card-name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.hd-card-path { color: var(--ui-text-tertiary); font-family: var(--dt-font-mono, monospace); font-size: 0.6875rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; direction: rtl; text-align: left; }
.hd-card-counts { display: flex; gap: 0.375rem; color: var(--ui-text-secondary); font-size: 0.75rem; }
.hd-card-actions { display: flex; gap: 0.5rem; padding-top: 0.25rem; }
.hd-count-needs { color: var(--ui-yellow); font-size: 0.6875rem; font-weight: 600; white-space: nowrap; }
.hd-composer { display: flex; flex-direction: column; gap: 0.375rem; padding-top: 0.25rem; }
.hd-composer-actions { display: flex; flex-wrap: wrap; align-items: center; justify-content: flex-end; gap: 0.375rem; }
.hd-composer-actions .hd-inline-error { flex: 1 1 100%; }

.hd-pill {
  display: inline-flex; flex-shrink: 0; align-items: center; gap: 0.3125rem;
  min-width: 5.25rem; padding: 0.0625rem 0.5rem;
  border-radius: var(--radius-pill, 9999px); font-size: 0.6875rem; line-height: 1rem; font-weight: 500;
  background: var(--ui-bg-quaternary); color: var(--ui-text-secondary);
}
.hd-pill-dot { width: 0.375rem; height: 0.375rem; border-radius: 9999px; background: var(--ui-text-quaternary); }
.hd-pill--running { color: var(--ui-accent); background: ${tint('--ui-accent', 10)}; }
.hd-pill--running .hd-pill-dot { background: var(--ui-accent); animation: hd-pulse 1.4s var(--hd-ease) infinite; }
.hd-pill--queued .hd-pill-dot { background: var(--ui-text-tertiary); }
.hd-pill--needs_you { color: var(--ui-yellow); background: ${tint('--ui-yellow', 12)}; }
.hd-pill--needs_you .hd-pill-dot { background: var(--ui-yellow); }
.hd-pill--done { color: var(--ui-green); background: ${tint('--ui-green', 10)}; }
.hd-pill--done .hd-pill-dot { background: var(--ui-green); }
.hd-pill--error { color: var(--ui-red); background: ${tint('--ui-red', 10)}; }
.hd-pill--error .hd-pill-dot { background: var(--ui-red); }
@keyframes hd-pulse { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.35; transform: scale(0.75); } }

.hd-kind { display: inline-block; min-width: 2.5rem; padding: 0 0.25rem; border-radius: var(--hd-radius-control); font-size: 0.625rem; font-weight: 600; text-align: center; letter-spacing: 0.02em; }
.hd-kind--docx { color: var(--ui-blue); background: ${tint('--ui-blue', 10)}; }
.hd-kind--xlsx { color: var(--ui-green); background: ${tint('--ui-green', 10)}; }
.hd-kind--pptx { color: var(--ui-orange); background: ${tint('--ui-orange', 10)}; }
.hd-kind--other { color: var(--ui-text-tertiary); background: var(--ui-bg-quaternary); }

.hd-skeleton { display: flex; flex-direction: column; gap: 0.5rem; padding: 0.25rem 0; }
.hd-skeleton-line { display: block; height: 0.75rem; border-radius: var(--hd-radius-control); background: var(--ui-bg-quaternary); animation: hd-shimmer 1.2s var(--hd-ease) infinite alternate; }
.hd-skeleton-line:nth-child(2n) { width: 70%; }
@keyframes hd-shimmer { from { opacity: 1; } to { opacity: 0.45; } }
.hd-error { display: flex; align-items: center; gap: 0.5rem; padding: 0.5rem 0.625rem; border-radius: var(--hd-radius-row); color: var(--ui-red); background: ${tint('--ui-red', 8)}; }
.hd-error > span { flex: 1; min-width: 0; overflow-wrap: anywhere; }
.hd-inline-error { flex-basis: 100%; color: var(--ui-red); font-size: 0.75rem; line-height: 1rem; overflow-wrap: anywhere; }
.hd-empty { padding: 0.75rem; border-radius: var(--hd-radius-row); color: var(--ui-text-tertiary); background: var(--ui-bg-quaternary); font-size: 0.75rem; }

.hd-back { align-self: flex-start; padding: 0.125rem 0.375rem; border: 0; border-radius: var(--hd-radius-control); background: transparent; color: var(--ui-text-secondary); font: inherit; cursor: pointer; transition: background-color var(--hd-fast) var(--hd-ease); }
.hd-back:hover { background: var(--ui-row-hover-background); color: var(--ui-text-primary); }

.hd-island {
  display: flex; flex-wrap: wrap; align-items: center; gap: 0.75rem;
  min-height: 3rem; padding: 0.625rem 0.875rem;
  border-radius: var(--hd-radius-panel); background: var(--ui-widget-surface-background);
  box-shadow: inset 0 0 0 1px var(--ui-stroke-tertiary);
  transition: box-shadow var(--hd-base) var(--hd-ease);
}
.hd-island--running { box-shadow: inset 0 0 0 1px ${tint('--ui-accent', 35)}; }
.hd-island--needs_you { box-shadow: inset 0 0 0 1px ${tint('--ui-yellow', 45)}; }
.hd-island--error { box-shadow: inset 0 0 0 1px ${tint('--ui-red', 45)}; }
.hd-island-text { display: flex; flex: 1; min-width: 0; flex-wrap: wrap; align-items: baseline; gap: 0.5rem; }
.hd-island-label { font-weight: 500; }
.hd-island-tool { font-family: var(--dt-font-mono, monospace); font-size: 0.6875rem; color: var(--ui-text-secondary); }
.hd-island-file { color: var(--ui-text-secondary); }
.hd-island-note { flex-basis: 100%; color: var(--ui-yellow); font-size: 0.75rem; }

.hd-ws-grid { display: grid; grid-template-columns: minmax(14rem, 1fr) minmax(12rem, 0.8fr) minmax(24rem, 2.4fr); gap: 0.75rem; align-items: start; }
@media (max-width: 980px) { .hd-ws-grid { grid-template-columns: 1fr; } }
.hd-col { display: flex; flex-direction: column; gap: 0.5rem; min-width: 0; padding: 0.75rem; border-radius: var(--hd-radius-panel); box-shadow: inset 0 0 0 1px var(--ui-stroke-tertiary); }
.hd-col-head { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 0.5rem; min-height: 1.75rem; }
.hd-col-head .hd-col-title { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.hd-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 0.375rem; }

.hd-chat { display: flex; flex-direction: column; gap: 0.5rem; max-height: 32rem; overflow: auto; }
.hd-bubble { padding: 0.375rem 0.5rem; border-radius: var(--hd-radius-row); background: var(--ui-bg-quaternary); }
.hd-bubble--user { background: var(--ui-chat-bubble-background, var(--ui-bg-tertiary)); box-shadow: inset 0 0 0 1px var(--ui-stroke-tertiary); }
.hd-bubble--tool { background: transparent; color: var(--ui-text-tertiary); padding-block: 0.125rem; }
.hd-bubble-role { font-size: 0.6875rem; color: var(--ui-text-tertiary); text-transform: capitalize; }
.hd-bubble-text { white-space: pre-wrap; overflow-wrap: anywhere; }

.hd-files { display: flex; flex-direction: column; gap: 0.125rem; }
.hd-file { display: flex; align-items: center; gap: 0.5rem; width: 100%; min-height: 2rem; padding: 0.25rem 0.375rem; border: 0; border-radius: var(--hd-radius-row); background: transparent; color: inherit; font: inherit; text-align: left; cursor: pointer; transition: background-color var(--hd-fast) var(--hd-ease); }
.hd-file:hover { background: var(--ui-row-hover-background); }
.hd-file--selected, .hd-file--selected:hover { background: var(--ui-row-active-background); }
.hd-file-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.hd-state { font-size: 0.6875rem; color: var(--ui-text-tertiary); }
.hd-state--working { color: var(--ui-accent); }
.hd-state--changed { color: var(--ui-yellow); }
.hd-state--created { color: var(--ui-green); }

.hd-diff, .hd-timeline { border-radius: var(--hd-radius-row); background: var(--ui-bg-quaternary); padding: 0.375rem 0.625rem; }
.hd-diff-summary { cursor: pointer; color: var(--ui-text-secondary); font-size: 0.75rem; }
.hd-diff-list { display: flex; flex-direction: column; gap: 0.125rem; padding-top: 0.375rem; font-size: 0.75rem; color: var(--ui-text-secondary); max-height: 12rem; overflow: auto; }
.hd-panes { display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; }
.hd-pane { display: flex; flex-direction: column; min-width: 0; border-radius: var(--hd-radius-card); box-shadow: inset 0 0 0 1px var(--ui-stroke-tertiary); overflow: hidden; }
.hd-pane-label { padding: 0.25rem 0.5rem; font-size: 0.6875rem; font-weight: 600; color: var(--ui-text-tertiary); box-shadow: inset 0 -1px 0 var(--ui-stroke-tertiary); }
.hd-pane-body { padding: 0.5rem; min-height: 12rem; max-height: 36rem; overflow: auto; animation: hd-fade var(--hd-base) var(--hd-ease); }
.hd-subtitle { font-size: 0.6875rem; font-weight: 600; color: var(--ui-text-tertiary); padding: 0.25rem 0; }

.hd-doc { display: flex; flex-direction: column; gap: 0.375rem; font-size: 0.8125rem; line-height: 1.375rem; }
.hd-doc-h { font-weight: 600; font-size: 0.9375rem; }
.hd-ins { color: var(--ui-accent); text-decoration: underline; text-decoration-color: var(--ui-accent); text-underline-offset: 2px; }
.hd-del { color: var(--ui-text-quaternary); text-decoration: line-through; }
.hd-comments { display: flex; flex-direction: column; gap: 0.375rem; padding-top: 0.5rem; box-shadow: inset 0 1px 0 var(--ui-stroke-tertiary); }
.hd-comment { padding: 0.25rem 0.5rem; border-radius: var(--hd-radius-row); background: ${tint('--ui-yellow', 8)}; font-size: 0.75rem; }
.hd-comment--resolved { opacity: 0.6; }
.hd-comment-meta { color: var(--ui-text-tertiary); font-size: 0.6875rem; }

.hd-grid-wrap { overflow: auto; }
.hd-grid { border-collapse: collapse; font-size: 0.6875rem; font-variant-numeric: tabular-nums; }
.hd-cell { min-width: 3.5rem; max-width: 10rem; padding: 0.125rem 0.375rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; box-shadow: inset 0 0 0 0.5px var(--ui-stroke-tertiary); }
.hd-cell--changed { background: ${tint('--ui-accent', 16)}; color: var(--ui-text-primary); font-weight: 600; }

.hd-deck { display: flex; flex-direction: column; gap: 0.5rem; }
.hd-slide { position: relative; width: 100%; overflow: hidden; border-radius: var(--hd-radius-control); background: var(--ui-bg-elevated); box-shadow: inset 0 0 0 1px var(--ui-stroke-secondary); container-type: inline-size; }
.hd-shape { position: absolute; overflow: hidden; padding: 0.125rem; border-radius: 2px; box-shadow: inset 0 0 0 1px ${tint('--ui-accent', 25)}; font-size: max(0.5rem, 2.2cqw); line-height: 1.2; color: var(--ui-text-secondary); white-space: pre-wrap; }

.hd-receipt { display: flex; flex-wrap: wrap; align-items: center; gap: 0.5rem; padding: 0.375rem 0.5rem; border-radius: var(--hd-radius-row); }
.hd-receipt:hover { background: var(--ui-row-hover-background); }
.hd-events { display: flex; flex-direction: column; padding-top: 0.375rem; }
.hd-event { display: grid; grid-template-columns: 6rem 7rem 10rem 1fr; gap: 0.5rem; padding: 0.125rem 0; font-size: 0.75rem; color: var(--ui-text-secondary); }
.hd-event-ts, .hd-event-kind { color: var(--ui-text-tertiary); }
.hd-event-tool { font-family: var(--dt-font-mono, monospace); font-size: 0.6875rem; overflow: hidden; text-overflow: ellipsis; }
.hd-event-summary { overflow-wrap: anywhere; }
.hd-event--refused, .hd-event--refused .hd-event-kind { color: var(--ui-red); }

.hd-enter { animation: hd-enter var(--hd-base) var(--hd-ease); }
@keyframes hd-enter { from { opacity: 0; transform: translateY(2px); } to { opacity: 1; transform: none; } }
@keyframes hd-fade { from { opacity: 0; } to { opacity: 1; } }

@media (prefers-reduced-motion: reduce) {
  .hd-page *, .hd-page *::before, .hd-page *::after { animation: none !important; transition: none !important; }
}
`

export function injectStyle() {
  document.getElementById(STYLE_ID)?.remove()
  const style = document.createElement('style')
  style.id = STYLE_ID
  style.textContent = DESK_CSS
  document.head.appendChild(style)

  return () => style.remove()
}
