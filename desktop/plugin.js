// hermes-office Desk desktop plugin. Built from desktop/src by desktop/build.mjs; do not edit.

// desktop/src/plugin.jsx
import { host as host4, queryClient, ROUTES_AREA, SIDEBAR_NAV_AREA, useValue } from "@hermes/plugin-sdk";

// desktop/src/api.js
var KEY = "office-desk";
var bound = { rest: null, os: null };
function bindContext(ctx) {
  bound.rest = ctx.rest;
  bound.os = ctx.os;
  return () => {
    bound.rest = null;
    bound.os = null;
  };
}
function osDoor() {
  return bound.os;
}
function rest(path, opts) {
  if (!bound.rest) {
    return Promise.reject(new Error("Desk is not loaded"));
  }
  return bound.rest(path, opts);
}
function query(params) {
  const search = new URLSearchParams();
  for (const [name, value] of Object.entries(params)) {
    if (value !== void 0 && value !== null && value !== "") {
      search.set(name, String(value));
    }
  }
  return search.toString();
}
function errorText(error) {
  const raw = error instanceof Error ? error.message : String(error ?? "");
  const body = raw.replace(/^\d{3}:\s*/, "");
  try {
    const parsed = JSON.parse(body);
    if (parsed && typeof parsed.detail === "string") {
      return parsed.detail;
    }
    if (parsed && parsed.detail !== void 0) {
      return JSON.stringify(parsed.detail);
    }
  } catch {
  }
  return body || "Request failed";
}
var api = {
  health: () => rest("/health"),
  desk: () => rest("/desk"),
  addFolder: (path) => rest("/folders", { method: "POST", body: { path } }),
  createWorkstream: (body) => rest("/workstreams", { method: "POST", body }),
  patchWorkstream: (id, body) => rest(`/workstreams/${encodeURIComponent(id)}`, { method: "PATCH", body }),
  workstream: (id) => rest(`/workstreams/${encodeURIComponent(id)}`),
  conversation: (id) => rest(`/workstreams/${encodeURIComponent(id)}/conversation`),
  preview: (path, receiptId) => rest(`/preview?${query({ path, receipt_id: receiptId })}`),
  diff: (path, receiptId) => rest(`/diff?${query({ path, receipt_id: receiptId })}`),
  redline: (body) => rest("/redline", { method: "POST", body }),
  revert: (receiptId) => rest(`/receipts/${encodeURIComponent(receiptId)}/revert`, { method: "POST", body: {} })
};
var keys = {
  all: [KEY],
  health: [KEY, "health"],
  desk: [KEY, "desk"],
  workstream: (id) => [KEY, "workstream", id],
  conversation: (id) => [KEY, "conversation", id],
  preview: (path, receiptId, version) => [KEY, "preview", path, receiptId ?? "current", version],
  diff: (path, receiptId, version) => [KEY, "diff", path, receiptId, version]
};
var ACTIVE_STATUSES = /* @__PURE__ */ new Set(["running", "queued"]);

// desktop/src/home.jsx
import { Button as Button2, host as host2, useQuery, useQueryClient } from "@hermes/plugin-sdk";
import { useState } from "react";

// desktop/src/session.js
import { host } from "@hermes/plugin-sdk";
function requestPrompt(folder, workstreamId, request) {
  return `${request}

Work only inside ${folder.path}. Use the office tools (toolset "office"). Workstream id: ${workstreamId}.`;
}
async function startWorkstream(folder, request) {
  const workstream = await api.createWorkstream({ folder_id: folder.id, title: request.slice(0, 60), request });
  if (typeof host.request !== "function") {
    return { workstream, started: false, error: "This Hermes build has no gateway request door (host.request)." };
  }
  let created;
  try {
    created = await host.request("session.create", { cwd: folder.path, cwd_explicit: true });
  } catch (error) {
    return { workstream, started: false, error: `Could not start a Hermes session: ${errorText(error)}` };
  }
  const storedId = created?.stored_session_id || created?.session_id;
  try {
    await host.request("prompt.submit", {
      session_id: created.session_id,
      text: requestPrompt(folder, workstream.id, request)
    });
  } catch (error) {
    if (storedId) {
      await api.patchWorkstream(workstream.id, { session_id: storedId }).catch(() => void 0);
    }
    return { workstream, started: false, error: `The session started but the request was not sent: ${errorText(error)}` };
  }
  try {
    const updated = await api.patchWorkstream(workstream.id, { session_id: storedId, status: "running" });
    return { workstream: updated, started: true };
  } catch (error) {
    return { workstream, started: true, error: `Started, but the Desk could not record the session: ${errorText(error)}` };
  }
}

// desktop/src/state.js
import { atom } from "@hermes/plugin-sdk";
var $selectedWorkstream = atom(null);

// desktop/src/ui.jsx
import { Button } from "@hermes/plugin-sdk";
import { jsx, jsxs } from "react/jsx-runtime";
var STATUS_LABELS = {
  running: "Running",
  queued: "Queued",
  needs_you: "Needs you",
  done: "Done",
  error: "Error"
};
function StatusPill({ status }) {
  const known = STATUS_LABELS[status] ? status : "unknown";
  return /* @__PURE__ */ jsxs("span", { className: `hd-pill hd-pill--${known}`, children: [
    /* @__PURE__ */ jsx("span", { "aria-hidden": true, className: "hd-pill-dot" }),
    STATUS_LABELS[status] ?? String(status ?? "Unknown")
  ] });
}
function toMillis(ts) {
  if (ts === null || ts === void 0 || ts === "") {
    return null;
  }
  if (typeof ts === "number") {
    return ts < 1e12 ? ts * 1e3 : ts;
  }
  const ms = Date.parse(ts);
  return Number.isNaN(ms) ? null : ms;
}
function timeAgo(ts) {
  const ms = toMillis(ts);
  if (ms === null) {
    return "";
  }
  const seconds = Math.max(0, Math.round((Date.now() - ms) / 1e3));
  if (seconds < 45) {
    return "just now";
  }
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) {
    return `${minutes} min ago`;
  }
  const hours = Math.round(minutes / 60);
  if (hours < 24) {
    return `${hours} h ago`;
  }
  const days = Math.round(hours / 24);
  return days < 30 ? `${days} d ago` : new Date(ms).toLocaleDateString();
}
function clockTime(ts) {
  const ms = toMillis(ts);
  return ms === null ? "" : new Date(ms).toLocaleString();
}
var KIND_LABELS = { docx: "DOCX", pptx: "PPTX", xlsx: "XLSX" };
function KindMark({ kind }) {
  return /* @__PURE__ */ jsx("span", { className: `hd-kind hd-kind--${KIND_LABELS[kind] ? kind : "other"}`, children: KIND_LABELS[kind] ?? "FILE" });
}
function baseName(path) {
  return String(path ?? "").split(/[\\/]/).filter(Boolean).pop() ?? "";
}
function Skeleton({ lines = 3 }) {
  return /* @__PURE__ */ jsx("div", { "aria-busy": "true", className: "hd-skeleton", children: Array.from({ length: lines }, (_, index) => /* @__PURE__ */ jsx("span", { className: "hd-skeleton-line" }, index)) });
}
function QueryError({ error, onRetry }) {
  return /* @__PURE__ */ jsxs("div", { className: "hd-error", role: "alert", children: [
    /* @__PURE__ */ jsx("span", { children: errorText(error) }),
    onRetry ? /* @__PURE__ */ jsx(Button, { onClick: () => onRetry(), size: "xs", variant: "outline", children: "Retry" }) : null
  ] });
}
function Empty({ children }) {
  return /* @__PURE__ */ jsx("div", { className: "hd-empty", children });
}
function InlineError({ message }) {
  return message ? /* @__PURE__ */ jsx("div", { className: "hd-inline-error", role: "alert", children: message }) : null;
}

// desktop/src/home.jsx
import { Fragment, jsx as jsx2, jsxs as jsxs2 } from "react/jsx-runtime";
function HealthLine() {
  const { data, error, isLoading } = useQuery({ queryKey: keys.health, queryFn: api.health, refetchInterval: 6e4 });
  if (isLoading) {
    return /* @__PURE__ */ jsx2("div", { className: "hd-health hd-health--muted", children: "Checking the office engines…" });
  }
  if (error) {
    return /* @__PURE__ */ jsxs2("div", { className: "hd-health hd-health--bad", children: [
      "Backend unreachable: ",
      errorText(error)
    ] });
  }
  const engines = data?.engines ?? {};
  const names = ["docx", "pptx", "xlsx"];
  const missing = names.filter((name) => !engines[name]);
  return /* @__PURE__ */ jsxs2("div", { className: "hd-health", children: [
    /* @__PURE__ */ jsx2("span", { className: missing.length ? "hd-health-bad" : "hd-health-ok", children: missing.length ? `Engines missing: ${missing.join(", ")}` : "Engines ok (docx, pptx, xlsx)" }),
    /* @__PURE__ */ jsx2("span", { className: "hd-sep", children: "·" }),
    /* @__PURE__ */ jsxs2("span", { children: [
      "LibreOffice ",
      data?.libreoffice ? "yes" : "no"
    ] }),
    data?.version ? /* @__PURE__ */ jsxs2(Fragment, { children: [
      /* @__PURE__ */ jsx2("span", { className: "hd-sep", children: "·" }),
      /* @__PURE__ */ jsxs2("span", { children: [
        "v",
        data.version
      ] })
    ] }) : null
  ] });
}
function openWorkstream(id) {
  $selectedWorkstream.set(id);
}
function NeedsYou({ items, folderName }) {
  if (!items.length) {
    return null;
  }
  return /* @__PURE__ */ jsxs2("section", { className: "hd-section hd-needs", children: [
    /* @__PURE__ */ jsx2("h2", { className: "hd-section-title hd-needs-title", children: "Needs you" }),
    /* @__PURE__ */ jsx2("ul", { className: "hd-list", children: items.map((ws) => /* @__PURE__ */ jsxs2("li", { className: "hd-row hd-row--needs hd-enter", children: [
      /* @__PURE__ */ jsxs2("div", { className: "hd-row-main", children: [
        /* @__PURE__ */ jsx2("div", { className: "hd-row-title", children: ws.title }),
        /* @__PURE__ */ jsxs2("div", { className: "hd-row-meta", children: [
          /* @__PURE__ */ jsx2("span", { children: folderName(ws.folder_id) }),
          ws.note ? /* @__PURE__ */ jsxs2(Fragment, { children: [
            /* @__PURE__ */ jsx2("span", { className: "hd-sep", children: "·" }),
            /* @__PURE__ */ jsx2("span", { className: "hd-row-note", children: ws.note })
          ] }) : null,
          /* @__PURE__ */ jsx2("span", { className: "hd-sep", children: "·" }),
          /* @__PURE__ */ jsx2("span", { children: timeAgo(ws.updated_at) })
        ] })
      ] }),
      /* @__PURE__ */ jsx2(Button2, { onClick: () => openWorkstream(ws.id), size: "sm", variant: "secondary", children: "Review" })
    ] }, ws.id)) })
  ] });
}
function AddFolder() {
  const client = useQueryClient();
  const [path, setPath] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const canPick = typeof host2.pickOpenPaths === "function";
  const add = async (candidate) => {
    const trimmed = String(candidate ?? "").trim();
    if (!trimmed) {
      setError("Enter the absolute path of a folder.");
      return;
    }
    setBusy(true);
    setError("");
    try {
      await api.addFolder(trimmed);
      setPath("");
      await client.invalidateQueries({ queryKey: keys.desk });
    } catch (err) {
      setError(errorText(err));
    } finally {
      setBusy(false);
    }
  };
  const pick = async () => {
    setError("");
    try {
      const picked = await host2.pickOpenPaths({ directories: true, title: "Add a folder to your Desk" });
      if (picked?.[0]) {
        await add(picked[0]);
      }
    } catch (err) {
      setError(errorText(err));
    }
  };
  return /* @__PURE__ */ jsxs2(
    "form",
    {
      className: "hd-add",
      onSubmit: (event) => {
        event.preventDefault();
        void add(path);
      },
      children: [
        /* @__PURE__ */ jsx2(
          "input",
          {
            "aria-label": "Folder path",
            className: "hd-input",
            disabled: busy,
            onChange: (event) => setPath(event.target.value),
            placeholder: "/path/to/a/folder",
            spellCheck: false,
            value: path
          }
        ),
        /* @__PURE__ */ jsx2(Button2, { disabled: busy, size: "sm", type: "submit", variant: "secondary", children: "Add folder" }),
        canPick ? /* @__PURE__ */ jsx2(Button2, { disabled: busy, onClick: () => void pick(), size: "sm", type: "button", variant: "ghost", children: "Browse…" }) : null,
        /* @__PURE__ */ jsx2(InlineError, { message: error })
      ]
    }
  );
}
function Composer({ folder, onClose }) {
  const client = useQueryClient();
  const [text, setText] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const submit = async () => {
    const request = text.trim();
    if (!request) {
      setError("Describe what you want done in this folder.");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const result = await startWorkstream(folder, request);
      await client.invalidateQueries({ queryKey: keys.desk });
      if (result.error) {
        setError(`${result.error} The request is saved on the Desk as queued.`);
        setText("");
      } else {
        onClose();
      }
    } catch (err) {
      setError(errorText(err));
    } finally {
      setBusy(false);
    }
  };
  return /* @__PURE__ */ jsxs2("div", { className: "hd-composer hd-enter", children: [
    /* @__PURE__ */ jsx2(
      "textarea",
      {
        "aria-label": `New request for ${folder.name}`,
        className: "hd-textarea",
        disabled: busy,
        onChange: (event) => setText(event.target.value),
        onKeyDown: (event) => {
          if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
            event.preventDefault();
            void submit();
          }
        },
        placeholder: "e.g. Tighten the executive summary in the board memo and track the changes.",
        rows: 3,
        value: text
      }
    ),
    /* @__PURE__ */ jsxs2("div", { className: "hd-composer-actions", children: [
      /* @__PURE__ */ jsx2(InlineError, { message: error }),
      /* @__PURE__ */ jsx2(Button2, { disabled: busy, onClick: onClose, size: "sm", variant: "ghost", children: "Cancel" }),
      /* @__PURE__ */ jsx2(Button2, { disabled: busy, onClick: () => void submit(), size: "sm", children: busy ? "Starting…" : "Start" })
    ] })
  ] });
}
function FolderCard({ folder }) {
  const [composing, setComposing] = useState(false);
  const counts = folder.counts ?? {};
  return /* @__PURE__ */ jsxs2("li", { className: "hd-card hd-enter", children: [
    /* @__PURE__ */ jsxs2("div", { className: "hd-card-head", children: [
      /* @__PURE__ */ jsx2("div", { className: "hd-card-name", children: folder.name }),
      folder.counts?.needs_you ? /* @__PURE__ */ jsxs2("span", { className: "hd-count-needs", children: [
        counts.needs_you,
        " need you"
      ] }) : null
    ] }),
    /* @__PURE__ */ jsx2("div", { className: "hd-card-path", title: folder.path, children: folder.path }),
    /* @__PURE__ */ jsxs2("div", { className: "hd-card-counts", children: [
      /* @__PURE__ */ jsxs2("span", { children: [
        counts.files ?? 0,
        " files"
      ] }),
      /* @__PURE__ */ jsx2("span", { className: "hd-sep", children: "·" }),
      /* @__PURE__ */ jsxs2("span", { children: [
        counts.workstreams ?? 0,
        " requests"
      ] })
    ] }),
    composing ? /* @__PURE__ */ jsx2(Composer, { folder, onClose: () => setComposing(false) }) : /* @__PURE__ */ jsx2("div", { className: "hd-card-actions", children: /* @__PURE__ */ jsx2(Button2, { onClick: () => setComposing(true), size: "sm", variant: "secondary", children: "New request" }) })
  ] });
}
function Recent({ items, folderName }) {
  return /* @__PURE__ */ jsxs2("section", { className: "hd-section", children: [
    /* @__PURE__ */ jsx2("h2", { className: "hd-section-title", children: "Recent" }),
    items.length ? /* @__PURE__ */ jsx2("ul", { className: "hd-list", children: items.map((ws) => /* @__PURE__ */ jsx2("li", { children: /* @__PURE__ */ jsxs2("button", { className: "hd-row hd-row--button", onClick: () => openWorkstream(ws.id), type: "button", children: [
      /* @__PURE__ */ jsx2(StatusPill, { status: ws.status }),
      /* @__PURE__ */ jsxs2("div", { className: "hd-row-main", children: [
        /* @__PURE__ */ jsx2("div", { className: "hd-row-title", children: ws.title }),
        /* @__PURE__ */ jsxs2("div", { className: "hd-row-meta", children: [
          /* @__PURE__ */ jsx2("span", { children: folderName(ws.folder_id) }),
          ws.last_event?.summary ? /* @__PURE__ */ jsxs2(Fragment, { children: [
            /* @__PURE__ */ jsx2("span", { className: "hd-sep", children: "·" }),
            /* @__PURE__ */ jsx2("span", { className: "hd-row-note", children: ws.last_event.summary })
          ] }) : null
        ] })
      ] }),
      /* @__PURE__ */ jsxs2("div", { className: "hd-row-side", children: [
        ws.files_changed ? /* @__PURE__ */ jsxs2("span", { children: [
          ws.files_changed,
          " changed"
        ] }) : null,
        /* @__PURE__ */ jsx2("span", { children: timeAgo(ws.updated_at) })
      ] })
    ] }) }, ws.id)) }) : /* @__PURE__ */ jsx2(Empty, { children: "No requests yet. Start one from a folder above." })
  ] });
}
function DeskHome() {
  const { data, error, isLoading, refetch } = useQuery({ queryKey: keys.desk, queryFn: api.desk, refetchInterval: 15e3 });
  const folders = data?.folders ?? [];
  const folderName = (id) => folders.find((folder) => folder.id === id)?.name ?? "Unknown folder";
  return /* @__PURE__ */ jsxs2("div", { className: "hd-page", children: [
    /* @__PURE__ */ jsxs2("header", { className: "hd-header", children: [
      /* @__PURE__ */ jsx2("h1", { className: "hd-title", children: "Desk" }),
      /* @__PURE__ */ jsx2(HealthLine, {})
    ] }),
    isLoading ? /* @__PURE__ */ jsx2(Skeleton, { lines: 4 }) : null,
    error ? /* @__PURE__ */ jsx2(QueryError, { error, onRetry: refetch }) : null,
    data ? /* @__PURE__ */ jsxs2(Fragment, { children: [
      /* @__PURE__ */ jsx2(NeedsYou, { folderName, items: data.needs_you ?? [] }),
      /* @__PURE__ */ jsxs2("section", { className: "hd-section", children: [
        /* @__PURE__ */ jsx2("h2", { className: "hd-section-title", children: "On your desk" }),
        /* @__PURE__ */ jsx2(AddFolder, {}),
        folders.length ? /* @__PURE__ */ jsx2("ul", { className: "hd-cards", children: folders.map((folder) => /* @__PURE__ */ jsx2(FolderCard, { folder }, folder.id)) }) : /* @__PURE__ */ jsx2(Empty, { children: "No folders yet. Add a folder of Word, Excel or PowerPoint files to start." })
      ] }),
      /* @__PURE__ */ jsx2(Recent, { folderName, items: data.recent ?? [] })
    ] }) : null
  ] });
}

// desktop/src/styles.js
var tint = (token, amount) => `color-mix(in srgb, var(${token}) ${amount}%, transparent)`;
var STYLE_ID = "hermes-office-desk-style";
var DESK_CSS = `
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
.hd-needs { padding: 0.75rem; border-radius: var(--hd-radius-panel); background: ${tint("--ui-yellow", 8)}; box-shadow: inset 0 0 0 1px ${tint("--ui-yellow", 30)}; }
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
.hd-row--button:focus-visible, .hd-file:focus-visible, .hd-back:focus-visible { outline: 2px solid ${tint("--ui-accent", 55)}; outline-offset: 1px; }
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
.hd-input:focus, .hd-textarea:focus { outline: none; box-shadow: inset 0 0 0 1px ${tint("--ui-accent", 60)}; }
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
.hd-pill--running { color: var(--ui-accent); background: ${tint("--ui-accent", 10)}; }
.hd-pill--running .hd-pill-dot { background: var(--ui-accent); animation: hd-pulse 1.4s var(--hd-ease) infinite; }
.hd-pill--queued .hd-pill-dot { background: var(--ui-text-tertiary); }
.hd-pill--needs_you { color: var(--ui-yellow); background: ${tint("--ui-yellow", 12)}; }
.hd-pill--needs_you .hd-pill-dot { background: var(--ui-yellow); }
.hd-pill--done { color: var(--ui-green); background: ${tint("--ui-green", 10)}; }
.hd-pill--done .hd-pill-dot { background: var(--ui-green); }
.hd-pill--error { color: var(--ui-red); background: ${tint("--ui-red", 10)}; }
.hd-pill--error .hd-pill-dot { background: var(--ui-red); }
@keyframes hd-pulse { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.35; transform: scale(0.75); } }

.hd-kind { display: inline-block; min-width: 2.5rem; padding: 0 0.25rem; border-radius: var(--hd-radius-control); font-size: 0.625rem; font-weight: 600; text-align: center; letter-spacing: 0.02em; }
.hd-kind--docx { color: var(--ui-blue); background: ${tint("--ui-blue", 10)}; }
.hd-kind--xlsx { color: var(--ui-green); background: ${tint("--ui-green", 10)}; }
.hd-kind--pptx { color: var(--ui-orange); background: ${tint("--ui-orange", 10)}; }
.hd-kind--other { color: var(--ui-text-tertiary); background: var(--ui-bg-quaternary); }

.hd-skeleton { display: flex; flex-direction: column; gap: 0.5rem; padding: 0.25rem 0; }
.hd-skeleton-line { display: block; height: 0.75rem; border-radius: var(--hd-radius-control); background: var(--ui-bg-quaternary); animation: hd-shimmer 1.2s var(--hd-ease) infinite alternate; }
.hd-skeleton-line:nth-child(2n) { width: 70%; }
@keyframes hd-shimmer { from { opacity: 1; } to { opacity: 0.45; } }
.hd-error { display: flex; align-items: center; gap: 0.5rem; padding: 0.5rem 0.625rem; border-radius: var(--hd-radius-row); color: var(--ui-red); background: ${tint("--ui-red", 8)}; }
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
.hd-island--running { box-shadow: inset 0 0 0 1px ${tint("--ui-accent", 35)}; }
.hd-island--needs_you { box-shadow: inset 0 0 0 1px ${tint("--ui-yellow", 45)}; }
.hd-island--error { box-shadow: inset 0 0 0 1px ${tint("--ui-red", 45)}; }
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
.hd-comment { padding: 0.25rem 0.5rem; border-radius: var(--hd-radius-row); background: ${tint("--ui-yellow", 8)}; font-size: 0.75rem; }
.hd-comment--resolved { opacity: 0.6; }
.hd-comment-meta { color: var(--ui-text-tertiary); font-size: 0.6875rem; }

.hd-grid-wrap { overflow: auto; }
.hd-grid { border-collapse: collapse; font-size: 0.6875rem; font-variant-numeric: tabular-nums; }
.hd-cell { min-width: 3.5rem; max-width: 10rem; padding: 0.125rem 0.375rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; box-shadow: inset 0 0 0 0.5px var(--ui-stroke-tertiary); }
.hd-cell--changed { background: ${tint("--ui-accent", 16)}; color: var(--ui-text-primary); font-weight: 600; }

.hd-deck { display: flex; flex-direction: column; gap: 0.5rem; }
.hd-slide { position: relative; width: 100%; overflow: hidden; border-radius: var(--hd-radius-control); background: var(--ui-bg-elevated); box-shadow: inset 0 0 0 1px var(--ui-stroke-secondary); container-type: inline-size; }
.hd-shape { position: absolute; overflow: hidden; padding: 0.125rem; border-radius: 2px; box-shadow: inset 0 0 0 1px ${tint("--ui-accent", 25)}; font-size: max(0.5rem, 2.2cqw); line-height: 1.2; color: var(--ui-text-secondary); white-space: pre-wrap; }

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
`;
function injectStyle() {
  document.getElementById(STYLE_ID)?.remove();
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = DESK_CSS;
  document.head.appendChild(style);
  return () => style.remove();
}

// desktop/src/workstream.jsx
import { Button as Button3, host as host3, useQuery as useQuery3, useQueryClient as useQueryClient2 } from "@hermes/plugin-sdk";
import { useEffect, useState as useState2 } from "react";

// desktop/src/preview.jsx
import { useQuery as useQuery2 } from "@hermes/plugin-sdk";
import { jsx as jsx3, jsxs as jsxs3 } from "react/jsx-runtime";
function DocView({ docx }) {
  const paragraphs = docx?.paragraphs ?? [];
  const comments = docx?.comments ?? [];
  if (!paragraphs.length) {
    return /* @__PURE__ */ jsx3(Empty, { children: "This document has no text paragraphs." });
  }
  return /* @__PURE__ */ jsxs3("div", { className: "hd-doc", children: [
    paragraphs.map((paragraph, index) => {
      const heading = /^(heading|title)/i.test(paragraph.style ?? "");
      return /* @__PURE__ */ jsx3("p", { className: heading ? "hd-doc-p hd-doc-h" : "hd-doc-p", title: paragraph.style || void 0, children: (paragraph.runs ?? []).map((run, runIndex) => /* @__PURE__ */ jsx3(
        "span",
        {
          className: run.rev === "ins" ? "hd-ins" : run.rev === "del" ? "hd-del" : void 0,
          title: run.rev ? `${run.rev === "ins" ? "Inserted" : "Deleted"}${run.author ? ` by ${run.author}` : ""}` : void 0,
          children: run.text
        },
        runIndex
      )) }, index);
    }),
    comments.length ? /* @__PURE__ */ jsxs3("div", { className: "hd-comments", children: [
      /* @__PURE__ */ jsx3("div", { className: "hd-subtitle", children: "Comments" }),
      comments.map((comment) => /* @__PURE__ */ jsxs3("div", { className: comment.resolved ? "hd-comment hd-comment--resolved" : "hd-comment", children: [
        /* @__PURE__ */ jsxs3("div", { className: "hd-comment-meta", children: [
          comment.author || "Unknown",
          comment.anchor ? /* @__PURE__ */ jsxs3("span", { className: "hd-comment-anchor", children: [
            " on “",
            comment.anchor,
            "”"
          ] }) : null,
          comment.resolved ? /* @__PURE__ */ jsx3("span", { children: " · resolved" }) : null
        ] }),
        /* @__PURE__ */ jsx3("div", { children: comment.text })
      ] }, comment.id))
    ] }) : null
  ] });
}
function normalizeAddress(address) {
  return String(address ?? "").replace(/\$/g, "").replace(/^'(.*)'!/, "$1!").toUpperCase();
}
function isChanged(changed, sheet, ref) {
  const cell = normalizeAddress(ref);
  return changed.has(`${String(sheet).toUpperCase()}!${cell}`) || changed.has(cell);
}
function SheetView({ xlsx, changed }) {
  const sheets = xlsx?.sheets ?? [];
  if (!sheets.length) {
    return /* @__PURE__ */ jsx3(Empty, { children: "This workbook has no sheets." });
  }
  return /* @__PURE__ */ jsx3("div", { className: "hd-sheets", children: sheets.map((sheet) => /* @__PURE__ */ jsxs3("div", { className: "hd-sheet", children: [
    /* @__PURE__ */ jsxs3("div", { className: "hd-subtitle", children: [
      sheet.name,
      sheet.dims ? /* @__PURE__ */ jsxs3("span", { className: "hd-muted", children: [
        " · ",
        sheet.dims
      ] }) : null
    ] }),
    /* @__PURE__ */ jsx3("div", { className: "hd-grid-wrap", children: /* @__PURE__ */ jsx3("table", { className: "hd-grid", children: /* @__PURE__ */ jsx3("tbody", { children: (sheet.rows ?? []).map((row, rowIndex) => /* @__PURE__ */ jsx3("tr", { children: row.map((cell, cellIndex) => /* @__PURE__ */ jsx3(
      "td",
      {
        className: isChanged(changed, sheet.name, cell.ref) ? "hd-cell hd-cell--changed" : "hd-cell",
        title: [cell.ref, cell.formula].filter(Boolean).join("  ") || void 0,
        children: cell.value === null || cell.value === void 0 ? "" : String(cell.value)
      },
      cell.ref ?? cellIndex
    )) }, rowIndex)) }) }) })
  ] }, sheet.name)) });
}
function pct(value, total) {
  return total ? `${Number(value || 0) / total * 100}%` : "0%";
}
function DeckView({ pptx }) {
  const slides = pptx?.slides ?? [];
  const width = Number(pptx?.slide_w) || 0;
  const height = Number(pptx?.slide_h) || 0;
  if (!slides.length) {
    return /* @__PURE__ */ jsx3(Empty, { children: "This deck has no slides." });
  }
  return /* @__PURE__ */ jsx3("div", { className: "hd-deck", children: slides.map((slide) => /* @__PURE__ */ jsxs3("div", { className: "hd-slide-wrap", children: [
    /* @__PURE__ */ jsxs3("div", { className: "hd-subtitle", children: [
      "Slide ",
      Number(slide.index) + 1,
      slide.title ? /* @__PURE__ */ jsxs3("span", { className: "hd-muted", children: [
        " · ",
        slide.title
      ] }) : null
    ] }),
    /* @__PURE__ */ jsx3("div", { className: "hd-slide", style: { aspectRatio: width && height ? `${width} / ${height}` : "16 / 9" }, children: (slide.shapes ?? []).map((shape, index) => /* @__PURE__ */ jsx3(
      "div",
      {
        className: "hd-shape",
        style: { left: pct(shape.x, width), top: pct(shape.y, height), width: pct(shape.w, width), height: pct(shape.h, height) },
        title: `${shape.name} (${shape.kind})`,
        children: shape.text
      },
      `${shape.name}-${index}`
    )) })
  ] }, slide.index)) });
}
function PreviewBody({ preview, changed }) {
  if (preview.kind === "docx") {
    return /* @__PURE__ */ jsx3(DocView, { docx: preview.docx });
  }
  if (preview.kind === "xlsx") {
    return /* @__PURE__ */ jsx3(SheetView, { changed, xlsx: preview.xlsx });
  }
  if (preview.kind === "pptx") {
    return /* @__PURE__ */ jsx3(DeckView, { pptx: preview.pptx });
  }
  return /* @__PURE__ */ jsx3(Empty, { children: "No preview for this file type." });
}
function PreviewPane({ label, path, receiptId, version, changed, enabled = true, emptyText }) {
  const { data, error, isLoading, refetch } = useQuery2({
    queryKey: keys.preview(path, receiptId, version),
    queryFn: () => api.preview(path, receiptId),
    enabled
  });
  return /* @__PURE__ */ jsxs3("div", { className: "hd-pane", children: [
    /* @__PURE__ */ jsx3("div", { className: "hd-pane-label", children: label }),
    /* @__PURE__ */ jsxs3("div", { className: "hd-pane-body", children: [
      !enabled ? /* @__PURE__ */ jsx3(Empty, { children: emptyText }) : null,
      enabled && isLoading ? /* @__PURE__ */ jsx3(Skeleton, { lines: 6 }) : null,
      enabled && error ? /* @__PURE__ */ jsx3(QueryError, { error, onRetry: refetch }) : null,
      enabled && data ? /* @__PURE__ */ jsx3(PreviewBody, { changed, preview: data }) : null
    ] })
  ] });
}
function texts(list) {
  return (list ?? []).map((part) => part?.text ?? "").join(" ");
}
function DiffList({ items }) {
  return items.length ? /* @__PURE__ */ jsx3("ul", { className: "hd-diff-list", children: items.map((item, index) => /* @__PURE__ */ jsx3("li", { children: item }, index)) }) : null;
}
function docxDiff(diff) {
  const revisions = diff.revisions ?? [];
  const changed = diff.paragraphs_changed ?? [];
  return {
    counts: [`${revisions.length} tracked change${revisions.length === 1 ? "" : "s"}`, `${changed.length} paragraph${changed.length === 1 ? "" : "s"} changed`],
    items: [
      ...revisions.map((rev) => `${rev.revision_type}${rev.author ? ` (${rev.author})` : ""}: ${rev.text}`),
      ...changed.map((change) => `“${change.before}” → “${change.after}”`)
    ]
  };
}
function xlsxDiff(diff) {
  const changed = diff.changed ?? [];
  const shifted = diff.shifted ?? [];
  const added = diff.added_sheets ?? [];
  const removed = diff.removed_sheets ?? [];
  return {
    counts: [
      `${changed.length} cell${changed.length === 1 ? "" : "s"} changed`,
      `${shifted.length} shifted`,
      ...added.length ? [`sheets added: ${added.join(", ")}`] : [],
      ...removed.length ? [`sheets removed: ${removed.join(", ")}`] : []
    ],
    items: [
      ...changed.map((cell) => `${cell.address}: ${String(cell.before ?? "∅")} → ${String(cell.after ?? "∅")}`),
      ...shifted.map((cell) => `${cell.from} → ${cell.to}${cell.value !== void 0 ? ` (${String(cell.value)})` : ""}`)
    ]
  };
}
function countOf(value) {
  return Array.isArray(value) ? value.length : Number(value) || 0;
}
function pptxDiff(diff) {
  const slides = diff.slide_changes ?? [];
  const changes = slides.flatMap((slide) => (slide.text_changes ?? []).map((change) => ({ slide, change })));
  return {
    counts: [
      `${countOf(diff.slides_added)} slides added`,
      `${countOf(diff.slides_removed)} removed`,
      `${countOf(diff.slides_moved)} moved`,
      `${changes.length} text change${changes.length === 1 ? "" : "s"}`
    ],
    items: changes.map(
      ({ slide, change }) => `Slide ${slide.slide_id} · ${change.shape?.name ?? "shape"} (${change.kind}): “${texts(change.before)}” → “${texts(change.after)}”`
    )
  };
}
var DIFF_SUMMARIES = { docx: docxDiff, xlsx: xlsxDiff, pptx: pptxDiff };
function DiffSummary({ diff }) {
  const summarize = DIFF_SUMMARIES[diff?.kind];
  if (!summarize) {
    return null;
  }
  const { counts, items } = summarize(diff);
  return /* @__PURE__ */ jsxs3("details", { className: "hd-diff", children: [
    /* @__PURE__ */ jsx3("summary", { className: "hd-diff-summary", children: counts.join(" · ") }),
    /* @__PURE__ */ jsx3(DiffList, { items })
  ] });
}
function changedCells(diff) {
  return new Set(diff?.kind === "xlsx" ? (diff.changed ?? []).map((cell) => normalizeAddress(cell.address)) : []);
}

// desktop/src/workstream.jsx
import { Fragment as Fragment2, jsx as jsx4, jsxs as jsxs4 } from "react/jsx-runtime";
var OPEN_LABELS = { docx: "Open in Word", xlsx: "Open in Excel", pptx: "Open in PowerPoint" };
var STATE_LABELS = { working: "Working", changed: "Changed", created: "Created" };
function ActivityIsland({ workstream, now, events }) {
  const last = workstream.last_event ?? events[events.length - 1] ?? null;
  return /* @__PURE__ */ jsxs4("div", { className: `hd-island hd-island--${workstream.status}`, children: [
    /* @__PURE__ */ jsx4(StatusPill, { status: workstream.status }),
    /* @__PURE__ */ jsx4("div", { className: "hd-island-text", children: now ? /* @__PURE__ */ jsxs4(Fragment2, { children: [
      /* @__PURE__ */ jsx4("span", { className: "hd-island-label", children: now.label }),
      now.tool ? /* @__PURE__ */ jsx4("span", { className: "hd-island-tool", children: now.tool }) : null,
      now.path ? /* @__PURE__ */ jsx4("span", { className: "hd-island-file", children: baseName(now.path) }) : null,
      now.since ? /* @__PURE__ */ jsxs4("span", { className: "hd-muted", children: [
        "since ",
        timeAgo(now.since)
      ] }) : null
    ] }) : last ? /* @__PURE__ */ jsxs4(Fragment2, { children: [
      /* @__PURE__ */ jsx4("span", { className: "hd-island-label", children: last.summary || last.kind }),
      /* @__PURE__ */ jsx4("span", { className: "hd-muted", children: timeAgo(last.ts) })
    ] }) : /* @__PURE__ */ jsx4("span", { className: "hd-muted", children: "No activity recorded yet." }) }),
    workstream.note ? /* @__PURE__ */ jsx4("div", { className: "hd-island-note", children: workstream.note }) : null
  ] });
}
function Conversation({ workstream, active }) {
  const { data, error, isLoading, refetch } = useQuery3({
    queryKey: keys.conversation(workstream.id),
    queryFn: () => api.conversation(workstream.id),
    refetchInterval: active ? 4e3 : false
  });
  const [openError, setOpenError] = useState2("");
  const continueInChat = async () => {
    setOpenError("");
    try {
      await host3.openSession(workstream.session_id);
    } catch (err) {
      setOpenError(errorText(err));
    }
  };
  return /* @__PURE__ */ jsxs4("section", { className: "hd-col hd-col--chat", children: [
    /* @__PURE__ */ jsxs4("div", { className: "hd-col-head", children: [
      /* @__PURE__ */ jsx4("h3", { className: "hd-col-title", children: "Conversation" }),
      workstream.session_id && typeof host3.openSession === "function" ? /* @__PURE__ */ jsx4(Button3, { onClick: () => void continueInChat(), size: "xs", variant: "outline", children: "Continue in chat" }) : null
    ] }),
    /* @__PURE__ */ jsx4(InlineError, { message: openError }),
    /* @__PURE__ */ jsxs4("div", { className: "hd-chat", children: [
      isLoading ? /* @__PURE__ */ jsx4(Skeleton, { lines: 5 }) : null,
      error ? /* @__PURE__ */ jsx4(QueryError, { error, onRetry: refetch }) : null,
      data && !data.available ? /* @__PURE__ */ jsx4(Empty, { children: data.reason === "missing_session_id" ? "No Hermes session is attached to this request yet." : data.reason || "The conversation is not available." }) : null,
      data?.available ? (data.messages ?? []).map((message, index) => /* @__PURE__ */ jsxs4("div", { className: `hd-bubble hd-bubble--${message.role}`, children: [
        /* @__PURE__ */ jsxs4("div", { className: "hd-bubble-role", children: [
          message.role,
          message.ts ? /* @__PURE__ */ jsxs4("span", { className: "hd-muted", children: [
            " · ",
            timeAgo(message.ts)
          ] }) : null
        ] }),
        /* @__PURE__ */ jsx4("div", { className: "hd-bubble-text", children: message.text })
      ] }, index)) : null
    ] })
  ] });
}
function FileList({ files, selected, onSelect }) {
  return /* @__PURE__ */ jsxs4("section", { className: "hd-col hd-col--files", children: [
    /* @__PURE__ */ jsx4("div", { className: "hd-col-head", children: /* @__PURE__ */ jsx4("h3", { className: "hd-col-title", children: "Files" }) }),
    files.length ? /* @__PURE__ */ jsx4("ul", { className: "hd-files", children: files.map((file) => /* @__PURE__ */ jsx4("li", { children: /* @__PURE__ */ jsxs4(
      "button",
      {
        "aria-pressed": file.path === selected,
        className: file.path === selected ? "hd-file hd-file--selected" : "hd-file",
        onClick: () => onSelect(file.path),
        title: file.path,
        type: "button",
        children: [
          /* @__PURE__ */ jsx4(KindMark, { kind: file.kind }),
          /* @__PURE__ */ jsx4("span", { className: "hd-file-name", children: file.rel || baseName(file.path) }),
          /* @__PURE__ */ jsx4("span", { className: `hd-state hd-state--${file.state}`, children: STATE_LABELS[file.state] ?? file.state })
        ]
      }
    ) }, file.path)) }) : /* @__PURE__ */ jsx4(Empty, { children: "No files touched yet." })
  ] });
}
function RedlineActions({ path, workstreamId, onDone }) {
  const [busy, setBusy] = useState2("");
  const [message, setMessage] = useState2("");
  const [error, setError] = useState2("");
  const run = async (action) => {
    setBusy(action);
    setError("");
    setMessage("");
    try {
      const result = await api.redline({ path, action, workstream_id: workstreamId });
      setMessage(`${action === "accept" ? "Accepted" : "Rejected"} ${result.resolved ?? 0} change${result.resolved === 1 ? "" : "s"}.`);
      await onDone();
    } catch (err) {
      setError(errorText(err));
    } finally {
      setBusy("");
    }
  };
  return /* @__PURE__ */ jsxs4(Fragment2, { children: [
    /* @__PURE__ */ jsx4(Button3, { disabled: Boolean(busy), onClick: () => void run("accept"), size: "sm", children: busy === "accept" ? "Accepting…" : "Accept changes" }),
    /* @__PURE__ */ jsx4(Button3, { disabled: Boolean(busy), onClick: () => void run("reject"), size: "sm", variant: "outline", children: busy === "reject" ? "Rejecting…" : "Reject changes" }),
    message ? /* @__PURE__ */ jsx4("span", { className: "hd-ok", children: message }) : null,
    /* @__PURE__ */ jsx4(InlineError, { message: error })
  ] });
}
function OpenExternal({ file }) {
  const os = osDoor();
  const [error, setError] = useState2("");
  if (typeof os?.openExternal !== "function") {
    return null;
  }
  const open = async () => {
    setError("");
    const opened = await os.openExternal(file.path);
    if (!opened) {
      setError("This window cannot open files with the system app.");
    }
  };
  return /* @__PURE__ */ jsxs4(Fragment2, { children: [
    /* @__PURE__ */ jsx4(Button3, { onClick: () => void open(), size: "sm", variant: "ghost", children: OPEN_LABELS[file.kind] ?? "Open" }),
    /* @__PURE__ */ jsx4(InlineError, { message: error })
  ] });
}
function hasTrackedRuns(preview) {
  return (preview?.docx?.paragraphs ?? []).some((paragraph) => (paragraph.runs ?? []).some((run) => run.rev));
}
function FileDetail({ file, receipts, workstream, version, onChanged }) {
  const fileReceipts = receipts.filter((receipt) => receipt.path === file.path);
  const baseline = fileReceipts.reduce((first, receipt) => first === null || receipt.id < first.id ? receipt : first, null);
  const baselineId = baseline?.id;
  const diffQuery = useQuery3({
    queryKey: keys.diff(file.path, baselineId, version),
    queryFn: () => api.diff(file.path, baselineId),
    enabled: baselineId !== void 0
  });
  const afterQuery = useQuery3({
    queryKey: keys.preview(file.path, void 0, version),
    queryFn: () => api.preview(file.path)
  });
  const changed = changedCells(diffQuery.data);
  const showRedline = file.kind === "docx" && ((diffQuery.data?.revisions ?? []).length > 0 || hasTrackedRuns(afterQuery.data));
  return /* @__PURE__ */ jsxs4("section", { className: "hd-col hd-col--preview", children: [
    /* @__PURE__ */ jsxs4("div", { className: "hd-col-head", children: [
      /* @__PURE__ */ jsx4("h3", { className: "hd-col-title", title: file.path, children: file.rel || baseName(file.path) }),
      /* @__PURE__ */ jsxs4("div", { className: "hd-actions", children: [
        showRedline ? /* @__PURE__ */ jsx4(RedlineActions, { onDone: onChanged, path: file.path, workstreamId: workstream.id }) : null,
        /* @__PURE__ */ jsx4(OpenExternal, { file })
      ] })
    ] }),
    diffQuery.error ? /* @__PURE__ */ jsx4(QueryError, { error: diffQuery.error, onRetry: diffQuery.refetch }) : null,
    diffQuery.data ? /* @__PURE__ */ jsx4(DiffSummary, { diff: diffQuery.data }) : null,
    /* @__PURE__ */ jsxs4("div", { className: "hd-panes", children: [
      /* @__PURE__ */ jsx4(
        PreviewPane,
        {
          changed,
          emptyText: "No receipt yet, so there is no saved original to compare against.",
          enabled: baselineId !== void 0,
          label: "Before",
          path: file.path,
          receiptId: baselineId,
          version
        }
      ),
      /* @__PURE__ */ jsx4(PreviewPane, { changed, label: "After", path: file.path, version })
    ] })
  ] });
}
function ReceiptRow({ receipt, onChanged }) {
  const [confirming, setConfirming] = useState2(false);
  const [busy, setBusy] = useState2(false);
  const [error, setError] = useState2("");
  const revert = async () => {
    setBusy(true);
    setError("");
    try {
      await api.revert(receipt.id);
      setConfirming(false);
      await onChanged();
    } catch (err) {
      setError(errorText(err));
    } finally {
      setBusy(false);
    }
  };
  return /* @__PURE__ */ jsxs4("li", { className: "hd-receipt", children: [
    /* @__PURE__ */ jsxs4("div", { className: "hd-row-main", children: [
      /* @__PURE__ */ jsxs4("div", { className: "hd-row-title", children: [
        receipt.tool,
        " · ",
        baseName(receipt.path)
      ] }),
      /* @__PURE__ */ jsxs4("div", { className: "hd-row-meta", children: [
        /* @__PURE__ */ jsx4("span", { children: receipt.op_summary }),
        /* @__PURE__ */ jsx4("span", { className: "hd-sep", children: "·" }),
        /* @__PURE__ */ jsx4("span", { title: clockTime(receipt.created_at), children: timeAgo(receipt.created_at) }),
        receipt.reverted_at ? /* @__PURE__ */ jsxs4(Fragment2, { children: [
          /* @__PURE__ */ jsx4("span", { className: "hd-sep", children: "·" }),
          /* @__PURE__ */ jsxs4("span", { children: [
            "reverted ",
            timeAgo(receipt.reverted_at)
          ] })
        ] }) : null
      ] }),
      /* @__PURE__ */ jsx4(InlineError, { message: error })
    ] }),
    receipt.reverted_at || !receipt.backup_path ? null : confirming ? /* @__PURE__ */ jsxs4("div", { className: "hd-actions", children: [
      /* @__PURE__ */ jsx4("span", { className: "hd-muted", children: "Restore the file from before this change?" }),
      /* @__PURE__ */ jsx4(Button3, { disabled: busy, onClick: () => void revert(), size: "xs", variant: "destructive", children: busy ? "Reverting…" : "Revert" }),
      /* @__PURE__ */ jsx4(Button3, { disabled: busy, onClick: () => setConfirming(false), size: "xs", variant: "ghost", children: "Cancel" })
    ] }) : /* @__PURE__ */ jsx4(Button3, { onClick: () => setConfirming(true), size: "xs", variant: "outline", children: "Revert" })
  ] });
}
function Receipts({ receipts, onChanged }) {
  if (!receipts.length) {
    return null;
  }
  return /* @__PURE__ */ jsxs4("section", { className: "hd-section", children: [
    /* @__PURE__ */ jsx4("h3", { className: "hd-col-title", children: "Receipts" }),
    /* @__PURE__ */ jsx4("ul", { className: "hd-list", children: receipts.map((receipt) => /* @__PURE__ */ jsx4(ReceiptRow, { onChanged, receipt }, receipt.id)) })
  ] });
}
function Timeline({ events }) {
  return /* @__PURE__ */ jsxs4("details", { className: "hd-timeline", children: [
    /* @__PURE__ */ jsxs4("summary", { className: "hd-diff-summary", children: [
      "Events (",
      events.length,
      ")"
    ] }),
    events.length ? /* @__PURE__ */ jsx4("ul", { className: "hd-events", children: events.map((event) => /* @__PURE__ */ jsxs4("li", { className: event.kind === "refused" ? "hd-event hd-event--refused" : "hd-event", children: [
      /* @__PURE__ */ jsx4("span", { className: "hd-event-ts", title: clockTime(event.ts), children: timeAgo(event.ts) }),
      /* @__PURE__ */ jsx4("span", { className: "hd-event-kind", children: event.kind }),
      /* @__PURE__ */ jsx4("span", { className: "hd-event-tool", children: event.tool ?? "" }),
      /* @__PURE__ */ jsx4("span", { className: "hd-event-summary", children: event.summary })
    ] }, event.id)) }) : /* @__PURE__ */ jsx4(Empty, { children: "No events yet." })
  ] });
}
function WorkstreamView({ id }) {
  const client = useQueryClient2();
  const [activeHint, setActiveHint] = useState2(true);
  const { data, error, isLoading, refetch } = useQuery3({
    queryKey: keys.workstream(id),
    queryFn: () => api.workstream(id),
    refetchInterval: activeHint ? 4e3 : false
  });
  const [selected, setSelected] = useState2(null);
  const workstream = data?.workstream;
  const active = ACTIVE_STATUSES.has(workstream?.status);
  const files = data?.files ?? [];
  const receipts = data?.receipts ?? [];
  const events = data?.events ?? [];
  const file = files.find((item) => item.path === selected) ?? files[0] ?? null;
  useEffect(() => {
    if (workstream) {
      setActiveHint(active);
    }
  }, [workstream, active]);
  const refreshAll = () => client.invalidateQueries({ queryKey: keys.all });
  const version = `${workstream?.updated_at ?? ""}:${receipts.length}`;
  return /* @__PURE__ */ jsxs4("div", { className: "hd-page", children: [
    /* @__PURE__ */ jsx4("button", { className: "hd-back", onClick: () => $selectedWorkstream.set(null), type: "button", children: "← Desk" }),
    isLoading ? /* @__PURE__ */ jsx4(Skeleton, { lines: 5 }) : null,
    error ? /* @__PURE__ */ jsx4(QueryError, { error, onRetry: refetch }) : null,
    workstream ? /* @__PURE__ */ jsxs4(Fragment2, { children: [
      /* @__PURE__ */ jsxs4("header", { className: "hd-ws-header", children: [
        /* @__PURE__ */ jsx4("h1", { className: "hd-title", children: workstream.title }),
        /* @__PURE__ */ jsxs4("div", { className: "hd-row-meta", children: [
          /* @__PURE__ */ jsxs4("span", { children: [
            "Started ",
            timeAgo(workstream.created_at)
          ] }),
          /* @__PURE__ */ jsx4("span", { className: "hd-sep", children: "·" }),
          /* @__PURE__ */ jsxs4("span", { children: [
            workstream.files_changed ?? 0,
            " files changed"
          ] })
        ] })
      ] }),
      /* @__PURE__ */ jsx4(ActivityIsland, { events, now: data.now, workstream }),
      /* @__PURE__ */ jsxs4("div", { className: "hd-ws-grid", children: [
        /* @__PURE__ */ jsx4(Conversation, { active, workstream }),
        /* @__PURE__ */ jsx4(FileList, { files, onSelect: setSelected, selected: file?.path ?? null }),
        file ? /* @__PURE__ */ jsx4(
          FileDetail,
          {
            file,
            onChanged: refreshAll,
            receipts,
            version,
            workstream
          },
          file.path
        ) : /* @__PURE__ */ jsx4("section", { className: "hd-col hd-col--preview", children: /* @__PURE__ */ jsx4(Empty, { children: "Select a file to preview it." }) })
      ] }),
      /* @__PURE__ */ jsx4(Receipts, { onChanged: refreshAll, receipts }),
      /* @__PURE__ */ jsx4(Timeline, { events })
    ] }) : null
  ] });
}

// desktop/src/plugin.jsx
import { jsx as jsx5 } from "react/jsx-runtime";
var WORKSTREAM_UPDATED = "plugin.office-desk.workstream.updated";
function DeskPage() {
  const selected = useValue($selectedWorkstream);
  return selected === null ? /* @__PURE__ */ jsx5(DeskHome, {}) : /* @__PURE__ */ jsx5(WorkstreamView, { id: selected }, selected);
}
var plugin_default = {
  id: "office-desk",
  name: "Desk",
  description: "Office work on your folders: Word, Excel and PowerPoint edits with tracked changes, previews and receipts.",
  register(ctx) {
    ctx.onDispose(bindContext(ctx));
    ctx.onDispose(injectStyle());
    ctx.onDispose(() => $selectedWorkstream.set(null));
    ctx.onDispose(host4.onEvent(WORKSTREAM_UPDATED, () => void queryClient.invalidateQueries({ queryKey: [KEY] })));
    ctx.registerMany([
      {
        id: "page",
        area: ROUTES_AREA,
        data: { path: "/desk" },
        render: () => /* @__PURE__ */ jsx5(DeskPage, {})
      },
      {
        id: "nav",
        area: SIDEBAR_NAV_AREA,
        data: { path: "/desk", label: "Desk", codicon: "briefcase" }
      }
    ]);
  }
};
export {
  plugin_default as default
};
