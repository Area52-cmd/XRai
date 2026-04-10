// XRai Studio dashboard — vanilla ES2022, no build step.
//
// Architecture:
//   1. Fetch /state once on load for initial paint.
//   2. Open WebSocket /events, apply each event to the corresponding panel.
//   3. That's it. No framework, no virtual DOM. Direct DOM mutation per event.
//
// Target: Chrome / Edge >= 111. Uses native modules, top-level await,
// structuredClone, WebSocket subprotocol. No polyfills.

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ── State ────────────────────────────────────────────────────────

const state = {
  attached: false,
  paused: false,
  frameCount: 0,
  eventCount: 0,
  model: {},            // current ViewModel property dict
  filesSeen: new Set(), // file paths that have flashed
};

// ── WebSocket connection ─────────────────────────────────────────

function openSocket() {
  const proto = location.protocol === "https:" ? "wss:" : "ws:";
  const ws = new WebSocket(`${proto}//${location.host}/events`);

  ws.addEventListener("open", () => {
    console.log("[studio] ws open");
    setAttachStatus("connecting");
  });

  ws.addEventListener("message", (e) => {
    try {
      const evt = JSON.parse(e.data);
      handleEvent(evt);
    } catch (err) {
      console.warn("[studio] bad event", err, e.data);
    }
  });

  ws.addEventListener("close", () => {
    console.log("[studio] ws closed, retry in 2s");
    setAttachStatus("disconnected");
    setTimeout(openSocket, 2000);
  });

  ws.addEventListener("error", (e) => {
    console.warn("[studio] ws error", e);
  });
}

// ── Initial state fetch ──────────────────────────────────────────

async function fetchInitialState() {
  try {
    const res = await fetch("/state");
    if (!res.ok) {
      console.warn("[studio] /state returned", res.status);
      return;
    }
    const s = await res.json();
    applyState(s);
  } catch (err) {
    console.warn("[studio] /state fetch failed", err);
  }
}

function applyState(s) {
  if (s.attached) {
    state.attached = true;
    setAttachStatus("attached", s);
  } else {
    setAttachStatus("not-attached");
  }

  if (s.build?.version) {
    $("#build-text").textContent = `build ${s.build.version}`;
  }

  // Initial screenshot if provided
  if (s.screenshot?.url) {
    const img = $("#screenshot-img");
    img.src = s.screenshot.url;
  }

  if (s.model) {
    state.model = { ...s.model };
    renderModel();
  }
}

// ── Event dispatch ───────────────────────────────────────────────

function handleEvent(evt) {
  state.eventCount++;
  $("#subs-text").textContent = `${state.eventCount} events`;

  switch (evt.kind) {
    case "frame": return onFrame(evt.data);
    case "model.change": return onModelChange(evt.data);
    case "model.exposed": return onModelExposed(evt.data);
    case "pane.exposed": return onPaneExposed(evt.data);
    case "control.change": return onControlChange(evt.data);
    case "file.changed": return onFileChanged(evt.data);
    case "rebuild.step": return onRebuildStep(evt.data);
    case "command.start":
    case "command.end":
      return onCommand(evt);
    case "log": return onLog(evt.data);
    case "error": return onError(evt.data);
    default:
      // Unknown event kinds flow into the command stream for visibility
      appendCommand(evt.ts, evt.kind, "meta", JSON.stringify(evt.data || {}).slice(0, 80));
      break;
  }
}

// ── Frame stream ─────────────────────────────────────────────────

function onFrame(data) {
  if (state.paused) return;
  const img = $("#screenshot-img");
  img.src = `data:${data.mime};base64,${data.b64}`;
  $("#screenshot-empty").classList.add("hidden");
  state.frameCount++;
}

// ── Model ────────────────────────────────────────────────────────

function onModelExposed(data) {
  state.model = { ...(data?.properties || {}) };
  renderModel();
}

function onModelChange(data) {
  if (!data || !data.property) return;
  state.model[data.property] = data.new;
  renderModel(data.property);
}

function renderModel(flashKey) {
  const tbody = $("#model-table tbody");
  const keys = Object.keys(state.model).sort();

  if (keys.length === 0) {
    $("#model-empty").classList.remove("hidden");
    tbody.innerHTML = "";
    return;
  }
  $("#model-empty").classList.add("hidden");

  // Rebuild table from scratch — simple and fast for < 100 properties.
  const frag = document.createDocumentFragment();
  for (const k of keys) {
    const tr = document.createElement("tr");
    if (k === flashKey) tr.classList.add("flash");
    const tdK = document.createElement("td");
    tdK.className = "k";
    tdK.textContent = k;
    const tdV = document.createElement("td");
    tdV.className = "v";
    tdV.textContent = formatValue(state.model[k]);
    tr.appendChild(tdK);
    tr.appendChild(tdV);
    frag.appendChild(tr);
  }
  tbody.innerHTML = "";
  tbody.appendChild(frag);
}

function formatValue(v) {
  if (v === null || v === undefined) return "—";
  if (typeof v === "string") return `"${v}"`;
  if (typeof v === "object") {
    try { return JSON.stringify(v).slice(0, 60); } catch { return "[object]"; }
  }
  return String(v);
}

// ── Pane (control tree) ─────────────────────────────────────────

function onPaneExposed(data) {
  const count = data?.controlCount ?? 0;
  const rootType = data?.rootType ?? "unknown";
  appendCommand(Date.now(), "pane.exposed", "ok", `${rootType} · ${count} controls`);
}

function onControlChange(data) {
  if (!data) return;
  appendCommand(Date.now(), `control ${data.controlName || "?"}.${data.propertyName || "?"}`,
    "meta", formatValue(data.newValue));
}

// ── Files ────────────────────────────────────────────────────────

function onFileChanged(data) {
  if (!data || !data.path) return;
  const path = data.path;
  $("#files-empty").classList.add("hidden");

  const list = $("#files-list");
  const existing = list.querySelector(`li[data-path="${CSS.escape(path)}"]`);
  if (existing) {
    existing.classList.remove("ok"); // trigger reflow for re-animation
    void existing.offsetWidth;
    existing.classList.add("ok");
    existing.querySelector(".ts").textContent = shortTime(Date.now());
    return;
  }

  const li = document.createElement("li");
  li.className = "ok";
  li.dataset.path = path;
  li.innerHTML = `
    <span class="ts">${shortTime(Date.now())}</span>
    <span class="name">${escapeHtml(path)}</span>
    <span class="meta">${data.kind || "change"}</span>
  `;
  list.insertBefore(li, list.firstChild);

  // Cap at 50 entries
  while (list.children.length > 50) list.removeChild(list.lastChild);
}

// ── Build console ───────────────────────────────────────────────

function onRebuildStep(data) {
  if (!data || !data.step) return;
  $("#build-empty").classList.add("hidden");
  const list = $("#build-list");

  const existing = list.querySelector(`li[data-step="${CSS.escape(data.step)}"]`);
  const li = existing ?? document.createElement("li");
  li.dataset.step = data.step;
  li.className = data.status || "meta";

  const elapsed = data.elapsedMs != null ? `${data.elapsedMs} ms` : "";
  li.innerHTML = `
    <span class="ts">${shortTime(Date.now())}</span>
    <span class="name">${escapeHtml(data.step)}</span>
    <span class="meta">${escapeHtml(data.status || "")} ${elapsed} ${data.detail ? `— ${escapeHtml(data.detail)}` : ""}</span>
  `;
  if (!existing) list.insertBefore(li, list.firstChild);

  // Cap at 30 entries
  while (list.children.length > 30) list.removeChild(list.lastChild);
}

// ── Command stream ──────────────────────────────────────────────

function onCommand(evt) {
  const data = evt.data || {};
  const cmd = data.cmd || data.name || "?";
  const status = data.ok === false ? "err" : (evt.kind === "command.start" ? "start" : "ok");
  const elapsed = data.elapsedMs != null ? `${data.elapsedMs} ms` : "";
  appendCommand(evt.ts, cmd, status, elapsed);
}

function onLog(data) {
  appendCommand(Date.now(), "log", "meta", (data?.message || "").slice(0, 100));
}

function onError(data) {
  appendCommand(Date.now(), data?.type || "error", "err", (data?.message || "").slice(0, 100));
}

function appendCommand(ts, name, status, meta) {
  $("#commands-empty").classList.add("hidden");
  const list = $("#commands-list");
  const li = document.createElement("li");
  li.className = status || "";
  li.innerHTML = `
    <span class="ts">${shortTime(ts)}</span>
    <span class="name">${escapeHtml(name)}</span>
    <span class="meta">${escapeHtml(meta || "")}</span>
  `;
  list.insertBefore(li, list.firstChild);
  while (list.children.length > 80) list.removeChild(list.lastChild);
}

// ── Status chip ─────────────────────────────────────────────────

function setAttachStatus(state, s = null) {
  const dot = $("#attach-dot");
  const txt = $("#attach-text");
  dot.className = "dot";
  switch (state) {
    case "attached":
      dot.classList.add("ok");
      txt.textContent = s?.excel?.workbook
        ? `attached · ${s.excel.workbook}`
        : "attached";
      break;
    case "not-attached":
      dot.classList.add("warn");
      txt.textContent = "not attached";
      break;
    case "connecting":
      dot.classList.add("warn");
      txt.textContent = "connecting…";
      break;
    case "disconnected":
      dot.classList.add("err");
      txt.textContent = "disconnected";
      break;
  }
}

// ── Controls ────────────────────────────────────────────────────

function wireControls() {
  $("#btn-pause").addEventListener("click", () => {
    state.paused = !state.paused;
    $("#btn-pause").classList.toggle("active", state.paused);
    $("#btn-pause").textContent = state.paused ? "▶ Resume" : "⏸ Pause";
  });

  document.addEventListener("keydown", (e) => {
    // Only hotkeys when no input has focus
    if (document.activeElement?.tagName === "INPUT") return;
    if (e.key === " ") {
      e.preventDefault();
      $("#btn-pause").click();
    }
  });
}

// ── Utilities ───────────────────────────────────────────────────

function shortTime(ms) {
  const d = new Date(ms);
  return d.toLocaleTimeString("en-US", { hour12: false }).split(" ")[0];
}

function escapeHtml(s) {
  if (s == null) return "";
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

// ── Boot ────────────────────────────────────────────────────────

wireControls();
await fetchInitialState();
openSocket();
