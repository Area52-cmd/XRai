// XRai Studio dashboard — vanilla ES2022, no build step.
//
// Responsibilities (all read-only; Studio never edits your code):
//   1. Offer an upfront "follow your IDE" onboarding experience.
//   2. Tail the agent's session transcript and render every thought,
//      edit, and tool call as a scrolling activity feed.
//   3. Show the live target app screenshot at ~4 fps.
//   4. When follow-mode is on, auto-launch the user's IDE on every
//      file edit so they watch the code land IN their real editor —
//      Studio never tries to replace VS Code / VS 2026 / Rider.
//
// Target: Chrome / Edge >= 111 — native ES modules, top-level await.

const $ = (sel) => document.querySelector(sel);

// ── State ────────────────────────────────────────────────────────

const state = {
  attached: false,
  paused: false,
  frameCount: 0,
  eventCount: 0,
  agentName: null,
  agentConnected: false,
  model: {},
  filesSeen: new Map(),  // path -> li element
  buildSteps: new Map(), // step -> li element
  ides: [],              // [{kind, name, installed, running, installUrl, ...}]
  preferences: null,     // loaded from /preferences
  recentFileEdits: new Map(), // filePath -> lastTs, for de-duping rapid edits
};

// ── Boot sequence ────────────────────────────────────────────────

async function boot() {
  wireControls();
  wireOverlay();
  await fetchInitialState();
  await loadPreferences();
  await loadIdes();
  maybeShowStartupOverlay();
  openSocket();
}

// ── WebSocket ────────────────────────────────────────────────────

function openSocket() {
  const proto = location.protocol === "https:" ? "wss:" : "ws:";
  const ws = new WebSocket(`${proto}//${location.host}/events`);

  ws.addEventListener("open", () => setStatus("connecting"));
  ws.addEventListener("message", (e) => {
    try {
      const evt = JSON.parse(e.data);
      handleEvent(evt);
    } catch (err) {
      console.warn("[studio] bad event", err, e.data);
    }
  });
  ws.addEventListener("close", () => {
    setStatus("disconnected");
    setTimeout(openSocket, 2000);
  });
}

// ── Initial state + preferences + IDE list ──────────────────────

async function fetchInitialState() {
  try {
    const s = await fetchJson("/state");
    if (s?.attached) {
      state.attached = true;
      setStatus("attached", s);
    } else {
      setStatus("not-attached");
    }
  } catch (err) {
    console.warn("[studio] /state fetch failed", err);
  }
}

async function loadPreferences() {
  try {
    state.preferences = await fetchJson("/preferences");
  } catch {
    state.preferences = { followMode: false, preferredIde: null, onboarded: false };
  }
  renderFollowChip();
}

async function savePreferences() {
  try {
    const res = await fetch("/preferences", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(state.preferences),
    });
    const saved = await res.json();
    state.preferences = saved;
  } catch (err) {
    toast("Failed to save preferences", "err");
  }
}

async function loadIdes() {
  try {
    state.ides = await fetchJson("/ides");
  } catch {
    state.ides = [];
  }
  renderIdeChip();
}

function renderIdeChip() {
  const dot = $("#ide-dot");
  const name = $("#ide-name");

  if (state.preferences?.preferredIde) {
    const chosen = state.ides.find(i => i.kind === state.preferences.preferredIde);
    if (chosen?.installed) {
      dot.className = "dot ok";
      name.textContent = chosen.name;
      return;
    }
  }

  const running = state.ides.find(i => i.running);
  if (running) {
    dot.className = "dot ok";
    name.textContent = running.name;
    return;
  }

  const installed = state.ides.find(i => i.installed);
  if (installed) {
    dot.className = "dot warn";
    name.textContent = `${installed.name} (idle)`;
    return;
  }

  dot.className = "dot";
  name.textContent = "no IDE";
}

function renderFollowChip() {
  const btn = $("#btn-follow");
  const label = $("#follow-label");
  const on = state.preferences?.followMode === true;
  btn.classList.toggle("active", on);
  label.textContent = on ? "Follow: on" : "Follow: off";
}

// ── Startup overlay ─────────────────────────────────────────────

function maybeShowStartupOverlay() {
  // Skip only if the user has already been onboarded AND a preferred IDE
  // is still valid (still installed).
  if (state.preferences?.onboarded && state.preferences?.preferredIde) {
    const chosen = state.ides.find(i => i.kind === state.preferences.preferredIde);
    if (chosen?.installed) return;
  }
  showStartupOverlay();
}

function showStartupOverlay() {
  const overlay = $("#startup-overlay");
  const choices = $("#ide-choices");
  choices.innerHTML = "";

  for (const ide of state.ides) {
    const row = document.createElement("div");
    row.className = `ide-choice ${ide.installed ? "" : "unavailable"}`;
    row.dataset.kind = ide.kind;

    const status = ide.running ? "running"
                  : ide.installed ? "installed"
                  : "not-installed";
    const statusLabel = ide.running ? "running"
                  : ide.installed ? "installed"
                  : "not installed";

    const action = ide.installed
      ? `<span class="ide-action">${ide.running ? "Use this" : "Launch & use"}</span>`
      : `<a class="ide-action-install" href="${escapeAttr(ide.installUrl)}" target="_blank" rel="noopener">Install →</a>`;

    row.innerHTML = `
      <div class="ide-icon">${ideInitial(ide.kind)}</div>
      <div class="ide-body">
        <div class="ide-name">${escapeHtml(ide.name)} <span class="ide-status ${status}">${statusLabel}</span></div>
        <div class="ide-tag">${escapeHtml(ide.installTagline || "")}</div>
      </div>
      <div class="ide-action-wrap">${action}</div>
    `;

    if (ide.installed) {
      row.addEventListener("click", (e) => {
        if (e.target.closest(".ide-action-install")) return; // let install link pass through
        selectIde(ide.kind);
      });
    }

    choices.appendChild(row);
  }

  // Pre-select: running IDE first, then installed
  const pre = state.ides.find(i => i.running) || state.ides.find(i => i.installed);
  if (pre) selectIde(pre.kind);

  overlay.classList.remove("hidden");
}

function selectIde(kind) {
  for (const el of document.querySelectorAll(".ide-choice")) {
    el.classList.toggle("selected", el.dataset.kind === kind);
  }
  state.pendingIdeSelection = kind;
}

function hideStartupOverlay() {
  $("#startup-overlay").classList.add("hidden");
}

function wireOverlay() {
  document.addEventListener("DOMContentLoaded", () => {}); // no-op placeholder

  document.getElementById("overlay-continue")?.addEventListener("click", async () => {
    const kind = state.pendingIdeSelection;
    const follow = $("#follow-toggle").checked;
    state.preferences = state.preferences || {};
    state.preferences.preferredIde = kind || null;
    state.preferences.followMode = follow && !!kind;
    state.preferences.onboarded = true;
    await savePreferences();

    // If an IDE was chosen and it's not running, politely launch it
    if (kind) {
      const chosen = state.ides.find(i => i.kind === kind);
      if (chosen?.installed && !chosen.running) {
        await launchIde(kind);
      }
    }

    renderFollowChip();
    renderIdeChip();
    hideStartupOverlay();
    toast(
      follow && kind
        ? `Following edits in ${state.ides.find(i => i.kind === kind)?.name || kind}`
        : "Just watching — no IDE follow",
      "ok"
    );
  });

  document.getElementById("overlay-skip")?.addEventListener("click", async () => {
    state.preferences = state.preferences || {};
    state.preferences.followMode = false;
    state.preferences.onboarded = true;
    await savePreferences();
    renderFollowChip();
    hideStartupOverlay();
    toast("Watch-only mode — no IDE follow", "info");
  });
}

// ── IDE launcher calls ─────────────────────────────────────────

async function launchIde(kind) {
  try {
    const res = await fetch("/ide/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ kind }),
    });
    return await res.json();
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

async function openFileInIde(filePath, line) {
  if (!filePath) return;
  try {
    const res = await fetch("/ide/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        filePath,
        line: line || null,
        kind: state.preferences?.preferredIde || null,
      }),
    });
    const result = await res.json();
    if (!result.ok) {
      toast(`IDE open failed: ${result.error}`, "err");
    }
    return result;
  } catch (err) {
    toast(`IDE open error: ${err}`, "err");
    return { ok: false };
  }
}

// ── Event dispatch ──────────────────────────────────────────────

function handleEvent(evt) {
  state.eventCount++;
  $("#events-count").textContent = state.eventCount.toLocaleString();

  switch (evt.kind) {
    case "agent.session":         return onAgentSession(evt.data);
    case "agent.message.user":    return onAgentUser(evt.data);
    case "agent.message.text":    return onAgentAssistant(evt.data);
    case "agent.message.think":   return onAgentThinking(evt.data);
    case "agent.tool.use":        return onAgentToolUse(evt.data);
    case "agent.tool.result":     return onAgentToolResult(evt.data);
    case "frame":                 return onFrame(evt.data);
    case "model.change":          return onModelChange(evt.data);
    case "model.exposed":         return onModelExposed(evt.data);
    case "pane.exposed":          return onPaneExposed(evt.data);
    case "control.change":        return onControlChange(evt.data);
    case "file.changed":          return onFileChanged(evt.data);
    case "rebuild.step":          return onRebuildStep(evt.data);
    case "command.start":
    case "command.end":           return onCommand(evt);
    default: break;
  }
}

// ── Screenshot stream ───────────────────────────────────────────

function onFrame(data) {
  if (state.paused) return;
  const img = $("#screenshot-img");
  img.src = `data:${data.mime};base64,${data.b64}`;
  $("#screenshot-empty").classList.add("hidden");
  state.frameCount++;
  const meta = $("#frame-meta");
  if (meta) {
    meta.textContent = `${state.frameCount} frames · ${Math.round((data.bytes || 0) / 1024)} KB`;
  }
}

// ── Agent feed ──────────────────────────────────────────────────

let currentAssistantItem = null;
let currentAssistantUuid = null;

function ensureAgentFeedVisible() {
  $("#agent-empty").classList.add("hidden");
}

function onAgentSession(data) {
  state.agentName = data?.agent || "agent";
  state.agentConnected = true;
  $("#agent-name").textContent = data?.agent || "connected";
  $("#agent-dot").className = "dot ok";
  const meta = $("#agent-meta");
  if (meta) {
    const file = data?.file || "";
    meta.textContent = file ? file.split(/[\\/]/).slice(-2).join("/") : "";
  }
}

function onAgentUser(data) {
  ensureAgentFeedVisible();
  appendAgentItem(createAgentItem({
    cls: "user",
    icon: "You",
    role: "User",
    text: data?.text || "",
  }));
}

function onAgentAssistant(data) {
  ensureAgentFeedVisible();
  const uuid = data?.uuid;
  const text = data?.text || "";

  if (currentAssistantItem && currentAssistantUuid === uuid) {
    const textEl = currentAssistantItem.querySelector(".text");
    textEl.textContent += "\n\n" + text;
    scrollFeedToBottom();
    return;
  }

  const item = createAgentItem({
    cls: "assistant",
    icon: state.agentName ? state.agentName.charAt(0) : "C",
    role: state.agentName || "Assistant",
    text,
  });
  appendAgentItem(item);
  currentAssistantItem = item;
  currentAssistantUuid = uuid;
}

function onAgentThinking(data) {
  ensureAgentFeedVisible();
  appendAgentItem(createAgentItem({
    cls: "thinking",
    icon: "⋯",
    role: "Thinking",
    text: data?.text || "",
  }));
  currentAssistantItem = null;
  currentAssistantUuid = null;
}

function onAgentToolUse(data) {
  ensureAgentFeedVisible();
  const toolName = data?.toolName || "?";
  const toolUseId = data?.toolUseId;
  const toolCls = classifyTool(toolName);
  const iconChar = iconForTool(toolName);
  const target = inferToolTarget(data);
  const description = data?.description || "";
  const detail = buildToolDetail(data);

  const item = document.createElement("div");
  item.className = `agent-item tool ${toolCls}`;
  if (toolUseId) item.dataset.toolUseId = toolUseId;
  item.innerHTML = `
    <div class="gutter">
      <div class="icon">${escapeHtml(iconChar)}</div>
    </div>
    <div class="body">
      <div class="role">${escapeHtml(friendlyToolLabel(toolName))}</div>
      <div class="tool-header">
        <span class="tool-target">${escapeHtml(target)}</span>
      </div>
      ${description ? `<div class="tool-desc">${escapeHtml(description)}</div>` : ""}
      ${detail.html}
      <div class="ts">${shortTime(Date.now())}</div>
    </div>
  `;
  appendAgentItem(item);

  currentAssistantItem = null;
  currentAssistantUuid = null;

  // Click → open file in IDE
  const clickable = item.querySelector(".tool-detail.clickable");
  if (clickable && detail.filePath) {
    clickable.addEventListener("click", () => {
      openFileInIde(detail.filePath, detail.line);
    });
  }

  // Auto-follow: if follow-mode is on AND this is an edit/write, open the
  // file in the user's IDE — but rate-limited so a burst of 50 edits doesn't
  // spam the IDE with 50 launches. Two layers of throttling:
  //   1. Per-file: max once per 500ms on the same path (de-dupe)
  //   2. Global: max one IDE launch per 1000ms across all paths
  // The second layer means a 50-edit burst opens at most one file per second,
  // and we drop the older queued edits in favor of the newest one.
  if (state.preferences?.followMode && detail.filePath) {
    const n = (toolName || "").toLowerCase();
    if (n === "edit" || n === "write" || n === "notebookedit") {
      const now = Date.now();
      const lastForFile = state.recentFileEdits.get(detail.filePath) || 0;
      if (now - lastForFile > 500) {
        state.recentFileEdits.set(detail.filePath, now);
        scheduleIdeOpen(detail.filePath, detail.line);
      }
    }
  }
}

// ── Global IDE launch throttle ─────────────────────────────────
// Schedules an IDE open with a 1-second global rate limit. If multiple
// edits arrive within the throttle window, only the most recent one is
// opened — older queued targets are silently dropped because the user
// almost always wants to see the LATEST edit, not the first one.
let lastIdeLaunchAt = 0;
let pendingIdeOpen = null;
let pendingIdeTimer = null;
const IDE_THROTTLE_MS = 1000;

function scheduleIdeOpen(filePath, line) {
  const now = Date.now();
  const timeSince = now - lastIdeLaunchAt;

  // Replace any pending open with the latest target
  pendingIdeOpen = { filePath, line };

  if (timeSince >= IDE_THROTTLE_MS) {
    // Throttle window has passed — open immediately
    flushPendingIdeOpen();
  } else if (!pendingIdeTimer) {
    // Schedule a flush at the end of the throttle window
    pendingIdeTimer = setTimeout(() => {
      pendingIdeTimer = null;
      flushPendingIdeOpen();
    }, IDE_THROTTLE_MS - timeSince);
  }
}

function flushPendingIdeOpen() {
  if (!pendingIdeOpen) return;
  const target = pendingIdeOpen;
  pendingIdeOpen = null;
  lastIdeLaunchAt = Date.now();

  openFileInIde(target.filePath, target.line).then(r => {
    if (r?.ok) {
      toast(`Opened ${shortPath(target.filePath)} in ${r.name || r.ide || "IDE"}`, "ok", 1800);
    }
  });
}

function classifyTool(name) {
  if (!name) return "";
  const n = name.toLowerCase();
  if (n === "edit" || n === "write" || n === "notebookedit") return "edit";
  if (n === "bash") return "bash";
  return "";
}

function iconForTool(name) {
  if (!name) return "?";
  const n = name.toLowerCase();
  if (n === "edit") return "✎";
  if (n === "write") return "+";
  if (n === "read") return "👁";
  if (n === "bash") return "$";
  if (n === "grep") return "/";
  if (n === "glob") return "*";
  if (n === "todowrite") return "☐";
  if (n === "webfetch" || n === "websearch") return "🌐";
  if (n === "task" || n === "agent") return "⚡";
  return name.charAt(0).toUpperCase();
}

// Friendly, plain-English label for a tool name. The raw transcript uses
// internal API names (Edit, Glob, NotebookEdit, etc.) which mean nothing
// to a non-expert. We translate those to natural language.
function friendlyToolLabel(name) {
  if (!name) return "Working";
  const n = name.toLowerCase();
  if (n === "edit") return "Editing file";
  if (n === "write") return "Creating file";
  if (n === "read") return "Reading file";
  if (n === "bash") return "Running command";
  if (n === "grep") return "Searching code";
  if (n === "glob") return "Finding files";
  if (n === "todowrite") return "Updating todo list";
  if (n === "webfetch") return "Fetching web page";
  if (n === "websearch") return "Searching the web";
  if (n === "task" || n === "agent") return "Running sub-agent";
  if (n === "notebookedit") return "Editing notebook";
  return name;
}

function inferToolTarget(data) {
  if (data?.filePath) return data.filePath.split(/[\\/]/).slice(-3).join("/");
  if (data?.path) return data.path.split(/[\\/]/).slice(-3).join("/");
  if (data?.command) return data.command.length > 80 ? data.command.slice(0, 77) + "…" : data.command;
  if (data?.pattern) return data.pattern;
  if (data?.url) return data.url;
  if (data?.query) return data.query;
  return "";
}

function buildToolDetail(data) {
  const detail = {
    filePath: data?.filePath || null,
    oldString: data?.oldString,
    newString: data?.newString,
    fullContent: data?.fullContent,
    line: null,
  };
  const name = (data?.toolName || "").toLowerCase();

  if (name === "edit" && data?.oldString !== undefined && data?.newString !== undefined) {
    detail.html = `<div class="tool-detail clickable">
      ${renderInlineDiff(data.oldString, data.newString)}
      <span class="open-hint">Click: open in IDE</span>
    </div>`;
    return detail;
  }

  if (name === "write" && data?.fullContent) {
    detail.html = `<div class="tool-detail clickable">
      ${renderInlineContent(data.fullContent)}
      <span class="open-hint">Click: open in IDE</span>
    </div>`;
    return detail;
  }

  if (name === "bash" && data?.command) {
    detail.html = `<div class="tool-detail">$ ${escapeHtml(data.command)}</div>`;
    return detail;
  }

  detail.html = "";
  return detail;
}

function renderInlineDiff(oldStr, newStr) {
  const oldLines = (oldStr || "").split("\n");
  const newLines = (newStr || "").split("\n");
  const maxPreview = 8;
  const parts = [];
  for (let i = 0; i < Math.min(oldLines.length, maxPreview); i++) {
    parts.push(`<span class="diff-line-del">- ${escapeHtml(oldLines[i])}</span>`);
  }
  if (oldLines.length > maxPreview) parts.push(`<span class="diff-line-ctx">  … ${oldLines.length - maxPreview} more</span>`);
  for (let i = 0; i < Math.min(newLines.length, maxPreview); i++) {
    parts.push(`<span class="diff-line-add">+ ${escapeHtml(newLines[i])}</span>`);
  }
  if (newLines.length > maxPreview) parts.push(`<span class="diff-line-ctx">  … ${newLines.length - maxPreview} more</span>`);
  return parts.join("");
}

function renderInlineContent(content) {
  const lines = (content || "").split("\n");
  const maxPreview = 10;
  const parts = [];
  for (let i = 0; i < Math.min(lines.length, maxPreview); i++) {
    parts.push(`<span class="diff-line-add">+ ${escapeHtml(lines[i])}</span>`);
  }
  if (lines.length > maxPreview) parts.push(`<span class="diff-line-ctx">  … ${lines.length - maxPreview} more</span>`);
  return parts.join("");
}

function onAgentToolResult(data) {
  const isError = data?.isError === true;
  const toolUseId = data?.toolUseId;

  if (toolUseId) {
    const existing = document.querySelector(`.agent-item.tool[data-tool-use-id="${CSS.escape(toolUseId)}"]`);
    if (existing && isError) existing.classList.add("err");
  }

  if (!isError) return;

  appendAgentItem(createAgentItem({
    cls: "result err",
    icon: "!",
    role: "Error",
    text: (data?.summary || "").slice(0, 600),
  }));
}

function createAgentItem({ cls, icon, role, text }) {
  const item = document.createElement("div");
  item.className = `agent-item ${cls}`;
  item.innerHTML = `
    <div class="gutter">
      <div class="icon">${escapeHtml(icon)}</div>
    </div>
    <div class="body">
      <div class="role">${escapeHtml(role)}</div>
      <div class="text">${escapeHtml(text)}</div>
      <div class="ts">${shortTime(Date.now())}</div>
    </div>
  `;
  return item;
}

function appendAgentItem(item) {
  const feed = $("#agent-feed");
  feed.appendChild(item);
  while (feed.children.length > 500) feed.removeChild(feed.firstChild);
  scrollFeedToBottom();
}

function scrollFeedToBottom() {
  const feed = $("#agent-feed");
  const threshold = 120;
  const nearBottom = feed.scrollHeight - feed.scrollTop - feed.clientHeight < threshold;
  if (nearBottom) {
    requestAnimationFrame(() => { feed.scrollTop = feed.scrollHeight; });
  }
}

// ── Files / Model / Build / Commands ───────────────────────────

function onFileChanged(data) {
  if (!data || !data.path) return;
  $("#files-empty").classList.add("hidden");
  const list = $("#files-list");
  const path = data.path;
  const absolute = data.absolute || path;

  const existing = state.filesSeen.get(path);
  if (existing) {
    existing.classList.remove("ok");
    void existing.offsetWidth;
    existing.classList.add("ok");
    existing.querySelector(".ts").textContent = shortTime(Date.now());
    list.insertBefore(existing, list.firstChild);
    return;
  }

  const li = document.createElement("li");
  li.className = "ok";
  li.dataset.path = absolute;
  li.innerHTML = `
    <span class="ts">${shortTime(Date.now())}</span>
    <span class="name">${escapeHtml(path)}</span>
    <span class="meta">${escapeHtml(data.kind || "change")}</span>
  `;
  li.addEventListener("click", () => openFileInIde(absolute, null));
  li.style.cursor = "pointer";
  list.insertBefore(li, list.firstChild);
  state.filesSeen.set(path, li);
  while (list.children.length > 80) list.removeChild(list.lastChild);
}

function onModelExposed(data) {
  state.model = { ...(data?.properties || {}) };
  renderModel();
}

function onModelChange(data) {
  if (!data || !data.property) return;
  state.model[data.property] = data.new ?? data.value ?? null;
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

function onPaneExposed() { /* quiet */ }
function onControlChange() { /* quiet */ }

function onRebuildStep(data) {
  if (!data || !data.step) return;
  $("#build-empty").classList.add("hidden");
  const list = $("#build-list");

  const existing = state.buildSteps.get(data.step);
  const li = existing ?? document.createElement("li");
  li.className = data.status || "meta";
  state.buildSteps.set(data.step, li);

  const friendly = friendlyBuildStep(data.step);
  const elapsed = data.elapsedMs != null && data.elapsedMs > 0
    ? `${formatMs(data.elapsedMs)}`
    : "";
  const friendlyStatus = friendlyStepStatus(data.status);
  li.innerHTML = `
    <span class="ts">${shortTime(Date.now())}</span>
    <span class="name">${escapeHtml(friendly)}</span>
    <span class="meta">${friendlyStatus} ${elapsed}</span>
  `;
  if (!existing) list.insertBefore(li, list.firstChild);
  while (list.children.length > 30) list.removeChild(list.lastChild);
}

function friendlyStepStatus(status) {
  if (!status) return "";
  const s = status.toLowerCase();
  if (s === "ok") return "✓ done";
  if (s === "error") return "✕ failed";
  if (s === "warning") return "⚠ warning";
  if (s === "skip") return "skipped";
  if (s === "start") return "running…";
  return status;
}

function formatMs(ms) {
  if (ms < 1000) return `${ms} ms`;
  return `${(ms / 1000).toFixed(1)} s`;
}

function onCommand(evt) {
  const data = evt.data || {};
  const stats = $("#timeline-stats");
  if (!stats) return;
  const cmd = data.cmd || "?";
  const kind = evt.kind === "command.start" ? "▸" : (data.ok === false ? "✕" : "✓");
  stats.textContent = `${kind} ${cmd}`;
}

// ── Status chip ─────────────────────────────────────────────────

function setStatus(status, s = null) {
  const dot = $("#attach-dot");
  const txt = $("#attach-text");
  dot.className = "dot";
  switch (status) {
    case "attached": {
      dot.classList.add("ok");
      const label = s?.target?.document || s?.target?.name || "ready";
      txt.textContent = label;
      break;
    }
    case "not-attached":
      dot.classList.add("warn");
      txt.textContent = "ready (no app open)";
      break;
    case "connecting":
      dot.classList.add("warn");
      txt.textContent = "connecting…";
      break;
    case "disconnected":
      dot.classList.add("err");
      txt.textContent = "offline";
      break;
  }
}

// Translate internal build step names (dotnet-restore, nuget-cache-clear)
// into plain-English labels for non-expert users.
function friendlyBuildStep(step) {
  if (!step) return "Working";
  const s = step.toLowerCase();
  if (s === "kill-excel") return "Closing app";
  if (s === "nuget-source") return "Setting up packages";
  if (s === "nuget-cache-clear") return "Clearing package cache";
  if (s === "dotnet-restore") return "Downloading dependencies";
  if (s === "dotnet-build") return "Compiling code";
  if (s === "xll-resolve") return "Locating add-in";
  if (s === "launch-excel") return "Launching app";
  if (s === "attach-com") return "Connecting to app";
  if (s === "hooks-connect") return "Linking to add-in";
  return step;
}

// ── Controls ────────────────────────────────────────────────────

function wireControls() {
  $("#btn-pause").addEventListener("click", () => {
    state.paused = !state.paused;
    $("#btn-pause").classList.toggle("active", state.paused);
    $("#btn-pause").innerHTML = state.paused
      ? '<span class="btn-icon">▶</span> Resume'
      : '<span class="btn-icon">⏸</span> Pause';
  });

  $("#btn-clear").addEventListener("click", () => {
    $("#agent-feed").innerHTML = "";
    $("#agent-empty").classList.remove("hidden");
    $("#files-list").innerHTML = "";
    $("#files-empty").classList.remove("hidden");
    $("#build-list").innerHTML = "";
    $("#build-empty").classList.remove("hidden");
    state.filesSeen.clear();
    state.buildSteps.clear();
    currentAssistantItem = null;
    currentAssistantUuid = null;
  });

  $("#btn-follow").addEventListener("click", async () => {
    if (!state.preferences) state.preferences = {};
    state.preferences.followMode = !state.preferences.followMode;
    await savePreferences();
    renderFollowChip();
    toast(state.preferences.followMode ? "Follow mode ON" : "Follow mode OFF", "info");
  });

  document.addEventListener("keydown", (e) => {
    if (document.activeElement?.tagName === "INPUT") return;
    if (e.key === " ") {
      e.preventDefault();
      $("#btn-pause").click();
    }
  });
}

// ── Toasts ──────────────────────────────────────────────────────

function toast(msg, kind = "info", ttlMs = 2500) {
  const host = $("#toast-host");
  if (!host) return;
  const el = document.createElement("div");
  el.className = `toast ${kind}`;
  el.innerHTML = `<span class="toast-dot"></span><span>${escapeHtml(msg)}</span>`;
  host.appendChild(el);
  setTimeout(() => {
    el.style.animation = "toast-out 220ms cubic-bezier(0.4, 0, 0.9, 0.3) forwards";
    setTimeout(() => el.remove(), 260);
  }, ttlMs);
}

// ── Utilities ───────────────────────────────────────────────────

async function fetchJson(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

function shortTime(ms) {
  return new Date(ms).toLocaleTimeString("en-US", { hour12: false }).split(" ")[0];
}

function shortPath(p) {
  if (!p) return "";
  const parts = p.replace(/\\/g, "/").split("/");
  return parts.slice(-2).join("/");
}

function ideInitial(kind) {
  if (kind === "VSCode") return "VS";
  if (kind === "VisualStudio") return "VS";
  if (kind === "Rider") return "R";
  return "?";
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

function escapeAttr(s) {
  return escapeHtml(s);
}

// ── Boot ────────────────────────────────────────────────────────

boot();
