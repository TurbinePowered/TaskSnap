// TaskSnap-WebApp: Notizen und Bilder in OneDrive/TaskSnap/Eingang ablegen, Aufgaben aus tasks/ anzeigen.
// Nur der PC schreibt in tasks/ – auch „erledigt“ geht als Befehlsdatei (cmd_*.json) über den Eingang.

import { CLIENT_ID, DATA_PATH } from "./config.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Files.ReadWrite"];
const REDIRECT = new URL("./", location.href).href;
const DEMO = new URLSearchParams(location.search).has("demo");
const PRIORITIES = { 0: "", 1: "Hoch", 2: "Mittel", 3: "Niedrig" };
const STATUSES = ["offen", "erledigt"];
const MAX_SIDE = 2400; // größere Bilder werden verkleinert (OCR am PC bleibt gut lesbar)
const KEEP_BYTES = 3.5e6; // bis dahin wird das Original unverändert hochgeladen
const SIMPLE_UPLOAD_BYTES = 4e6; // Graph: darüber Upload-Session nötig
const CMD_NAME = /^cmd_.+?_(offen|erledigt)_([\w-]+)\.json$/;

const $ = (sel) => document.querySelector(sel);
const state = {
  tasks: [],
  attachments: new Map(), // Dateiname → { thumb, large, full }
  inbox: [], // Dateien im Eingang, die der PC noch nicht übernommen hat
  outbox: [], // noch nicht hochgeladen
  view: "offen",
  query: "",
  loadedAt: null,
  offline: false,
  loginNeeded: false,
};

class LoginNeeded extends Error {}
class Offline extends Error {}
class GraphError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

// ---------- Speicher im Browser (alles optional, App muss auch ohne funktionieren) ----------

const local = {
  get(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; }
  },
  set(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* voll oder gesperrt */ }
  },
};

// Warteschlange für Uploads: IndexedDB, damit Notizen ohne Netz nicht verloren gehen
const memoryOutbox = new Map();
const outboxDb = DEMO ? Promise.reject(new Error("demo")) : new Promise((resolve, reject) => {
  const req = indexedDB.open("tasksnap", 1);
  req.onupgradeneeded = () => req.result.createObjectStore("outbox", { keyPath: "name" });
  req.onsuccess = () => resolve(req.result);
  req.onerror = () => reject(req.error);
});
outboxDb.catch(() => {});

async function outbox(mode, fn) {
  try {
    const db = await outboxDb;
    return await new Promise((resolve, reject) => {
      const tx = db.transaction("outbox", mode);
      const req = fn(tx.objectStore("outbox"));
      tx.oncomplete = () => resolve(req.result);
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    return null;
  }
}
const outboxAll = async () => (await outbox("readonly", (s) => s.getAll())) ?? [...memoryOutbox.values()];
async function outboxPut(entry) {
  memoryOutbox.set(entry.name, entry);
  if ((await outbox("readwrite", (s) => s.put(entry))) !== null) memoryOutbox.delete(entry.name);
}
async function outboxDelete(name) {
  memoryOutbox.delete(name);
  await outbox("readwrite", (s) => s.delete(name));
}

// ---------- Anmeldung ----------

const pca = DEMO ? null : new msal.PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/consumers", // nur private Konten – nicht versehentlich das Firmenkonto
    redirectUri: REDIRECT,
    postLogoutRedirectUri: REDIRECT,
  },
  cache: { cacheLocation: "localStorage" }, // PWA auf dem Home-Bildschirm hat eigenen Speicher
});

let msalReady = false; // MSAL wirft, wenn es vor initialize() gefragt wird
const account = () => (msalReady ? pca.getActiveAccount() ?? pca.getAllAccounts()[0] ?? null : null);

async function accessToken() {
  const acc = account();
  if (!acc) throw new LoginNeeded();
  try {
    const result = await pca.acquireTokenSilent({
      scopes: SCOPES,
      account: acc,
      // Kein verstecktes iframe: das scheitert in Safari ohnehin (Tracking-Schutz) und kostet nur Zeit
      cacheLookupPolicy: msal.CacheLookupPolicy.AccessTokenAndRefreshToken,
    });
    return result.accessToken;
  } catch (e) {
    if (!navigator.onLine) throw new Offline();
    throw new LoginNeeded(e.message);
  }
}

function login() {
  const acc = account();
  if (acc) pca.acquireTokenRedirect({ scopes: SCOPES, account: acc, loginHint: acc.username });
  else pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
}

// ---------- Microsoft Graph ----------

const encodePath = (path) => path.split("/").map(encodeURIComponent).join("/");
const itemUrl = (relative) => `/me/drive/root:/${encodePath(`${DATA_PATH}/${relative}`)}`;

async function graph(url, options = {}) {
  const token = await accessToken();
  let res;
  try {
    res = await fetch(url.startsWith("https://") ? url : GRAPH + url, {
      ...options,
      headers: { Authorization: `Bearer ${token}`, ...options.headers },
    });
  } catch {
    throw new Offline();
  }
  if (res.status === 401) throw new LoginNeeded();
  if (!res.ok) {
    let message = res.statusText;
    try { message = (await res.json()).error?.message || message; } catch { /* kein JSON */ }
    throw new GraphError(res.status, `${res.status} ${message}`);
  }
  return res;
}

async function listFolder(relative, query = "") {
  const items = [];
  let url = `${itemUrl(relative)}:/children?$top=200${query}`;
  try {
    while (url) {
      const data = await (await graph(url)).json();
      items.push(...data.value.filter((i) => i.file));
      url = data["@odata.nextLink"];
    }
  } catch (e) {
    if (e instanceof GraphError && e.status === 404) return []; // Ordner gibt es (noch) nicht
    throw e;
  }
  return items;
}

async function uploadFile(name, blob, type) {
  const target = itemUrl(`Eingang/${name}`);
  if (blob.size <= SIMPLE_UPLOAD_BYTES) {
    await graph(`${target}:/content?@microsoft.graph.conflictBehavior=rename`, {
      method: "PUT", headers: { "Content-Type": type }, body: blob,
    });
    return;
  }
  const session = await (await graph(`${target}:/createUploadSession`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
  })).json();
  const chunk = 320 * 1024 * 16; // Vielfaches von 320 KiB
  for (let start = 0; start < blob.size; start += chunk) {
    const end = Math.min(start + chunk, blob.size);
    let res;
    try {
      res = await fetch(session.uploadUrl, {
        method: "PUT", headers: { "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end),
      });
    } catch {
      throw new Offline();
    }
    if (!res.ok) throw new GraphError(res.status, `Upload ${res.status}`);
  }
}

// Mehrere Downloads parallel, aber nicht alle auf einmal
async function pool(items, size, fn) {
  const queue = [...items];
  await Promise.all(Array.from({ length: size }, async () => {
    while (queue.length) await fn(queue.shift());
  }));
}

// Wie Task.from_file am PC: fehlende/abweichende Felder auffüllen
function normalizeTask(data, fileName) {
  let tags = data.tags ?? [];
  if (typeof tags === "string") tags = tags.split(",");
  tags = tags.map((t) => String(t).trim().replace(/^#/, "")).filter(Boolean);
  let priority = parseInt(data.priority, 10);
  if (!(priority in PRIORITIES)) priority = 0;
  let attachments = data.attachments ?? [];
  if (typeof attachments === "string") attachments = attachments ? [attachments] : [];
  return {
    id: fileName.replace(/\.json$/i, ""),
    title: String(data.title ?? "").trim(),
    text: String(data.text ?? "").trim(),
    tags,
    priority,
    status: STATUSES.includes(data.status) ? data.status : "offen",
    attachments,
    created: String(data.created ?? ""),
    source: String(data.source ?? ""),
  };
}

const thumbOf = (item, size) => item.thumbnails?.[0]?.[size]?.url ?? "";

let refreshing = null;
function refresh() {
  refreshing ??= doRefresh().finally(() => { refreshing = null; });
  return refreshing;
}

async function doRefresh() {
  if (DEMO) { render(); return; }
  setBusy(true);
  try {
    const [taskItems, attachmentItems, inboxItems] = await Promise.all([
      listFolder("tasks"),
      listFolder("attachments", "&$expand=thumbnails"),
      listFolder("Eingang"),
    ]);
    const cache = local.get("tasksnap.tasks", {});
    const next = {};
    await pool(taskItems.filter((i) => i.name.toLowerCase().endsWith(".json")), 6, async (item) => {
      const hit = cache[item.id];
      if (hit && hit.etag === item.eTag) { next[item.id] = hit; return; }
      try {
        const res = await fetch(item["@microsoft.graph.downloadUrl"]);
        if (!res.ok) throw new Error(res.status);
        next[item.id] = { etag: item.eTag, task: normalizeTask(await res.json(), item.name) };
      } catch {
        if (hit) next[item.id] = hit; // z.B. halb synchronisiert – alte Fassung behalten
      }
    });
    local.set("tasksnap.tasks", next);
    state.tasks = Object.values(next).map((e) => e.task);
    state.attachments = new Map(attachmentItems.map((i) => [i.name, {
      thumb: thumbOf(i, "medium"), large: thumbOf(i, "large"), full: i["@microsoft.graph.downloadUrl"],
    }]));
    state.inbox = inboxItems.map((i) => i.name);
    state.loadedAt = new Date();
    state.offline = false;
    state.loginNeeded = false;
    local.set("tasksnap.loadedAt", state.loadedAt.toISOString());
    // Beschriftungen hochgeladener Dateien nur so lange merken, wie sie im Eingang liegen
    const labels = local.get("tasksnap.labels", {});
    local.set("tasksnap.labels", Object.fromEntries(Object.entries(labels).filter(([n]) => state.inbox.includes(n))));
  } catch (e) {
    handleError(e);
  } finally {
    setBusy(false);
    render();
  }
}

function handleError(e) {
  if (e instanceof LoginNeeded) state.loginNeeded = true;
  else if (e instanceof Offline) state.offline = true;
  else toast(`Fehler: ${e.message}`);
}

// ---------- Hochladen ----------

function stamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  const rand = Math.random().toString(16).slice(2, 6);
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}_${rand}`;
}

async function enqueue(entry) {
  await outboxPut({ ...entry, queued: Date.now() });
  await loadOutbox();
  render();
  flush();
}

let flushing = null;
function flush() {
  flushing ??= doFlush().finally(() => { flushing = null; });
  return flushing;
}

async function doFlush() {
  if (DEMO) return;
  let uploaded = false;
  for (const entry of await outboxAll()) {
    try {
      await uploadFile(entry.name, entry.blob, entry.type);
      await outboxDelete(entry.name);
      const labels = local.get("tasksnap.labels", {});
      labels[entry.name] = entry.label;
      local.set("tasksnap.labels", labels);
      state.inbox.push(entry.name);
      uploaded = true;
    } catch (e) {
      handleError(e);
      break; // Reihenfolge einhalten, später erneut versuchen
    }
  }
  await loadOutbox();
  render();
  if (uploaded) setTimeout(refresh, 1500);
}

async function loadOutbox() {
  state.outbox = (await outboxAll()).map(({ name, label }) => ({ name, label }));
}

async function sendNote() {
  const note = $("#note");
  const text = note.value.trim();
  if (!text) return;
  await enqueue({
    name: `web_${stamp()}.txt`,
    blob: new Blob([text], { type: "text/plain" }),
    type: "text/plain; charset=utf-8",
    label: text.split("\n")[0].slice(0, 80),
  });
  note.value = "";
  local.set("tasksnap.draft", "");
  autosize();
  toast("Notiz gespeichert – wird am PC zur Aufgabe");
}

const IMAGE_EXT = { "image/png": "png", "image/jpeg": "jpg", "image/heic": "heic", "image/heif": "heif", "image/gif": "gif", "image/webp": "webp" };

async function prepareImage(file) {
  const type = file.type || "image/jpeg";
  const original = { blob: file, type, ext: IMAGE_EXT[type] ?? (file.name.split(".").pop() || "jpg").toLowerCase() };
  let bitmap;
  try { bitmap = await createImageBitmap(file); } catch { return original; } // z.B. HEIC – wandelt der PC um
  const side = Math.max(bitmap.width, bitmap.height);
  if ((type === "image/png" || type === "image/jpeg") && file.size <= KEEP_BYTES && side <= 3000) {
    bitmap.close();
    return original; // Screenshots unverändert lassen – beste Texterkennung
  }
  const scale = Math.min(1, MAX_SIDE / side);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.88));
  return blob ? { blob, type: "image/jpeg", ext: "jpg" } : original;
}

async function addImages(files) {
  const images = [...files].filter((f) => !f.type || f.type.startsWith("image/"));
  for (const file of images) {
    const { blob, type, ext } = await prepareImage(file);
    await enqueue({ name: `web_${stamp()}.${ext}`, blob, type, label: "Bild" });
  }
  if (images.length) toast(images.length > 1 ? `${images.length} Bilder gespeichert` : "Bild gespeichert – wird am PC zur Aufgabe");
}

async function pasteImage() {
  if (!navigator.clipboard?.read) {
    toast("Einfügen geht hier nicht – lange ins Notizfeld tippen → Einfügen");
    return;
  }
  try {
    for (const item of await navigator.clipboard.read()) {
      const type = item.types.find((t) => t.startsWith("image/"));
      if (type) {
        const blob = await item.getType(type);
        await addImages([new File([blob], "zwischenablage", { type })]);
        return;
      }
    }
    toast("Kein Bild in der Zwischenablage");
  } catch {
    toast("Zwischenablage nicht lesbar");
  }
}

function setStatus(task, status) {
  const previous = effectiveStatus(task);
  if (previous === status) return;
  enqueue({
    name: `cmd_${stamp()}_${status}_${task.id}.json`,
    blob: new Blob([JSON.stringify({ id: task.id, set: { status } })], { type: "application/json" }),
    type: "application/json",
    label: `${task.title}: ${status}`,
  });
  toast(status === "erledigt" ? "Erledigt" : "Wieder offen", "Rückgängig", () => setStatus(task, previous));
}

// ---------- Anzeige ----------

// Statusänderungen, die der PC noch nicht übernommen hat (jüngste gewinnt)
function pendingStatuses() {
  const map = new Map();
  const names = [...state.inbox, ...state.outbox.map((e) => e.name)].sort();
  for (const name of names) {
    const m = CMD_NAME.exec(name);
    if (m) map.set(m[2], m[1]);
  }
  return map;
}
let pending = new Map();
const effectiveStatus = (task) => pending.get(task.id) ?? task.status;

function h(tag, props = {}, ...children) {
  const el = document.createElement(tag);
  for (const [key, value] of Object.entries(props)) {
    if (key.startsWith("on")) el.addEventListener(key.slice(2), value);
    else if (value !== false && value != null) el.setAttribute(key, value === true ? "" : value);
  }
  el.append(...children.flat().filter((c) => c != null && c !== false));
  return el;
}

function formatDate(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return "";
  const sameYear = d.getFullYear() === new Date().getFullYear();
  return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: sameYear ? undefined : "2-digit" });
}

function metaLine(task) {
  return [
    task.priority ? h("span", { class: `prio-${task.priority}` }, PRIORITIES[task.priority]) : null,
    ...task.tags.map((t) => h("span", { class: "tag" }, `#${t}`)),
    formatDate(task.created),
    task.source === "iphone" ? "📱" : null,
  ].filter(Boolean).flatMap((part, i) => (i ? [" · ", part] : [part]));
}

const attachmentOf = (task) => state.attachments.get((task.attachments[0] ?? "").split("/").pop());

function taskRow(task) {
  const status = effectiveStatus(task);
  const waiting = pending.has(task.id) && pending.get(task.id) !== task.status;
  const att = attachmentOf(task);
  return h("li", { class: `task${status === "erledigt" ? " done" : ""}${waiting ? " waiting" : ""}` },
    h("button", {
      class: "check",
      "aria-label": status === "erledigt" ? "Wieder öffnen" : "Erledigt",
      title: waiting ? "Wird am PC übernommen" : "",
      onclick: () => setStatus(task, status === "erledigt" ? "offen" : "erledigt"),
    }, h("i", {}, status === "erledigt" ? "✓" : "")),
    h("div", { class: "body", onclick: () => showDetail(task) },
      h("div", { class: "title" }, task.title || "(ohne Titel)"),
      h("div", { class: "meta" }, metaLine(task))),
    att?.thumb ? h("img", { src: att.thumb, alt: "", loading: "lazy", onclick: () => showDetail(task) })
      : task.attachments.length ? h("div", { class: "thumb" }) : null);
}

function render() {
  pending = pendingStatuses();
  $("#login").hidden = DEMO || !state.loginNeeded;
  $("#login p").innerHTML = account()
    ? "Die Anmeldung ist abgelaufen. Notizen werden bis dahin auf dem Gerät gesammelt."
    : "Mit dem <b>privaten</b> Microsoft-Konto anmelden, in dessen OneDrive der Ordner <code>TaskSnap</code> liegt.";
  $("#login-btn").textContent = account() ? "Erneut anmelden" : "Anmelden";

  const loaded = state.loadedAt ?? (local.get("tasksnap.loadedAt", null) && new Date(local.get("tasksnap.loadedAt")));
  const time = loaded ? loaded.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }) : "";
  $("#stamp").textContent = [state.offline ? "offline" : "", time && `Stand ${time}`].filter(Boolean).join(" · ");

  renderPending();

  const query = state.query.toLowerCase();
  const open = state.tasks.filter((t) => effectiveStatus(t) === "offen");
  const done = state.tasks.filter((t) => effectiveStatus(t) === "erledigt");
  $("#count-open").textContent = open.length || "";
  $("#count-done").textContent = done.length || "";
  const shown = (state.view === "offen" ? open : done)
    .filter((t) => !query || [t.title, t.text, ...t.tags].some((s) => s.toLowerCase().includes(query)))
    .sort((a, b) => b.created.localeCompare(a.created));
  $("#list").replaceChildren(...shown.map(taskRow));
  const empty = $("#empty");
  empty.hidden = shown.length > 0;
  empty.textContent = !state.tasks.length && !state.loadedAt ? (DEMO ? "" : "Noch nichts geladen.")
    : query ? "Nichts gefunden." : state.view === "offen" ? "Keine offenen Aufgaben 🎉" : "Noch nichts erledigt.";
}

function renderPending() {
  const labels = local.get("tasksnap.labels", {});
  const notCmd = (name) => !name.startsWith("cmd_");
  const waitingUpload = state.outbox.filter((e) => notCmd(e.name));
  const waitingPc = state.inbox.filter(notCmd);
  const commands = state.outbox.length - waitingUpload.length + state.inbox.length - waitingPc.length;
  const box = $("#pending");
  box.hidden = !waitingUpload.length && !waitingPc.length && !commands;
  if (box.hidden) return;
  const describe = (name, label) => (name.endsWith(".txt") ? `📝 ${label || "Notiz"}` : `🖼 ${label || name}`);
  const lines = [];
  if (waitingUpload.length) lines.push(`📤 ${waitingUpload.length} noch nicht hochgeladen (keine Verbindung oder Anmeldung)`);
  if (waitingPc.length) lines.push(`⏳ ${waitingPc.length} im Eingang – wird zur Aufgabe, sobald TaskSnap am PC läuft`);
  if (commands) lines.push(`↻ ${commands} Statusänderung${commands > 1 ? "en" : ""} noch nicht am PC übernommen`);
  const items = [...waitingUpload.map((e) => describe(e.name, e.label)), ...waitingPc.map((n) => describe(n, labels[n]))];
  box.replaceChildren(
    ...lines.map((l) => h("div", {}, l)),
    items.length ? h("ul", {}, items.slice(0, 5).map((t) => h("li", {}, t)), items.length > 5 ? h("li", {}, "…") : null) : null);
}

function showDetail(task) {
  const status = effectiveStatus(task);
  const att = attachmentOf(task);
  const text = task.text && task.text !== task.title ? task.text : "";
  $("#detail-body").replaceChildren(
    h("h2", {}, task.title || "(ohne Titel)"),
    h("div", { class: "meta" }, metaLine(task)),
    att ? h("a", { class: "full", href: att.full, target: "_blank", rel: "noopener" },
      h("img", { src: att.large || att.thumb, alt: "Anhang" })) : null,
    text ? h("pre", {}, text) : null,
    h("button", {
      class: "primary",
      onclick: () => { setStatus(task, status === "erledigt" ? "offen" : "erledigt"); $("#detail").close(); },
    }, status === "erledigt" ? "Wieder öffnen" : "Erledigt ✓"));
  $("#detail").showModal();
}

let toastTimer;
function toast(message, actionLabel, action) {
  const el = $("#toast");
  el.replaceChildren(h("span", {}, message),
    actionLabel ? h("button", { onclick: () => { el.hidden = true; action(); } }, actionLabel) : null);
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, actionLabel ? 5000 : 3000);
}

function setBusy(busy) {
  $("#refresh").classList.toggle("spin", busy);
}

function autosize() {
  const note = $("#note");
  note.style.height = "auto";
  note.style.height = `${note.scrollHeight + 2}px`;
}

// ---------- Start ----------

function bindUi() {
  const note = $("#note");
  note.value = local.get("tasksnap.draft", "");
  autosize();
  note.addEventListener("input", () => { autosize(); local.set("tasksnap.draft", note.value); });
  note.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); sendNote(); }
  });
  note.addEventListener("paste", (e) => {
    const files = [...(e.clipboardData?.files ?? [])].filter((f) => f.type.startsWith("image/"));
    if (files.length) { e.preventDefault(); addImages(files); }
  });
  $("#send").addEventListener("click", sendNote);
  $("#file").addEventListener("change", (e) => { addImages(e.target.files); e.target.value = ""; });
  $("#paste").addEventListener("click", pasteImage);
  $("#refresh").addEventListener("click", () => { flush(); refresh(); });
  $("#login-btn").addEventListener("click", login);
  $("#account").addEventListener("click", () => {
    if (DEMO) { toast("Demo-Modus – ohne OneDrive"); return; }
    const acc = account();
    if (!acc) { login(); return; }
    if (confirm(`Angemeldet als ${acc.username}.\n\nAbmelden?`)) pca.logoutRedirect({ account: acc });
  });
  $("#search").addEventListener("input", (e) => { state.query = e.target.value.trim(); render(); });
  for (const btn of document.querySelectorAll(".seg button")) {
    btn.addEventListener("click", () => {
      state.view = btn.dataset.view;
      document.querySelectorAll(".seg button").forEach((b) => b.classList.toggle("on", b === btn));
      render();
    });
  }
  const detail = $("#detail");
  detail.addEventListener("click", (e) => { if (e.target === detail) detail.close(); });

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && (!state.loadedAt || Date.now() - state.loadedAt > 30_000)) {
      flush();
      refresh();
    }
  });
  window.addEventListener("online", () => { state.offline = false; flush(); refresh(); });
}

function loadDemo() {
  const day = (n) => new Date(Date.now() - n * 864e5).toISOString();
  state.tasks = [
    { id: "d1", title: "Angebot für Müller bis Freitag prüfen", text: "Angebot für Müller bis Freitag prüfen\nPositionen 3–7 mit Vertrieb klären.", tags: ["kunde"], priority: 1, status: "offen", attachments: [], created: day(0.1), source: "capture" },
    { id: "d2", title: "Release 1.0.4 testen", text: "", tags: [], priority: 2, status: "offen", attachments: [], created: day(1), source: "manual" },
    { id: "d3", title: "Foto Whiteboard Sprint-Planung", text: "Sprint 42\n- Login\n- Export", tags: ["team", "sprint"], priority: 0, status: "offen", attachments: ["attachments/d3.png"], created: day(3), source: "iphone" },
    { id: "d4", title: "Reisekosten einreichen", text: "", tags: [], priority: 3, status: "erledigt", attachments: [], created: day(9), source: "manual" },
  ];
  state.inbox = ["web_2026-10-12_090000_abcd.txt"];
  local.set("tasksnap.labels", { "web_2026-10-12_090000_abcd.txt": "Rückruf Frau Schmidt" });
  state.loadedAt = new Date();
}

async function start() {
  bindUi();
  if ("serviceWorker" in navigator && !DEMO) navigator.serviceWorker.register("sw.js").catch(() => {});
  if (DEMO) { loadDemo(); render(); return; }

  // Zuletzt geladene Aufgaben sofort zeigen (auch offline), dann aktualisieren
  state.tasks = Object.values(local.get("tasksnap.tasks", {})).map((e) => e.task);
  await loadOutbox();
  render();

  await pca.initialize();
  msalReady = true;
  try {
    const result = await pca.handleRedirectPromise();
    if (result?.account) pca.setActiveAccount(result.account);
  } catch (e) {
    toast(`Anmeldung fehlgeschlagen: ${e.message}`);
  }
  if (!account()) {
    state.loginNeeded = true;
    render();
    return;
  }
  pca.setActiveAccount(account());
  await flush();
  await refresh();
}

start().catch((e) => toast(`Startfehler: ${e.message}`));
