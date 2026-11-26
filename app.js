// Nube BI: interfaz, almacenamiento local, Google Drive y auto-actualización.
(() => {
  "use strict";

  const CFG = window.NUBE_CONFIG || {};
  const HAS_DRIVE = Boolean(CFG.GOOGLE_CLIENT_ID);
  const SCOPES = "https://www.googleapis.com/auth/drive.file";
  const API = "https://www.googleapis.com/drive/v3";
  const UPLOAD = "https://www.googleapis.com/upload/drive/v3";
  const APP_KEY = "nubebi";
  const CAN_WATCH = "showOpenFilePicker" in window;

  const $ = (sel, el = document) => el.querySelector(sel);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = new Intl.NumberFormat("es", { maximumFractionDigits: 2 });
  const fmtCompact = new Intl.NumberFormat("es", { notation: "compact", maximumFractionDigits: 1 });

  const state = {
    datasets: new Map(),   // id -> registro
    handles: new Map(),    // id -> FileSystemFileHandle (Excel vigilados)
    selected: localStorage.getItem("nube.selected"),
    sheetIdx: 0,
    filter: "",
    token: null,
    tokenExp: 0,
    tokenClient: null,
    syncing: false,
    lastSync: 0,
    syncError: null,
    busy: 0,
    charts: new Map(),
    structureKey: "",
    needsPermission: new Set(),
    installPrompt: null,
  };

  // ---------- IndexedDB ----------
  const idb = (() => {
    let dbp;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const r = indexedDB.open("nubebi", 1);
      r.onupgradeneeded = () => {
        const db = r.result;
        db.createObjectStore("datasets", { keyPath: "id" });
        db.createObjectStore("blobs");
        db.createObjectStore("handles");
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }));
    const tx = async (store, mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction(store, mode);
        const req = fn(t.objectStore(store));
        t.oncomplete = () => res(req && req.result);
        t.onerror = () => rej(t.error);
      });
    };
    return {
      all: (s) => tx(s, "readonly", (st) => st.getAll()),
      keys: (s) => tx(s, "readonly", (st) => st.getAllKeys()),
      get: (s, k) => tx(s, "readonly", (st) => st.get(k)),
      put: (s, v, k) => tx(s, "readwrite", (st) => (k === undefined ? st.put(v) : st.put(v, k))),
      del: (s, k) => tx(s, "readwrite", (st) => st.delete(k)),
    };
  })();

  // ---------- Worker de análisis ----------
  const worker = new Worker("worker.js");
  let reqSeq = 0;
  const pending = new Map();
  worker.onmessage = ({ data }) => {
    const p = pending.get(data.id);
    if (!p) return;
    pending.delete(data.id);
    data.ok ? p.resolve(data.result) : p.reject(new Error(data.error));
  };
  const runWorker = (msg, transfer = []) => new Promise((resolve, reject) => {
    const id = ++reqSeq;
    pending.set(id, { resolve, reject });
    worker.postMessage({ id, ...msg }, transfer);
  });
  const analyze = (buffer, name) => runWorker({ buffer: buffer.slice(0), name });
  const combine = (parts, name) => runWorker({ op: "combine", parts, name });

  // ---------- Avisos y estado ----------
  function toast(msg, kind = "") {
    const el = document.createElement("div");
    el.className = `toast ${kind}`;
    el.textContent = msg;
    $("#toasts").appendChild(el);
    setTimeout(() => el.remove(), kind === "err" ? 7000 : 3500);
  }

  function ago(t) {
    if (!t) return "";
    const s = Math.round((Date.now() - t) / 1000);
    if (s < 10) return "ahora";
    if (s < 60) return `hace ${s} s`;
    const m = Math.round(s / 60);
    if (m < 60) return `hace ${m} min`;
    const h = Math.round(m / 60);
    if (h < 24) return `hace ${h} h`;
    return new Date(t).toLocaleDateString("es");
  }

  function renderStatus() {
    const el = $("#status");
    const text = $(".status-text", el);
    el.className = "status";
    el.onclick = null;
    if (!navigator.onLine) { el.classList.add("warn"); text.textContent = "Sin conexión · usando copia guardada"; return; }
    if (state.busy || state.syncing) { el.classList.add("sync"); text.textContent = state.busy ? "Analizando…" : "Sincronizando…"; return; }
    if (!HAS_DRIVE) { text.textContent = "Modo local"; return; }
    if (!state.token) {
      if (localStorage.getItem("nube.connected")) {
        el.classList.add("warn", "clickable"); text.textContent = "Toca para reconectar Drive";
        el.onclick = () => signIn(true);
      } else text.textContent = "Drive sin conectar";
      return;
    }
    if (state.syncError) { el.classList.add("err"); text.textContent = state.syncError; return; }
    el.classList.add("ok");
    text.textContent = `Drive al día · ${ago(state.lastSync)}`;
  }

  const setBusy = (delta) => { state.busy = Math.max(0, state.busy + delta); renderStatus(); };

  // ---------- Datos ----------
  async function loadLocal() {
    const [all, handleKeys] = await Promise.all([idb.all("datasets"), idb.keys("handles")]);
    all.forEach((d) => state.datasets.set(d.id, d));
    for (const k of handleKeys) {
      const h = await idb.get("handles", k);
      if (h && state.datasets.has(k)) state.handles.set(k, h);
    }
  }

  async function saveDataset(rec, blob) {
    state.datasets.set(rec.id, rec);
    await idb.put("datasets", rec);
    if (blob) await idb.put("blobs", blob, rec.id);
  }

  async function removeDataset(id) {
    state.datasets.delete(id);
    state.handles.delete(id);
    await Promise.all([idb.del("datasets", id), idb.del("blobs", id), idb.del("handles", id)]);
    if (state.selected === id) select(null);
  }

  const newId = () => "l_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const folderOf = (d) => `${d.category || "General"} / ${d.year || new Date().getFullYear()}`;

  async function ingestFile(file, handle) {
    setBusy(1);
    try {
      const buffer = await file.arrayBuffer();
      const analysis = await analyze(buffer, file.name);
      // Si ya existe uno con el mismo nombre, se actualiza en vez de duplicarse.
      const existing = [...state.datasets.values()].find((d) => d.name === file.name && d.source !== "combined");
      const rec = {
        ...(existing || { id: newId(), createdAt: Date.now() }),
        name: file.name,
        source: existing?.source || "local",
        category: analysis.category,
        year: analysis.year,
        analysis,
        localModified: file.lastModified,
        updatedAt: Date.now(),
      };
      await saveDataset(rec, new Blob([buffer], { type: file.type }));
      if (handle) { state.handles.set(rec.id, handle); await idb.put("handles", handle, rec.id); }
      toast(existing ? `"${file.name}" actualizado` : `"${file.name}" listo · guardado en ${folderOf(rec)}`);
      if (state.token) await pushToDrive(rec, new Blob([buffer], { type: file.type }));
      await refreshCombined();
      return rec;
    } catch (err) {
      toast(`No pude leer "${file.name}": ${err.message}`, "err");
      return null;
    } finally {
      setBusy(-1);
    }
  }

  async function handleFiles(files, handles = []) {
    let last = null;
    for (let i = 0; i < files.length; i++) {
      const r = await ingestFile(files[i], handles[i]);
      if (r) last = r;
    }
    if (last) select(last.id);
    render();
  }

  // ---------- Archivos combinados ----------
  async function createCombined(ids) {
    const sources = ids.map((id) => state.datasets.get(id)).filter(Boolean);
    const base = sources[0].name.replace(/\.[^.]+$/, "");
    const rec = {
      id: "c_" + Date.now().toString(36),
      name: `Combinado: ${base} + ${sources.length - 1} más`,
      source: "combined",
      sources: ids,
      updatedAt: 0,
      createdAt: Date.now(),
    };
    state.datasets.set(rec.id, rec);
    await refreshCombined(rec.id);
    select(rec.id);
  }

  async function refreshCombined(onlyId) {
    for (const rec of [...state.datasets.values()]) {
      if (rec.source !== "combined" || (onlyId && rec.id !== onlyId)) continue;
      const srcs = rec.sources.map((id) => state.datasets.get(id)).filter(Boolean);
      if (srcs.length < 2) { await removeDataset(rec.id); continue; }
      const newest = Math.max(...srcs.map((s) => s.updatedAt || 0));
      if (rec.analysis && newest <= rec.updatedAt) continue;
      const parts = [];
      for (const s of srcs) {
        const blob = await idb.get("blobs", s.id);
        if (blob) parts.push({ buffer: await blob.arrayBuffer(), name: s.name });
      }
      if (parts.length < 2) continue;
      setBusy(1);
      try {
        const analysis = await combine(parts, rec.name);
        await saveDataset({ ...rec, analysis, category: analysis.category, year: analysis.year, updatedAt: Date.now() });
        if (rec.analysis) flashFile(rec.id);
      } catch (err) {
        toast(`No pude combinar: ${err.message}`, "err");
      } finally {
        setBusy(-1);
      }
    }
  }

  // ---------- Google: inicio de sesión ----------
  function initGoogle() {
    if (!HAS_DRIVE) {
      $("#btnAuth").title = "Agrega tu ID de cliente de Google en config.js para activar Drive (ver GUIA.md)";
      $("#btnAuth").onclick = () => toast("Para usar Drive, pega tu ID de cliente de Google en config.js (ver GUIA.md).");
      return;
    }
    const ready = () => {
      state.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CFG.GOOGLE_CLIENT_ID,
        scope: SCOPES,
        callback: onToken,
        error_callback: (e) => {
          if (e && e.type !== "popup_closed") toast("No se pudo conectar con Google. Intenta de nuevo.", "err");
          renderStatus();
        },
      });
      $("#btnAuth").onclick = () => (state.token ? signOut() : signIn(false));
      renderAuth();
    };
    if (window.google?.accounts?.oauth2) ready();
    else {
      const t = setInterval(() => { if (window.google?.accounts?.oauth2) { clearInterval(t); ready(); } }, 150);
    }
  }

  function signIn(silent) {
    if (!state.tokenClient) return;
    state.tokenClient.requestAccessToken({ prompt: silent || localStorage.getItem("nube.connected") ? "" : "consent" });
  }

  async function onToken(resp) {
    if (resp.error) { toast("Google rechazó el acceso.", "err"); return; }
    state.token = resp.access_token;
    state.tokenExp = Date.now() + (resp.expires_in - 60) * 1000;
    state.syncError = null;
    localStorage.setItem("nube.connected", "1");
    renderAuth();
    loadAbout();
    await uploadPendingLocals();
    syncDrive();
  }

  function signOut() {
    if (state.token) google.accounts.oauth2.revoke(state.token, () => {});
    state.token = null;
    localStorage.removeItem("nube.connected");
    renderAuth();
    renderStatus();
    toast("Drive desconectado. Tus datos guardados siguen en este dispositivo.");
  }

  // El token de Google dura 1 hora. Lo renovamos en el siguiente toque del usuario
  // (los navegadores solo permiten la ventanita de Google tras un gesto).
  document.addEventListener("pointerdown", () => {
    if (!HAS_DRIVE || !state.tokenClient || !localStorage.getItem("nube.connected")) return;
    if (!state.token || Date.now() > state.tokenExp - 5 * 60e3) {
      if (Date.now() - (state.lastRenew || 0) < 30e3) return;
      state.lastRenew = Date.now();
      signIn(true);
    }
  }, true);

  function renderAuth() {
    const btn = $("#btnAuth");
    btn.textContent = state.token ? "Desconectar" : "Conectar Google Drive";
    btn.classList.toggle("primary", !state.token);
    btn.classList.toggle("ghost", Boolean(state.token));
    $("#user").classList.toggle("hidden", !state.token);
    $("#storage").classList.toggle("hidden", !state.token);
  }

  // ---------- Google Drive ----------
  class AuthError extends Error {}

  async function api(url, opts = {}) {
    if (!state.token || Date.now() > state.tokenExp) { state.token = null; renderAuth(); throw new AuthError("Sesión vencida"); }
    const res = await fetch(url, { ...opts, headers: { Authorization: `Bearer ${state.token}`, ...(opts.headers || {}) } });
    if (res.status === 401) { state.token = null; renderAuth(); throw new AuthError("Sesión vencida"); }
    if (!res.ok) {
      let msg = `Drive respondió ${res.status}`;
      try { msg = (await res.json()).error.message || msg; } catch {}
      throw new Error(msg);
    }
    return res;
  }

  async function loadAbout() {
    try {
      const a = await (await api(`${API}/about?fields=user(displayName,photoLink),storageQuota`)).json();
      const u = $("#user");
      u.innerHTML = `${a.user.photoLink ? `<img src="${esc(a.user.photoLink)}" alt="" referrerpolicy="no-referrer">` : ""}<span>${esc(a.user.displayName)}</span>`;
      const q = a.storageQuota;
      if (q && q.limit) {
        const used = +q.usage, lim = +q.limit;
        const gb = (b) => (b >= 1e12 ? `${fmt.format(b / 1e12)} TB` : `${fmt.format(b / 1e9)} GB`);
        $("#storage").innerHTML = `Drive: ${gb(used)} de ${gb(lim)} usados<div class="bar"><span style="width:${Math.min(100, (used / lim) * 100).toFixed(1)}%"></span></div>`;
      }
    } catch {}
  }

  const folderCache = JSON.parse(localStorage.getItem("nube.folders") || "{}");
  async function ensureFolder(name, parentId) {
    const key = `${parentId}/${name}`;
    if (folderCache[key]) return folderCache[key];
    const q = `name='${name.replace(/'/g, "\\'")}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`;
    const found = await (await api(`${API}/files?q=${encodeURIComponent(q)}&fields=files(id)&spaces=drive`)).json();
    let id = found.files[0]?.id;
    if (!id) {
      const created = await (await api(`${API}/files?fields=id`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, mimeType: "application/vnd.google-apps.folder", parents: [parentId] }),
      })).json();
      id = created.id;
    }
    folderCache[key] = id;
    localStorage.setItem("nube.folders", JSON.stringify(folderCache));
    return id;
  }

  async function pushToDrive(rec, blob) {
    try {
      const fields = "id,name,modifiedTime,md5Checksum,appProperties";
      let initUrl, method, meta;
      if (!rec.driveId) {
        // Si en Drive ya hay un archivo de la app con ese nombre (subido desde otro
        // dispositivo), se actualiza ese en vez de crear un duplicado.
        const q = `name='${rec.name.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}' and appProperties has { key='${APP_KEY}' and value='1' } and trashed=false`;
        const found = await (await api(`${API}/files?q=${encodeURIComponent(q)}&fields=files(id)&spaces=drive`)).json();
        if (found.files[0]) rec = { ...rec, driveId: found.files[0].id };
      }
      if (rec.driveId) {
        initUrl = `${UPLOAD}/files/${rec.driveId}?uploadType=resumable&fields=${fields}`;
        method = "PATCH";
        meta = { appProperties: { cat: rec.category, year: String(rec.year) } };
      } else {
        const root = await ensureFolder(CFG.ROOT_FOLDER || "Nube BI", "root");
        const cat = await ensureFolder(rec.category, root);
        const year = await ensureFolder(String(rec.year), cat);
        initUrl = `${UPLOAD}/files?uploadType=resumable&fields=${fields}`;
        method = "POST";
        meta = { name: rec.name, parents: [year], appProperties: { [APP_KEY]: "1", cat: rec.category, year: String(rec.year) } };
      }
      const init = await api(initUrl, {
        method,
        headers: { "Content-Type": "application/json; charset=UTF-8", "X-Upload-Content-Type": blob.type || "application/octet-stream" },
        body: JSON.stringify(meta),
      });
      const loc = init.headers.get("Location");
      const res = await fetch(loc, { method: "PUT", body: blob });
      if (!res.ok) throw new Error(`Subida falló (${res.status})`);
      const f = await res.json();
      // Si sync ya había traído ese mismo archivo como registro aparte, se quita el duplicado.
      for (const other of [...state.datasets.values()]) {
        if (other.id !== rec.id && other.driveId === f.id) await removeDataset(other.id);
      }
      const updated = { ...state.datasets.get(rec.id), source: "drive", driveId: f.id, md5: f.md5Checksum, modifiedTime: f.modifiedTime };
      await saveDataset(updated);
      render();
    } catch (err) {
      if (!(err instanceof AuthError)) toast(`No pude subir "${rec.name}" a Drive: ${err.message}`, "err");
    }
  }

  async function uploadPendingLocals() {
    const locals = [...state.datasets.values()].filter((d) => d.source === "local");
    for (const d of locals) {
      const blob = await idb.get("blobs", d.id);
      if (blob) await pushToDrive(d, blob);
    }
  }

  async function listAppFiles() {
    const q = `appProperties has { key='${APP_KEY}' and value='1' } and trashed=false`;
    const out = [];
    let pageToken = "";
    do {
      const url = `${API}/files?q=${encodeURIComponent(q)}&spaces=drive&pageSize=1000&fields=nextPageToken,files(id,name,modifiedTime,md5Checksum,appProperties,mimeType)${pageToken ? `&pageToken=${pageToken}` : ""}`;
      const page = await (await api(url)).json();
      out.push(...page.files);
      pageToken = page.nextPageToken || "";
    } while (pageToken);
    return out;
  }

  // Revisa Drive y trae lo que cambió (desde otro dispositivo, o editado en Drive).
  async function syncDrive() {
    if (!state.token || state.syncing || state.busy || !navigator.onLine) { renderStatus(); return; }
    state.syncing = true;
    renderStatus();
    let changed = false;
    try {
      const files = await listAppFiles();
      const known = new Map([...state.datasets.values()].filter((d) => d.driveId).map((d) => [d.driveId, d]));
      for (const f of files) {
        const d = known.get(f.id);
        known.delete(f.id);
        if (d && d.md5 && d.md5 === f.md5Checksum) continue;
        const buf = await (await api(`${API}/files/${f.id}?alt=media`)).arrayBuffer();
        const analysis = await analyze(buf, f.name);
        await saveDataset({
          ...(d || { id: "d_" + f.id, createdAt: Date.now() }),
          name: f.name,
          source: "drive",
          driveId: f.id,
          md5: f.md5Checksum,
          modifiedTime: f.modifiedTime,
          category: f.appProperties?.cat || analysis.category,
          year: +(f.appProperties?.year || analysis.year),
          analysis,
          updatedAt: Date.now(),
        }, new Blob([buf]));
        changed = true;
        if (d) { flashFile(d.id); toast(`"${f.name}" se actualizó desde Drive`); }
      }
      // Lo que se borró en Drive desaparece aquí también.
      for (const d of known.values()) { await removeDataset(d.id); changed = true; }
      state.lastSync = Date.now();
      state.syncError = null;
      if (changed) await refreshCombined();
    } catch (err) {
      if (!(err instanceof AuthError)) state.syncError = "Error al sincronizar";
    } finally {
      state.syncing = false;
      if (changed) render(); else renderStatus();
    }
  }

  async function trashInDrive(rec) {
    try {
      await api(`${API}/files/${rec.driveId}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ trashed: true }) });
    } catch (err) {
      toast(`No pude moverlo a la papelera de Drive: ${err.message}`, "err");
      return false;
    }
    return true;
  }

  // ---------- Excel vigilados en la PC ----------
  async function pickAndWatch() {
    try {
      const handles = await window.showOpenFilePicker({
        multiple: true,
        types: [{ description: "Hojas de cálculo", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx", ".xlsm"], "application/vnd.ms-excel": [".xls"], "text/csv": [".csv"] } }],
      });
      const files = await Promise.all(handles.map((h) => h.getFile()));
      await handleFiles(files, handles);
      toast("Vigilando: cada vez que guardes el Excel, el dashboard se actualiza solo.");
    } catch (err) {
      if (err.name !== "AbortError") toast(err.message, "err");
    }
  }

  async function checkWatched() {
    for (const [id, handle] of state.handles) {
      const rec = state.datasets.get(id);
      if (!rec) continue;
      try {
        const perm = await handle.queryPermission({ mode: "read" });
        if (perm !== "granted") { state.needsPermission.add(id); continue; }
        state.needsPermission.delete(id);
        const file = await handle.getFile();
        if (file.lastModified !== rec.localModified) {
          await ingestFile(file);
          flashFile(id);
          render();
        }
      } catch {
        // El archivo puede estar bloqueado mientras Excel guarda; se reintenta en la próxima vuelta.
      }
    }
  }

  async function reactivateWatch(id) {
    const h = state.handles.get(id);
    if (!h) return;
    const perm = await h.requestPermission({ mode: "read" });
    if (perm === "granted") { state.needsPermission.delete(id); await checkWatched(); render(); }
  }

  // ---------- Interfaz ----------
  function select(id) {
    state.selected = id;
    state.sheetIdx = 0;
    if (id) localStorage.setItem("nube.selected", id); else localStorage.removeItem("nube.selected");
    document.body.classList.remove("show-list");
    render();
    $("#main").scrollTop = 0;
  }

  function flashFile(id) {
    requestAnimationFrame(() => {
      const el = document.querySelector(`.file[data-id="${id}"]`);
      if (!el) return;
      el.classList.remove("flash");
      void el.offsetWidth;
      el.classList.add("flash");
    });
  }

  function render() {
    renderList();
    renderDashboard();
    renderStatus();
  }

  function renderList() {
    const list = $("#fileList");
    const q = state.filter.trim().toLowerCase();
    const items = [...state.datasets.values()].filter((d) => d.analysis && (!q || d.name.toLowerCase().includes(q)));
    if (!items.length) {
      list.innerHTML = `<div class="empty-list">${q ? "Nada coincide con la búsqueda." : "Aún no hay archivos. Sube un Excel para empezar."}</div>`;
      return;
    }
    const groups = new Map();
    for (const d of items) {
      const key = d.source === "combined" ? "Combinados" : folderOf(d);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(d);
    }
    const keys = [...groups.keys()].sort((a, b) => (a === "Combinados") - (b === "Combinados") || a.localeCompare(b, "es"));
    list.innerHTML = keys.map((k) => `
      <div class="folder">
        <div class="folder-name"><svg width="13" height="13" viewBox="0 0 24 24"><path d="M3 6a2 2 0 012-2h4l2 2h8a2 2 0 012 2v10a2 2 0 01-2 2H5a2 2 0 01-2-2z" fill="currentColor" opacity=".5"/></svg>${esc(k)}</div>
        ${groups.get(k).sort((a, b) => b.updatedAt - a.updatedAt).map((d) => `
          <button class="file ${d.id === state.selected ? "active" : ""}" data-id="${d.id}">
            <span class="file-title"><span style="overflow:hidden;text-overflow:ellipsis">${esc(d.name)}</span>${badge(d)}</span>
            <span class="file-meta">${fmt.format(d.analysis.sheets[0].rowCount)} filas · ${ago(d.updatedAt)}</span>
          </button>`).join("")}
      </div>`).join("");
  }

  function badge(d) {
    if (d.source === "combined") return `<span class="badge combo">combinado</span>`;
    const parts = [];
    if (state.handles.has(d.id)) parts.push(state.needsPermission.has(d.id) ? `<span class="badge">pausado</span>` : `<span class="badge watch">vigilando</span>`);
    parts.push(d.source === "drive" ? `<span class="badge cloud">Drive</span>` : `<span class="badge">local</span>`);
    return parts.join("");
  }

  function renderWelcome() {
    destroyCharts();
    state.structureKey = "";
    $("#main").innerHTML = `
      <section class="welcome">
        <h1>Sube un Excel y listo</h1>
        <p>La app detecta qué tiene cada columna, limpia los datos, te recomienda los gráficos y ordena el archivo en tu Drive.</p>
        <div class="drop" id="dropBig" role="button" tabindex="0">
          <strong>Toca para elegir un archivo</strong>
          <span style="color:var(--muted);font-size:14px">o arrástralo aquí · .xlsx, .xls, .csv</span>
        </div>
        <div class="steps">
          <div class="step"><b>1 · Análisis automático</b>Fechas, montos y categorías se detectan solos.</div>
          <div class="step"><b>2 · Gráficos sugeridos</b>Tendencias, comparaciones y participación.</div>
          <div class="step"><b>3 · Siempre al día</b>Si el archivo cambia en Drive o en tu PC, el dashboard se actualiza.</div>
        </div>
      </section>`;
    const big = $("#dropBig");
    big.onclick = () => $("#fileInput").click();
    big.onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") $("#fileInput").click(); };
  }

  function renderDashboard() {
    const d = state.datasets.get(state.selected);
    if (!d || !d.analysis) { renderWelcome(); return; }
    const a = d.analysis;
    const sheet = a.sheets[Math.min(state.sheetIdx, a.sheets.length - 1)];
    const key = [d.id, state.sheetIdx, a.sheets.length, sheet.charts.map((c) => c.type + c.title).join(","), sheet.columns.length].join("|");

    if (key === state.structureKey) { updateDashboard(d, sheet); return; }
    state.structureKey = key;
    destroyCharts();

    const crumbs = d.source === "combined"
      ? `Combina: ${d.sources.map((id) => esc(state.datasets.get(id)?.name || "?")).join(", ")}`
      : `${esc(CFG.ROOT_FOLDER || "Nube BI")} / ${esc(d.category)} / ${esc(d.year)} · ${d.source === "drive" ? "en Google Drive" : "solo en este dispositivo"}`;

    $("#main").innerHTML = `
      <div class="dash-head">
        <div class="dash-title">
          <h1>${esc(d.name)}</h1>
          <div class="crumbs">${crumbs} · <span id="updatedAgo">actualizado ${ago(d.updatedAt)}</span></div>
        </div>
        <div class="dash-actions">
          ${state.needsPermission.has(d.id) ? `<button class="btn small primary" id="btnReWatch">Reactivar vigilancia</button>` : ""}
          <button class="btn small danger" id="btnDelete">Quitar</button>
        </div>
      </div>
      ${a.sheets.length > 1 ? `<div class="tabs">${a.sheets.map((s, i) => `<button class="tab ${i === state.sheetIdx ? "active" : ""}" data-sheet="${i}">${esc(s.sheet)}</button>`).join("")}</div>` : ""}
      <section class="kpis">${sheet.kpis.map((k, i) => `
        <div class="kpi"><div class="kpi-label" title="${esc(k.label)}">${esc(k.label)}</div><div class="kpi-value" data-kpi="${i}" data-v="${k.value}">${formatKpi(k)}</div></div>`).join("")}
      </section>
      <section class="grid">${sheet.charts.map((c, i) => `
        <article class="card ${c.type === "line" ? "wide" : ""}">
          <h3>${esc(c.title)}</h3>
          <p class="why">${esc(c.why)}</p>
          <div class="chart-box"><canvas data-chart="${i}" aria-label="${esc(c.title)}"></canvas></div>
        </article>`).join("")}
        ${sheet.charts.length === 0 ? `<article class="card wide"><h3>Sin gráficos sugeridos</h3><p class="why">No encontré columnas numéricas ni categorías claras. Revisa que la primera fila tenga los títulos de las columnas.</p></article>` : ""}
      </section>
      <section class="recos">
        <article class="card">
          <h3>Recomendaciones</h3>
          <ul class="reco-list">${sheet.notes.length ? sheet.notes.map((n) => `<li>${esc(n)}</li>`).join("") : "<li>Los datos se ven limpios.</li>"}</ul>
          ${relatedHtml(d)}
        </article>
        <article class="card">
          <h3>Columnas detectadas</h3>
          <div class="cols">${sheet.columns.map((c) => `<span class="col-chip">${esc(c.name)}<i>${typeLabel(c.type)}</i></span>`).join("")}</div>
        </article>
        <article class="card wide">
          <h3>Vista previa</h3>
          <p class="why">Primeras ${sheet.preview.rows.length} de ${fmt.format(sheet.rowCount)} filas.</p>
          <div class="table-wrap" id="previewTable">${tableHtml(sheet)}</div>
        </article>
      </section>`;

    $("#main").querySelectorAll("[data-sheet]").forEach((b) => (b.onclick = () => { state.sheetIdx = +b.dataset.sheet; renderDashboard(); }));
    $("#btnDelete").onclick = () => confirmDelete(d);
    const rw = $("#btnReWatch");
    if (rw) rw.onclick = () => reactivateWatch(d.id);
    $("#main").querySelectorAll("[data-open]").forEach((b) => (b.onclick = () => select(b.dataset.open)));
    $("#main").querySelectorAll("[data-combine]").forEach((b) => (b.onclick = () => createCombined(b.dataset.combine.split(","))));

    sheet.charts.forEach((c, i) => {
      const canvas = $(`canvas[data-chart="${i}"]`);
      state.charts.set(i, new Chart(canvas, chartConfig(c)));
    });
  }

  // Actualización suave: no rearma la página, solo mueve números y gráficos.
  function updateDashboard(d, sheet) {
    const ua = $("#updatedAgo");
    if (ua) ua.textContent = `actualizado ${ago(d.updatedAt)}`;
    sheet.kpis.forEach((k, i) => {
      const el = $(`[data-kpi="${i}"]`);
      if (el && el.dataset.v !== String(k.value)) animateNumber(el, k);
    });
    sheet.charts.forEach((c, i) => {
      const ch = state.charts.get(i);
      if (!ch) return;
      const fresh = chartConfig(c).data;
      ch.data.labels = fresh.labels;
      fresh.datasets.forEach((ds, j) => {
        if (ch.data.datasets[j]) ch.data.datasets[j].data = ds.data;
        else ch.data.datasets.push(ds);
      });
      ch.data.datasets.length = fresh.datasets.length;
      ch.update();
    });
    const pt = $("#previewTable");
    if (pt) pt.innerHTML = tableHtml(sheet);
  }

  function animateNumber(el, k) {
    const from = +(el.dataset.v || 0), to = k.value, t0 = performance.now();
    el.dataset.v = String(to);
    const step = (t) => {
      const p = Math.min(1, (t - t0) / 600), e = 1 - Math.pow(1 - p, 3);
      el.textContent = formatKpi({ ...k, value: from + (to - from) * e });
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  function formatKpi(k) {
    const v = k.format === "int" ? Math.round(k.value) : k.value;
    return Math.abs(v) >= 1e6 ? fmtCompact.format(v) : fmt.format(k.format === "int" ? v : Math.round(v * 100) / 100);
  }

  function typeLabel(t) {
    return { number: "número", date: "fecha", category: "categoría", text: "texto", id: "código", empty: "vacía" }[t] || t;
  }

  function tableHtml(sheet) {
    const { headers, rows } = sheet.preview;
    return `<table><thead><tr>${headers.map((h) => `<th>${esc(h)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${headers.map((_, i) => `<td>${esc(typeof r[i] === "number" ? fmt.format(r[i]) : r[i] ?? "")}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  }

  // Recomienda archivos con las mismas columnas para verlos juntos.
  function relatedHtml(d) {
    if (d.source === "combined") return "";
    const sig = new Set(d.analysis.signature);
    const similar = [...state.datasets.values()].filter((o) => {
      if (o.id === d.id || o.source === "combined" || !o.analysis) return false;
      const other = new Set(o.analysis.signature);
      const inter = [...sig].filter((x) => other.has(x)).length;
      return inter / new Set([...sig, ...other]).size >= 0.7;
    });
    if (!similar.length) return "";
    const ids = [d.id, ...similar.map((s) => s.id)];
    const already = [...state.datasets.values()].find((c) => c.source === "combined" && c.sources.length === ids.length && ids.every((id) => c.sources.includes(id)));
    return `<div class="related">
      <b>Archivos relacionados:</b> ${similar.map((s) => `<a href="#" data-open="${s.id}">${esc(s.name)}</a>`).join(", ")} tienen las mismas columnas.
      ${already
        ? `<br><button class="btn small" data-open="${already.id}">Ver combinado</button>`
        : `<br><button class="btn small primary" data-combine="${ids.join(",")}">Combinar y comparar</button>`}
    </div>`;
  }

  function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
  const palette = () => [1, 2, 3, 4, 5, 6, 7, 8].map((i) => cssVar(`--c${i}`));

  function chartConfig(c) {
    const colors = palette();
    const text = cssVar("--muted"), grid = cssVar("--border");
    const many = c.labels.length > 60;
    const datasets = c.datasets.map((ds, j) => {
      const color = colors[j % colors.length];
      if (c.type === "doughnut") return { label: ds.label, data: ds.data, backgroundColor: c.labels.map((_, i) => colors[i % colors.length]), borderColor: cssVar("--surface"), borderWidth: 2 };
      if (c.type === "line") return { label: ds.label, data: ds.data, borderColor: color, backgroundColor: color + "22", fill: c.datasets.length === 1, cubicInterpolationMode: "monotone", pointRadius: many ? 0 : 3, pointHoverRadius: 5, borderWidth: 2 };
      return { label: ds.label, data: ds.data, backgroundColor: color, borderRadius: 5, maxBarThickness: 42 };
    });
    const tick = (v) => (Math.abs(v) >= 1000 ? fmtCompact.format(v) : fmt.format(v));
    const scales = c.type === "doughnut" ? {} : {
      x: { ticks: { color: text, maxRotation: 0, autoSkip: true }, grid: { display: Boolean(c.horizontal), color: grid } },
      y: { ticks: { color: text }, grid: { display: !c.horizontal, color: grid }, beginAtZero: true },
    };
    if (c.type !== "doughnut") {
      const valueAxis = c.horizontal ? scales.x : scales.y;
      const labelAxis = c.horizontal ? scales.y : scales.x;
      valueAxis.ticks.callback = (v) => tick(v);
      labelAxis.ticks.callback = function (v) { const l = String(this.getLabelForValue(v)); return l.length > 22 ? l.slice(0, 21) + "…" : l; };
    }
    return {
      type: c.type,
      data: { labels: c.labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: c.horizontal ? "y" : "x",
        animation: { duration: 450, easing: "easeOutCubic" },
        interaction: { mode: c.type === "line" ? "index" : "nearest", intersect: false },
        plugins: {
          legend: { display: c.type === "doughnut" || c.datasets.length > 1, position: "bottom", labels: { color: text, boxWidth: 12, usePointStyle: true } },
          tooltip: { callbacks: { label: (ctx) => ` ${ctx.dataset.label}: ${fmt.format(ctx.raw)}` } },
        },
        scales,
        cutout: c.type === "doughnut" ? "62%" : undefined,
      },
    };
  }

  function destroyCharts() {
    state.charts.forEach((c) => c.destroy());
    state.charts.clear();
  }

  async function confirmDelete(d) {
    const msg = d.source === "drive"
      ? `¿Quitar "${d.name}"?\n\nSe moverá a la papelera de Google Drive (puedes recuperarlo desde ahí durante 30 días).`
      : `¿Quitar "${d.name}" de este dispositivo?`;
    if (!confirm(msg)) return;
    if (d.source === "drive") {
      if (!state.token) { toast("Conecta Drive para poder moverlo a la papelera.", "err"); return; }
      if (!(await trashInDrive(d))) return;
    }
    await removeDataset(d.id);
    await refreshCombined();
    render();
    toast(`"${d.name}" quitado`);
  }

  // ---------- Eventos ----------
  function bindEvents() {
    $("#btnUpload").onclick = () => $("#fileInput").click();
    $("#fileInput").onchange = (e) => { handleFiles([...e.target.files]); e.target.value = ""; };
    if (CAN_WATCH) { $("#btnWatch").classList.remove("hidden"); $("#btnWatch").onclick = pickAndWatch; }
    $("#search").oninput = (e) => { state.filter = e.target.value; renderList(); };
    $("#fileList").onclick = (e) => { const b = e.target.closest(".file"); if (b) select(b.dataset.id); };
    $("#btnList").onclick = () => document.body.classList.toggle("show-list");
    $("#main").addEventListener("click", (e) => { if (e.target.closest("a[data-open]")) e.preventDefault(); });
    document.addEventListener("click", (e) => {
      if (document.body.classList.contains("show-list") && !e.target.closest(".sidebar") && !e.target.closest("#btnList")) document.body.classList.remove("show-list");
    });

    // Arrastrar y soltar (en Windows además queda vigilado el archivo).
    let depth = 0;
    window.addEventListener("dragenter", (e) => { if (e.dataTransfer?.types.includes("Files")) { depth++; document.body.classList.add("dragging"); } });
    window.addEventListener("dragleave", () => { if (--depth <= 0) { depth = 0; document.body.classList.remove("dragging"); } });
    window.addEventListener("dragover", (e) => e.preventDefault());
    window.addEventListener("drop", async (e) => {
      e.preventDefault();
      depth = 0;
      document.body.classList.remove("dragging");
      const items = [...e.dataTransfer.items].filter((i) => i.kind === "file");
      const handles = await Promise.all(items.map((i) => (i.getAsFileSystemHandle ? i.getAsFileSystemHandle().catch(() => null) : null)));
      const files = [...e.dataTransfer.files];
      handleFiles(files, handles.map((h) => (h && h.kind === "file" ? h : undefined)));
    });

    window.addEventListener("online", () => { renderStatus(); syncDrive(); });
    window.addEventListener("offline", renderStatus);
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") { syncDrive(); checkWatched(); } });
    window.addEventListener("focus", () => { checkWatched(); });

    matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { state.structureKey = ""; renderDashboard(); });

    window.addEventListener("beforeinstallprompt", (e) => {
      e.preventDefault();
      state.installPrompt = e;
      $("#btnInstall").classList.remove("hidden");
    });
    $("#btnInstall").onclick = async () => {
      if (!state.installPrompt) return;
      state.installPrompt.prompt();
      await state.installPrompt.userChoice;
      state.installPrompt = null;
      $("#btnInstall").classList.add("hidden");
    };
  }

  // Bucles de auto-actualización
  function startLoops() {
    setInterval(() => { if (document.visibilityState === "visible") syncDrive(); }, (CFG.POLL_SECONDS || 20) * 1000);
    setInterval(() => { if (document.visibilityState === "visible") checkWatched(); }, (CFG.LOCAL_WATCH_SECONDS || 4) * 1000);
    setInterval(() => {
      renderStatus();
      document.querySelectorAll(".file").forEach((el) => {
        const d = state.datasets.get(el.dataset.id);
        const meta = el.querySelector(".file-meta");
        if (d && meta && d.analysis) meta.textContent = `${fmt.format(d.analysis.sheets[0].rowCount)} filas · ${ago(d.updatedAt)}`;
      });
      const d = state.datasets.get(state.selected), ua = $("#updatedAgo");
      if (d && ua) ua.textContent = `actualizado ${ago(d.updatedAt)}`;
    }, 5000);
  }

  async function main() {
    bindEvents();
    await loadLocal();
    for (const id of state.handles.keys()) state.needsPermission.add(id);
    render();
    initGoogle();
    startLoops();
    checkWatched();
    if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js").catch(() => {});
  }

  window.__nube = { handleFiles, state };
  main();
})();
