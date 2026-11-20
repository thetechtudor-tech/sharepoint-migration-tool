"""
Browser UI for sp_migrate.py. Keep both files in the same folder.

Run:   python sp_migrate_ui.py
A page opens in your browser. The UI is served only to this computer
(127.0.0.1); press Ctrl+C in this window to stop it.

Settings are saved to config.json in the current folder. Passwords are kept
in memory only while this program runs and are never written to disk.
"""
import argparse
import copy
import json
import os
import re
import secrets
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import sp_migrate as eng

CONFIG_PATH = "config.json"
UPLOAD_DIR = "uploads"
LOG_PATH = "migration_log.csv"
MAX_UPLOAD = 50 * 1024 * 1024
TOKEN = secrets.token_urlsafe(24)       # page must present this on every call
LOCK = threading.Lock()
STATE = {"password": None, "columns": None, "job": None}


# -------------------------------------------------------------- helpers ----
def load_cfg():
    cfg = {}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding="utf-8") as f:
            cfg = json.load(f)
    for section, defaults in eng.STARTER_CONFIG.items():
        target = cfg.setdefault(section, {})
        for k, v in defaults.items():
            target.setdefault(k, copy.deepcopy(v))
    return cfg


def excel_info(cfg):
    """Sheets and headers of the chosen workbook, or None if there is none."""
    xl = cfg["excel"]
    if not xl.get("path") or not os.path.exists(xl["path"]):
        return None
    sheets = eng.open_workbook(xl).sheetnames
    if xl.get("sheet") not in sheets:
        xl["sheet"] = sheets[0]
    return {"name": os.path.basename(xl["path"]), "sheets": sheets,
            "sheet": xl["sheet"], "headers": [h for h in eng.read_headers(xl) if h]}


def source_password():
    return STATE["password"] or ""      # never None: None would prompt in the console


def job_running():
    job = STATE["job"]
    return bool(job and job["running"])


def run_job(overwrite, move):
    job = STATE["job"]
    try:
        cfg = load_cfg()
        source = eng.open_source(cfg, source_password())
        try:
            files, metadata, problems = eng.plan(cfg, source)
            with LOCK:
                job.update(total=len(files), problems=problems, phase="Signing in")

            def on_start(n, f):
                with LOCK:
                    job.update(done=n - 1, current=f, phase="Uploading")

            def on_result(n, f, status, detail, url):
                with LOCK:
                    job["done"] = n
                    job["counts"][status] += 1
                    job["rows"].append({"file": f, "status": status,
                                        "detail": detail, "url": url})

            eng.migrate(cfg, source, files, metadata, overwrite, move, LOG_PATH,
                        on_start, on_result, lambda: job["stop"])
        finally:
            source.close()
    except BaseException as e:          # includes SystemExit from the engine
        with LOCK:
            job["error"] = str(e) or type(e).__name__
    finally:
        with LOCK:
            job.update(running=False, current="",
                       phase="Stopped" if job["stop"] else "Finished")


# --------------------------------------------------------------- routes ----
def api_state(_q, _body):
    cfg = load_cfg()
    try:
        excel = excel_info(cfg)
    except BaseException as e:
        excel = None
        cfg["excel"]["path"] = ""
        print(f"Could not read the saved Excel file: {e}")
    return {"config": cfg, "has_password": bool(STATE["password"]),
            "excel": excel, "columns": STATE["columns"],
            "running": job_running()}


def api_config(_q, body):
    data = json.loads(body or b"{}")
    cfg = load_cfg()
    incoming = data.get("config", {})
    for section in ("sharepoint", "source", "excel"):
        for k, v in incoming.get(section, {}).items():
            if section == "excel" and k == "path":
                continue                # the path is only set by an upload
            cfg[section][k] = v
    if data.get("password"):
        STATE["password"] = data["password"]
    eng.save_config(CONFIG_PATH, cfg)
    return {"ok": True}


def api_test_source(_q, _body):
    source = eng.open_source(load_cfg(), source_password())
    try:
        files = source.list()
        return {"label": source.label, "count": len(files), "sample": files[:8]}
    finally:
        source.close()


def api_columns(_q, _body):
    sp = load_cfg()["sharepoint"]
    token = eng.get_token(sp)
    STATE["columns"] = eng.get_columns(token, eng.get_drive_id(token, sp))
    return {"columns": STATE["columns"]}


def api_excel(q, body):
    name = os.path.basename(q.get("name", ["metadata.xlsx"])[0])
    name = re.sub(r"[^A-Za-z0-9._ -]", "_", name) or "metadata.xlsx"
    if not name.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("Please choose an .xlsx file.")
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    path = os.path.join(UPLOAD_DIR, name)
    with open(path, "wb") as f:
        f.write(body)
    cfg = load_cfg()
    cfg["excel"].update(path=path, sheet="")
    info = excel_info(cfg)
    eng.save_config(CONFIG_PATH, cfg)
    return info


def api_headers(_q, _body):
    cfg = load_cfg()
    info = excel_info(cfg)
    if info is None:
        raise ValueError("No Excel file chosen yet.")
    return info


def api_template(_q, _body):
    """Create a starter workbook listing every source file."""
    cfg = load_cfg()
    source = eng.open_source(cfg, source_password())
    try:
        files = source.list()
    finally:
        source.close()
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    path = os.path.join(UPLOAD_DIR, "metadata_template.xlsx")
    cfg["excel"].update(path=path, sheet="", file_column="FileName")
    eng.write_template(cfg, files)
    eng.save_config(CONFIG_PATH, cfg)
    return {"path": os.path.abspath(path), "count": len(files)}


def api_preview(_q, _body):
    cfg = load_cfg()
    source = eng.open_source(cfg, source_password())
    try:
        files, metadata, problems = eng.plan(cfg, source)
    finally:
        source.close()
    rows = eng.preview_rows(cfg, files, metadata)
    return {"total": len(rows), "with_metadata": sum(1 for r in rows if r["columns"]),
            "problems": problems, "rows": rows[:500],
            "library": cfg["sharepoint"].get("library", "")}


def api_run(_q, body):
    opts = json.loads(body or b"{}")
    with LOCK:
        if job_running():
            raise ValueError("A migration is already running.")
        STATE["job"] = {"running": True, "stop": False, "total": 0, "done": 0,
                        "current": "", "phase": "Connecting to the source",
                        "counts": {"ok": 0, "skipped": 0, "failed": 0},
                        "rows": [], "problems": [], "error": "",
                        "log": os.path.abspath(LOG_PATH)}
    threading.Thread(target=run_job, daemon=True,
                     args=(bool(opts.get("overwrite")), bool(opts.get("move")))).start()
    return {"ok": True}


def api_stop(_q, _body):
    with LOCK:
        if STATE["job"]:
            STATE["job"]["stop"] = True
    return {"ok": True}


def api_status(q, _body):
    since = int(q.get("since", ["0"])[0])
    with LOCK:
        job = STATE["job"]
        if not job:
            return {"job": None}
        out = {k: v for k, v in job.items() if k != "rows"}
        out["rows"] = job["rows"][since:since + 500]
        out["next"] = since + len(out["rows"])
        return {"job": out}


GET_ROUTES = {"/api/state": api_state, "/api/status": api_status,
              "/api/headers": api_headers}
POST_ROUTES = {"/api/config": api_config, "/api/test-source": api_test_source,
               "/api/columns": api_columns, "/api/excel": api_excel,
               "/api/template": api_template, "/api/preview": api_preview,
               "/api/run": api_run, "/api/stop": api_stop}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass                            # keep the console quiet

    def _send(self, code, body, ctype="application/json"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _allowed(self):
        """Only this computer's browser, on the address we are serving."""
        port = self.server.server_address[1]
        host = self.headers.get("Host", "")
        return host in (f"127.0.0.1:{port}", f"localhost:{port}")

    def _route(self, routes):
        url = urlparse(self.path)
        if not self._allowed():
            return self._send(403, json.dumps({"error": "Forbidden"}))
        if url.path == "/" and routes is GET_ROUTES:
            return self._send(200, PAGE.replace("__TOKEN__", TOKEN), "text/html")
        handler = routes.get(url.path)
        if handler is None:
            return self._send(404, json.dumps({"error": "Not found"}))
        if not secrets.compare_digest(self.headers.get("X-Token", ""), TOKEN):
            return self._send(403, json.dumps({"error": "Forbidden"}))
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD:
            return self._send(413, json.dumps({"error": "File is too large."}))
        body = self.rfile.read(length) if length else b""
        try:
            result = handler(parse_qs(url.query), body)
            self._send(200, json.dumps(result))
        except BaseException as e:      # includes SystemExit from the engine
            self._send(400, json.dumps({"error": str(e) or type(e).__name__}))

    def do_GET(self):
        self._route(GET_ROUTES)

    def do_POST(self):
        self._route(POST_ROUTES)


# ----------------------------------------------------------------- page ----
PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SharePoint file migration</title>
<style>
  :root { --ink:#1c2430; --muted:#5d6b7c; --line:#d9dfe7; --bg:#f3f5f8; --card:#fff;
          --accent:#0b6bcb; --ok:#1a7f4b; --warn:#9a6200; --err:#b3261e; }
  * { box-sizing:border-box; }
  body { margin:0; font:15px/1.45 "Segoe UI",system-ui,-apple-system,sans-serif;
         color:var(--ink); background:var(--bg); }
  header { padding:28px 24px 8px; max-width:980px; margin:0 auto; }
  h1 { font-size:24px; margin:0 0 4px; }
  header p { margin:0; color:var(--muted); }
  main { max-width:980px; margin:0 auto; padding:12px 24px 60px; }
  section { background:var(--card); border:1px solid var(--line); border-radius:10px;
            padding:20px 22px; margin-top:16px; }
  h2 { font-size:17px; margin:0 0 14px; display:flex; align-items:center; gap:10px; }
  h2 .n { width:26px; height:26px; border-radius:50%; background:var(--accent); color:#fff;
          font-size:14px; display:inline-flex; align-items:center; justify-content:center; }
  .grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); gap:12px 16px; }
  label { display:flex; flex-direction:column; gap:4px; font-size:13px; color:var(--muted); }
  label.check { flex-direction:row; align-items:center; gap:8px; font-size:15px; color:var(--ink); }
  input:not([type=checkbox]):not([type=file]), select {
    font:inherit; color:var(--ink); padding:8px 10px; border:1px solid var(--line);
    border-radius:6px; background:#fff; width:100%; }
  input:focus, select:focus, button:focus-visible { outline:2px solid var(--accent); outline-offset:1px; }
  button { font:inherit; padding:8px 16px; border-radius:6px; border:1px solid var(--accent);
           background:var(--accent); color:#fff; cursor:pointer; }
  button.secondary { background:#fff; color:var(--accent); }
  button.danger { background:#fff; color:var(--err); border-color:var(--err); }
  button:disabled { opacity:.5; cursor:default; }
  .row { display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-top:14px; }
  .msg { margin-top:10px; font-size:14px; color:var(--muted); white-space:pre-wrap; overflow-wrap:anywhere; }
  .msg.ok { color:var(--ok); } .msg.err { color:var(--err); } .msg.warn { color:var(--warn); }
  .hint { font-size:13px; color:var(--muted); margin:6px 0 0; }
  .hidden { display:none !important; }
  .tablewrap { overflow:auto; max-height:420px; margin-top:12px; border:1px solid var(--line); border-radius:6px; }
  table { border-collapse:collapse; width:100%; font-size:14px; }
  th, td { text-align:left; padding:7px 10px; border-bottom:1px solid var(--line); vertical-align:top; }
  th { background:#f7f9fb; position:sticky; top:0; font-weight:600; }
  td.path { overflow-wrap:anywhere; }
  #map td:first-child { font-weight:600; width:40%; }
  .tag { display:inline-block; padding:1px 8px; border-radius:10px; font-size:12px; font-weight:600; }
  .tag.ok { background:#e2f4ea; color:var(--ok); } .tag.skipped { background:#fdf0d5; color:var(--warn); }
  .tag.failed { background:#fbe4e2; color:var(--err); }
  progress { width:100%; height:14px; margin-top:14px; }
  ul.problems { margin:8px 0 0; padding-left:20px; color:var(--warn); font-size:14px; }
</style>
</head>
<body>
<header>
  <h1>SharePoint file migration</h1>
  <p>Copy the files listed in an Excel sheet from a server or folder into a SharePoint library, and fill in their columns from the same sheet.</p>
</header>
<main>

<section>
  <h2><span class="n">1</span>Where the files are</h2>
  <div class="grid">
    <label>Connection type
      <select id="src_type">
        <option value="sftp">SFTP</option>
        <option value="ftps">FTPS (FTP over TLS)</option>
        <option value="ftp">FTP (not encrypted)</option>
        <option value="folder">Folder on this computer or network</option>
      </select></label>
    <label class="remote">Server address <input id="src_host" placeholder="files.example.com"></label>
    <label class="remote">Port <input id="src_port" inputmode="numeric" placeholder="Default"></label>
    <label class="remote">User name <input id="src_user" autocomplete="off"></label>
    <label class="remote">Password <input id="src_password" type="password" autocomplete="new-password"></label>
    <label class="sftp">Private key file (optional) <input id="src_key_file" placeholder="C:\Users\you\.ssh\id_ed25519"></label>
    <label class="remote">Folder on the server <input id="src_remote_dir" placeholder="/"></label>
    <label class="folder">Folder path <input id="src_folder_path" placeholder="C:\Files\ToMigrate"></label>
  </div>
  <p class="hint remote">The password is kept only while this tool is running. It is never saved.</p>
  <p class="hint ftpwarn">Plain FTP sends the password and files without encryption. Use SFTP or FTPS if the server offers them.</p>
  <div class="row"><button id="btn_test">Test connection</button></div>
  <div class="msg" id="msg1"></div>
</section>

<section>
  <h2><span class="n">2</span>Where they are going</h2>
  <div class="grid">
    <label>Application (client) ID <input id="sp_client_id"></label>
    <label>Directory (tenant) ID <input id="sp_tenant_id"></label>
    <label>SharePoint site address <input id="sp_site_url" placeholder="https://contoso.sharepoint.com/sites/YourSite"></label>
    <label>Document library <input id="sp_library" placeholder="Documents"></label>
    <label>Folder in the library (optional) <input id="sp_target_folder" placeholder="Imported"></label>
  </div>
  <p class="hint">The two IDs come from your IT team's app registration.</p>
  <div class="row"><button id="btn_columns">Sign in and load columns</button></div>
  <div class="msg" id="msg2"></div>
</section>

<section>
  <h2><span class="n">3</span>Excel sheet and column matching</h2>
  <div class="row" style="margin-top:0">
    <input type="file" id="excel_file" accept=".xlsx,.xlsm" aria-label="Excel file">
    <button class="secondary" id="btn_template">Create a starter sheet from the source files</button>
  </div>
  <div class="msg" id="msg3"></div>
  <div id="excel_opts" class="hidden">
    <div class="grid" style="margin-top:14px">
      <label>Worksheet <select id="xl_sheet"></select></label>
      <label>Column that holds the file name <select id="xl_file_column"></select></label>
    </div>
    <p class="hint">Only the files listed in this sheet are moved. Other files in the source are left alone. File names can be a path such as <b>reports/q1.pdf</b>, or just the name when it is unique.</p>
    <div class="tablewrap"><table id="map">
      <thead><tr><th>Excel column</th><th>SharePoint column</th></tr></thead><tbody></tbody>
    </table></div>
    <p class="hint" id="map_hint"></p>
  </div>
</section>

<section>
  <h2><span class="n">4</span>Check and run</h2>
  <label class="check"><input type="checkbox" id="opt_overwrite"> Replace files that already exist in SharePoint (otherwise they are skipped)</label>
  <label class="check" style="margin-top:6px"><input type="checkbox" id="opt_move"> Delete each file from the source after it uploads successfully</label>
  <div class="row">
    <button class="secondary" id="btn_preview">Preview</button>
    <button id="btn_run">Start migration</button>
    <button class="danger hidden" id="btn_stop">Stop</button>
  </div>
  <div class="msg" id="msg4"></div>
  <ul class="problems" id="problems"></ul>
  <progress id="bar" class="hidden" value="0" max="1"></progress>
  <div class="tablewrap hidden" id="out_wrap"><table id="out"><thead></thead><tbody></tbody></table></div>
</section>

</main>
<script>
const TOKEN = "__TOKEN__";
const $ = id => document.getElementById(id);
let cfg = null, spCols = null, headers = [], pollTimer = null, nextRow = 0;

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g,
    c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}
function msg(id, text, kind) { const e = $(id); e.textContent = text; e.className = "msg " + (kind || ""); }

async function api(path, body, raw) {
  const opt = {headers: {"X-Token": TOKEN}};
  if (body !== undefined) { opt.method = "POST"; opt.body = raw ? body : JSON.stringify(body); }
  const r = await fetch(path, opt);
  let j = {};
  try { j = await r.json(); } catch (e) {}
  if (!r.ok) throw new Error(j.error || ("Request failed (" + r.status + ")"));
  return j;
}
async function act(btn, msgId, working, fn) {
  btn.disabled = true; msg(msgId, working);
  try { await fn(); } catch (e) { msg(msgId, e.message, "err"); } finally { btn.disabled = false; }
}

/* ---------- form <-> config ---------- */
function showFields() {
  const t = $("src_type").value;
  document.querySelectorAll(".remote").forEach(e => e.classList.toggle("hidden", t === "folder"));
  document.querySelectorAll(".sftp").forEach(e => e.classList.toggle("hidden", t !== "sftp"));
  document.querySelectorAll(".folder").forEach(e => e.classList.toggle("hidden", t !== "folder"));
  document.querySelectorAll(".ftpwarn").forEach(e => e.classList.toggle("hidden", t !== "ftp"));
}
function fill(hasPassword) {
  const s = cfg.source, p = cfg.sharepoint;
  $("src_type").value = ["sftp","ftps","ftp","folder"].includes(s.type) ? s.type : "sftp";
  $("src_host").value = s.host || ""; $("src_port").value = s.port || "";
  $("src_user").value = s.user || ""; $("src_key_file").value = s.key_file || "";
  $("src_remote_dir").value = s.remote_dir || ""; $("src_folder_path").value = s.folder_path || "";
  if (hasPassword) $("src_password").placeholder = "(entered)";
  $("sp_client_id").value = p.client_id || ""; $("sp_tenant_id").value = p.tenant_id || "";
  $("sp_site_url").value = p.site_host ? "https://" + p.site_host + (p.site_path || "") : "";
  $("sp_library").value = p.library || ""; $("sp_target_folder").value = p.target_folder || "";
  showFields();
}
function parseSite(v) {
  v = v.trim();
  if (!v) return {site_host: "", site_path: ""};
  if (!/^https?:\/\//i.test(v)) v = "https://" + v;
  let u;
  try { u = new URL(v); } catch (e) { throw new Error("The SharePoint site address is not a valid web address."); }
  let parts = u.pathname.split("/").filter(Boolean);
  if (parts.length && ["sites","teams"].includes(parts[0].toLowerCase())) parts = parts.slice(0, 2);
  else parts = [];
  return {site_host: u.hostname, site_path: parts.length ? "/" + parts.map(decodeURIComponent).join("/") : ""};
}
function stashMapping() {
  document.querySelectorAll("#map select").forEach(sel => { cfg.excel.columns[sel.dataset.h] = sel.value; });
}
function collect() {
  const site = parseSite($("sp_site_url").value);
  cfg.source = Object.assign(cfg.source, {
    type: $("src_type").value, host: $("src_host").value.trim(), port: $("src_port").value.trim(),
    user: $("src_user").value.trim(), key_file: $("src_key_file").value.trim(),
    remote_dir: $("src_remote_dir").value.trim() || "/", folder_path: $("src_folder_path").value.trim()});
  cfg.sharepoint = Object.assign(cfg.sharepoint, site, {
    client_id: $("sp_client_id").value.trim(), tenant_id: $("sp_tenant_id").value.trim(),
    library: $("sp_library").value.trim() || "Documents", target_folder: $("sp_target_folder").value.trim()});
  stashMapping();
  if (headers.length) { cfg.excel.sheet = $("xl_sheet").value; cfg.excel.file_column = $("xl_file_column").value; }
  const mapping = {};
  headers.forEach(h => { if (h !== cfg.excel.file_column) mapping[h] = cfg.excel.columns[h] || ""; });
  if (headers.length) cfg.excel.columns = mapping;
}
async function save() {
  collect();
  const body = {config: cfg};
  if ($("src_password").value) body.password = $("src_password").value;
  await api("/api/config", body);
}

/* ---------- Excel and mapping ---------- */
function setExcel(info, keepChoices) {
  headers = info.headers;
  $("excel_opts").classList.remove("hidden");
  $("xl_sheet").innerHTML = info.sheets.map(s => `<option>${esc(s)}</option>`).join("");
  $("xl_sheet").value = info.sheet;
  let fileCol = cfg.excel.file_column;
  if (!headers.includes(fileCol)) {
    fileCol = headers.find(h => /file|name|path|document/i.test(h)) || headers[0] || "";
  }
  cfg.excel.file_column = fileCol; cfg.excel.sheet = info.sheet;
  $("xl_file_column").innerHTML = headers.map(h => `<option>${esc(h)}</option>`).join("");
  $("xl_file_column").value = fileCol;
  renderMapping();
}
function autoMatch(h) {
  if (!spCols) return "";
  const k = h.toLowerCase();
  const c = spCols.find(c => c.display.toLowerCase() === k) || spCols.find(c => c.name.toLowerCase() === k);
  return c ? c.name : "";
}
function renderMapping() {
  const body = $("map").querySelector("tbody");
  const fileCol = $("xl_file_column").value;
  const rows = headers.filter(h => h !== fileCol).map(h => {
    let cur = cfg.excel.columns[h];
    if (cur === undefined) { cur = autoMatch(h); cfg.excel.columns[h] = cur; }
    let opts = `<option value="">(do not import)</option>`;
    let found = !cur;
    (spCols || []).forEach(c => {
      if (c.name === cur) found = true;
      const extra = (c.required ? ", required" : "") + (c.special ? ", not supported yet" : "");
      opts += `<option value="${esc(c.name)}"${c.name === cur ? " selected" : ""}>${esc(c.display)} (${esc(c.type)}${extra})</option>`;
    });
    if (!found) opts += `<option value="${esc(cur)}" selected>${esc(cur)}</option>`;
    return `<tr><td>${esc(h)}</td><td><select data-h="${esc(h)}" aria-label="SharePoint column for ${esc(h)}">${opts}</select></td></tr>`;
  });
  body.innerHTML = rows.join("") || `<tr><td colspan="2">This sheet has no other columns to import.</td></tr>`;
  $("map_hint").textContent = spCols
    ? "Columns with matching names were paired automatically. Change any of them here."
    : "Sign in under step 2 to load the SharePoint columns you can choose from.";
}

/* ---------- actions ---------- */
$("src_type").addEventListener("change", showFields);

$("btn_test").addEventListener("click", e => act(e.target, "msg1", "Connecting...", async () => {
  await save();
  const r = await api("/api/test-source", {});
  msg("msg1", `Connected to ${r.label}. Found ${r.count} file${r.count === 1 ? "" : "s"}.` +
      (r.sample.length ? "\nFor example: " + r.sample.join(", ") : ""), "ok");
}));

$("btn_columns").addEventListener("click", e => act(e.target, "msg2",
  "Waiting for sign-in. A Microsoft sign-in page may open in another tab...", async () => {
  await save();
  const r = await api("/api/columns", {});
  spCols = r.columns;
  headers.forEach(h => { if (!cfg.excel.columns[h]) delete cfg.excel.columns[h]; });
  if (headers.length) renderMapping();
  msg("msg2", `Signed in. The library has ${spCols.length} columns you can fill in.`, "ok");
}));

$("excel_file").addEventListener("change", async e => {
  const f = e.target.files[0];
  if (!f) return;
  msg("msg3", "Reading the sheet...");
  try {
    await save();
    const info = await api("/api/excel?name=" + encodeURIComponent(f.name), f, true);
    setExcel(info, false);
    msg("msg3", `Loaded ${info.name}.`, "ok");
  } catch (err) { msg("msg3", err.message, "err"); }
});

$("btn_template").addEventListener("click", e => act(e.target, "msg3", "Listing the source files...", async () => {
  await save();
  const r = await api("/api/template", {});
  msg("msg3", `Created a sheet listing ${r.count} files:\n${r.path}\n` +
      "Open it in Excel, add a column for each piece of information, save it, then choose it here.", "ok");
}));

$("xl_sheet").addEventListener("change", async () => {
  try {
    cfg.excel.sheet = $("xl_sheet").value;
    await api("/api/config", {config: {excel: {sheet: cfg.excel.sheet}}});
    setExcel(await api("/api/headers"), false);
  } catch (err) { msg("msg3", err.message, "err"); }
});
$("xl_file_column").addEventListener("change", () => { stashMapping(); renderMapping(); });

function showProblems(list) {
  $("problems").innerHTML = (list || []).slice(0, 50).map(p => `<li>${esc(p)}</li>`).join("");
}
function showTable(head) {
  $("out").querySelector("thead").innerHTML = "<tr>" + head.map(h => `<th>${h}</th>`).join("") + "</tr>";
  $("out").querySelector("tbody").innerHTML = "";
  $("out_wrap").classList.remove("hidden");
}

$("btn_preview").addEventListener("click", e => act(e.target, "msg4", "Building the preview...", async () => {
  await save();
  const r = await api("/api/preview", {});
  showProblems(r.problems);
  showTable(["File", "Goes to", "Columns that will be filled in"]);
  $("out").querySelector("tbody").innerHTML = r.rows.map(x =>
    `<tr><td class="path">${esc(x.file)}${x.invalid ? ' <span class="tag failed">invalid name</span>' : ""}</td>` +
    `<td class="path">${esc(r.library + "/" + x.remote)}</td>` +
    `<td>${x.columns.length ? esc(x.columns.join(", ")) : "<i>none</i>"}</td></tr>`).join("");
  msg("msg4", `${r.total} file${r.total === 1 ? "" : "s"} listed in the sheet will be copied.` +
      (r.total > r.rows.length ? ` Showing the first ${r.rows.length}.` : "") + " Nothing has been changed yet.", "ok");
}));

function setRunning(on) {
  $("btn_run").disabled = on; $("btn_preview").disabled = on;
  $("btn_stop").classList.toggle("hidden", !on);
  $("bar").classList.toggle("hidden", !on);
}
async function poll() {
  let r;
  try { r = await api("/api/status?since=" + nextRow); }
  catch (e) { msg("msg4", "Lost contact with the tool. Is its window still open?", "err"); return; }
  const j = r.job;
  if (!j) return;
  nextRow = j.next;
  showProblems(j.problems);
  const tb = $("out").querySelector("tbody");
  tb.insertAdjacentHTML("beforeend", j.rows.map(x =>
    `<tr><td class="path">${x.url ? `<a href="${esc(x.url)}" target="_blank" rel="noopener">${esc(x.file)}</a>` : esc(x.file)}</td>` +
    `<td><span class="tag ${esc(x.status)}">${esc(x.status === "ok" ? "uploaded" : x.status)}</span></td>` +
    `<td class="path">${esc(x.detail)}</td></tr>`).join(""));
  $("bar").max = j.total || 1; $("bar").value = j.done;
  const c = j.counts, sum = `${c.ok} uploaded, ${c.skipped} skipped, ${c.failed} failed`;
  if (j.running) {
    msg("msg4", `${j.phase}: ${j.done} of ${j.total}` + (j.current ? ` (${j.current})` : "") + `\n${sum}`);
    pollTimer = setTimeout(poll, 1000);
  } else {
    setRunning(false);
    if (j.error) msg("msg4", `Stopped because of a problem: ${j.error}\n${sum}`, "err");
    else msg("msg4", `${j.phase}. ${sum}.\nFull log: ${j.log}`, c.failed ? "warn" : "ok");
  }
}
$("btn_run").addEventListener("click", async () => {
  try {
    await save();
    await api("/api/run", {overwrite: $("opt_overwrite").checked, move: $("opt_move").checked});
  } catch (e) { msg("msg4", e.message, "err"); return; }
  nextRow = 0;
  showTable(["File", "Result", "Details"]);
  setRunning(true);
  poll();
});
$("btn_stop").addEventListener("click", async () => {
  await api("/api/stop", {});
  msg("msg4", "Stopping after the current file...");
});

/* ---------- start ---------- */
(async () => {
  try {
    const s = await api("/api/state");
    cfg = s.config; spCols = s.columns;
    fill(s.has_password);
    if (s.excel) { setExcel(s.excel, true); msg("msg3", `Using ${s.excel.name}.`); }
    if (s.running) { showTable(["File", "Result", "Details"]); setRunning(true); poll(); }
  } catch (e) { msg("msg1", e.message, "err"); }
})();
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--port", type=int, default=0, help="default: any free port")
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()

    if getattr(sys, "frozen", False):
        # Running as a built .exe: keep settings and logs next to the program.
        os.chdir(os.path.dirname(os.path.abspath(sys.executable)))

    server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"SharePoint migration UI is running at {url}\nKeep this window open while you use the tool. Close it to stop.", flush=True)
    if not args.no_browser:
        threading.Timer(0.5, webbrowser.open, args=(url,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
