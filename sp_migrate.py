"""
Copy (or move) files from an SFTP / FTPS / FTP server or a folder into a
SharePoint Online document library, and set each file's metadata from an
Excel sheet. Runs as the signed-in user (delegated auth via Microsoft Graph).

Install:  pip install msal requests openpyxl paramiko
          (paramiko is only needed for SFTP)

Most people should use the UI instead:   python sp_migrate_ui.py

Command line (all commands read config.json unless --config says otherwise)
  init       write a starter config.json to fill in
  template   create an Excel sheet listing every file in the source
  columns    list the SharePoint library's columns
  map        match Excel headers to SharePoint columns and save to the config
  run        do the migration   (--dry-run, --overwrite, --move)

source.type is one of: "sftp", "ftps" (FTP over TLS), "ftp" (unencrypted),
"folder". Passwords are never stored: set the SOURCE_PASSWORD environment
variable or you will be prompted. SFTP can use source.key_file instead.

The Excel sheet decides what moves: only files that have a row in the sheet
are copied. Other files in the source are left alone.

Column matching lives in config.json under excel.columns:
    "columns": { "Excel header": "SharePointInternalName", ... }
Headers mapped to "" (or not listed) are ignored. excel.file_column names
the header that identifies the file: a path relative to the source root
("sub/a.pdf"), or just the file name when it is unique.
"""
import argparse
import csv
import datetime as dt
import ftplib
import getpass
import json
import os
import posixpath
import shutil
import stat
import sys
import tempfile
import time
from urllib.parse import quote

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.ReadWrite.All"]
CACHE_FILE = "token_cache.json"
KNOWN_HOSTS_FILE = "sftp_known_hosts"
SMALL_FILE_LIMIT = 4 * 1024 * 1024        # above this, use an upload session
CHUNK_SIZE = 320 * 1024 * 16              # 5 MiB; must be a multiple of 320 KiB
BAD_CHARS = set('"*:<>?|')                # not allowed in SharePoint names
CLUTTER = ("Thumbs.db", ".DS_Store")

STARTER_CONFIG = {
    "sharepoint": {
        "client_id": "",
        "tenant_id": "",
        "site_host": "",
        "site_path": "",
        "library": "Documents",
        "target_folder": "",
    },
    "source": {
        "type": "sftp",
        "host": "",
        "port": "",
        "user": "",
        "key_file": "",
        "remote_dir": "/",
        "folder_path": "",
    },
    "excel": {
        "path": "",
        "sheet": "",
        "file_column": "FileName",
        "columns": {},
    },
}


class GraphError(Exception):
    def __init__(self, status, text):
        super().__init__(f"{status}: {text[:300]}")
        self.status = status


def skip_name(name):
    return name.startswith("~$") or name in CLUTTER


def resolve_password(cfg, password, optional=False):
    """Password passed in (UI) > environment variable > console prompt."""
    if password is not None:
        return password
    for var in ("SOURCE_PASSWORD", "FTP_PASSWORD"):
        if os.environ.get(var) is not None:
            return os.environ[var]
    if optional:
        return ""
    return getpass.getpass(f"Password for {cfg.get('user')}@{cfg.get('host')}: ")


# ------------------------------------------------------------- sources ----
class FolderSource:
    """Files in a local folder, mapped drive or UNC path."""

    def __init__(self, cfg, password=None):
        self.root = cfg.get("folder_path", "")
        if not self.root or not os.path.isdir(self.root):
            raise SystemExit(f"Source folder not found: {self.root!r}")
        self.label = self.root

    def list(self):
        out = []
        for root, _dirs, names in os.walk(self.root):
            for n in names:
                if not skip_name(n):
                    rel = os.path.relpath(os.path.join(root, n), self.root)
                    out.append(rel.replace(os.sep, "/"))
        return sorted(out, key=str.lower)

    def fetch(self, rel):
        return os.path.join(self.root, *rel.split("/"))

    def release(self, local_path):
        pass

    def delete(self, rel):
        os.remove(self.fetch(rel))

    def close(self):
        pass


class RemoteSource:
    """Shared behaviour for FTP and SFTP: each file is downloaded to a
    temporary folder, uploaded to SharePoint, then the copy is removed."""

    def _setup(self, cfg, default_port, scheme):
        if not cfg.get("host"):
            raise SystemExit("Source host is not set.")
        self.host = cfg["host"]
        self.port = int(cfg.get("port") or default_port)
        self.user = cfg.get("user") or "anonymous"
        self.root = cfg.get("remote_dir") or "/"
        self.tmp = tempfile.mkdtemp(prefix="sp_migrate_")
        self.label = f"{scheme}://{self.host}{self.root}"

    def list(self):
        out = []

        def walk(directory, rel):
            for name, is_dir in list(self._entries(directory)):
                child_rel = f"{rel}/{name}" if rel else name
                if is_dir:
                    walk(posixpath.join(directory, name), child_rel)
                elif not skip_name(name):
                    out.append(child_rel)

        walk(self.root, "")
        return sorted(out, key=str.lower)

    def fetch(self, rel):
        local = os.path.join(self.tmp, "download.tmp")
        self._retry(lambda: self._download(posixpath.join(self.root, rel), local))
        return local

    def release(self, local_path):
        if os.path.exists(local_path):
            os.remove(local_path)

    def delete(self, rel):
        self._retry(lambda: self._remove(posixpath.join(self.root, rel)))

    def close(self):
        try:
            self._disconnect()
        except Exception:
            pass
        shutil.rmtree(self.tmp, ignore_errors=True)


class FtpSource(RemoteSource):
    """FTP, or FTPS (explicit TLS) when tls=True."""

    def __init__(self, cfg, password=None, tls=False):
        self.tls = tls
        self._setup(cfg, 21, "ftps" if tls else "ftp")
        self.password = resolve_password(cfg, password)
        self.ftp = None
        self._connect()

    def _connect(self):
        ftp = ftplib.FTP_TLS(timeout=120) if self.tls else ftplib.FTP(timeout=120)
        ftp.connect(self.host, self.port)
        ftp.login(self.user, self.password)
        if self.tls:
            ftp.prot_p()                  # encrypt the data channel too
        self.ftp = ftp

    def _disconnect(self):
        self.ftp.quit()

    def _retry(self, action):
        """Run an action; reconnect once if the connection dropped."""
        try:
            return action()
        except ftplib.error_perm:
            raise                         # a real refusal, not a dropped link
        except ftplib.all_errors:
            self._connect()
            return action()

    def _entries(self, directory):
        try:
            for name, facts in self.ftp.mlsd(directory):
                kind = facts.get("type", "")
                if kind in ("cdir", "pdir") or name in (".", ".."):
                    continue
                yield name, kind == "dir"
            return
        except ftplib.error_perm:
            pass                          # server has no MLSD: fall back
        for path in self.ftp.nlst(directory):
            name = posixpath.basename(path.rstrip("/"))
            if name in ("", ".", ".."):
                continue
            try:
                self.ftp.cwd(posixpath.join(directory, name))  # dirs only
                yield name, True
            except ftplib.error_perm:
                yield name, False

    def _download(self, remote, local):
        with open(local, "wb") as f:
            self.ftp.retrbinary("RETR " + remote, f.write)

    def _remove(self, remote):
        self.ftp.delete(remote)


class SftpSource(RemoteSource):
    """SFTP (file transfer over SSH). Signs in with a password, a private
    key file, or both. The server's host key is remembered on first
    connection (sftp_known_hosts) and a later mismatch is refused."""

    def __init__(self, cfg, password=None):
        try:
            import paramiko
        except ImportError:
            raise SystemExit("SFTP needs the 'paramiko' package:  pip install paramiko")
        self.paramiko = paramiko
        self._setup(cfg, 22, "sftp")
        self.key_file = os.path.expanduser(cfg.get("key_file") or "")
        if self.key_file and not os.path.exists(self.key_file):
            raise SystemExit(f"Private key file not found: {self.key_file}")
        self.password = resolve_password(cfg, password, optional=bool(self.key_file))
        self.client = None
        self.sftp = None
        self._connect()

    def _connect(self):
        p = self.paramiko
        client = p.SSHClient()
        client.load_system_host_keys()
        if not os.path.exists(KNOWN_HOSTS_FILE):
            open(KNOWN_HOSTS_FILE, "a").close()
        client.load_host_keys(KNOWN_HOSTS_FILE)   # new keys are saved here
        client.set_missing_host_key_policy(p.AutoAddPolicy())
        client.connect(
            self.host, port=self.port, username=self.user,
            password=self.password or None,       # also unlocks an encrypted key
            key_filename=self.key_file or None,
            timeout=60, banner_timeout=60, auth_timeout=60)
        self.client = client
        self.sftp = client.open_sftp()

    def _disconnect(self):
        self.client.close()

    def _retry(self, action):
        try:
            return action()
        except (PermissionError, FileNotFoundError):
            raise                         # a real refusal, not a dropped link
        except (self.paramiko.SSHException, EOFError, OSError):
            self._connect()
            return action()

    def _entries(self, directory):
        for attr in self.sftp.listdir_attr(directory):
            if attr.filename in (".", ".."):
                continue
            mode = attr.st_mode or 0
            if stat.S_ISDIR(mode):
                yield attr.filename, True
            elif stat.S_ISREG(mode) or not mode:
                yield attr.filename, False    # symlinks etc. are skipped

    def _download(self, remote, local):
        self.sftp.get(remote, local)

    def _remove(self, remote):
        self.sftp.remove(remote)


def open_source(cfg, password=None):
    """password=None means 'ask' (env var or console); the UI passes a string."""
    src = cfg["source"]
    kind = (src.get("type") or "folder").lower()
    if kind == "sftp":
        return SftpSource(src, password)
    if kind == "ftps":
        return FtpSource(src, password, tls=True)
    if kind == "ftp":
        return FtpSource(src, password, tls=bool(src.get("tls", False)))
    if kind == "folder":
        return FolderSource(src)
    raise SystemExit(f"source.type must be sftp, ftps, ftp or folder, not {kind!r}")


# ---------------------------------------------------------------- auth ----
_app = None
_app_key = None
_cache = None


def get_token(sp):
    """Return a valid token. Browser sign-in the first time, silent after.
    Called before every file so long runs survive token expiry."""
    global _app, _app_key, _cache
    try:
        import msal
        import requests  # noqa: F401  (needed for every SharePoint call)
    except ImportError as e:
        raise SystemExit(f"Missing package ({e.name}). Run:  pip install msal requests openpyxl")
    if not sp.get("client_id") or not sp.get("tenant_id"):
        raise SystemExit("SharePoint client ID and tenant ID are not set.")
    key = (sp["client_id"], sp["tenant_id"])
    if _app is None or _app_key != key:
        _cache = msal.SerializableTokenCache()
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE) as f:
                _cache.deserialize(f.read())
        _app = msal.PublicClientApplication(
            sp["client_id"],
            authority=f"https://login.microsoftonline.com/{sp['tenant_id']}",
            token_cache=_cache,
        )
        _app_key = key
    result = None
    accounts = _app.get_accounts()
    if accounts:
        result = _app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:
        result = _app.acquire_token_interactive(SCOPES)
    if "access_token" not in result:
        raise SystemExit(f"Sign-in failed: {result.get('error_description', result)}")
    if _cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(_cache.serialize())
    return result["access_token"]


# ---------------------------------------------------------------- http ----
def call(method, url, token=None, retries=6, **kwargs):
    """HTTP call with throttling handling. Raises GraphError on failure."""
    import requests
    headers = kwargs.pop("headers", {})
    if token:
        headers["Authorization"] = f"Bearer {token}"
    for attempt in range(retries):
        resp = requests.request(method, url, headers=headers, timeout=300, **kwargs)
        if resp.status_code in (429, 503, 504):
            time.sleep(int(resp.headers.get("Retry-After", 2 ** attempt)))
            continue
        if not resp.ok:
            raise GraphError(resp.status_code, resp.text)
        return resp
    raise GraphError(429, "gave up after repeated throttling")


# ---------------------------------------------------------- sharepoint ----
def get_drive_id(token, sp):
    if not sp.get("site_host"):
        raise SystemExit("SharePoint site is not set.")
    site = call("GET", f"{GRAPH}/sites/{sp['site_host']}:{sp.get('site_path', '')}",
                token).json()
    drives = call("GET", f"{GRAPH}/sites/{site['id']}/drives", token).json()["value"]
    for d in drives:
        if d["name"].lower() == sp["library"].lower():
            return d["id"]
    names = ", ".join(d["name"] for d in drives)
    raise SystemExit(f"Library '{sp['library']}' not found. Available: {names}")


TYPE_KEYS = ["text", "choice", "dateTime", "number", "boolean", "currency",
             "personOrGroup", "lookup", "term", "hyperlinkOrPicture", "calculated"]
SPECIAL = {"personOrGroup", "lookup", "term", "hyperlinkOrPicture"}


def get_columns(token, drive_id):
    """Return the library's editable, visible columns."""
    url = f"{GRAPH}/drives/{drive_id}/list/columns"
    cols = []
    while url:
        page = call("GET", url, token).json()
        for c in page["value"]:
            if c.get("readOnly") or c.get("hidden"):
                continue
            kind = next((k for k in TYPE_KEYS if k in c), "other")
            if kind == "calculated":
                continue
            cols.append({"name": c["name"], "display": c.get("displayName", c["name"]),
                         "type": kind, "required": bool(c.get("required")),
                         "special": kind in SPECIAL})
        url = page.get("@odata.nextLink")
    return sorted(cols, key=lambda c: c["display"].lower())


_folders_done = set()


def ensure_folder(token, drive_id, folder):
    """Create the folder path in the library if it does not exist yet."""
    parts = [p for p in folder.split("/") if p]
    for i in range(len(parts)):
        path = "/".join(parts[: i + 1])
        if (drive_id, path.lower()) in _folders_done:
            continue
        parent = "/".join(parts[:i])
        url = (f"{GRAPH}/drives/{drive_id}/root:/{quote(parent)}:/children"
               if parent else f"{GRAPH}/drives/{drive_id}/root/children")
        try:
            call("POST", url, token, json={
                "name": parts[i], "folder": {},
                "@microsoft.graph.conflictBehavior": "fail"})
        except GraphError as e:
            if e.status != 409:           # 409 = already exists, which is fine
                raise
        _folders_done.add((drive_id, path.lower()))


def upload(token, drive_id, local_path, remote_path, overwrite):
    """Upload one file; returns the driveItem JSON. Raises GraphError 409 if
    the file exists and overwrite is False."""
    size = os.path.getsize(local_path)
    behavior = "replace" if overwrite else "fail"
    base = f"{GRAPH}/drives/{drive_id}/root:/{quote(remote_path)}:"

    if size <= SMALL_FILE_LIMIT:
        with open(local_path, "rb") as f:
            return call(
                "PUT", f"{base}/content?@microsoft.graph.conflictBehavior={behavior}",
                token, data=f.read(),
                headers={"Content-Type": "application/octet-stream"}).json()

    session = call("POST", f"{base}/createUploadSession", token, json={
        "item": {"@microsoft.graph.conflictBehavior": behavior}}).json()
    upload_url = session["uploadUrl"]     # pre-authenticated: no bearer token
    resp = None
    with open(local_path, "rb") as f:
        start = 0
        while start < size:
            chunk = f.read(CHUNK_SIZE)
            end = start + len(chunk) - 1
            resp = call("PUT", upload_url, data=chunk, headers={
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{size}"})
            start = end + 1
    return resp.json()


def set_metadata(token, drive_id, item_id, fields):
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/listItem/fields"
    call("PATCH", url, token, json=fields)


# --------------------------------------------------------------- excel ----
def open_workbook(xl):
    import openpyxl
    if not xl.get("path") or not os.path.exists(xl["path"]):
        raise SystemExit(f"Excel file not found: {xl.get('path') or '(none chosen)'}")
    return openpyxl.load_workbook(xl["path"], read_only=True, data_only=True)


def open_sheet(xl):
    wb = open_workbook(xl)
    if xl.get("sheet"):
        if xl["sheet"] not in wb.sheetnames:
            raise SystemExit(f"Sheet '{xl['sheet']}' not found. "
                             f"Available: {', '.join(wb.sheetnames)}")
        return wb[xl["sheet"]]
    return wb.worksheets[0]


def read_headers(xl):
    first = next(open_sheet(xl).iter_rows(values_only=True), ())
    return [str(h).strip() if h is not None else "" for h in first]


def cell_value(v):
    """Convert an Excel cell to something Graph accepts."""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%dT%H:%M:%SZ")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%dT00:00:00Z")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        return v.strip()
    return v


def read_excel(xl, files):
    """Return ({relpath: {internal_name: value}}, [problem strings])."""
    rows = open_sheet(xl).iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else "" for h in next(rows, ())]
    file_col = xl.get("file_column", "")
    if file_col not in headers or not file_col:
        raise SystemExit(f"Excel sheet has no '{file_col}' column in row 1. "
                         f"Choose the file-name column. Headers found: "
                         f"{', '.join(h for h in headers if h)}")
    name_idx = headers.index(file_col)
    mapping = {h: t for h, t in xl.get("columns", {}).items() if t}

    problems = []
    for h in mapping:
        if h not in headers:
            problems.append(f"'{h}' is mapped but the sheet has no such header")
    ignored = [h for h in headers if h and h != file_col and h not in mapping]
    if ignored:
        problems.append(f"Ignored (unmapped) Excel columns: {', '.join(ignored)}")

    by_rel = {f.lower(): f for f in files}
    by_name = {}
    for f in files:
        by_name.setdefault(f.rsplit("/", 1)[-1].lower(), []).append(f)

    metadata = {}
    for rownum, row in enumerate(rows, start=2):
        raw = row[name_idx] if name_idx < len(row) else None
        if raw is None or str(raw).strip() == "":
            continue
        key = str(raw).strip().replace("\\", "/").strip("/").lower()
        if key in by_rel:
            target = by_rel[key]
        elif len(by_name.get(key, [])) == 1:
            target = by_name[key][0]
        elif len(by_name.get(key, [])) > 1:
            problems.append(f"Row {rownum}: '{raw}' matches several files; "
                            f"use the relative path instead")
            continue
        else:
            problems.append(f"Row {rownum}: '{raw}' not found in the source")
            continue
        if target in metadata:
            problems.append(f"Row {rownum}: duplicate row for '{target}' (later row wins)")
        fields = {}
        for i, h in enumerate(headers):
            if h in mapping and i < len(row):
                v = cell_value(row[i])
                if v is not None and v != "":
                    fields[mapping[h]] = v
        metadata[target] = fields
    return metadata, problems


def auto_map(headers, cols, existing=None):
    """Match Excel headers to SharePoint columns by display or internal name.
    Choices already present in `existing` are kept."""
    existing = existing or {}
    by_display = {c["display"].lower(): c for c in cols}
    by_name = {c["name"].lower(): c for c in cols}
    mapping = {}
    for h in headers:
        if h in existing:
            mapping[h] = existing[h]
        else:
            match = by_display.get(h.lower()) or by_name.get(h.lower())
            mapping[h] = match["name"] if match else ""
    return mapping


# ------------------------------------------------------------ pipeline ----
def target_folder(cfg):
    return cfg["sharepoint"].get("target_folder", "").strip("/").replace("\\", "/")


def plan(cfg, source):
    """List the source and read the Excel sheet. No SharePoint access.
    Only files that have a row in the sheet are returned for migration."""
    all_files = source.list()
    metadata, problems = read_excel(cfg["excel"], all_files)
    files = [f for f in all_files if f in metadata]   # the sheet decides what moves
    left = len(all_files) - len(files)
    if left:
        problems.insert(0, f"{left} files in the source are not listed in the "
                           f"sheet and will be left alone")
    return files, metadata, problems


def preview_rows(cfg, files, metadata):
    target = target_folder(cfg)
    return [{"file": f,
             "remote": f"{target}/{f}" if target else f,
             "columns": list(metadata.get(f, {})),
             "invalid": bool(BAD_CHARS & set(f))} for f in files]


def migrate(cfg, source, files, metadata, overwrite=False, move=False,
            log_path="migration_log.csv", on_start=None, on_result=None,
            should_stop=None):
    """Upload every file and set its metadata. Returns the status counts."""
    sp = cfg["sharepoint"]
    target = target_folder(cfg)
    drive_id = get_drive_id(get_token(sp), sp)
    counts = {"ok": 0, "skipped": 0, "failed": 0}
    with open(log_path, "w", newline="", encoding="utf-8") as logf:
        log = csv.writer(logf)
        log.writerow(["file", "status", "detail", "url"])
        for n, f in enumerate(files, 1):
            if should_stop and should_stop():
                break
            if on_start:
                on_start(n, f)
            remote = f"{target}/{f}" if target else f
            status, notes, url, local = "ok", [], "", None
            try:
                if BAD_CHARS & set(f):
                    raise ValueError("name contains a character SharePoint does not allow")
                local = source.fetch(f)
                local_size = os.path.getsize(local)
                token = get_token(sp)
                if "/" in remote:
                    ensure_folder(token, drive_id, remote.rsplit("/", 1)[0])
                item = upload(token, drive_id, local, remote, overwrite)
                url = item.get("webUrl", "")
                if metadata.get(f):
                    try:
                        set_metadata(token, drive_id, item["id"], metadata[f])
                        notes.append("metadata set")
                    except GraphError as e:
                        status = "failed"
                        notes.append(f"uploaded, but metadata failed: {e}")
                if move and status == "ok":
                    # SharePoint can rewrite Office files on upload and
                    # change their size; a mismatch keeps the source.
                    if item.get("size") == local_size:
                        source.delete(f)
                        notes.append("source deleted")
                    else:
                        notes.append("source KEPT (size differs after upload, check manually)")
            except GraphError as e:
                if e.status == 409:
                    status, notes = "skipped", ["already exists"]
                else:
                    status, notes = "failed", [str(e)]
            except Exception as e:        # one bad file should not stop the run
                status, notes = "failed", [f"{type(e).__name__}: {e}"]
            finally:
                if local:
                    source.release(local)
            detail = "; ".join(notes)
            counts[status] += 1
            log.writerow([f, status, detail, url])
            logf.flush()
            if on_result:
                on_result(n, f, status, detail, url)
    return counts


def write_template(cfg, files):
    import openpyxl
    xl = cfg["excel"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = xl.get("sheet") or "Metadata"
    ws.append([xl.get("file_column") or "FileName"])
    for f in files:
        ws.append([f])
    ws.column_dimensions["A"].width = 60
    ws.freeze_panes = "A2"
    wb.save(xl["path"])


# ------------------------------------------------------------ commands ----
def load_config(path):
    if not os.path.exists(path):
        raise SystemExit(f"{path} not found. Run:  python sp_migrate.py init")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_config(path, cfg):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def cmd_init(args):
    if os.path.exists(args.config):
        raise SystemExit(f"{args.config} already exists; not overwriting it.")
    save_config(args.config, STARTER_CONFIG)
    print(f"Wrote {args.config}. Fill in the sharepoint, source and excel sections.")


def cmd_template(args):
    cfg = load_config(args.config)
    xl = cfg["excel"]
    if not xl.get("path"):
        xl["path"] = "metadata.xlsx"
        save_config(args.config, cfg)
    if os.path.exists(xl["path"]):
        raise SystemExit(f"{xl['path']} already exists; not overwriting it.")
    source = open_source(cfg)
    try:
        files = source.list()
    finally:
        source.close()
    write_template(cfg, files)
    print(f"Wrote {xl['path']} listing {len(files)} files. Add a column for each "
          f"piece of metadata, then run:  python sp_migrate.py map")


def cmd_columns(args):
    cfg = load_config(args.config)
    token = get_token(cfg["sharepoint"])
    cols = get_columns(token, get_drive_id(token, cfg["sharepoint"]))
    print(f"{'Display name':32} {'Internal name':32} Type")
    for c in cols:
        note = "  (required)" if c["required"] else ""
        note += "  (needs special handling)" if c["special"] else ""
        print(f"{c['display'][:31]:32} {c['name'][:31]:32} {c['type']}{note}")


def cmd_map(args):
    cfg = load_config(args.config)
    xl = cfg["excel"]
    headers = [h for h in read_headers(xl) if h and h != xl["file_column"]]
    token = get_token(cfg["sharepoint"])
    cols = get_columns(token, get_drive_id(token, cfg["sharepoint"]))
    kept = {h: t for h, t in xl.get("columns", {}).items() if t}
    xl["columns"] = auto_map(headers, cols, kept)
    by_name = {c["name"]: c for c in cols}
    print(f"{'Excel header':32} -> SharePoint column")
    for h, t in xl["columns"].items():
        c = by_name.get(t)
        warn = "  ! needs special handling" if c and c["special"] else ""
        print(f"{h[:31]:32} -> {t or '(no match: will be ignored)'}{warn}")
    save_config(args.config, cfg)
    print(f"\nSaved to {args.config}. Edit excel.columns there to change any match.")


def cmd_run(args):
    cfg = load_config(args.config)
    source = open_source(cfg)
    try:
        files, metadata, problems = plan(cfg, source)
        print(f"{len(files)} files listed in the sheet were found in {source.label}")
        for p in problems:
            print(f"  ! {p}")

        if args.dry_run:
            print("\nDRY RUN: nothing will be uploaded or deleted\n")
            for r in preview_rows(cfg, files, metadata):
                cols = ", ".join(r["columns"]) or "(no metadata)"
                flag = "  ! invalid character in name" if r["invalid"] else ""
                print(f"  {r['file']}  ->  {cfg['sharepoint']['library']}/"
                      f"{r['remote']}   [{cols}]{flag}")
            return

        total = len(files)

        def on_result(n, f, status, detail, url):
            print(f"[{n}/{total}] {f}: {status}" + (f" ({detail})" if detail else ""))

        counts = migrate(cfg, source, files, metadata, args.overwrite, args.move,
                         args.log, on_result=on_result)
        print(f"\nDone. {counts['ok']} uploaded, {counts['skipped']} skipped, "
              f"{counts['failed']} failed. Details in {args.log}")
        if counts["failed"]:
            sys.exit(1)
    finally:
        source.close()


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default="config.json")
    sub = ap.add_subparsers(dest="command", required=True)
    sub.add_parser("init").set_defaults(func=cmd_init)
    sub.add_parser("template").set_defaults(func=cmd_template)
    sub.add_parser("columns").set_defaults(func=cmd_columns)
    sub.add_parser("map").set_defaults(func=cmd_map)
    run = sub.add_parser("run")
    run.add_argument("--dry-run", action="store_true")
    run.add_argument("--overwrite", action="store_true",
                     help="replace files that already exist (default: skip)")
    run.add_argument("--move", action="store_true",
                     help="delete each source file after a verified upload")
    run.add_argument("--log", default="migration_log.csv")
    run.set_defaults(func=cmd_run)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
