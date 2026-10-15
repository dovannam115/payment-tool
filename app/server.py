# Local web UI for Payment Tool. Binds 127.0.0.1 only, serves app\ui and runs each step as a job.
# Started by "Payment Tool.bat" with pyembed\pythonw.exe (no console window).
#   python.exe app\server.py --diag   -> same server with the log in a console (for troubleshooting)
import os, re, sys, json, time, glob, codecs, shutil, socket, zipfile, ctypes, tempfile, threading, subprocess, webbrowser
import urllib.request, urllib.error
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from layout import APP, ROOT, DATA, DONE, people_dirs, scan, read_check_log, check_problem

PORT    = 8711
HOST    = "127.0.0.1"
UI      = os.path.join(APP, "ui")
DUMP    = os.path.join(APP, "dump")
VERFILE = os.path.join(ROOT, "VERSION")
CFGFILE = os.path.join(APP, "update.json")
TOKFILE = os.path.join(DUMP, "ui_token.txt")
LOGFILE = os.path.join(DUMP, "server.log")
INPUT   = os.path.join(DATA, "Request_Input.xlsx")
CHECKLOG= os.path.join(DATA, "Check_Log.xlsx")
GUIDE   = os.path.join(ROOT, "Operating_Guide.html")
BAT     = os.path.join(ROOT, "Payment Tool.bat")
DIAG    = "--diag" in sys.argv

PY   = [sys.executable] if sys.executable else ["py", "-3.12"]
NODE = os.path.join(APP, "node.exe")
if not os.path.exists(NODE): NODE = "node"
NO_WINDOW = 0x08000000 if os.name == "nt" else 0
ANSI = re.compile(r"\x1b\[[0-9;]*[A-Za-z]")   # colour codes in child output


def log(msg):
    line = time.strftime("%H:%M:%S ") + str(msg)
    if DIAG: print(line, flush=True)
    try:
        os.makedirs(DUMP, exist_ok=True)
        with open(LOGFILE, "a", encoding="utf-8") as f: f.write(line + "\n")
    except Exception: pass


def box(msg, title="Payment Tool"):
    """Fatal error popup: pythonw has no console, so a message box is the only way to be seen."""
    if DIAG: print(msg, flush=True); return
    try: ctypes.windll.user32.MessageBoxW(0, str(msg), title, 0x10)
    except Exception: pass


def version():
    try: return open(VERFILE, encoding="utf-8").read().strip() or "0.0.0"
    except Exception: return "0.0.0"


def cfg():
    out = {"repo": "", "branch": "main", "token": ""}
    try: out.update(json.load(open(CFGFILE, encoding="utf-8")))
    except Exception: pass
    return out


# ------------------------------------------------------------------ jobs
class Job:
    """One step of the tool = a chain of child processes, output streamed to the page.
    A child that prints '[[ASK]] <text>' is waiting for the user to click a button here;
    the answer is a newline written to its stdin."""
    def __init__(self, action, steps, label):
        self.action, self.steps, self.label = action, steps, label
        self.lines, self.ask, self.rc = [], None, None
        self.progress = ""              # last \r line (a progress bar), shown on its own row
        self.state = "running"          # running | waiting | login | done
        self.started = time.time()
        self.proc = None
        self.stopped = False
        self.lock = threading.Lock()

    def put(self, text):
        with self.lock:
            self.lines.append(text); self.progress = ""

    def answer(self, text=""):
        p = self.proc
        self.ask = None
        if self.state in ("waiting", "login"): self.state = "running"
        if p and p.poll() is None:
            try:
                p.stdin.write(((text or "") + "\n").encode("utf-8")); p.stdin.flush(); return True
            except Exception: return False
        return False

    def feed(self, text, ended_with_cr):
        """One chunk of child output. A chunk that ended with \\r is a progress bar redrawing
        itself on the same line, so it replaces the progress row instead of piling up."""
        text = ANSI.sub("", text).rstrip()      # Invoice Check colours its bar, the page shows plain text
        if not text: return
        if ended_with_cr:
            with self.lock: self.progress = text
            return
        if text.startswith("[[ASK]]"):
            self.ask = text[7:].strip(); self.state = "waiting"
        elif text.startswith("[[LOGIN]]"):
            self.state = "login"; self.put(text[9:].strip())
        else:
            if self.state == "login": self.state = "running"   # user is through, work went on
            self.put(text)

    def kill_tree(self, p):
        """/T takes the children too, else the Edge window playwright opened stays behind."""
        if not p or p.poll() is not None: return
        try: subprocess.call(["taskkill", "/F", "/T", "/PID", str(p.pid)], creationflags=NO_WINDOW,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception: pass
        try: p.kill()
        except Exception: pass

    def stop(self):
        self.stopped = True
        self.kill_tree(self.proc)

    def run(self):
        try:
            for cmd, env in self.steps:
                if self.stopped: break
                e = dict(os.environ); e.update(env or {}); e["PYTHONIOENCODING"] = "utf-8"
                script = next((a for a in reversed(cmd) if a.lower().endswith((".py", ".js"))), cmd[0])
                self.put("> " + os.path.basename(script))
                # raw bytes, not text mode: text mode turns every \r into \n, and then a progress
                # bar redrawing itself would look like hundreds of separate lines.
                p = subprocess.Popen(cmd, cwd=APP, env=e, stdin=subprocess.PIPE,
                                     stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                     bufsize=0, creationflags=NO_WINDOW)
                self.proc = p
                dec = codecs.getincrementaldecoder("utf-8")(errors="replace")
                buf, cr = "", False
                while True:
                    chunk = p.stdout.read(4096)      # unbuffered pipe: returns as soon as there is data
                    if not chunk: break
                    for ch in dec.decode(chunk):
                        if ch == "\r":
                            cr = True                # wait for the next character to know what it meant
                        elif ch == "\n":
                            self.feed(buf, False); buf = ""; cr = False      # \r\n and \n are both a line
                        else:
                            if cr:                   # \r then more text = a progress bar redrawing itself
                                self.feed(buf, True); buf = ""; cr = False
                            buf += ch
                if buf: self.feed(buf, cr)
                try:
                    rc = p.wait(timeout=25)          # stdout is closed, so this is only a safety net
                except subprocess.TimeoutExpired:
                    self.put("[step finished but did not close, closing it]")
                    self.kill_tree(p); rc = p.wait()
                self.rc = rc
                if rc != 0:
                    self.put("[exit code %d]" % rc)
                    if not self.stopped: break
        except Exception as ex:
            self.put("ERROR: %s" % ex); log("job %s failed: %r" % (self.action, ex))
        self.state = "done"; self.ask = None
        if self.stopped: self.put("[stopped]")

    def view(self, since=0):
        with self.lock: lines = self.lines[since:]
        return {"action": self.action, "label": self.label, "state": self.state,
                "ask": self.ask, "rc": self.rc, "total": len(self.lines),
                "progress": self.progress, "since": since, "lines": lines}


class Runner:
    def __init__(self): self.job = None
    @property
    def busy(self): return bool(self.job and self.job.state != "done")
    def start(self, job):
        if self.busy: return False
        self.job = job
        threading.Thread(target=job.run, daemon=True).start()
        return True


RUN = Runner()


def step_check(names, redo):
    args = list(names) + (["--all"] if redo else [])
    return [(PY + ["-u", os.path.join(APP, "check_invoices.py")] + args, {})]


def step_fill(names):
    env = {"TB_WEB": "1"}
    if names: env["TB_ONLY"] = ";".join(names)
    return [(PY + ["-u", os.path.join(APP, "build_input.py")], {}),
            (PY + ["-u", os.path.join(APP, "don_to_json.py")], {}),
            ([NODE, os.path.join(APP, "fill_don.js")], env)]


def step_archive(names):
    return [(PY + ["-u", os.path.join(APP, "archive_done.py")] + list(names), {})]


def step_login():
    return [([NODE, os.path.join(APP, "login.js")], {"TB_WEB": "1"})]


# ------------------------------------------------------------------ state
def person_rows():
    rows = []
    for name in people_dirs():
        invs, totrinh = scan(name)
        if not invs and not totrinh: continue
        logrows = read_check_log(name)
        cdir = os.path.join(DATA, name, "check")
        have = [f.lower() for f in os.listdir(cdir)] if os.path.isdir(cdir) else []
        checks = len([f for f in have if f.endswith(".pdf")])
        # same rule as check_invoices.pending(): in log.csv, or a check PDF whose name starts
        # with the invoice name (people rename them to ...-044check.pdf)
        def is_checked(iv):
            if os.path.basename(iv["xml"]) in logrows: return True
            s = iv["stem"].lower()
            return any(f.startswith(s) and f.endswith(".pdf") for f in have)
        detail, problems, pending = [], [], 0
        for iv in invs:
            r = logrows.get(os.path.basename(iv["xml"]))
            prob = check_problem(r) if r else ""
            done = is_checked(iv)
            if not done: pending += 1
            if prob: problems.append("%s: %s" % (iv["stem"], prob))
            detail.append({"stem": iv["stem"], "so": iv["sh"], "kh": iv["kh"], "ngay": iv["ngay"],
                           "tong": iv["tong"], "seller": iv["seller"],
                           "pdf": bool(iv["pdf"]), "check": done, "problem": prob})
        rows.append({"name": name, "invoices": len(invs),
                     "pdf": sum(1 for i in invs if i["pdf"]), "checks": checks,
                     "totrinh": len(totrinh), "checked": bool(logrows),
                     "pending_check": pending,
                     "problems": problems, "detail": detail})
    return rows


def state():
    j = RUN.job
    return {"version": version(), "people": person_rows(),
            "has_input": os.path.exists(INPUT), "has_checklog": os.path.exists(CHECKLOG),
            "data_dir": DATA, "busy": RUN.busy,
            "job": ({"action": j.action, "label": j.label, "state": j.state} if j else None),
            "update": UPD.get("cache")}


# ------------------------------------------------------------------ update from GitHub
UPD = {"cache": None, "checked": 0}
SKIP_DIRS  = {"node_modules", "pyembed", "dump", "data", "done", "dist", ".git",
              "__pycache__", ".edge-profile", ".vscode"}
KEEP_FILES = {"people.xlsx"}
ALLOW_EXT  = {".py", ".js", ".html", ".css", ".bat", ".txt", ".json", ".md"}
ALLOW_NAME = {"VERSION"}   # no extension, but it IS the file that says which version this is


def wanted(name):
    return name in ALLOW_NAME or os.path.splitext(name)[1].lower() in ALLOW_EXT


def http_get(url, token="", timeout=25):
    req = urllib.request.Request(url, headers={"User-Agent": "PaymentTool", "Cache-Control": "no-cache"})
    if token: req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()


def cmp_ver(a, b):
    na = [int(x) for x in str(a).replace("v", "").split(".") if x.isdigit()]
    nb = [int(x) for x in str(b).replace("v", "").split(".") if x.isdigit()]
    for i in range(max(len(na), len(nb))):
        x, y = (na[i] if i < len(na) else 0), (nb[i] if i < len(nb) else 0)
        if x != y: return 1 if x > y else -1
    return 0


def read_remote(c, name):
    """One file from the repo. The API answers with what was just pushed; raw.githubusercontent
    sits behind a CDN that keeps serving the old file for a few minutes, so it is only the fallback
    (the API allows 60 calls an hour per address)."""
    repo, branch, token = c["repo"], c["branch"] or "main", c["token"]
    try:
        api = "https://api.github.com/repos/%s/contents/%s?ref=%s" % (repo, name, branch)
        j = json.loads(http_get(api, token).decode("utf-8", "replace"))
        import base64
        return base64.b64decode(j.get("content", "")).decode("utf-8", "replace")
    except urllib.error.HTTPError as ex:
        if ex.code == 404: raise                       # repo or file really is not there
    except Exception:
        pass
    raw = "https://raw.githubusercontent.com/%s/%s/%s" % (repo, branch, name)
    return http_get(raw, token).decode("utf-8", "replace")


def update_check():
    c = cfg(); cur = version()
    if not c["repo"] or "/" not in c["repo"]:
        out = {"configured": False, "current": cur}
    else:
        try:
            latest = read_remote(c, "VERSION").strip()
            try: notes = read_remote(c, "RELEASE-NOTES.txt")
            except Exception: notes = ""
            out = {"configured": True, "current": cur, "latest": latest,
                   "available": cmp_ver(latest, cur) > 0, "notes": notes[:4000], "repo": c["repo"]}
        except urllib.error.HTTPError as ex:
            msg = ("chưa thấy repo %s trên GitHub" % c["repo"]) if ex.code in (403, 404) \
                  else "GitHub trả về lỗi %s" % ex.code
            out = {"configured": True, "current": cur, "error": msg, "repo": c["repo"]}
        except Exception as ex:
            out = {"configured": True, "current": cur,
                   "error": "không nối được GitHub (%s)" % str(ex)[:120], "repo": c["repo"]}
    UPD["cache"] = out; UPD["checked"] = time.time()
    return out


def update_apply(job):
    c = cfg()
    if not c["repo"]: job.put("update.json has no repo yet."); return
    url = "https://codeload.github.com/%s/zip/refs/heads/%s" % (c["repo"], c["branch"] or "main")
    job.put("Downloading %s ..." % url)
    data = http_get(url, c["token"], timeout=120)
    tmp = tempfile.mkdtemp(prefix="pt_upd_")
    try:
        zp = os.path.join(tmp, "src.zip")
        open(zp, "wb").write(data)
        with zipfile.ZipFile(zp) as z: z.extractall(tmp)
        roots = [os.path.join(tmp, d) for d in os.listdir(tmp) if os.path.isdir(os.path.join(tmp, d))]
        if not roots: job.put("Downloaded file has no source folder."); return
        src = roots[0]
        backup = os.path.join(ROOT, "app", "_backup_" + time.strftime("%Y%m%d_%H%M%S"))
        n = 0
        for base, dirs, files in os.walk(src):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            for f in files:
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src).replace("\\", "/")
                if f in KEEP_FILES or not wanted(f): continue
                dst = os.path.join(ROOT, rel.replace("/", os.sep))
                if os.path.exists(dst):
                    b = os.path.join(backup, rel.replace("/", os.sep))
                    os.makedirs(os.path.dirname(b), exist_ok=True); shutil.copy2(dst, b)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(full, dst); n += 1
                job.put("  " + rel)
        job.put("")
        job.put("%d file(s) updated. Old copies kept in %s" % (n, os.path.basename(backup) if n else "-"))
        job.put("Version on disk is now %s. Restart the tool to use it." % version())
        UPD["cache"] = None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def update_local(zip_path, job):
    if not os.path.exists(zip_path): job.put("File not found: %s" % zip_path); return
    job.put("Installing from %s" % zip_path)
    tmp = tempfile.mkdtemp(prefix="pt_upd_")
    try:
        with zipfile.ZipFile(zip_path) as z: z.extractall(tmp)
        roots = [os.path.join(tmp, d) for d in os.listdir(tmp) if os.path.isdir(os.path.join(tmp, d))]
        src = roots[0] if len(roots) == 1 and not glob.glob(os.path.join(tmp, "*.py")) else tmp
        n = 0
        for base, dirs, files in os.walk(src):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            for f in files:
                if f in KEEP_FILES or not wanted(f): continue
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src)
                dst = os.path.join(ROOT, rel)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(full, dst); n += 1; job.put("  " + rel.replace("\\", "/"))
        job.put("")
        job.put("%d file(s) updated. Restart the tool to use it." % n)
        UPD["cache"] = None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def job_update(mode, zip_path=""):
    j = Job("update", [], "Cập nhật")
    def body():
        try:
            if mode == "local": update_local(zip_path, j)
            else: update_apply(j)
        except urllib.error.HTTPError as ex:
            j.put("GitHub returned %s. Check the repo name and branch in app\\update.json." % ex.code)
        except Exception as ex:
            j.put("Update failed: %s" % ex)
        j.state = "done"
    j.run = body
    return j


# ------------------------------------------------------------------ http
TOKEN = ""
LAST_PING = [time.time()]
SHUTDOWN = []


class Handler(BaseHTTPRequestHandler):
    server_version = "PaymentTool"
    protocol_version = "HTTP/1.1"

    def log_message(self, *a): pass

    def _send(self, code, body=b"", ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try: self.wfile.write(body)
        except Exception: pass

    def json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def ok_token(self, q):
        t = self.headers.get("X-Token") or (q.get("t", [""])[0])
        return t == TOKEN

    def body(self):
        try: n = int(self.headers.get("Content-Length") or 0)
        except ValueError: n = 0
        if not n: return {}
        try: return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception: return {}

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query); p = u.path
        if p in ("/", "/index.html"):
            f = os.path.join(UI, "index.html")
            if not os.path.exists(f): return self._send(404, b"ui missing", "text/plain")
            return self._send(200, open(f, "rb").read(), "text/html; charset=utf-8")
        if p == "/guide":
            if not os.path.exists(GUIDE): return self._send(404, b"no guide", "text/plain")
            return self._send(200, open(GUIDE, "rb").read(), "text/html; charset=utf-8")
        if p == "/api/state":
            if not self.ok_token(q): return self.json({"error": "token"}, 403)
            return self.json(state())
        if p == "/api/job":
            if not self.ok_token(q): return self.json({"error": "token"}, 403)
            j = RUN.job
            if not j: return self.json({"job": None})
            since = int((q.get("since", ["0"])[0]) or 0)
            return self.json({"job": j.view(since)})
        if p == "/api/update/check":
            if not self.ok_token(q): return self.json({"error": "token"}, 403)
            return self.json(update_check())
        return self._send(404, b"not found", "text/plain")

    def do_POST(self):
        u = urlparse(self.path); q = parse_qs(u.query); p = u.path
        if not self.ok_token(q): return self.json({"error": "token"}, 403)
        d = self.body()
        if p == "/api/ping":
            LAST_PING[0] = time.time(); return self.json({"ok": True})
        if p == "/api/run":
            action = d.get("action"); names = d.get("names") or []
            if RUN.busy: return self.json({"error": "busy"}, 409)
            if action == "check":
                job = Job(action, step_check(names, bool(d.get("redo"))), "Kiểm tra hóa đơn")
            elif action == "fill":
                job = Job(action, step_fill(names), "Điền đơn lên cổng")
            elif action == "archive":
                if not names: return self.json({"error": "no names"}, 400)
                job = Job(action, step_archive(names), "Chuyển sang done")
            elif action == "login":
                job = Job(action, step_login(), "Đăng nhập cổng")
            elif action == "update":
                job = job_update(d.get("mode") or "github", d.get("zip") or "")
            else:
                return self.json({"error": "unknown action"}, 400)
            RUN.start(job)
            return self.json({"ok": True})
        if p == "/api/answer":
            j = RUN.job
            return self.json({"ok": bool(j and j.answer(d.get("text") or ""))})
        if p == "/api/stop":
            j = RUN.job
            if j: j.stop()
            return self.json({"ok": True})
        if p == "/api/open":
            what = d.get("what"); target = {"input": INPUT, "checklog": CHECKLOG, "data": DATA,
                                            "done": DONE, "guide": GUIDE}.get(what)
            if what == "person": target = os.path.join(DATA, d.get("name") or "")
            if not target or not os.path.exists(target):
                return self.json({"error": "Chưa có file này"}, 404)
            try: os.startfile(target)
            except Exception as ex: return self.json({"error": str(ex)}, 500)
            return self.json({"ok": True})
        if p == "/api/restart":
            restart()
            return self.json({"ok": True})
        if p == "/api/quit":
            j = RUN.job
            if j: j.stop()
            threading.Timer(0.4, lambda: SHUTDOWN and SHUTDOWN[0]()).start()
            return self.json({"ok": True})
        return self.json({"error": "not found"}, 404)


def restart():
    """Start a detached helper that waits for this process to free the port, then runs the .bat."""
    helper = ("import os,time,subprocess\n"
              "time.sleep(2.0)\n"
              "os.startfile(r'%s')\n" % BAT)
    exe = sys.executable or "pythonw.exe"
    pw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    subprocess.Popen([pw if os.path.exists(pw) else exe, "-c", helper],
                     creationflags=NO_WINDOW | 0x00000008)
    threading.Timer(0.5, lambda: SHUTDOWN and SHUTDOWN[0]()).start()


def open_ui():
    url = "http://%s:%d/?t=%s" % (HOST, PORT, TOKEN)
    if os.environ.get("PT_NO_OPEN"): log("PT_NO_OPEN set, window not opened: " + url); return
    for exe in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
        if os.path.exists(exe):
            try:
                subprocess.Popen([exe, "--app=" + url, "--window-size=1280,900"],
                                 creationflags=NO_WINDOW)
                return
            except Exception: pass
    webbrowser.open(url)


def already_running():
    s = socket.socket()
    try:
        s.connect((HOST, PORT)); return True
    except Exception:
        return False
    finally:
        s.close()


def watchdog():
    """No page talking to us for 2 minutes and nothing running -> close the server."""
    while True:
        time.sleep(10)
        if RUN.busy: LAST_PING[0] = time.time(); continue
        if time.time() - LAST_PING[0] > 120:
            log("no page for 2 minutes, closing"); SHUTDOWN and SHUTDOWN[0](); return


def main():
    global TOKEN
    os.makedirs(DUMP, exist_ok=True); os.makedirs(DATA, exist_ok=True); os.makedirs(DONE, exist_ok=True)
    if already_running():
        try: TOKEN = open(TOKFILE, encoding="utf-8").read().strip()
        except Exception: TOKEN = ""
        log("already running, opening the window"); open_ui(); return
    TOKEN = os.urandom(16).hex()
    open(TOKFILE, "w", encoding="utf-8").write(TOKEN)
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as ex:
        box("Không mở được cổng %d (%s).\nĐóng tool đang chạy rồi mở lại." % (PORT, ex)); return
    SHUTDOWN.append(lambda: threading.Thread(target=srv.shutdown, daemon=True).start())
    log("listening on %s:%d  v%s" % (HOST, PORT, version()))
    threading.Thread(target=watchdog, daemon=True).start()
    threading.Thread(target=lambda: (time.sleep(1.5), update_check()), daemon=True).start()
    open_ui()
    try: srv.serve_forever()
    except KeyboardInterrupt: pass
    log("closed")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        import traceback
        log(traceback.format_exc())
        box("Tool gặp lỗi khi khởi động:\n\n%s\n\nLog: %s" % (ex, LOGFILE))
