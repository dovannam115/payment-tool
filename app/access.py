# Who may use this copy. The list lives in app/access.json on GitHub and is read again at every start,
# so a machine can be allowed or cut off by pushing that one file, without a new version.
#   enforce false -> everybody runs, each start is only reported (machine, Windows user, version)
#   enforce true  -> only machines whose hash is in "allowed" can run a step
#   "blocked"     -> machines locked either way (one PC can be cut off before the list is enforced)
# A copy that cannot reach GitHub uses the newer ("rev") of dump\access_cache.json and app\access.json.
import os, json, time, base64, socket, ctypes, hashlib, threading, urllib.parse, urllib.request

APP   = os.path.dirname(os.path.abspath(__file__))
LOCAL = os.path.join(APP, "access.json")
CACHE = os.path.join(APP, "dump", "access_cache.json")
SALT  = "pvi-apps|"
STATE = {"ok": True, "enforce": False, "machine": "", "user": "", "contact": "", "source": ""}
READY = threading.Event()


def machine():
    return (socket.gethostname() or os.environ.get("COMPUTERNAME", "")).upper()


def user(kind=2):
    """Windows sign-in: 2 = DOMAIN\\user, 3 = display name (needs the domain, can be empty)."""
    try:
        n = ctypes.c_ulong(256); buf = ctypes.create_unicode_buffer(256)
        if ctypes.windll.secur32.GetUserNameExW(kind, buf, ctypes.byref(n)): return buf.value
    except Exception:
        pass
    return "%s\\%s" % (os.environ.get("USERDOMAIN", ""), os.environ.get("USERNAME", "")) if kind == 2 else ""


def key(name):
    return hashlib.sha256((SALT + name.strip().upper()).encode("utf-8")).hexdigest()


def _read(path):
    try: return json.load(open(path, encoding="utf-8"))
    except Exception: return None


def _get(url, token, timeout=8):
    req = urllib.request.Request(url, headers={"User-Agent": "PaymentTool", "Cache-Control": "no-cache"})
    if token: req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()


def _remote():
    """API first (raw.githubusercontent keeps the old file for a few minutes), raw as the fallback."""
    c = {"repo": "", "branch": "main", "token": ""}
    c.update(_read(os.path.join(APP, "update.json")) or {})
    if not c["repo"]: return None
    repo, branch, token = c["repo"], c["branch"] or "main", c["token"]
    try:
        j = json.loads(_get("https://api.github.com/repos/%s/contents/app/access.json?ref=%s" % (repo, branch), token))
        return json.loads(base64.b64decode(j["content"]).decode("utf-8"))
    except Exception:
        pass
    try:
        return json.loads(_get("https://raw.githubusercontent.com/%s/%s/app/access.json" % (repo, branch), token))
    except Exception:
        return None


def _apply(j, source):
    m = machine()
    STATE.update(machine=m, user=user(), contact=j.get("contact", ""), source=source,
                 enforce=bool(j.get("enforce")))
    k = key(m)
    STATE["ok"] = k not in set(j.get("blocked") or []) and (not STATE["enforce"] or k in set(j.get("allowed") or []))


def _report(j, app, ver):
    r = j.get("report") or {}
    if not r.get("url"): return
    row = {"machine": STATE["machine"], "user": STATE["user"], "name": user(3), "app": app,
           "version": ver, "allowed": "yes" if STATE["ok"] else "no", "time": time.strftime("%Y-%m-%d %H:%M:%S")}
    try:
        if r.get("type") == "gform":       # Google Form: field -> entry.NNN
            data = urllib.parse.urlencode({f: row.get(k, "") for k, f in (r.get("fields") or {}).items()}).encode()
            req = urllib.request.Request(r["url"], data=data, headers={"User-Agent": "Mozilla/5.0"})
        else:                              # anything that takes a JSON POST (Power Automate HTTP trigger...)
            req = urllib.request.Request(r["url"], data=json.dumps(row).encode("utf-8"),
                                         headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=15).read()
    except Exception:
        pass


def check(app, ver, report=True):
    """Saved list first (instant), then GitHub; reports the start (not a re-check). Never shows anything."""
    have = [(x, n) for x, n in ((_read(CACHE), "cache"), (_read(LOCAL), "local")) if x]
    j, src = max(have, key=lambda h: h[0].get("rev", 0)) if have else ({}, "none")
    _apply(j, src)
    rj = _remote()
    if rj is not None:
        j = rj; _apply(j, "github")
        try:
            os.makedirs(os.path.dirname(CACHE), exist_ok=True)
            json.dump(rj, open(CACHE, "w", encoding="utf-8"))
        except Exception:
            pass
    READY.set()
    if report: _report(j, app, ver)


def start(app, ver):
    READY.clear()
    threading.Thread(target=check, args=(app, ver), daemon=True).start()


def status(wait=8):
    READY.wait(wait)
    return dict(STATE)
