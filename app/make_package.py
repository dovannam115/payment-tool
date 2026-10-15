# Build dist\Payment_Tool.zip to hand over. Run: app\pyembed\python.exe app\make_package.py
# Ships code + pyembed + node + Invoice Check + people.xlsx, with EMPTY data\ and done\.
# Left out: .edge-profile (Nam's portal login), data\ and done\ contents (real invoices),
# app\dump\* (old runs), __pycache__, this script.
import os, zipfile
from layout import ROOT, APP

TOP   = "Payment-Tool"
OUT   = os.path.join(ROOT, "dist", "Payment_Tool.zip")
FILES = ["Payment Tool.bat", "Operating_Guide.html", "people.xlsx", "VERSION", "RELEASE-NOTES.txt"]
EMPTY = ["data/", "data/_hold/", "done/", "app/dump/"]

def skip_app(rel):
    parts = rel.replace("\\", "/").split("/")
    if "__pycache__" in parts: return True
    if parts[0] == "dump": return True
    if parts[0].startswith("_backup_"): return True
    if rel == "make_package.py": return True
    return False

def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    n = 0
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for f in FILES:
            z.write(os.path.join(ROOT, f), f"{TOP}/{f}"); n += 1
        for d in EMPTY:
            z.writestr(f"{TOP}/{d}", "");
        for root, dirs, files in os.walk(APP):
            for f in files:
                full = os.path.join(root, f); rel = os.path.relpath(full, APP)
                if skip_app(rel): continue
                z.write(full, f"{TOP}/app/{rel.replace(os.sep, '/')}"); n += 1
    # sanity: files the tool cannot start without
    must = ["app/pyembed/python.exe", "app/pyembed/python312.zip", "app/pyembed/python312._pth",
            "app/pyembed/Lib/site-packages/openpyxl/__init__.py", "app/node.exe",
            "app/node_modules/playwright-core/package.json", "app/node_modules/pdf-lib/package.json",
            "app/invoice_check/check-print.js", "app/run.py", "app/layout.py",
            "app/server.py", "app/ui/index.html", "app/login.js", "app/update.json",
            "app/fill_don.js", "app/build_input.py", "VERSION"]
    names = set(zipfile.ZipFile(OUT).namelist())
    miss = [m for m in must if f"{TOP}/{m}" not in names]
    if any(".edge-profile" in x for x in names): miss.append("LEAK: .edge-profile")
    print(f"{n} files -> {OUT} ({os.path.getsize(OUT)/1e6:.1f} MB)")
    print("MISSING: " + ", ".join(miss) if miss else "Check OK: all required files inside.")

if __name__ == "__main__":
    main()
