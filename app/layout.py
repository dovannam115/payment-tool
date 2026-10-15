# Folder layout shared by every step.
#   data\<Ho ten>\          current batch, ONE folder per person, files dropped flat:
#       *.xml               invoice XML
#       *.pdf               invoice PDF: same name as an XML, else its number in the file name, else
#                           its number + serial/seller tax ID inside the PDF text; a PDF that reads
#                           like an invoice but fits no XML is "stray" (warned, never sent as to trinh);
#                           any other PDF -> to trinh nguon
#       check\*.pdf         invoice-check result (made by check_invoices.py)
#   data\_<name>\           folders starting with "_" are skipped (e.g. _hold = not this batch)
#   data\Request_Input.xlsx one row per person (made by build_input.py)
#   done\YYYY-MM\<Ho ten>\  people already saved on the portal (moved by archive_done.py)
import os, re, glob, sys, json, threading

APP    = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.dirname(APP)
DATA   = os.path.join(ROOT, "data")
DONE   = os.path.join(ROOT, "done")
PEOPLE = os.path.join(ROOT, "people.xlsx")
DUMP   = os.path.join(APP, "dump")

sys.path.insert(0, os.path.join(APP, "vendor"))   # pypdf ships as plain .py so the GitHub update carries it
try:
    import logging; logging.getLogger("pypdf").setLevel(logging.CRITICAL)
    import pypdf
except Exception:
    pypdf = None

def people_dirs():
    """Person folders in data\\ (skip _hold etc.)."""
    if not os.path.isdir(DATA): return []
    return [d for d in sorted(os.listdir(DATA))
            if os.path.isdir(os.path.join(DATA, d)) and not d.startswith(("_", "."))]

def read_invoice(xml_path):
    try:
        x=open(xml_path,encoding="utf-8",errors="ignore").read()
        def g(tag):
            m=re.search(rf"<{tag}>(.*?)</{tag}>", x, re.S); return m.group(1).strip() if m else ""
        mnt=re.search(r"<NBan>.*?<Ten>(.*?)</Ten>", x, re.S)
        mst=re.search(r"<NBan>.*?<MST>(.*?)</MST>", x, re.S)
        return {"kh":g("KHHDon"),"sh":g("SHDon"),"ngay":g("NLap"),
                "seller":(mnt.group(1).strip() if mnt else "")[:45],
                "mst":(mst.group(1).strip() if mst else ""),
                "tong":g("TgTTTBSo") or g("TgTCThueGTGT")}
    except Exception:
        return {"kh":"","sh":"","ngay":"","seller":"","mst":"","tong":""}

# ---- PDF text, cached by path + size + mtime (the UI rescans every few seconds)
_TXT, _TXT_LOCK, _TXT_FILE = {}, threading.Lock(), os.path.join(DUMP, "pdf_text.json")

def _load_cache():
    if _TXT or not os.path.exists(_TXT_FILE): return
    try: _TXT.update(json.load(open(_TXT_FILE, encoding="utf-8")))
    except Exception: pass

def pdf_text(path):
    """Upper-cased text of the first 2 pages, whitespace squeezed; '' for scans or without pypdf."""
    if not pypdf: return ""
    try: st=os.stat(path)
    except OSError: return ""
    key=os.path.abspath(path).lower(); sig="%d:%d" % (st.st_size, int(st.st_mtime))
    with _TXT_LOCK:
        _load_cache()
        hit=_TXT.get(key)
        if hit and hit[0]==sig: return hit[1]
    try:
        r=pypdf.PdfReader(path)
        txt=" ".join((pg.extract_text() or "") for pg in r.pages[:2])
    except Exception:
        txt=""
    txt=re.sub(r"\s+", " ", txt).upper()[:8000]
    with _TXT_LOCK:
        _TXT[key]=[sig, txt]
        try:
            os.makedirs(DUMP, exist_ok=True)
            json.dump(_TXT, open(_TXT_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception: pass
    return txt

def pdf_kind(txt):
    """'check' = an invoice-check result, 'invoice' = an invoice, '' = anything else (to trinh, scan)."""
    if not txt: return ""
    if "KẾT QUẢ KIỂM TRA HÓA ĐƠN" in txt or "KET QUA KIEM TRA HOA DON" in txt: return "check"
    if "TỜ TRÌNH" in txt: return ""
    if re.search(r"H[OÓ]A ?Đ[OƠ]N", txt) and re.search(r"KÝ HIỆU|SERIAL|MÃ CỦA CƠ QUAN THUẾ", txt)        and re.search(r"MÃ SỐ THUẾ|MST|TAX CODE", txt):
        return "invoice"
    return ""

def pdf_fits(iv, txt):
    """True when the PDF text carries this invoice's number plus its serial or seller tax ID."""
    no=iv.get("sh","").lstrip("0")
    if not no or not txt: return False
    if not re.search(rf"(?<!\d)0*{re.escape(no)}(?!\d)", txt): return False
    flat=txt.replace(" ", "")
    kh=iv.get("kh","").upper(); mst=iv.get("mst","")
    return bool((kh and kh in flat) or (mst and mst in flat))

def _stem(p): return os.path.splitext(os.path.basename(p))[0]

def scan(person):
    """Classify one person's folder -> invoices (xml + pdf + check) and to trinh PDFs."""
    r=scan_full(person)
    return r["invoices"], r["totrinh"]

def scan_full(person):
    """scan() plus what is off in the folder: 'stray' invoice-looking PDFs no XML claims,
    'misplaced' check results lying next to the invoices, 'other' files the tool ignores,
    and 'warnings' (short Vietnamese lines for the UI)."""
    pdir=os.path.join(DATA, person)
    xmls=sorted(glob.glob(os.path.join(pdir,"*.xml")))
    pdfs={_stem(p).lower():p for p in sorted(glob.glob(os.path.join(pdir,"*.pdf")))}
    checks={_stem(p).lower():p for p in glob.glob(os.path.join(pdir,"check","*.pdf"))}
    invs=[]
    for xp in xmls:
        s=_stem(xp)
        invs.append({"xml":xp,"stem":s,"pdf":pdfs.pop(s.lower(),""),"paired":"",
                     "check":checks.get(s.lower(),""),**read_invoice(xp)})
        if invs[-1]["pdf"]: invs[-1]["paired"]="name"
    # invoice PDF named differently from its XML: pair by invoice number found in the PDF name
    for iv in invs:
        if iv["pdf"]: continue
        no=iv["sh"].lstrip("0")
        if len(no)<3: continue
        hit=[k for k in pdfs if re.search(rf"(?<!\d)0*{re.escape(no)}(?!\d)", k)]
        if len(hit)==1: iv["pdf"]=pdfs.pop(hit[0]); iv["paired"]="filename"
    # still unpaired: read the leftover PDFs and pair by number + serial / tax ID inside them
    text={k:pdf_text(p) for k,p in pdfs.items()}
    for iv in invs:
        if iv["pdf"]: continue
        hit=[k for k in pdfs if pdf_fits(iv, text[k])]
        if len(hit)==1: iv["pdf"]=pdfs.pop(hit[0]); iv["paired"]="content"
    stray, misplaced, totrinh=[], [], []
    for k,p in sorted(pdfs.items(), key=lambda kv: kv[1]):
        kind=pdf_kind(text.get(k,""))
        (stray if kind=="invoice" else misplaced if kind=="check" else totrinh).append(p)
    other=[f for f in sorted(os.listdir(pdir)) if os.path.isfile(os.path.join(pdir,f))
           and os.path.splitext(f)[1].lower() not in (".xml",".pdf")
           and not f.startswith(("~$",".")) and f.lower()!="desktop.ini"]

    warn=[]
    for iv in invs:
        if not iv["pdf"]: warn.append(f"{iv['stem']}.xml chưa có PDF hóa đơn")
    for p in stray:
        warn.append(f"{os.path.basename(p)} giống hóa đơn nhưng không khớp XML nào (thiếu XML hoặc sai người)")
    for p in misplaced:
        warn.append(f"{os.path.basename(p)} là kết quả kiểm tra, nên nằm trong thư mục check")
    if invs and not totrinh: warn.append("Chưa có PDF tờ trình nguồn")
    if len(totrinh)>1:
        warn.append(f"{len(totrinh)} file được coi là tờ trình: " + ", ".join(os.path.basename(p) for p in totrinh))
    if other: warn.append("Bỏ qua file không phải XML/PDF: " + ", ".join(other))
    return {"invoices":invs, "totrinh":totrinh, "stray":stray, "misplaced":misplaced,
            "other":other, "warnings":warn}

LOG_COLS = ["File","Status","Seller","Seller tax ID","Company","Tax status","Output","Note"]

def read_check_log(person):
    """data\\<person>\\check\\log.csv (written by check_invoices.py) -> {xml file name: row dict}."""
    import csv
    f=os.path.join(DATA, person, "check", "log.csv")
    if not os.path.exists(f): return {}
    with open(f, encoding="utf-8-sig", newline="") as fh:
        return {r["File"]: r for r in csv.DictReader(fh) if r.get("File")}

def check_problem(row):
    """Short warning for one log row, '' if the invoice is valid and the seller is active."""
    out=[]
    st=row.get("Status") or "?"
    if st in ("ERROR","?"): return "chưa kiểm tra được (trang kiểm tra chậm), chạy lại bước Kiểm tra"
    if st!="VALID": out.append(f"hóa đơn {'không hợp lệ' if st=='INVALID' else st}")
    if row.get("Company")=="Inactive": out.append("người bán ngừng hoạt động"+(f" ({row['Tax status']})" if row.get("Tax status") else ""))
    elif row.get("Company")!="Active": out.append("chưa đọc được tình trạng MST người bán, tra tay trên tracuunnt.gdt.gov.vn")
    return ", ".join(out)

def done_names(person):
    """File names (lower) of this person already moved to done\\ in any month."""
    out=set()
    for p in glob.glob(os.path.join(DONE,"*",person,"**","*"), recursive=True):
        if os.path.isfile(p): out.add(os.path.basename(p).lower())
    return out
