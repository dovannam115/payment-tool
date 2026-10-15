# Folder layout shared by every step.
#   data\<Ho ten>\          current batch, ONE folder per person, files dropped flat:
#       *.xml               invoice XML
#       *.pdf               same name as an XML -> invoice PDF; any other PDF -> to trinh nguon
#       check\*.pdf         invoice-check result (made by check_invoices.py)
#   data\_<name>\           folders starting with "_" are skipped (e.g. _hold = not this batch)
#   data\Request_Input.xlsx one row per person (made by build_input.py)
#   done\YYYY-MM\<Ho ten>\  people already saved on the portal (moved by archive_done.py)
import os, re, glob

APP    = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.dirname(APP)
DATA   = os.path.join(ROOT, "data")
DONE   = os.path.join(ROOT, "done")
PEOPLE = os.path.join(ROOT, "people.xlsx")

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
        return {"kh":g("KHHDon"),"sh":g("SHDon"),"ngay":g("NLap"),
                "seller":(mnt.group(1).strip() if mnt else "")[:45],
                "tong":g("TgTTTBSo") or g("TgTCThueGTGT")}
    except Exception:
        return {"kh":"","sh":"","ngay":"","seller":"","tong":""}

def _stem(p): return os.path.splitext(os.path.basename(p))[0]

def scan(person):
    """Classify one person's folder -> invoices (xml + pdf + check) and to trinh PDFs."""
    pdir=os.path.join(DATA, person)
    xmls=sorted(glob.glob(os.path.join(pdir,"*.xml")))
    pdfs={_stem(p).lower():p for p in sorted(glob.glob(os.path.join(pdir,"*.pdf")))}
    checks={_stem(p).lower():p for p in glob.glob(os.path.join(pdir,"check","*.pdf"))}
    invs=[]
    for xp in xmls:
        s=_stem(xp)
        invs.append({"xml":xp,"stem":s,"pdf":pdfs.pop(s.lower(),""),
                     "check":checks.get(s.lower(),""),**read_invoice(xp)})
    # invoice PDF named differently from its XML: pair by invoice number found in the PDF name
    for iv in invs:
        if iv["pdf"]: continue
        no=iv["sh"].lstrip("0")
        if len(no)<3: continue
        hit=[k for k in pdfs if re.search(rf"(?<!\d)0*{re.escape(no)}(?!\d)", k)]
        if len(hit)==1: iv["pdf"]=pdfs.pop(hit[0])
    totrinh=sorted(pdfs.values())   # whatever PDF is left = to trinh nguon
    return invs, totrinh

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
    if row.get("Status")!="VALID": out.append(f"invoice {row.get('Status') or '?'}")
    if row.get("Company")=="Inactive": out.append("seller INACTIVE"+(f" ({row['Tax status']})" if row.get("Tax status") else ""))
    elif row.get("Company")!="Active": out.append("seller status not read (tax site busy) - run step 2 again with redo")
    return ", ".join(out)

def done_names(person):
    """File names (lower) of this person already moved to done\\ in any month."""
    out=set()
    for p in glob.glob(os.path.join(DONE,"*",person,"**","*"), recursive=True):
        if os.path.isfile(p): out.add(os.path.basename(p).lower())
    return out
