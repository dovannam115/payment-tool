# Convert data\Request_Input.xlsx -> dump\don.json for Tool B (fill_don.js).
# One row = one request. Special path columns:
#   XML_HoaDon  : invoice XML files -> "Import File XML". Several files separated by ;
#   HoaDon_PDF  : invoice PDF files -> per-row "Sua" in the invoice list (matching invoice). Several separated by ;
#   ToTrinh_PDF + Check_PDF -> main "FILE DINH KEM" (#file_attachs)
#   Ten_Hang_Hoa: item name written into every invoice row ("Sua" -> TEN HANG HOA), default "Ăn uống"
#   Check_PDF   : invoice-check PDF files -> per-row "Sua" button (one per invoice, same order). Separated by ;
# Other columns -> fields (exact names: Ngay_De_Nghi, Nguoi_De_Nghi, Nhom_Ke_Toan,
#   Chung_Tu_Kem_Theo, Dien_Giai, Ghi_Chu, Hinh_Thuc_Thanh_Toan, Ma_Tien_Te,
#   Nguoi_Thu_Huong, Ma_Citad).
import sys, os, json, datetime, re
from openpyxl import load_workbook
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

def inv_meta(xml_path):
    # read invoice number (SHDon), serial (KHHDon), form (KHMSHDon) so checks map to the right row
    try:
        x=open(xml_path,encoding="utf-8",errors="ignore").read()
        g=lambda t:(re.search(rf"<{t}>(.*?)</{t}>",x) or [None,""])[1].strip()
        return {"so_hd":g("SHDon"), "ky_hieu":g("KHHDon"), "mau":g("KHMSHDon")}
    except Exception: return {"so_hd":"","ky_hieu":"","mau":""}
stem=lambda p: os.path.splitext(os.path.basename(p))[0]

from layout import DATA, APP
XLSX=sys.argv[1] if len(sys.argv)>1 else os.path.join(DATA,"Request_Input.xlsx")
OUT =os.path.join(APP,"dump","don.json")
PATHCOLS={"xml_hoadon":"xml","hoadon_pdf":"attachments","totrinh_pdf":"totrinh","check_pdf":"checks"}

def paths(val, base):
    out=[]
    for part in str(val).split(";"):
        p=part.strip()
        if not p: continue
        if not os.path.isabs(p): p=os.path.join(base,p)
        out.append(p)
    return out

if not os.path.exists(XLSX):
    print("Not found:",XLSX); sys.exit(0)
wb=load_workbook(XLSX,data_only=True)
ws=wb["DATA"] if "DATA" in wb.sheetnames else wb.worksheets[0]
header=[str(c.value).strip() if c.value is not None else "" for c in ws[1]]
base=os.path.dirname(os.path.abspath(XLSX))
dons=[]
for r in ws.iter_rows(min_row=2,values_only=True):
    if not any(v not in (None,"") for v in r): continue
    don={"fields":{},"xml":[],"attachments":[],"totrinh":[],"checks":[]}
    for h,v in zip(header,r):
        if not h or v in (None,""): continue
        key=PATHCOLS.get(h.lower())
        if key:
            don[key]=paths(v,base)
        else:
            if isinstance(v,(datetime.date,datetime.datetime)):
                v=f"{v.day:02d}/{v.month:02d}/{v.year}"
            don["fields"][h]=v
    # pair each invoice (xml) with its number + its check PDF + its invoice PDF (by file stem, not position)
    checks_by_stem={stem(c).lower():c for c in don["checks"]}
    pdf_by_stem={stem(p).lower():p for p in don["attachments"]}
    don["invoices"]=[]
    for x in don["xml"]:
        m=inv_meta(x); s=stem(x)
        pdf=pdf_by_stem.pop(s.lower(),"")
        if not pdf and len(m["so_hd"].lstrip("0"))>=3:   # PDF named differently: find the invoice number in its name
            no=m["so_hd"].lstrip("0")
            hit=[k for k in pdf_by_stem if re.search(rf"(?<!\d)0*{re.escape(no)}(?!\d)", k)]
            if len(hit)==1: pdf=pdf_by_stem.pop(hit[0])
        don["invoices"].append({"xml":x, "stem":s, **m, "check":checks_by_stem.get(s.lower(),""), "pdf":pdf})
    dons.append(don)
json.dump(dons,open(OUT,"w",encoding="utf-8"),ensure_ascii=False,indent=2)
print(f"Wrote {len(dons)} request(s) -> {OUT}")
