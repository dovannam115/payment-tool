# Step 3: the client-meeting approval slips (Giấy đề nghị duyệt chi phí tiếp khách hàng), one per invoice,
# all of a person's slips in one Word file: data\<person>\to_trinh\To_Trinh_<Name>_<YYYYMMDD>.docx.
# What the user typed is kept in data\<person>\to_trinh\input.json, so the window opens with it next time.
import os, re, io, json, glob, base64, zipfile, datetime as dt, unicodedata
from layout import DATA, DONE, DUMP, scan, people_dirs

import slip_tpl as TPL

DEPTS = ["Phòng Kinh doanh gia tăng (Upselling)", "Phòng Bancassurance and Online"]


def fix_dept(d):
    """input.json saved before 20/11/2026 may hold the name without 'Phòng'; a department taken off the list
    falls back to the first one."""
    names = depts()
    d = str(d or "").strip()
    if d and d not in names and "Phòng " + d in names: d = "Phòng " + d
    return d if d in names else names[0]
HOST = "Lãnh đạo phòng và Chuyên viên"
# the wordings of the sample slips; {sp} = the products. The text made from them can still be edited freely.
TOPICS = ["Trao đổi về các yêu cầu & chào phí đơn {sp}.",
          "Trao đổi, thảo luận điều khoản và nội dung hợp đồng {sp}.",
          "Gặp gỡ, trao đổi triển khai nội dung hợp đồng {sp}.",
          "Gặp gỡ, trao đổi, thảo luận điều khoản và phương án triển khai nội dung hợp đồng {sp}."]
CUSTOMERS = os.path.join(DUMP, "customers.json")      # imported per machine, never in the repo

PRODUCTS = [
    "Sức khỏe bổ sung", "Bệnh ung thư", "Bệnh hiểm nghèo", "Chăm sóc sức khỏe (CSSK) toàn diện",
    "Chi phí y tế và vận chuyển cấp cứu", "PVI Care (Tai nạn nhóm + Sức khỏe nhóm)", "PVI Premier Care",
    "Tai nạn cá nhân", "Tai nạn hành khách", "Tai nạn toàn diện", "Tai nạn hộ sử dụng điện", "Tai nạn thuyền viên",
    "Sức khỏe kết hợp", "Du lịch nội địa", "Du lịch quốc tế", "Sức khỏe Học sinh – sinh viên", "Sức khỏe tương trợ",
    "Trợ cấp điều trị dài ngày – An phúc ưu việt", "Bồi thường cho người lao động",
    "Bồi thường cho người lao động trong doanh nghiệp xây lắp", "Mọi rủi ro xây dựng", "Mọi rủi ro lắp đặt",
    "Đổ vỡ máy móc", "Máy móc thiết bị chủ thầu", "Thiết bị điện tử", "Năng lượng (dầu khí ngoài khơi)",
    "Toàn diện nhà tư nhân", "Bảo hiểm mọi rủi ro tài sản", "Cháy nổ bắt buộc theo NĐ 105", "Mọi rủi ro Tiền",
    "Trộm cướp", "Lòng trung thực", "Chi phí trả hàng", "Trộm cắp định danh cá nhân", "Thiết bị điện – điện tử",
    "Một đổi một (1 đổi 1)", "Tiền và tài sản trong giao dịch thẻ ngân hàng",
    "Bảo hành mở rộng (EW) cho thiết bị điện – điện tử", "Nứt vỡ màn hình",
    "Thiệt hại thiết bị điện – điện tử do tai nạn và chất lỏng", "Sự cố an ninh mạng cá nhân",
    "Bồi thường giải thưởng HIO (Hole in One)", "TN bồi thường giải câu cá", "TN nghề nghiệp công chứng viên",
    "TN cá nhân", "TN tiêm chủng", "TN bồi thường giải thưởng Golf", "TN công cộng và sản phẩm",
    "TN chung toàn diện (QBE)", "TN gây ô nhiễm tại cơ sở sản xuất kinh doanh", "TN nghề nghiệp",
    "TN người quản lý kho hàng", "TN pháp lý nghề nghiệp và dân sự", "TN chủ nuôi chó", "TN nghề nghiệp bác sỹ",
    "TN kiến trúc sư và kỹ sư tư vấn", "TN luật sư – công chứng viên", "TN nghề nghiệp trong khám chữa bệnh",
    "Trách nhiệm cảng", "Trách nhiệm của người sửa chữa tàu", "Trách nhiệm giao nhận", "Gián đoạn kinh doanh",
    "Hàng hóa vận chuyển nội địa (VCNĐ)", "Hàng hóa vận chuyển đường hàng không (CL259)",
    "Thực phẩm đông lạnh A (CL263)", "Thực phẩm đông lạnh C (CL264)", "Dầu chở rời (CL273)", "Thân tàu biển",
    "TNDS chủ tàu biển", "Máy móc, thiết bị hàng hải", "Thân tàu và TNDS chủ tàu (điều khoản tiếng Anh)",
    "Bắt cóc và tiền chuộc", "Đóng tàu", "Du thuyền", "Tàu thủy nội địa – tàu cá", "Tự nguyện xe ô tô",
    "Tự nguyện xe máy", "TNDS bắt buộc chủ xe cơ giới", "Tín dụng thương mại Tradeliner",
]


# ------------------------------------------------------------------ customer list
PRODUCTS_FILE = os.path.join(DUMP, "products.json")
DEPTS_FILE = os.path.join(DUMP, "depts.json")
LISTS = {"cust": (CUSTOMERS, "Tên khách hàng"), "prod": (PRODUCTS_FILE, "Tên sản phẩm"), "dept": (DEPTS_FILE, "Tên phòng")}
DEFAULTS = {"prod": PRODUCTS, "dept": DEPTS}


def clean(names):
    out, seen = [], set()
    for n in names:
        n = " ".join(unicodedata.normalize("NFC", str(n or "")).split())
        if n and n.casefold() not in seen: seen.add(n.casefold()); out.append(n)
    return out


def get_list(kind):
    """The customer list starts empty (imported per machine); products and departments start from the built-in lists."""
    try:
        with open(LISTS[kind][0], encoding="utf-8") as f: return json.load(f)
    except Exception:
        return list(DEFAULTS.get(kind, []))


def save_list(kind, names):
    names = clean(names)
    os.makedirs(DUMP, exist_ok=True)
    tmp = LISTS[kind][0] + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f: json.dump(names, f, ensure_ascii=False)
    os.replace(tmp, LISTS[kind][0])
    return names


def import_list(kind, data):
    """An Excel list replaces the saved one. The header row ('Tên khách hàng' / 'Tên sản phẩm') may sit a few
    rows down. Customers keep only the rows whose 'Mã n.khách TĐ' is TN, the filter of the sample file."""
    from openpyxl import load_workbook
    col = LISTS[kind][1]
    ws = load_workbook(io.BytesIO(data), read_only=True, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    for h, r in enumerate(rows[:15]):
        hdr = [" ".join(str(c or "").split()) for c in r]
        if col in hdr: break
    else:
        raise ValueError("Không thấy cột '%s' trong file" % col)
    ni = hdr.index(col); gi = hdr.index("Mã n.khách TĐ") if kind == "cust" and "Mã n.khách TĐ" in hdr else None
    names = [r[ni] for r in rows[h + 1:] if r and len(r) > ni and r[ni]
             and (gi is None or str(r[gi] or "").strip().upper() == "TN")]
    if not clean(names): raise ValueError("Không có dòng nào trong cột '%s'" % col)
    return save_list(kind, names)


def import_customers(data): return len(import_list("cust", data))
def customers(): return get_list("cust")
def products(): return get_list("prod")
def depts(): return get_list("dept") or list(DEPTS)


# ------------------------------------------------------------------ dates and money
def vn(d): return d.strftime("%d/%m/%Y")


def parse(s):
    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", str(s or ""))
    if not m: return None
    try: return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError: return None


def inv_date(iv):
    try: return dt.date.fromisoformat(str(iv.get("ngay") or "")[:10])
    except ValueError: return None


def inv_total(iv):
    try: return int(round(float(str(iv.get("tong") or "0").replace(",", ""))))
    except ValueError: return 0


def budget(total):
    """The next whole million above the invoice total: 4.5M -> 5M, 2.3M -> 3M, 5M -> 6M."""
    return (total // 1000000 + 1) * 1000000


def guests(total):
    """Guests from the bill, about 500,000 VND a head, at least 2: 0.8M -> 2, 3.9M -> 8, 4.9M -> 10."""
    return max(2, int(round(total / 500000.0)))


def dots(n): return "{:,}".format(int(n)).replace(",", ".")


def topic_text(i, products):
    sp = ", ".join(p if p.lower().startswith("bảo hiểm") else "Bảo hiểm " + p for p in products)
    try: t = TOPICS[int(i)]
    except (ValueError, IndexError, TypeError): t = TOPICS[0]
    return t.format(sp=sp)


# ------------------------------------------------------------------ what the window shows
def folder(person): return os.path.join(DATA, person, "to_trinh")


def saved(person):
    try:
        with open(os.path.join(folder(person), "input.json"), encoding="utf-8") as f: return json.load(f)
    except Exception: return {}


def ascii_name(person):
    s = unicodedata.normalize("NFD", person.replace("đ", "d").replace("Đ", "D"))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return "_".join(re.sub(r"[^A-Za-z0-9 ]", " ", s).split())


def last_file(person):
    d = folder(person)
    if not os.path.isdir(d): return ""
    f = sorted((x for x in os.listdir(d) if x.lower().endswith(".docx") and not x.startswith("~$")),
               key=lambda x: os.path.getmtime(os.path.join(d, x)))
    return os.path.relpath(os.path.join(d, f[-1]), DATA).replace("\\", "/") if f else ""


def form(person):
    """The person's invoices with the slip fields: what was saved, else the defaults."""
    invs, _ = scan(person)
    old = saved(person)
    keep = {r.get("xml"): r for r in old.get("rows", [])}
    rows = []
    for n, iv in enumerate(sorted(invs, key=lambda x: (str(x.get("ngay")), x["stem"]))):
        d, tot = inv_date(iv), inv_total(iv)
        k = os.path.basename(iv["xml"])
        r = {"xml": k, "so": iv["sh"], "ngay": vn(d) if d else "", "tong": tot, "seller": iv["seller"],
             "kh": "", "sp": [], "topic": n % len(TOPICS), "nd": "", "khach": str(guests(tot)),
             "tu": vn(d - dt.timedelta(days=1)) if d else "", "den": vn(d + dt.timedelta(days=1)) if d else "",
             "lap": vn(d - dt.timedelta(days=2)) if d else "", "cp": budget(tot) if tot else ""}
        r.update({x: keep[k][x] for x in ("kh", "sp", "topic", "nd", "khach", "tu", "den", "lap", "cp") if k in keep and x in keep[k]})
        rows.append(r)
    dept = fix_dept(old.get("dept"))
    return {"name": person, "dept": dept, "rows": rows, "file": last_file(person)}


# ------------------------------------------------------------------ checks
def problems(r, iv_total, iv_day, names):
    out = []
    if not str(r.get("kh") or "").strip(): out.append("chưa chọn khách hàng")
    elif names and r["kh"] not in names: out.append("khách hàng không có trong danh mục")
    sp = [p for p in r.get("sp") or [] if p]
    if not sp: out.append("chưa chọn sản phẩm")
    elif any(p not in products() for p in sp): out.append("sản phẩm không có trong danh mục")
    try: k = int(str(r.get("khach") or "").strip())
    except ValueError: k = 0
    if k < 1: out.append("chưa nhập số khách")
    tu, den, lap = parse(r.get("tu")), parse(r.get("den")), parse(r.get("lap"))
    if not tu or not den: out.append("thời gian dự kiến chưa đúng dạng dd/mm/yyyy")
    elif tu > den: out.append("thời gian dự kiến: từ ngày sau đến ngày")
    elif iv_day and not tu <= iv_day <= den: out.append("ngày hóa đơn %s nằm ngoài thời gian dự kiến" % vn(iv_day))
    if not lap: out.append("ngày lập chưa đúng dạng dd/mm/yyyy")
    elif tu and lap >= tu: out.append("ngày lập phải trước thời gian dự kiến")
    try: cp = int(str(r.get("cp") or "0").replace(".", "").replace(",", ""))
    except ValueError: cp = 0
    if cp <= iv_total: out.append("chi phí dự kiến phải lớn hơn hóa đơn (%s)" % dots(iv_total))
    return out


# ------------------------------------------------------------------ the Word file
def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def unit(dept, r, person, num):
    lap, tu, den = parse(r["lap"]), parse(r["tu"]), parse(r["den"])
    tg = vn(tu) if tu == den else "%s-%s" % (vn(tu), vn(den))
    cp = int(str(r["cp"]).replace(".", "").replace(",", ""))
    v = {"PHONG": dept, "DD": "%02d" % lap.day, "MM": "%02d" % lap.month, "YYYY": str(lap.year),
         "KH": r["kh"].strip(), "ND": r.get("nd") or topic_text(r.get("topic"), r["sp"]), "TG": tg,
         "SK": "%02d" % int(r["khach"]), "CP": dots(cp), "NL": person, "NUM": str(num)}
    return re.sub(r"\{(PHONG|DD|MM|YYYY|KH|ND|TG|SK|CP|NL|NUM)\}", lambda m: esc(v[m.group(1)]), TPL.UNIT)


PAGE = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
NUM = ('<w:num w:numId="%d"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0">'
       '<w:startOverride w:val="1"/></w:lvlOverride></w:num>')


def build(person, dept, rows):
    """rows already checked. Every slip gets its own list number so each one counts 1-6 again."""
    src = zipfile.ZipFile(io.BytesIO(base64.b64decode("".join(TPL.DOCX))))
    units = [unit(dept, r, person, 5001 + i) for i, r in enumerate(rows)]
    os.makedirs(folder(person), exist_ok=True)
    path = os.path.join(folder(person), "To_Trinh_%s_%s.docx" % (ascii_name(person), dt.date.today().strftime("%Y%m%d")))
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in src.namelist():
            d = src.read(n)
            if n == "word/document.xml":
                d = d.decode("utf-8").replace("{UNITS}", PAGE.join(units)).encode("utf-8")
            elif n == "word/numbering.xml":
                x = d.decode("utf-8")
                d = x.replace("</w:numbering>", "".join(NUM % (5001 + i) for i in range(len(rows))) + "</w:numbering>").encode("utf-8")
            out.writestr(n, d)
    try:
        os.replace(tmp, path)
    except PermissionError:
        os.remove(tmp)
        raise ValueError("%s đang mở trong Word, đóng file rồi bấm Tạo lại" % os.path.basename(path))
    return path


# ------------------------------------------------------------------ one customer per department and month
# Two different people of the same department may not meet the same customer in the same month (the month of
# the invoice). Checked against the people being made, the other people in data\ (what they saved) and the
# people already moved to done\ (Hoàn tất). The same person meeting a customer twice is fine.
def month_of(r):
    d = parse(r.get("ngay")) or parse(r.get("tu"))
    return "%02d/%d" % (d.month, d.year) if d else ""


def key(s): return " ".join(str(s or "").split()).casefold()


def records(skip):
    """(person, dept, month, customer, finished) of every saved slip outside the people in skip."""
    out = []
    srcs = [(p, os.path.join(folder(p), "input.json"), False) for p in people_dirs() if p not in skip]
    srcs += [(os.path.basename(os.path.dirname(os.path.dirname(f))), f, True)
             for f in glob.glob(os.path.join(DONE, "*", "*", "to_trinh", "input*.json"))]
    for person, f, fin in srcs:
        try:
            with open(f, encoding="utf-8") as fh: d = json.load(fh)
        except Exception: continue
        dept = fix_dept(d.get("dept"))
        for r in d.get("rows") or []:
            if str(r.get("kh") or "").strip() and month_of(r):
                out.append((person, dept, month_of(r), r["kh"].strip(), fin))
    return out


def taken(skip):
    """For the window: customers other people of each department already have, so Ngẫu nhiên can avoid them."""
    return [{"name": p, "dept": d, "month": m, "kh": k} for p, d, m, k, _ in records(set(skip))]


def clashes(people):
    """people: [{name, dept, rows}] being made. {(name, xml): [problem]} for each slip whose customer another
    person of the same department already has in that month."""
    mine = []
    for pp in people:
        invs, _ = scan(pp["name"])
        day = {os.path.basename(iv["xml"]): inv_date(iv) for iv in invs}
        dept = fix_dept(pp.get("dept"))
        for r in pp.get("rows") or []:
            d = day.get(r.get("xml"))
            if d and str(r.get("kh") or "").strip():
                mine.append((pp["name"], dept, "%02d/%d" % (d.month, d.year), r["kh"].strip(), r["xml"]))
    others = records({pp["name"] for pp in people})
    out = {}
    for name, dept, m, kh, xml in mine:
        hit = sorted({p + (" (đã hoàn tất)" if fin else "") for p, d2, m2, k2, fin in others
                      if p != name and d2 == dept and m2 == m and key(k2) == key(kh)} |
                     {p for p, d2, m2, k2, _ in mine if p != name and d2 == dept and m2 == m and key(k2) == key(kh)})
        if hit:
            out[(name, xml)] = ["khách hàng trùng với %s, cùng phòng trong tháng %s" % (", ".join(hit), m)]
    return out


def make(person, dept, rows_in, clash=None):
    """Check every slip, save what was typed (even with problems), build the file when all are right."""
    invs, _ = scan(person)
    by = {os.path.basename(iv["xml"]): iv for iv in invs}
    names = set(customers())
    rows, bad = [], []
    for r in rows_in:
        iv = by.get(r.get("xml"))
        if not iv: continue
        r = {k: r.get(k) for k in ("xml", "kh", "sp", "topic", "nd", "khach", "tu", "den", "lap", "cp")}
        if inv_date(iv): r["ngay"] = vn(inv_date(iv))          # the month the one-customer rule counts in
        r["nd"] = " ".join(str(r["nd"] or "").split())
        if not str(r["khach"] or "").strip(): r["khach"] = str(guests(inv_total(iv)))
        r["kh"] = " ".join(str(r["kh"] or "").split())
        r["sp"] = [p for p in (r["sp"] or []) if isinstance(p, str) and p]
        rows.append(r)
        p = problems(r, inv_total(iv), inv_date(iv), names) + (clash or {}).get((person, r["xml"]), [])
        if p: bad.append({"xml": r["xml"], "so": iv["sh"], "problems": p})
    dept = fix_dept(dept)
    os.makedirs(folder(person), exist_ok=True)
    with open(os.path.join(folder(person), "input.json"), "w", encoding="utf-8") as f:
        json.dump({"dept": dept, "rows": rows}, f, ensure_ascii=False, indent=1)
    if bad or not rows: return {"name": person, "bad": bad, "error": "" if rows else "Không có hóa đơn"}
    rows.sort(key=lambda r: (inv_date(by[r["xml"]]) or dt.date.min, r["xml"]))
    path = build(person, dept, rows)
    return {"name": person, "bad": [], "file": os.path.relpath(path, DATA).replace("\\", "/"), "count": len(rows)}

