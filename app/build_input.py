# Build data\Request_Input.xlsx for fill_don.js: one row (= one request) per folder data\<Ho ten>\.
# File rules are in layout.py (XML + same-name PDF = invoice, other PDF = to trinh, check\ = check PDF).
import sys, os, re, glob
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

from layout import DATA, PEOPLE, people_dirs, scan_full, read_check_log, check_problem
NAVY="183665"

def load_people():
    """people.xlsx: Ten / Ngan_Hang / So_Tai_Khoan (+ Ma_CBCNV optional) -> dict by name."""
    f=PEOPLE
    if not os.path.exists(f): return {}
    from openpyxl import load_workbook
    ws=load_workbook(f,data_only=True).active
    hdr=[str(c.value).strip() if c.value else "" for c in ws[1]]
    idx={n:hdr.index(n) for n in ["Ten","Ngan_Hang","So_Tai_Khoan","Ma_CBCNV"] if n in hdr}
    out={}
    if "Ten" not in idx: return out
    for r in ws.iter_rows(min_row=2,values_only=True):
        name=r[idx["Ten"]]
        if not name: continue
        out[str(name).strip()]={
            "ngan_hang": str(r[idx["Ngan_Hang"]]).strip() if "Ngan_Hang" in idx and r[idx["Ngan_Hang"]] else "",
            "so_tk": str(r[idx["So_Tai_Khoan"]]).strip() if "So_Tai_Khoan" in idx and r[idx["So_Tai_Khoan"]] else "",
            "ma": str(r[idx["Ma_CBCNV"]]).strip() if "Ma_CBCNV" in idx and r[idx["Ma_CBCNV"]] else "",
        }
    return out

def rel(p): return os.path.relpath(p, DATA) if p else ""

def main():
    out_xlsx=os.path.join(DATA,"Request_Input.xlsx")
    scans={p:scan_full(p) for p in people_dirs()}
    groups=[(p,r["invoices"],r["totrinh"]) for p,r in scans.items() if r["invoices"]]
    if not groups: sys.exit("No invoice .xml found in data\\<name>\\.")
    people=load_people()
    today=date.today()

    cols=["Ngay_De_Nghi","Nguoi_De_Nghi","Nhom_Ke_Toan","Chung_Tu_Kem_Theo","Dien_Giai","Ghi_Chu","Ten_Hang_Hoa",
          "Hinh_Thuc_Thanh_Toan","Loai_Chuyen_Khoan","Ngan_Hang","So_Tai_Khoan","Ma_Tien_Te",
          "Nguoi_Thu_Huong","Ma_Citad","XML_HoaDon","HoaDon_PDF","ToTrinh_PDF","Check_PDF"]
    wb=Workbook(); ws=wb.active; ws.title="DATA"; ws.append(cols)
    hfill=PatternFill("solid",start_color=NAVY); thin=Side(style="thin",color="BFBFBF"); bd=Border(thin,thin,thin,thin)
    print("Built requests:")
    for name,invs,tt in groups:
        xmls=";".join(rel(i["xml"]) for i in invs)
        pdfs=";".join(rel(i["pdf"]) for i in invs)
        # Check PDF = EVERY PDF in data\<name>\check\ (any file name): same-name ones first in invoice order, then the rest
        all_c=sorted(glob.glob(os.path.join(DATA,name,"check","*.pdf")))
        first=[i["check"] for i in invs if i["check"]]
        check_list=first+[c for c in all_c if c not in first]
        checks=";".join(rel(c) for c in check_list)
        miss_c=[] if len(check_list)>=len(invs) else [f"{len(check_list)} check PDF for {len(invs)} invoice(s)"]
        miss_p=[i["stem"] for i in invs if not i["pdf"]]
        # To trinh nguon = every PDF in the person folder that is not an invoice PDF
        totrinh_s=";".join(rel(p) for p in tt)
        # thang trong dien giai = thang lam de nghi (hom nay), khong theo ngay hoa don
        thang=today.strftime("%m/%Y")
        # name cua thu muc con = TEN NGUOI; dien giai theo tung nguoi
        person=name
        if person:
            dien=f"Thanh toán chi phí giao dịch khách hàng Tháng {thang}-{person}"
        else:
            dien=f"Thanh toán chi phí giao dịch khách hàng Tháng {thang}".rstrip()
        info=people.get(person, {})
        ngan_hang=info.get("ngan_hang",""); so_tk=info.get("so_tk","")
        has_bank = bool(ngan_hang and so_tk)   # can DU ngan hang + so TK moi auto-dien
        # Co SO TK trong people.xlsx -> dien san thu huong + bank (full auto).
        # Chua co -> de trong -> Tool B nhac bam "Chon Nguoi Thu Huong" tren cong (bank tu nhay).
        thu_huong = person if has_bank else ""
        ws.append([today, person, "Kế toán Đơn vị thành viên","", dien, "", "Ăn uống",
                   "Chuyển khoản","Nội địa", ngan_hang, so_tk, "VND", thu_huong, "", xmls, pdfs, totrinh_s, checks])
        label=person or "(all)"
        print(f"  {label}: {len(invs)} invoice(s)"
              + (" | bank auto (people.xlsx)" if has_bank else " | pick beneficiary on portal")
              + (f" | to trinh: {', '.join(os.path.basename(t) for t in tt)}" if tt else " | NO to trinh PDF")
              + (f" | MISSING check: {miss_c}" if miss_c else "")
              + (f" | MISSING pdf: {miss_p}" if miss_p else ""))
        for w in scans[name]["warnings"]:
            print(f"      !! {w}")
        log=read_check_log(name)
        for i in invs:
            r=log.get(os.path.basename(i["xml"]))
            flag=(f"  !! {check_problem(r)}" if r and check_problem(r) else "")
            print(f"      {i['stem']}: {i['kh']} no.{i['sh']} date {i['ngay']} total {i['tong']}  [{i['seller']}]{flag}")

    for c in ws[1]:
        c.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c.fill=hfill
        c.alignment=Alignment(horizontal="center",wrap_text=True); c.border=bd
    for i,w in enumerate([13,16,22,16,36,16,14,16,14,18,16,9,18,11,34,34,34,34],1): ws.column_dimensions[chr(64+i)].width=w
    for row in ws.iter_rows(min_row=2,max_row=ws.max_row):
        for c in row: c.font=Font(name="Arial",size=9); c.border=bd; c.alignment=Alignment(vertical="top",wrap_text=True)
        row[0].number_format="dd/mm/yyyy"
    ws.freeze_panes="A2"

    hd=wb.create_sheet("Hướng dẫn")
    for ln in [
     ["HƯỚNG DẪN - File đơn tự sinh từ thư mục data"],[""],
     ["Mỗi thư mục data\\<Họ tên>\\ thành 1 đơn. XML + PDF cùng tên = hóa đơn, PDF khác = tờ trình, check\\ = kết quả kiểm tra."],
     ["Mỗi dòng = 1 đơn. Kiểm tra lại rồi điền các cột còn trống:"],[""],
     ["Cột","Việc cần làm"],
     ["Ngay_De_Nghi","Ngày lúc tool điền đơn (Bước 2 luôn lấy ngày hiện tại)."],
     ["Nguoi_De_Nghi","Đã điền theo tên thư mục. Kiểm tra lại cho khớp tên trên cổng."],
     ["Nguoi_Thu_Huong","Đã điền theo tên. Nếu thụ hưởng là tài khoản thì sửa cho đúng."],
     ["Dien_Giai","Đã sinh 'Thanh toán chi phí giao dịch khách hàng Tháng MM/YYYY-Tên' (tháng hiện tại). Sửa nếu cần."],
     ["Ten_Hang_Hoa","Tên hàng hóa ghi vào mọi dòng trong danh sách hóa đơn (nút Sửa). Mặc định 'Ăn uống'."],
     ["Nhom_Ke_Toan / Ma_Tien_Te","Đã điền mặc định (Kế toán Đơn vị thành viên / VND)."],
     ["Hinh_Thuc_Thanh_Toan","Mặc định Chuyển khoản. Đổi sang Tiền mặt nếu trả tiền mặt."],
     ["Loai_Chuyen_Khoan","Nội địa | Quốc tế (chỉ khi Chuyển khoản)."],
     ["Ngan_Hang / So_Tai_Khoan","Tự lấy từ people.xlsx theo tên. Trống = người này chưa có trong people.xlsx."],
     ["XML_HoaDon / HoaDon_PDF / Check_PDF","Tool tự ghép theo mã hóa đơn. Trống = thiếu file (xem log). Check_PDF = MỌI file PDF trong thư mục check (tên gì cũng được). FILE ĐÍNH KÈM chính = tờ trình + Check_PDF; HoaDon_PDF đính vào nút Sửa từng dòng hóa đơn."],
     ["ToTrinh_PDF","Tờ trình nguồn của phòng: thả PDF vào cùng thư mục người đó (tên khác hóa đơn), tool tự đính vào FILE ĐÍNH KÈM."],[""],
     ["Chưa muốn nộp người nào: chuyển thư mục người đó vào data\\_hold\\ (thư mục bắt đầu bằng _ được bỏ qua)."],
    ]: hd.append(ln)
    hd.column_dimensions["A"].width=34; hd.column_dimensions["B"].width=60
    hd["A1"].font=Font(name="Arial",bold=True,size=12,color=NAVY)
    for r in hd.iter_rows():
        for c in r:
            if c.value=="Cột":
                for cc in r:
                    if cc.value: cc.font=Font(name="Arial",bold=True,color="FFFFFF",size=10); cc.fill=hfill
    for r in hd.iter_rows(min_row=2):
        for c in r:
            if c.value and not (c.font and c.font.bold):
                c.font=Font(name="Arial",size=10); c.alignment=Alignment(vertical="top",wrap_text=True)

    wb.save(out_xlsx)
    print(f"\nWrote {out_xlsx}")
    print("Open it, fill Ngay_De_Nghi / Nguoi_De_Nghi / Nguoi_Thu_Huong, then run Tool B.")

if __name__=="__main__":
    main()
