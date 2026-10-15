# PAYMENT TOOL - console fallback. The normal way in is "Payment Tool.bat" (window interface).
# Use this only when the window does not open. Files: data\<Ho ten>\ -> done\YYYY-MM\<Ho ten>\.
import os, sys, subprocess
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass
from layout import ROOT, DATA, DONE

APP  = os.path.dirname(os.path.abspath(__file__))
PY   = [sys.executable] if sys.executable else ["py","-3.12"]
NODE = os.path.join(APP,"node.exe")
if not os.path.exists(NODE): NODE="node"
INPUT = os.path.join(DATA,"Request_Input.xlsx")
CHECKLOG = os.path.join(DATA,"Check_Log.xlsx")

def run_py(script, *args): subprocess.call(PY+[os.path.join(APP,script), *args])
def run_node(script, *args): subprocess.call([NODE, os.path.join(APP,script), *args])
def openfile(path):
    try: os.startfile(path)
    except Exception as e: print("Cannot open:",path,e)
def ask(prompt):
    try: return input(prompt).strip()
    except EOFError: return ""

def main():
    os.makedirs(DATA, exist_ok=True); os.makedirs(DONE, exist_ok=True)
    while True:
        print("\n==================================================")
        print("     PAYMENT TOOL - Nhập đơn đề nghị thanh toán")
        print("==================================================")
        print("  (Bản cửa sổ: chạy 'Payment Tool.bat' ở thư mục gốc)")
        print()
        print("  1. Kiểm tra hóa đơn")
        print("  2. Điền đơn lên cổng (chưa đăng nhập thì tool tự mở cổng để đăng nhập)")
        print("  3. Chuyển người đã Lưu sang done")
        print()
        print("  4. Mở file đơn          5. Mở kết quả kiểm tra")
        print("  6. Mở thư mục data      7. Đăng nhập cổng")
        print("  0. Thoát")
        c=ask("Chọn: ")
        if c=="1":
            who=ask("Kiểm tra cho ai? gõ tên (Enter = tất cả): ").strip('"')
            redo=ask("Kiểm tra lại cả hóa đơn đã có kết quả? (y/N): ").lower()=="y"
            run_py("check_invoices.py", *([who] if who else []), *(["--all"] if redo else []))
        elif c=="2":
            run_py("build_input.py")      # luôn tạo lại file đơn từ data hiện tại
            run_py("don_to_json.py")
            run_node("fill_don.js")
        elif c=="3": run_py("archive_done.py")
        elif c=="4": openfile(INPUT) if os.path.exists(INPUT) else print("No Request_Input.xlsx yet (run 2 first).")
        elif c=="5": openfile(CHECKLOG) if os.path.exists(CHECKLOG) else print("No Check_Log.xlsx yet (run 1 first).")
        elif c=="6": openfile(DATA)
        elif c=="7": run_node("login.js")
        elif c=="0": print("Bye."); return

if __name__=="__main__":
    main()
