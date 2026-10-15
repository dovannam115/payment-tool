# Move people whose request is SAVED on the portal from data\<Ho ten>\ to done\YYYY-MM\<Ho ten>\.
# Keeps data\ = only what is still pending; done\ = history by month, so the same invoices are not sent twice.
import sys, os, shutil
from datetime import date
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass
from layout import DATA, DONE, people_dirs

def ask(q):
    try: return input(q).strip()
    except EOFError: return ""

def move(person, month):
    src=os.path.join(DATA,person); dst=os.path.join(DONE,month,person)
    for root,_,files in os.walk(src):
        tdir=os.path.join(dst, os.path.relpath(root,src)); os.makedirs(tdir, exist_ok=True)
        for f in files:
            t=os.path.join(tdir,f)
            if os.path.exists(t):
                b,e=os.path.splitext(f); k=2
                while os.path.exists(os.path.join(tdir,f"{b}_{k}{e}")): k+=1
                t=os.path.join(tdir,f"{b}_{k}{e}")
            shutil.move(os.path.join(root,f), t)
    shutil.rmtree(src, ignore_errors=True)
    print(f"  {person} -> done\\{month}\\{person}")

def main():
    people=sys.argv[1:] or None
    if not people:
        allp=people_dirs()
        if not allp: sys.exit("data\\ has no person folder.")
        print("\nPeople in data\\:")
        for i,p in enumerate(allp,1): print(f"  {i}. {p}")
        ans=ask("\nWho is saved on the portal? (numbers split by space, 0 = all, Enter = cancel): ")
        if not ans: print("Cancelled."); return
        if ans=="0": people=allp
        else: people=[allp[int(x)-1] for x in ans.replace(","," ").split() if x.isdigit() and 0<int(x)<=len(allp)]
    if not people: print("Nothing chosen."); return
    month=date.today().strftime("%Y-%m")
    for p in people: move(p, month)
    print(f"\nDone. {len(people)} person(s) moved to done\\{month}.")

if __name__=="__main__":
    main()
