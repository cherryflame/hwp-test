# -*- coding: utf-8 -*-
import os, re, csv, time, queue, hashlib, zipfile, difflib, threading
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP="백업 문서 중복·유사본 검사기 v2"
EXTS={".hwp",".hwpx",".docx",".txt"}

def fhash(p):
    h=hashlib.sha256()
    with open(p,"rb") as f:
        for b in iter(lambda:f.read(1024*1024),b""): h.update(b)
    return h.hexdigest()

def decode_txt(b):
    for enc in ("utf-8-sig","utf-8","cp949","euc-kr"):
        try:return b.decode(enc),enc
        except UnicodeDecodeError:pass
    return b.decode("cp949",errors="replace"),"cp949(replace)"

def xmltext(data):
    root=ET.fromstring(data); out=[]
    for e in root.iter():
        tag=e.tag.split("}")[-1].lower()
        if tag in ("t","text") and e.text: out.append(e.text)
        elif tag in ("p","paragraph","br","linebreak"): out.append("\n")
        elif tag=="tab": out.append("\t")
    return "".join(out)

def read_docx(p):
    out=[]
    with zipfile.ZipFile(p) as z:
        names=z.namelist()
        wanted=["word/document.xml"]+[n for n in names if re.match(r"word/(header|footer)\d+\.xml$",n)]
        wanted += ["word/footnotes.xml","word/endnotes.xml"]
        for n in wanted:
            if n in names:
                try: out.append(xmltext(z.read(n)))
                except: pass
    return "\n".join(out),"DOCX"

def read_hwpx(p):
    out=[]
    with zipfile.ZipFile(p) as z:
        sec=sorted(n for n in z.namelist() if re.search(r"Contents/section\d+\.xml$",n,re.I))
        for n in sec: out.append(xmltext(z.read(n)))
    return "\n".join(out),"HWPX"

def read_hwp(p):
    # 검증된 syhwp를 사용. 버전 차이에 대비해 공개 API 후보를 순차 시도.
    import syhwp
    errors=[]
    candidates=[]
    for name in ("extract_text","read","parse"):
        fn=getattr(syhwp,name,None)
        if callable(fn): candidates.append(fn)
    try:
        from syhwp import Hwp
        candidates.append(lambda x:Hwp(x).text)
    except Exception: pass
    for fn in candidates:
        try:
            v=fn(str(p))
            if isinstance(v,str): return v,"HWP5"
            if hasattr(v,"text"): return str(v.text),"HWP5"
        except Exception as e: errors.append(str(e))
    # syhwp CLI/API가 패키지 버전에 따라 다를 경우 모듈 내부 탐색
    for modname in ("syhwp.hwp","syhwp.parser","syhwp.reader"):
        try:
            mod=__import__(modname,fromlist=["*"])
            for name in ("extract_text","read_hwp","read","parse"):
                fn=getattr(mod,name,None)
                if callable(fn):
                    try:
                        v=fn(str(p))
                        if isinstance(v,str): return v,"HWP5"
                        if hasattr(v,"text"): return str(v.text),"HWP5"
                    except Exception as e: errors.append(str(e))
        except Exception: pass
    raise RuntimeError("HWP 본문 추출 실패: "+("; ".join(errors[-3:]) if errors else "syhwp API를 찾지 못했습니다."))

def read_text(p):
    e=Path(p).suffix.lower()
    if e==".txt":
        t,enc=decode_txt(Path(p).read_bytes()); return t,"TXT/"+enc
    if e==".docx":return read_docx(p)
    if e==".hwpx":return read_hwpx(p)
    if e==".hwp":return read_hwp(p)
    raise RuntimeError("지원하지 않는 형식")

def norm(s):
    s=s.replace("\r\n","\n").replace("\r","\n").replace("\u00a0"," ")
    s=re.sub(r"[ \t]+"," ",s); s=re.sub(r" *\n *","\n",s); s=re.sub(r"\n{3,}","\n\n",s)
    return s.strip()
def compact(s):return re.sub(r"\s+","",norm(s))
def thash(s):return hashlib.sha256(compact(s).encode()).hexdigest()
def shingles(s,k=9,limit=5000):
    s=compact(s)
    if len(s)<=k:return {s} if s else set()
    step=max(1,(len(s)-k+1)//limit)
    return {s[i:i+k] for i in range(0,len(s)-k+1,step)}
def jac(a,b):return len(a&b)/len(a|b) if a and b else (1 if not a and not b else 0)

class R:
    def __init__(self,p):
        self.path=str(p); st=os.stat(p); self.name=os.path.basename(p); self.ext=Path(p).suffix.lower()
        self.size=st.st_size; self.mtime=st.st_mtime; self.hash=""; self.text=""; self.th=""; self.kind=""; self.err=""; self.sh=None

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP); self.geometry("1220x780"); self.minsize(900,600)
        self.folders=[]; self.records=[]; self.groups=[]; self.q=queue.Queue(); self.running=False
        self.ui(); self.after(100,self.poll)
    def ui(self):
        head=ttk.Frame(self,padding=10);head.pack(fill="x")
        ttk.Label(head,text=APP,font=("",15,"bold")).pack(side="left")
        ttk.Label(head,text="원본 파일은 읽기만 하며 삭제·이동·수정하지 않습니다.",foreground="#6f7b86").pack(side="right")

        self.tabs=ttk.Notebook(self);self.tabs.pack(fill="both",expand=True,padx=10,pady=(0,10))
        folder_tab=ttk.Frame(self.tabs);pair_tab=ttk.Frame(self.tabs)
        self.tabs.add(folder_tab,text="  폴더 중복 검사  ")
        self.tabs.add(pair_tab,text="  파일 2개 비교  ")

        # 폴더 중복 검사
        box=ttk.LabelFrame(folder_tab,text="검사할 폴더",padding=8);box.pack(fill="x",padx=8,pady=(8,0))
        top=ttk.Frame(box);top.pack(fill="x",pady=(0,6))
        ttk.Button(top,text="폴더 추가",command=self.add).pack(side="left")
        ttk.Button(top,text="선택 폴더 제거",command=self.remove).pack(side="left",padx=6)
        self.lb=tk.Listbox(box,height=4,selectmode="extended");self.lb.pack(fill="x")

        opt=ttk.Frame(folder_tab,padding=8);opt.pack(fill="x")
        ttk.Label(opt,text="유사도 기준").pack(side="left")
        self.cut=tk.IntVar(value=90);ttk.Spinbox(opt,from_=70,to=99,width=5,textvariable=self.cut).pack(side="left",padx=5)
        ttk.Label(opt,text="%").pack(side="left")
        self.start=ttk.Button(opt,text="검사 시작",command=self.go);self.start.pack(side="right")
        self.save=ttk.Button(opt,text="CSV 저장",command=self.csv,state="disabled");self.save.pack(side="right",padx=6)

        self.pb=ttk.Progressbar(folder_tab);self.pb.pack(fill="x",padx=8)
        self.status=ttk.Label(folder_tab,text="폴더를 추가해 주세요.",padding=(8,5));self.status.pack(fill="x")

        cols=("judge","score","name","type","size","date","path")
        self.tree=ttk.Treeview(folder_tab,columns=cols,show="tree headings",selectmode="extended")
        self.tree.heading("#0",text="그룹")
        specs=[("judge","판정",95),("score","유사도",70),("name","파일명",210),("type","형식",70),("size","크기",80),("date","수정일",125),("path","경로",430)]
        for c,t,w in specs:self.tree.heading(c,text=t);self.tree.column(c,width=w)
        self.tree.column("#0",width=65)
        self.tree.pack(fill="both",expand=True,padx=8)

        foot=ttk.Frame(folder_tab,padding=8);foot.pack(fill="x")
        ttk.Button(foot,text="선택한 두 파일 차이 보기",command=self.diffwin).pack(side="left")

        # 파일 2개 비교
        pairtop=ttk.Frame(pair_tab,padding=12);pairtop.pack(fill="x")
        self.pair_paths=[tk.StringVar(),tk.StringVar()]
        for idx,label in enumerate(("A","B")):
            row=ttk.Frame(pairtop);row.pack(fill="x",pady=5)
            ttk.Label(row,text=label,width=3,font=("",11,"bold")).pack(side="left")
            ttk.Entry(row,textvariable=self.pair_paths[idx]).pack(side="left",fill="x",expand=True,padx=(0,7))
            ttk.Button(row,text="파일 선택",command=lambda i=idx:self.pick_pair(i)).pack(side="left")
        action=ttk.Frame(pairtop);action.pack(fill="x",pady=(8,0))
        ttk.Label(action,text="HWP 5.x · HWPX · DOCX · TXT / 서로 다른 형식도 본문 비교 가능",foreground="#6f7b86").pack(side="left")
        ttk.Button(action,text="두 파일 비교",command=self.compare_pair).pack(side="right")

        self.pair_summary=ttk.Label(pair_tab,text="비교할 파일 두 개를 선택해 주세요.",padding=(12,8),font=("",10,"bold"))
        self.pair_summary.pack(fill="x")
        pane=ttk.Panedwindow(pair_tab,orient="horizontal");pane.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.pair_text=[]
        for label in ("A","B"):
            f=ttk.Frame(pane);pane.add(f,weight=1)
            h=ttk.Label(f,text=label,padding=7,font=("",10,"bold"));h.pack(fill="x")
            t=tk.Text(f,wrap="word",font=("Malgun Gothic",10),undo=False)
            y=ttk.Scrollbar(f,orient="vertical",command=t.yview);t.configure(yscrollcommand=y.set)
            y.pack(side="right",fill="y");t.pack(fill="both",expand=True)
            self.pair_text.append(t)
        self.pair_note=ttk.Label(pair_tab,text="",padding=(10,5),foreground="#6f7b86");self.pair_note.pack(fill="x")

    def pick_pair(self,idx):
        p=filedialog.askopenfilename(filetypes=[("지원 문서","*.hwp *.hwpx *.docx *.txt"),("모든 파일","*.*")])
        if p:self.pair_paths[idx].set(p)

    def compare_pair(self):
        a,b=[x.get().strip() for x in self.pair_paths]
        if not a or not b:return messagebox.showinfo(APP,"A와 B 파일을 모두 선택해 주세요.")
        try:
            ra,rb=R(a),R(b)
            for r in (ra,rb):
                r.hash=fhash(r.path);r.text,r.kind=read_text(r.path);r.th=thash(r.text)
            ca,cb=compact(ra.text),compact(rb.text)
            exact=ra.hash==rb.hash
            content=ra.th==rb.th
            if exact:sim=1.0;judge="완전 동일"
            elif content:sim=1.0;judge="내용 동일"
            elif len(ca)+len(cb)<600000:sim=difflib.SequenceMatcher(None,ca,cb,autojunk=False).ratio();judge="유사/상이"
            else:sim=jac(shingles(ca),shingles(cb));judge="유사/상이"

            sm=difflib.SequenceMatcher(None,norm(ra.text).splitlines(),norm(rb.text).splitlines(),autojunk=False)
            added=removed=changed=0
            for tag,i1,i2,j1,j2 in sm.get_opcodes():
                if tag=="insert":added+=j2-j1
                elif tag=="delete":removed+=i2-i1
                elif tag=="replace":changed+=max(i2-i1,j2-j1)

            self.pair_summary["text"]=(f"{judge}  ·  유사도 {sim*100:.2f}%  ·  "
                f"A {len(ca):,}자 / B {len(cb):,}자  ·  추가 {added}줄 / 삭제 {removed}줄 / 변경 {changed}줄")
            self.show_side_diff(ra,rb,sm)
            self.pair_note["text"]=f"A: {ra.kind} · {self.sz(ra.size)}     B: {rb.kind} · {self.sz(rb.size)}"
        except Exception as e:
            messagebox.showerror(APP,"비교하지 못했습니다.\n\n"+str(e))

    def show_side_diff(self,a,b,sm):
        ta,tb=self.pair_text
        for t in (ta,tb):
            t.configure(state="normal");t.delete("1.0","end")
            t.tag_configure("same")
            t.tag_configure("del",background="#ffdede")
            t.tag_configure("add",background="#dff3df")
            t.tag_configure("chg",background="#fff0a8")
        la,lb=norm(a.text).splitlines(),norm(b.text).splitlines()
        for tag,i1,i2,j1,j2 in sm.get_opcodes():
            taga=tagb="same"
            if tag=="delete":taga="del"
            elif tag=="insert":tagb="add"
            elif tag=="replace":taga=tagb="chg"
            for line in la[i1:i2]:ta.insert("end",line+"\n",taga)
            for line in lb[j1:j2]:tb.insert("end",line+"\n",tagb)
        for t in (ta,tb):t.configure(state="disabled")

    def add(self):
        p=filedialog.askdirectory()
        if p and p not in self.folders:self.folders.append(p);self.lb.insert("end",p)
    def remove(self):
        for i in reversed(self.lb.curselection()):
            self.folders.pop(i);self.lb.delete(i)
    def go(self):
        if not self.folders:return messagebox.showinfo(APP,"검사할 폴더를 추가해 주세요.")
        if self.running:return
        self.running=True;self.start["state"]="disabled";self.save["state"]="disabled";self.tree.delete(*self.tree.get_children())
        threading.Thread(target=self.worker,daemon=True).start()
    def worker(self):
        try:
            paths=[];seen=set()
            for folder in self.folders:
                for base,ds,fs in os.walk(folder):
                    for n in fs:
                        p=os.path.join(base,n)
                        if Path(n).suffix.lower() in EXTS and p not in seen:seen.add(p);paths.append(p)
            self.q.put(("max",max(1,len(paths))))
            rec=[]
            for i,p in enumerate(paths,1):
                r=R(p)
                try:r.hash=fhash(p);r.text,r.kind=read_text(p);r.th=thash(r.text)
                except Exception as e:r.err=str(e)
                rec.append(r);self.q.put(("p",i,f"읽는 중 {i}/{len(paths)} · {r.name}"))
            groups=[]; claimed=set()
            byhash=defaultdict(list)
            for r in rec:byhash[r.hash].append(r)
            for h,x in byhash.items():
                if h and len(x)>1:
                    groups.append(("완전 동일",1.0,x))
                    for a in x:
                        for b in x:
                            if a.path<b.path:claimed.add((a.path,b.path))
            byt=defaultdict(list)
            for r in rec:
                if r.th and not r.err:byt[r.th].append(r)
            for h,x in byt.items():
                if len(x)>1 and len({r.hash for r in x})>1:groups.append(("내용 동일",1.0,x))
                if len(x)>1:
                    for a in x:
                        for b in x:
                            if a.path<b.path:claimed.add((a.path,b.path))
            valid=[r for r in rec if r.text and not r.err]; cut=max(.70,min(.99,self.cut.get()/100)); edges=[]
            for i,a in enumerate(valid):
                ca=compact(a.text);la=len(ca)
                for b in valid[i+1:]:
                    pair=tuple(sorted((a.path,b.path)))
                    if pair in claimed:continue
                    cb=compact(b.text);lb=len(cb)
                    if not la or not lb or min(la,lb)/max(la,lb)<.70:continue
                    if a.sh is None:a.sh=shingles(a.text)
                    if b.sh is None:b.sh=shingles(b.text)
                    rough=jac(a.sh,b.sh)
                    if rough<max(.30,cut-.40):continue
                    if la+lb<600000:s=difflib.SequenceMatcher(None,ca,cb,autojunk=False).ratio()
                    else:s=rough
                    if s>=cut:edges.append((a,b,s))
            # pair groups: avoids misleading transitive similarity percentages
            for a,b,s in sorted(edges,key=lambda x:-x[2]):groups.append(("유사",s,[a,b]))
            self.records=rec;self.groups=groups;self.q.put(("done",))
        except Exception as e:self.q.put(("fatal",str(e)))
    def poll(self):
        try:
            while 1:
                m=self.q.get_nowait()
                if m[0]=="max":self.pb["maximum"]=m[1];self.pb["value"]=0
                elif m[0]=="p":self.pb["value"]=m[1];self.status["text"]=m[2]
                elif m[0]=="done":self.render()
                elif m[0]=="fatal":self.running=False;self.start["state"]="normal";messagebox.showerror(APP,m[1])
        except queue.Empty:pass
        self.after(100,self.poll)
    def render(self):
        for gi,(j,s,rs) in enumerate(self.groups,1):
            root=self.tree.insert("","end",text=str(gi),values=(j,f"{s*100:.1f}%","","","","",""),open=True)
            for r in rs:self.tree.insert(root,"end",values=("","",r.name,r.kind,self.sz(r.size),time.strftime("%Y-%m-%d %H:%M",time.localtime(r.mtime)),r.path))
        errs=[r for r in self.records if r.err]
        if errs:
            root=self.tree.insert("","end",text="오류",values=("읽기 실패","",f"{len(errs)}개","","","",""),open=False)
            for r in errs:self.tree.insert(root,"end",values=("","",r.name,r.ext,self.sz(r.size),"",r.path+" | "+r.err))
        self.status["text"]=f"완료 · 대상 {len(self.records)}개 · 결과 그룹 {len(self.groups)}개 · 읽기 실패 {len(errs)}개"
        self.running=False;self.start["state"]="normal";self.save["state"]="normal"
    def selected(self):
        d={r.path:r for r in self.records};out=[]
        for x in self.tree.selection():
            v=self.tree.item(x,"values")
            if len(v)>=7 and v[6] in d:out.append(d[v[6]])
        return out
    def diffwin(self):
        rs=self.selected()
        if len(rs)!=2:return messagebox.showinfo(APP,"파일 두 개를 Ctrl+클릭으로 선택해 주세요.")
        a,b=rs;w=tk.Toplevel(self);w.title("문서 차이 보기");w.geometry("1100x700")
        t=tk.Text(w,wrap="none",font=("Consolas",10));t.pack(fill="both",expand=True)
        ca,cb=compact(a.text),compact(b.text)
        sim=difflib.SequenceMatcher(None,ca,cb,autojunk=False).ratio()*100 if len(ca)+len(cb)<600000 else jac(shingles(ca),shingles(cb))*100
        t.insert("end",f"유사도: {sim:.2f}%\nA: {a.path}\nB: {b.path}\n\n")
        d="\n".join(difflib.unified_diff(norm(a.text).splitlines(),norm(b.text).splitlines(),fromfile=a.name,tofile=b.name,lineterm="",n=3))
        t.insert("end",d or "정규화한 본문 내용이 동일합니다.")
    def csv(self):
        p=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv")],initialfile="문서_중복_유사본_결과.csv")
        if not p:return
        with open(p,"w",encoding="utf-8-sig",newline="") as f:
            w=csv.writer(f);w.writerow(["그룹","판정","유사도","파일명","형식","크기","수정일","경로"])
            for gi,(j,s,rs) in enumerate(self.groups,1):
                for r in rs:w.writerow([gi,j,f"{s*100:.2f}%",r.name,r.kind,r.size,time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(r.mtime)),r.path])
        messagebox.showinfo(APP,"저장했습니다.")
    @staticmethod
    def sz(n):
        for u in ("B","KB","MB","GB"):
            if n<1024 or u=="GB":return f"{int(n)} B" if u=="B" else f"{n:.1f} {u}"
            n/=1024
if __name__=="__main__":App().mainloop()
