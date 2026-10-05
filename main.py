# -*- coding: utf-8 -*-
import os, sys, re, csv, time, queue, hashlib, zipfile, difflib, threading
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import unquote
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

APP="문서 중복·유사성 검사기 v1.0"

# Windows 작업표시줄이 Python/Tk 기본 아이콘이 아니라 이 앱의 EXE 아이콘을 사용하도록 고유 ID를 지정합니다.
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "SianK.DocumentSimilarityChecker.1.0"
        )
    except Exception:
        pass

def resource_path(name):
    base=getattr(sys,"_MEIPASS",os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base,name)
EXTS={".hwp",".hwpx",".docx",".txt",".epub"}

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
                except Exception: pass
    return "\n".join(out),"DOCX"

def read_hwpx(p):
    out=[]
    with zipfile.ZipFile(p) as z:
        sec=sorted(n for n in z.namelist() if re.search(r"Contents/section\d+\.xml$",n,re.I))
        for n in sec: out.append(xmltext(z.read(n)))
    return "\n".join(out),"HWPX"

def _epub_local(tag):
    return tag.split("}")[-1].lower()

def _epub_html_text(data):
    root=ET.fromstring(data)
    out=[]
    block_tags={"p","div","section","article","header","footer","aside","nav","blockquote","pre","li","dt","dd","h1","h2","h3","h4","h5","h6","tr"}
    skip_tags={"script","style","svg","head"}
    def walk(e):
        tag=_epub_local(e.tag)
        if tag in skip_tags:return
        if tag in block_tags and out and not out[-1].endswith("\n"):out.append("\n")
        if e.text:out.append(e.text)
        for c in e:
            ctag=_epub_local(c.tag)
            if ctag=="br":out.append("\n")
            else:walk(c)
            if c.tail:out.append(c.tail)
        if tag in block_tags and (not out or not out[-1].endswith("\n")):out.append("\n")
    body=None
    for e in root.iter():
        if _epub_local(e.tag)=="body":body=e;break
    walk(body if body is not None else root)
    return "".join(out)

def read_epub(p):
    """EPUB container.xml -> OPF -> spine 순서로 선형 본문을 추출한다."""
    with zipfile.ZipFile(p) as z:
        try:
            container=ET.fromstring(z.read("META-INF/container.xml"))
        except KeyError:
            raise RuntimeError("EPUB container.xml이 없습니다.")
        rootfile=None
        for e in container.iter():
            if _epub_local(e.tag)=="rootfile":
                rootfile=e.attrib.get("full-path")
                if rootfile: rootfile=unquote(rootfile)
                if rootfile:break
        if not rootfile:raise RuntimeError("EPUB OPF 경로를 찾지 못했습니다.")
        try:opf=ET.fromstring(z.read(rootfile))
        except KeyError:raise RuntimeError("EPUB OPF 파일을 찾지 못했습니다.")
        manifest={}
        for e in opf.iter():
            if _epub_local(e.tag)=="item":
                iid=e.attrib.get("id");href=e.attrib.get("href")
                if iid and href:manifest[iid]=(href,e.attrib.get("media-type","").lower(),e.attrib.get("properties",""))
        spine=[]
        for e in opf.iter():
            if _epub_local(e.tag)=="itemref":
                iid=e.attrib.get("idref")
                if iid:spine.append(iid)
        base=Path(rootfile).parent
        out=[]
        for iid in spine:
            item=manifest.get(iid)
            if not item:continue
            href,media,props=item
            if "html" not in media and not href.lower().endswith((".xhtml",".html",".htm")):continue
            # nav 전용 문서는 일반적으로 본문 spine에 없지만, 포함된 경우 중복 목차를 피한다.
            if "nav" in props.split():continue
            name=(base / unquote(href.split("#",1)[0].split("?",1)[0])).as_posix()
            try:data=z.read(name)
            except KeyError:continue
            try:t=_epub_html_text(data)
            except ET.ParseError:
                # 드문 비정상 HTML은 텍스트 손실보다 읽기 실패가 안전하다.
                raise RuntimeError(f"EPUB XHTML 파싱 실패: {name}")
            if t.strip():out.append(t)
        if not out:raise RuntimeError("EPUB spine에서 읽을 본문을 찾지 못했습니다.")
        return "\n".join(out),"EPUB"

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
    if e==".epub":return read_epub(p)
    raise RuntimeError("지원하지 않는 형식")

def norm(s):
    s=s.replace("\r\n","\n").replace("\r","\n").replace("\u00a0"," ")
    s=re.sub(r"[ \t]+"," ",s); s=re.sub(r" *\n *","\n",s); s=re.sub(r"\n{3,}","\n\n",s)
    return s.strip()
def compact(s):return re.sub(r"\s+","",norm(s))
def pair_norm(s):
    # A/B 상세 비교에서는 사용자가 저장한 공백·탭·줄바꿈을 보존하고 개행 코드만 통일한다.
    return s.replace("\r\n","\n").replace("\r","\n").replace("\u00a0"," ")
def display_safe(s):
    if not isinstance(s,str): return str(s)
    # Tk/Windows 표시용: 고립 surrogate만 대체문자로 바꾸고 정상 한글/유니코드는 보존.
    return s.encode("utf-8",errors="replace").decode("utf-8")
def safe_utf8(s):
    # 일부 오래된 문서/파일명에서 고립 surrogate가 들어와도 검사 전체가 실패하지 않게 처리.
    return s.encode("utf-8",errors="surrogatepass")
def thash(s):return hashlib.sha256(safe_utf8(compact(s))).hexdigest()
def shingles(s,k=9,limit=5000):
    s=compact(s)
    if len(s)<=k:return {s} if s else set()
    step=max(1,(len(s)-k+1)//limit)
    return {s[i:i+k] for i in range(0,len(s)-k+1,step)}
def semantic_stream(s):
    # 형식별 줄바꿈/공백/문단 장식 차이에 흔들리지 않는 후보 탐색용 문자 흐름.
    # 한글/영문/숫자만 남기며 원문과 상세 비교에는 사용하지 않는다.
    return "".join(ch.lower() for ch in s if ch.isalnum())

def semantic_fingerprint(s,k=11,mod=97,limit=2600):
    """내용 기반 rolling fingerprint. 위치 기반 샘플링이 아니어서 앞부분 삽입/삭제에도 안정적이다."""
    s=semantic_stream(s)
    n=len(s)
    if not s:return set()
    if n<=k:return {s}
    base=257; mask=(1<<64)-1
    power=pow(base,k-1,1<<64)
    vals=[ord(c)+1 for c in s]
    h=0
    for x in vals[:k]:h=((h*base)+x)&mask
    out=set()
    def add(i,hv):
        if hv%mod==0: out.add(s[i:i+k])
    add(0,h)
    for i in range(1,n-k+1):
        h=(h-(vals[i-1]*power))&mask
        h=((h*base)+vals[i+k-1])&mask
        add(i,h)
        if len(out)>=limit:break
    # 매우 짧거나 우연히 선택점이 없는 문서도 후보가 될 수 있게 양끝 앵커 보강.
    out.add(s[:k]); out.add(s[-k:])
    return out

def dice(a,b):
    if not a and not b:return 1.0
    if not a or not b:return 0.0
    return 2*len(a&b)/(len(a)+len(b))

def semantic_length_ratio(la,lb):
    """Fast whole-document coverage penalty for the final folder similarity score.
    Fingerprints find matching content; this ratio makes one-sided omitted/added text reduce the final score.
    """
    if not la and not lb:return 1.0
    if not la or not lb:return 0.0
    return min(la,lb)/max(la,lb)

def shown_group_score(judge,score):
    # Display the real calculated score. A similar pair may legitimately be 100%
    # under the folder-search similarity metric even when detailed comparison finds
    # small word/whitespace/line-break differences.
    return score*100

def group_score_text(judge,score):
    pct=shown_group_score(judge,score)
    if pct>=100.0-1e-9:
        return "100%"
    return f"{pct:.1f}%"


def _detail_tokens(s):
    """Whitespace runs are one logical token; visible characters remain individual tokens."""
    out=[]; buf=[]
    def flush():
        nonlocal buf
        if buf:
            out.append(("WS","".join(buf))); buf=[]
    for ch in s:
        if ch.isspace():
            buf.append(ch)
        else:
            flush(); out.append((ch,ch))
    flush()
    return out

def _merge_detail_segments(parts):
    out=[]
    for tag,a,b in parts:
        if not a and not b: continue
        if out and out[-1][0]==tag:
            ot,oa,ob=out[-1]; out[-1]=(ot,oa+a,ob+b)
        else:
            out.append((tag,a,b))
    return out

def _ws_delta_segments(a,b):
    """Whitespace runs: keep the common part uncoloured and mark only added/removed whitespace."""
    if a==b:
        return [('equal',a,b)] if a else []
    sm=difflib.SequenceMatcher(None,list(a),list(b),autojunk=False)
    out=[]
    for tag,i1,i2,j1,j2 in sm.get_opcodes():
        aa=a[i1:i2]; bb=b[j1:j2]
        if tag=='equal': out.append(('equal',aa,bb))
        elif tag=='delete': out.append(('ws',aa,''))
        elif tag=='insert': out.append(('ws','',bb))
        else:
            # Different whitespace characters are both formatting-only differences.
            if aa: out.append(('ws',aa,''))
            if bb: out.append(('ws','',bb))
    return _merge_detail_segments(out)

def _local_detail_segments(a,b):
    ta,tb=_detail_tokens(a),_detail_tokens(b)
    ka=[x[0] for x in ta]; kb=[x[0] for x in tb]
    # Resynchronised regions are normally only a few hundred characters.
    # autojunk=False gives correct punctuation insert/delete classification on these small regions.
    sm=difflib.SequenceMatcher(None,ka,kb,autojunk=False)
    out=[]
    for tag,i1,i2,j1,j2 in sm.get_opcodes():
        if tag=='equal':
            for (akey,av),(bkey,bv) in zip(ta[i1:i2],tb[j1:j2]):
                if akey=='WS' and av!=bv: out.extend(_ws_delta_segments(av,bv))
                else: out.append(('equal',av,bv))
        elif tag=='delete':
            chunk=ta[i1:i2]; ws_only=all(key=='WS' for key,val in chunk)
            for key,val in chunk: out.append(('ws',val,'') if ws_only else ('delete',val,''))
        elif tag=='insert':
            chunk=tb[j1:j2]; ws_only=all(key=='WS' for key,val in chunk)
            for key,val in chunk: out.append(('ws','',val) if ws_only else ('insert','',val))
        else:
            aa=ta[i1:i2]; bb=tb[j1:j2]
            # Keep genuine substitutions yellow, but do not absorb whitespace into them.
            m=min(len(aa),len(bb))
            for n in range(m):
                ak,av=aa[n]; bk,bv=bb[n]
                if ak=='WS' or bk=='WS':
                    if ak=='WS' and bk=='WS': out.extend(_ws_delta_segments(av,bv))
                    else:
                        if ak=='WS': out.append(('ws',av,'')); out.append(('insert','',bv))
                        else: out.append(('delete',av,'')); out.append(('ws','',bv))
                else: out.append(('replace',av,bv))
            tail=aa[m:]; ws_only=bool(tail) and all(key=='WS' for key,val in tail)
            for key,val in tail: out.append(('ws',val,'') if ws_only else ('delete',val,''))
            tail=bb[m:]; ws_only=bool(tail) and all(key=='WS' for key,val in tail)
            for key,val in tail: out.append(('ws','',val) if ws_only else ('insert','',val))
    return _merge_detail_segments(out)

def _visible_for_block_score(s):
    """Comparison text for deciding whether a changed block is a rewrite or a small edit."""
    # Whitespace is formatting here; punctuation remains meaningful enough to distinguish edits.
    return "".join(ch.casefold() for ch in s if not ch.isspace())

def _changed_block_similarity(a,b):
    aa=_visible_for_block_score(a); bb=_visible_for_block_score(b)
    if not aa and not bb: return 1.0
    if not aa or not bb: return 0.0
    # Length mismatch is evidence too, but SequenceMatcher catches retained phrases in partial edits.
    return difflib.SequenceMatcher(None,aa,bb,autojunk=True).ratio()

def _detail_block_segments(a,b,rewrite_threshold=0.48,min_rewrite_chars=20):
    """Hierarchical gate: classify a changed block before character-level diffing.

    One-sided blocks stay wholly red/green (including their whitespace). Two substantial,
    low-similarity blocks are a replacement and stay wholly yellow. Only genuinely similar
    blocks are allowed into the fine-grained character/whitespace analyser.
    """
    if not a and not b: return []
    if not b: return [('delete',a,'')]
    if not a: return [('insert','',b)]
    if a==b: return [('equal',a,b)]
    va=_visible_for_block_score(a); vb=_visible_for_block_score(b)
    if min(len(va),len(vb))>=min_rewrite_chars:
        if _changed_block_similarity(a,b)<rewrite_threshold:
            return [('replace',a,b)]
    return _local_detail_segments(a,b)

def semantic_positions(s):
    s=pair_norm(s); chars=[]; pos=[]
    for i,ch in enumerate(s):
        if ch.isalnum(): chars.append(ch.casefold()); pos.append(i)
    return s,"".join(chars),pos

def resync_detail_segments(a,b,anchor_len=32,step=260,window=1600):
    """Fast cross-format detail diff with progressive semantic re-synchronisation."""
    oa,sa,pa=semantic_positions(a); ob,sb,pb=semantic_positions(b)
    if not sa or not sb: return _detail_block_segments(oa,ob)
    anchors=[]; last_b=0
    for ai in range(0,max(0,len(sa)-anchor_len)+1,step):
        key=sa[ai:ai+anchor_len]
        if len(key)<anchor_len: break
        expected=int(ai*len(sb)/max(1,len(sa)))
        lo=max(last_b,expected-window); hi=min(len(sb),expected+window+anchor_len)
        bi=sb.find(key,lo,hi)
        if bi<0: bi=sb.find(key,last_b)
        if bi>=last_b:
            anchors.append((ai,bi)); last_b=bi+anchor_len
    bounds=[(0,0)]
    lasta=lastb=0
    for ai,bi in anchors:
        if ai<=0 or bi<=0: continue
        ao,bo=pa[ai],pb[bi]
        if ao>lasta and bo>lastb:
            bounds.append((ao,bo)); lasta,lastb=ao,bo
    bounds.append((len(oa),len(ob)))
    parts=[]
    for (a1,b1),(a2,b2) in zip(bounds,bounds[1:]):
        xa,xb=oa[a1:a2],ob[b1:b2]
        if xa==xb: parts.append(('equal',xa,xb))
        else: parts.extend(_detail_block_segments(xa,xb))
    return _merge_detail_segments(parts)

def jac(a,b):return len(a&b)/len(a|b) if a and b else (1 if not a and not b else 0)

class R:
    def __init__(self,p):
        self.path=str(p); st=os.stat(p); self.name=os.path.basename(p); self.ext=Path(p).suffix.lower()
        self.size=st.st_size; self.mtime=st.st_mtime; self.hash=""; self.text=""; self.th=""; self.kind=""; self.err=""; self.sh=None

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP); self.geometry("1220x780"); self.minsize(900,600)
        self._app_icon_images=[]
        try:
            self.iconbitmap(default=resource_path("app.ico"))
        except Exception:
            pass
        try:
            self._app_icon_images=[tk.PhotoImage(file=resource_path(f"app_icon_{n}.png")) for n in (64,48,32,24,16)]
            self.iconphoto(True,*self._app_icon_images)
        except Exception:
            self._app_icon_images=[]
        self.folders=[]; self.records=[]; self.groups=[]; self.q=queue.Queue(); self.running=False
        self.active_filter="전체"; self.result_counts={"전체":0,"완전 동일":0,"내용 동일":0,"유사":0,"읽기 실패":0}
        self.skip_diff_transfer_notice=False
        self.pair_compare_seq=0
        self.configure_named_fonts()
        self.ui(); self.after(100,self.poll)


    def configure_named_fonts(self):
        specs={
            "TkDefaultFont": ("Malgun Gothic",9),
            "TkTextFont": ("Malgun Gothic",10),
            "TkFixedFont": ("Malgun Gothic",9),
            "TkMenuFont": ("Malgun Gothic",9),
            "TkHeadingFont": ("Malgun Gothic",9,"bold"),
            "TkCaptionFont": ("Malgun Gothic",9,"bold"),
            "TkSmallCaptionFont": ("Malgun Gothic",9),
            "TkIconFont": ("Malgun Gothic",9),
            "TkTooltipFont": ("Malgun Gothic",9),
        }
        for name,spec in specs.items():
            try:
                tkfont.nametofont(name).configure(
                    family=spec[0],size=spec[1],
                    weight=spec[2] if len(spec)>2 else "normal")
            except tk.TclError:
                pass

    def setup_style(self):
        self.configure(background="#f5f7fa")
        style=ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass
        style.configure(".",font=("Malgun Gothic",9),background="#f5f7fa",foreground="#243447")
        style.configure("TFrame",background="#f5f7fa")
        style.configure("Card.TFrame",background="#ffffff")
        style.configure("TLabel",background="#f5f7fa",foreground="#243447")
        style.configure("Card.TLabel",background="#ffffff",foreground="#243447")
        style.configure("Muted.TLabel",background="#f5f7fa",foreground="#718096")
        style.configure("CardMuted.TLabel",background="#ffffff",foreground="#718096")
        style.configure("Title.TLabel",background="#f5f7fa",foreground="#172b4d",font=("Malgun Gothic",16,"bold"))
        style.configure("PanelTitle.TLabel",background="#ffffff",foreground="#172b4d",font=("Malgun Gothic",10,"bold"))
        style.configure("TButton",padding=(12,7),relief="flat",borderwidth=1,
                        background="#ffffff",foreground="#17365d",bordercolor="#304a68")
        style.map("TButton",background=[("active","#e7eef8")])
        style.configure("Primary.TButton",background="#214f78",foreground="white",padding=(14,8),
                        relief="flat",borderwidth=1,bordercolor="#17365d")
        style.map("Primary.TButton",background=[("active","#245dcc"),("disabled","#a9b8cf")],foreground=[("disabled","#f5f7fa")])
        style.configure("TNotebook",background="#f5f7fa",borderwidth=0,tabmargins=(0,0,0,0))
        style.configure("TNotebook.Tab",padding=(18,9),background="#e9eef5",foreground="#52606d",borderwidth=0)
        style.map("TNotebook.Tab",background=[("selected","#ffffff")],foreground=[("selected","#245dcc")])
        style.configure("Treeview",background="#ffffff",fieldbackground="#ffffff",rowheight=27,borderwidth=1,relief="flat",bordercolor="#304a68")
        style.configure("Flat.Treeview",background="#ffffff",fieldbackground="#ffffff",rowheight=27,borderwidth=1,relief="flat",bordercolor="#304a68")
        style.layout("Flat.Treeview.Heading",[
            ("Treeheading.padding",{"sticky":"nswe","children":[
                ("Treeheading.image",{"side":"right","sticky":""}),
                ("Treeheading.text",{"sticky":"we"})
            ]})
        ])
        style.configure("Flat.Treeview.Heading",background="#eef3f8",foreground="#17365d",
                        relief="flat",borderwidth=0,padding=(7,7))
        style.map("Treeview",background=[("selected","#dceaff")],foreground=[("selected","#172b4d")])
        style.configure("TEntry",fieldbackground="#ffffff",padding=7,bordercolor="#304a68")
        style.configure("TLabelframe",background="#ffffff",borderwidth=1,relief="flat",bordercolor="#304a68")
        style.configure("TLabelframe.Label",background="#ffffff",foreground="#415466",font=("Malgun Gothic",9,"bold"))
        style.configure("Horizontal.TProgressbar",troughcolor="#f7f8fa",background="#214f78",
                        bordercolor="#17365d",lightcolor="#214f78",darkcolor="#214f78",thickness=10)
        style.configure("Navy.Vertical.TScrollbar",background="#ffffff",troughcolor="#ffffff",
                        bordercolor="#17365d",arrowcolor="#17365d",lightcolor="#ffffff",darkcolor="#ffffff")
        style.configure("Navy.Horizontal.TScrollbar",background="#ffffff",troughcolor="#ffffff",
                        bordercolor="#17365d",arrowcolor="#17365d",lightcolor="#ffffff",darkcolor="#ffffff")
        style.configure("ScrollCorner.TFrame",background="#e8edf2",borderwidth=1,relief="solid")


    def secondary_button(self,parent,text,command,state="normal"):
        return tk.Button(parent,text=text,command=command,state=state,
                         font=("Malgun Gothic",9),fg="#17365d",bg="#ffffff",
                         activeforeground="#17365d",activebackground="#eef4fa",
                         disabledforeground="#9aa4ae",
                         relief="solid",bd=1,highlightthickness=0,
                         padx=13,pady=5,cursor="arrow")

    @staticmethod
    def short_path(p):
        if not p:return ""
        q=Path(p)
        parent=q.parent.name
        return f"{parent} / {q.name}" if parent else q.name

    def ui(self):
        self.setup_style()
        head=ttk.Frame(self,padding=(18,14,18,10));head.pack(fill="x")
        ttk.Label(head,text=APP,style="Title.TLabel").pack(side="left")
        ttk.Label(head,text="원본 파일은 읽기만 하며 삭제·이동·수정하지 않습니다.",style="Muted.TLabel").pack(side="right")

        self.tabs=ttk.Notebook(self);self.tabs.pack(fill="both",expand=True,padx=10,pady=(0,10))
        folder_tab=ttk.Frame(self.tabs);pair_tab=ttk.Frame(self.tabs)
        self.folder_tab=folder_tab; self.pair_tab=pair_tab
        self.tabs.add(folder_tab,text="  폴더 중복 검사  ")
        self.tabs.add(pair_tab,text="  파일 2개 비교  ")

        # 폴더 중복 검사
        box=ttk.LabelFrame(folder_tab,text="검사할 폴더",padding=8);box.pack(fill="x",padx=8,pady=(8,0))
        top=ttk.Frame(box);top.pack(fill="x",pady=(0,6))
        self.secondary_button(top,"폴더 추가",self.add).pack(side="left")
        self.secondary_button(top,"선택 폴더 제거",self.remove).pack(side="left",padx=6)
        self.lb=tk.Listbox(box,height=4,selectmode="extended");self.lb.pack(fill="x")

        opt=ttk.Frame(folder_tab,padding=8);opt.pack(fill="x")
        ttk.Label(opt,text="유사도 기준").pack(side="left")
        self.cut=tk.IntVar(value=90);ttk.Spinbox(opt,from_=70,to=99,width=5,textvariable=self.cut).pack(side="left",padx=5)
        ttk.Label(opt,text="%").pack(side="left")
        self.start=ttk.Button(opt,text="검사 시작",command=self.go,style="Primary.TButton");self.start.pack(side="right")
        self.save=self.secondary_button(opt,"CSV 저장",self.csv,state="disabled");self.save.pack(side="right",padx=6)

        self.pb=ttk.Progressbar(folder_tab);self.pb.pack(fill="x",padx=8)
        self.status=ttk.Label(folder_tab,text="폴더를 추가해 주세요.",padding=(8,5));self.status.pack(fill="x")

        # 결과 요약 = 필터 버튼
        filterbar=ttk.Frame(folder_tab,padding=(8,5));filterbar.pack(fill="x")
        self.filter_buttons={}
        for key in ("전체","완전 동일","내용 동일","유사","읽기 실패"):
            b=tk.Button(filterbar,text=f"{key} 0",relief="flat",bd=0,padx=11,pady=5,
                        background="#eef2f5",foreground="#344553",
                        activebackground="#dceaf3",command=lambda k=key:self.apply_filter(k))
            b.pack(side="left",padx=(0,5));self.filter_buttons[key]=b
        ttk.Label(filterbar,text="버튼을 누르면 재검사 없이 결과만 필터링합니다.",foreground="#75828c").pack(side="right")

        cols=("judge","score","name","type","size","date","path")
        treebox=ttk.Frame(folder_tab);treebox.pack(fill="both",expand=True,padx=8)
        self.tree=ttk.Treeview(treebox,columns=cols,show="tree headings",selectmode="extended",style="Flat.Treeview")
        self.tree.heading("#0",text="그룹")
        specs=[("judge","판정",105),("score","유사도",72),("name","파일명",245),("type","형식",76),("size","크기",82),("date","수정일",135),("path","경로",520)]
        for c,t,w in specs:self.tree.heading(c,text=t);self.tree.column(c,width=w,minwidth=55)
        self.tree.column("#0",width=68,minwidth=55)
        self.vscroll=ttk.Scrollbar(treebox,orient="vertical",command=self.tree.yview,style="Navy.Vertical.TScrollbar")
        self.hscroll=ttk.Scrollbar(treebox,orient="horizontal",command=self.tree.xview,style="Navy.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=self.vscroll.set,xscrollcommand=self.hscroll.set)
        self.tree.grid(row=0,column=0,sticky="nsew")
        self.vscroll.grid(row=0,column=1,sticky="ns")
        self.hscroll.grid(row=1,column=0,sticky="ew")
        ttk.Frame(treebox,style="ScrollCorner.TFrame").grid(row=1,column=1,sticky="nsew")
        treebox.rowconfigure(0,weight=1);treebox.columnconfigure(0,weight=1)

        # 결과 그룹 시각 구분
        self.tree.tag_configure("exact",background="#f2f7fb")
        self.tree.tag_configure("content",background="#f1f8f3")
        self.tree.tag_configure("similar",background="#fff9e9")
        self.tree.tag_configure("error",background="#fff1f1")

        foot=ttk.Frame(folder_tab,padding=8);foot.pack(fill="x")
        ttk.Button(foot,text="선택한 두 파일 차이 보기",command=self.diffwin,style="Primary.TButton").pack(side="left")

        # 파일 2개 비교
        pairtop=ttk.Frame(pair_tab,padding=(16,14),style="Card.TFrame");pairtop.pack(fill="x",padx=8,pady=(8,6))
        self.pair_paths=[tk.StringVar(),tk.StringVar()]
        for idx,label in enumerate(("A","B")):
            row=ttk.Frame(pairtop,style="Card.TFrame");row.pack(fill="x",pady=5)
            ttk.Label(row,text=label,width=3,font=("Malgun Gothic",11,"bold"),style="Card.TLabel").pack(side="left")
            ttk.Entry(row,textvariable=self.pair_paths[idx]).pack(side="left",fill="x",expand=True,padx=(0,7))
            self.secondary_button(row,"파일 선택",lambda i=idx:self.pick_pair(i)).pack(side="left")
        action=ttk.Frame(pairtop,style="Card.TFrame");action.pack(fill="x",pady=(8,0))
        ttk.Label(action,text="HWP 5.x · HWPX · DOCX · TXT · EPUB / 모든 지원 형식 상호 교차검증",style="CardMuted.TLabel").pack(side="left")
        self.pair_compare_btn=ttk.Button(action,text="두 파일 비교",command=self.compare_pair,style="Primary.TButton");self.pair_compare_btn.pack(side="right")
        self.secondary_button(action,"초기화",self.reset_pair).pack(side="right",padx=(18,8))

        self.pair_summary=ttk.Label(pair_tab,text="비교할 파일 두 개를 선택해 주세요.",padding=(12,8),font=("Malgun Gothic",10,"bold"))
        self.pair_summary.pack(fill="x")
        pairprog=ttk.Frame(pair_tab,padding=(12,0,12,7));pairprog.pack(fill="x")
        self.pair_progress=tk.DoubleVar(value=0)
        self.pair_progress_bar=ttk.Progressbar(pairprog,maximum=100,variable=self.pair_progress,style="Horizontal.TProgressbar")
        self.pair_progress_bar.pack(side="left",fill="x",expand=True)
        self.pair_progress_text=ttk.Label(pairprog,text="",width=28,anchor="e",foreground="#6f7b86")
        self.pair_progress_text.pack(side="right",padx=(10,0))

        self.pair_note=ttk.Label(pair_tab,text="",padding=(10,5),foreground="#6f7b86")
        self.pair_note.pack(side="bottom",fill="x")
        legend=ttk.Frame(pair_tab,padding=(10,5))
        legend.pack(side="bottom",fill="x")
        ttk.Label(legend,text="차이 표시:").pack(side="left")
        tk.Label(legend,text=" A에만 있음 ",background="#ffdede").pack(side="left",padx=(6,3))
        tk.Label(legend,text=" B에만 있음 ",background="#dff3df").pack(side="left",padx=3)
        tk.Label(legend,text=" 양쪽 내용 변경 ",background="#fff0a8").pack(side="left",padx=3)
        tk.Label(legend,text=" 공백·줄바꿈 차이 ",background="#dcecff").pack(side="left",padx=3)

        pane=ttk.Panedwindow(pair_tab,orient="horizontal");pane.pack(fill="both",expand=True,padx=8,pady=(0,4))
        self.pair_text=[]; self.pair_headers=[]
        for label in ("A","B"):
            f=ttk.Frame(pane,style="Card.TFrame");pane.add(f,weight=1)
            h=ttk.Frame(f,style="Card.TFrame",padding=(10,8));h.pack(fill="x")
            ttk.Label(h,text=label,width=2,font=("Malgun Gothic",10,"bold"),style="PanelTitle.TLabel").pack(side="left")
            hv=tk.StringVar(value="")
            ttk.Label(h,textvariable=hv,style="CardMuted.TLabel").pack(side="left",fill="x",expand=True,padx=(5,0))
            self.pair_headers.append(hv)
            t=tk.Text(f,wrap="word",font=("Malgun Gothic",10),undo=False,relief="flat",bd=1,
                      background="#f7f6f2",foreground="#243447",padx=12,pady=10,
                      highlightthickness=1,highlightbackground="#304a68",highlightcolor="#304a68")
            y=ttk.Scrollbar(f,orient="vertical",command=t.yview,style="Navy.Vertical.TScrollbar");t.configure(yscrollcommand=y.set)
            y.pack(side="right",fill="y");t.pack(fill="both",expand=True)
            self.pair_text.append(t)

    def pick_pair(self,idx):
        p=filedialog.askopenfilename(filetypes=[("지원 문서","*.hwp *.hwpx *.docx *.txt *.epub"),("모든 파일","*.*")])
        if p:
            self.pair_paths[idx].set(p)
            self.pair_headers[idx].set(self.short_path(p))

    def reset_pair(self):
        for v in self.pair_paths:v.set("")
        for v in self.pair_headers:v.set("")
        self.pair_summary["text"]="비교할 파일 두 개를 선택해 주세요."
        self.pair_progress.set(0); self.pair_progress_text["text"]=""
        self.pair_note["text"]=""
        for t in self.pair_text:
            t.configure(state="normal");t.delete("1.0","end");t.configure(state="disabled")
        # 큰 비교 결과에 대한 참조도 즉시 해제
        self.pair_result=None

    def compare_pair(self):
        a,b=[x.get().strip() for x in self.pair_paths]
        if not a or not b:return messagebox.showinfo(APP,"A와 B 파일을 모두 선택해 주세요.")
        self.pair_headers[0].set(self.short_path(a))
        self.pair_headers[1].set(self.short_path(b))
        self.pair_compare_btn["state"]="disabled"
        self.pair_summary["text"]="두 파일을 읽고 비교하는 중입니다…"
        self.pair_progress.set(0); self.pair_progress_text["text"]="준비 중 · 0%"
        for t in self.pair_text:
            t.configure(state="normal");t.delete("1.0","end");t.configure(state="disabled")
        self.pair_compare_seq+=1
        seq=self.pair_compare_seq
        threading.Thread(target=self._compare_pair_worker,args=(a,b,seq),daemon=True).start()

    def _pair_progress(self,seq,value,label):
        self.q.put(("pairprogress",seq,value,label))

    @staticmethod
    def _ws_gaps(text):
        """Return non-whitespace characters and the whitespace gap before each char plus trailing gap."""
        chars=[]; gaps=[]; buf=[]
        for ch in text:
            if ch.isspace():
                buf.append(ch)
            else:
                gaps.append("".join(buf)); buf=[]; chars.append(ch)
        gaps.append("".join(buf))
        return "".join(chars),gaps

    @staticmethod
    def _ws_visible(ws):
        return ws.replace("\t","→").replace(" ","·").replace("\n","↵\n")

    @staticmethod
    def _ws_similarity(a,b):
        """Fast character-like ratio for texts whose non-whitespace content is identical."""
        core_a,gaps_a=App._ws_gaps(a); core_b,gaps_b=App._ws_gaps(b)
        if core_a!=core_b:
            return None,0
        matched=len(core_a); diff_gaps=0
        for ga,gb in zip(gaps_a,gaps_b):
            if ga==gb:
                matched+=len(ga)
            else:
                diff_gaps+=1
                # Same whitespace characters can still count as matches; order is irrelevant for this light metric.
                ca=defaultdict(int)
                for ch in ga: ca[ch]+=1
                for ch in gb:
                    if ca[ch]>0: matched+=1; ca[ch]-=1
        denom=len(a)+len(b)
        return (2*matched/denom if denom else 1.0),diff_gaps

    def _compare_pair_worker(self,a,b,seq):
        try:
            self._pair_progress(seq,5,"A 파일 확인 중")
            ra=R(a); ra.hash=fhash(ra.path)
            self._pair_progress(seq,15,"A 본문 추출 중")
            ra.text,ra.kind=read_text(ra.path); ra.th=thash(ra.text)
            self._pair_progress(seq,30,"B 파일 확인 중")
            rb=R(b); rb.hash=fhash(rb.path)
            self._pair_progress(seq,40,"B 본문 추출 중")
            rb.text,rb.kind=read_text(rb.path); rb.th=thash(rb.text)

            na,nb=pair_norm(ra.text),pair_norm(rb.text)
            ca,cb=compact(ra.text),compact(rb.text)
            exact=ra.hash==rb.hash; cross=(ra.ext!=rb.ext)
            self._pair_progress(seq,55,"유사도 계산 중")
            ws_only=(ca==cb and na!=nb); ws_gap_count=0
            if exact:
                sim=1.0; judge="완전 동일"
            elif na==nb:
                sim=1.0; judge="내용 동일"
            elif ws_only:
                sim,_=self._ws_similarity(na,nb); judge="공백·줄바꿈 차이"
            elif cross:
                # Folder scan and A/B comparison use the same cross-format fingerprint metric.
                sim=dice(semantic_fingerprint(ra.text),semantic_fingerprint(rb.text)); judge="유사/상이"
            elif len(ca)+len(cb)<600000:
                sim=difflib.SequenceMatcher(None,ca,cb,autojunk=False).ratio(); judge="유사/상이"
            else:
                sim=jac(shingles(ca),shingles(cb)); judge="유사/상이"

            self._pair_progress(seq,70,"본문 재정렬 중")
            if cross:
                segments=resync_detail_segments(na,nb)
            elif ws_only:
                segments=_local_detail_segments(na,nb)
            else:
                # Preserve the proven v23 line path for same-format documents.
                la0,lb0=na.splitlines(True),nb.splitlines(True)
                sm=difflib.SequenceMatcher(None,la0,lb0,autojunk=False)
                segments=[]
                for tag,i1,i2,j1,j2 in sm.get_opcodes():
                    aa=''.join(la0[i1:i2]); bb=''.join(lb0[j1:j2])
                    if tag=='equal': segments.append(('equal',aa,bb))
                    else: segments.extend(_detail_block_segments(aa,bb))
                segments=_merge_detail_segments(segments)

            self._pair_progress(seq,88,"차이 분류 중")
            added=removed=changed=0
            for tag,aa,bb in segments:
                if tag=='insert': added+=max(1,bb.count('\n') or 1)
                elif tag=='delete': removed+=max(1,aa.count('\n') or 1)
                elif tag=='replace': changed+=max(1,aa.count('\n'),bb.count('\n'))
                elif tag=='ws': ws_gap_count+=1
            self._pair_progress(seq,95,"결과 정리 중")
            self.q.put(("pairdone",seq,ra,rb,sim,judge,added,removed,changed,segments,[],[],ws_gap_count))
        except Exception as e:
            self.q.put(("pairfatal",seq,str(e)))

    def finish_pair(self,ra,rb,sim,judge,added,removed,changed,ops,la,lb,ws_gap_count):
        self.pair_result=(ra,rb)
        extra=f" · 공백·줄바꿈 {ws_gap_count}곳" if ws_gap_count else ""
        self.pair_summary["text"]=(f"{judge}  ·  유사도 {sim*100:.2f}%  ·  "
            f"A {len(pair_norm(ra.text)):,}자 / B {len(pair_norm(rb.text)):,}자  ·  "
            f"추가 {added}줄 / 삭제 {removed}줄 / 변경 {changed}줄{extra}")
        self.show_detail_segments(ops) if not la and not lb else self.show_side_diff_mixed(ops,la,lb)
        self.pair_note["text"]=(f"A: {ra.kind} · {self.sz(ra.size)}     B: {rb.kind} · {self.sz(rb.size)}  ·  "
                                "공백은 ·, 줄바꿈은 ↵로 표시")
        self.pair_progress.set(100); self.pair_progress_text["text"]="완료 · 100%"
        self.pair_compare_btn["state"]="normal"

    def show_detail_segments(self,segments):
        ta,tb=self.pair_text
        for t in (ta,tb):
            t.configure(state="normal"); t.delete("1.0","end")
            t.tag_configure("same")
            t.tag_configure("del",background="#ffdede")
            t.tag_configure("add",background="#dff3df")
            t.tag_configure("chg",background="#fff0a8")
            t.tag_configure("ws",background="#dcecff",foreground="#315b7b")
        for tag,a,b in segments:
            if tag=='equal': ta.insert('end',a,'same'); tb.insert('end',b,'same')
            elif tag=='delete': ta.insert('end',a,'del')
            elif tag=='insert': tb.insert('end',b,'add')
            elif tag=='replace': ta.insert('end',a,'chg'); tb.insert('end',b,'chg')
            elif tag=='ws':
                ta.insert('end',self._ws_visible(a),'ws'); tb.insert('end',self._ws_visible(b),'ws')
        for t in (ta,tb): t.configure(state="disabled")

    def _insert_ws_pair(self,ta,tb,a,b):
        core_a,gaps_a=self._ws_gaps(a); core_b,gaps_b=self._ws_gaps(b)
        if core_a!=core_b:
            ta.insert("end",a,"chg"); tb.insert("end",b,"chg"); return 0
        diffs=0
        for idx,ch in enumerate(core_a):
            ga,gb=gaps_a[idx],gaps_b[idx]
            if ga==gb:
                ta.insert("end",ga,"same"); tb.insert("end",gb,"same")
            else:
                diffs+=1
                ta.insert("end",self._ws_visible(ga),"ws")
                tb.insert("end",self._ws_visible(gb),"ws")
            ta.insert("end",ch,"same"); tb.insert("end",ch,"same")
        ga,gb=gaps_a[-1],gaps_b[-1]
        if ga==gb:
            ta.insert("end",ga,"same"); tb.insert("end",gb,"same")
        else:
            diffs+=1
            ta.insert("end",self._ws_visible(ga),"ws"); tb.insert("end",self._ws_visible(gb),"ws")
        return diffs

    def show_side_diff_mixed(self,ops,la,lb):
        ta,tb=self.pair_text
        for t in (ta,tb):
            t.configure(state="normal"); t.delete("1.0","end")
            t.tag_configure("same")
            t.tag_configure("del",background="#ffdede")
            t.tag_configure("add",background="#dff3df")
            t.tag_configure("chg",background="#fff0a8")
            t.tag_configure("ws",background="#dcecff",foreground="#315b7b")
        for tag,i1,i2,j1,j2 in ops:
            if tag=="equal":
                for line in la[i1:i2]: ta.insert("end",line+"\n","same")
                for line in lb[j1:j2]: tb.insert("end",line+"\n","same")
            elif tag=="delete":
                for line in la[i1:i2]: ta.insert("end",line+"\n","del")
            elif tag=="insert":
                for line in lb[j1:j2]: tb.insert("end",line+"\n","add")
            else:
                a="\n".join(la[i1:i2]); b="\n".join(lb[j1:j2])
                if compact(a)==compact(b):
                    self._insert_ws_pair(ta,tb,a,b)
                    ta.insert("end","\n","same"); tb.insert("end","\n","same")
                else:
                    for line in la[i1:i2]: ta.insert("end",line+"\n","chg")
                    for line in lb[j1:j2]: tb.insert("end",line+"\n","chg")
        for t in (ta,tb): t.configure(state="disabled")


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
        self.scan_cut=max(.70,min(.99,self.cut.get()/100))
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
            valid=[r for r in rec if r.text and not r.err]; cut=self.scan_cut; edges=[]
            # v29: 형식별 예외망 없이 하나의 공통 검색용 fingerprint를 사용한다.
            # v23의 검증된 후보망은 그대로 두고, 공통 fingerprint는 문서별 상위 소수 후보만 보강한다.
            self.q.put(("phase",f"유사 문서 후보를 만드는 중 · {len(valid)}개 문서"))
            inv=defaultdict(list); finv=defaultdict(list)
            sigs=[]; fps=[]; lengths=[]; sem_lengths=[]
            for idx,r in enumerate(valid):
                c=compact(r.text); lengths.append(len(c))
                sg=shingles(c,k=9,limit=1200); sigs.append(sg)
                fp=semantic_fingerprint(r.text,k=11,mod=97,limit=1800); fps.append(fp)
                sem_lengths.append(len(semantic_stream(r.text)))
                for token in sg: inv[token].append(idx)
                for token in fp: finv[token].append(idx)

            # 1) v23 후보망: 기존 성능 특성을 유지한다.
            pair_hits=defaultdict(int)
            for ids in inv.values():
                if len(ids)>80: continue
                for x in range(len(ids)):
                    for y in range(x+1,len(ids)):
                        i,j=ids[x],ids[y]
                        if lengths[i] and lengths[j] and min(lengths[i],lengths[j])/max(lengths[i],lengths[j])>=.70:
                            pair_hits[(i,j)]+=1
            base_candidates=set()
            for pair,hits in pair_hits.items():
                i,j=pair
                if hits>=2 or min(lengths[i],lengths[j])<120: base_candidates.add(pair)

            # 2) 공통 교차형식 보강: fingerprint 조합을 전부 만들지 않는다.
            # 각 문서가 공유하는 희귀 fingerprint만 세고 상위 6개 후보만 추가한다.
            # 따라서 700개 문서라도 보강 후보는 이론상 약 4,200쌍 이하이다.
            extra_candidates=set(); TOP_N=6; MAX_DF=14
            for i,fp in enumerate(fps):
                hits=defaultdict(int)
                for token in fp:
                    ids=finv.get(token,())
                    if len(ids)>MAX_DF: continue
                    for j in ids:
                        if j==i: continue
                        # 극단적인 길이 차이만 제외. 70% 절대조건은 사용하지 않는다.
                        if sem_lengths[i] and sem_lengths[j] and min(sem_lengths[i],sem_lengths[j])/max(sem_lengths[i],sem_lengths[j])<.20:
                            continue
                        hits[j]+=1
                ranked=sorted(hits.items(),key=lambda x:(-x[1],x[0]))[:TOP_N]
                for j,h in ranked:
                    if h>=3 or min(sem_lengths[i],sem_lengths[j])<120:
                        extra_candidates.add((min(i,j),max(i,j)))

            candidate_source={}
            for pair in base_candidates: candidate_source[pair]='base'
            for pair in extra_candidates: candidate_source.setdefault(pair,'fingerprint')
            candidates=[]
            for (i,j),source in candidate_source.items():
                pair=tuple(sorted((valid[i].path,valid[j].path)))
                if pair not in claimed:candidates.append((i,j,source))
            candidates.sort(key=lambda x:(x[0],x[1]))
            self.q.put(("phase",f"유사도 정밀 비교 중 · 후보 {len(candidates):,}쌍 (기존 {len(base_candidates):,} / 보강 {len(extra_candidates):,})"))

            for n,(i,j,source) in enumerate(candidates,1):
                a,b=valid[i],valid[j]
                sem=dice(fps[i],fps[j])
                coverage=semantic_length_ratio(sem_lengths[i],sem_lengths[j])
                if source=='fingerprint':
                    # Fingerprints are for locating matching content. Final score also reflects
                    # one-sided added/omitted text through whole-document semantic length coverage.
                    score=min(sem,coverage)
                else:
                    rough=jac(sigs[i],sigs[j])
                    ca,cb=compact(a.text),compact(b.text)
                    # Keep v29's fast path, but never let sampled fingerprints hide omitted text.
                    if sem>=cut:
                        score=min(sem,coverage)
                    elif rough<max(.30,cut-.40):
                        continue
                    elif len(ca)+len(cb)<600000:
                        aligned=difflib.SequenceMatcher(None,ca,cb,autojunk=False).ratio()
                        score=min(aligned,coverage)
                    else:
                        score=min(max(rough,sem),coverage)
                if score>=cut:edges.append((a,b,score))
                if n%100==0:self.q.put(("phase",f"유사도 정밀 비교 중 · {n:,}/{len(candidates):,}쌍"))
            for a,b,score in sorted(edges,key=lambda x:-x[2]):groups.append(("유사",score,[a,b]))

            # 같은 파일 조합은 순서가 뒤집히거나 여러 경로에서 다시 발견돼도 한 번만 표시.
            # 판정 우선순위: 완전 동일 > 내용 동일 > 유사
            priority={"완전 동일":3,"내용 동일":2,"유사":1}
            unique={}
            for judge,score,items in groups:
                key=tuple(sorted(os.path.normcase(os.path.abspath(display_safe(r.path))) for r in items))
                old=unique.get(key)
                if old is None or priority[judge]>priority[old[0]] or (priority[judge]==priority[old[0]] and score>old[1]):
                    unique[key]=(judge,score,items)
            groups=list(unique.values())
            groups.sort(key=lambda g:(-priority[g[0]],-g[1],tuple(r.name.lower() for r in g[2])))

            self.records=rec;self.groups=groups;self.q.put(("done",))
        except Exception as e:self.q.put(("fatal",str(e)))
    def poll(self):
        try:
            while 1:
                m=self.q.get_nowait()
                if m[0]=="max":self.pb["maximum"]=m[1];self.pb["value"]=0
                elif m[0]=="p":self.pb["value"]=m[1];self.status["text"]=m[2]
                elif m[0]=="phase":self.status["text"]=m[1]
                elif m[0]=="done":self.render()
                elif m[0]=="pairprogress":
                    if m[1]==self.pair_compare_seq:
                        self.pair_progress.set(m[2]); self.pair_progress_text["text"]=f"{m[3]} · {m[2]}%"
                elif m[0]=="pairdone":
                    if m[1]==self.pair_compare_seq:self.finish_pair(*m[2:])
                elif m[0]=="pairfatal":
                    if m[1]==self.pair_compare_seq:
                        self.pair_compare_btn["state"]="normal"
                        self.pair_summary["text"]="비교에 실패했습니다."
                        self.pair_progress_text["text"]="실패"
                        messagebox.showerror(APP,"비교하지 못했습니다.\n\n"+m[2])
                elif m[0]=="fatal":self.running=False;self.start["state"]="normal";messagebox.showerror(APP,m[1])
        except queue.Empty:pass
        self.after(100,self.poll)
    def render(self):
        errs=[r for r in self.records if r.err]
        counts={"완전 동일":0,"내용 동일":0,"유사":0}
        for j,score,rs in self.groups:counts[j]=counts.get(j,0)+1
        exact_n=counts.get("완전 동일",0)
        content_n=counts.get("내용 동일",0)
        similar_n=counts.get("유사",0)
        failed_n=len(errs)
        self.result_counts={
            "전체":exact_n+content_n+similar_n+failed_n,
            "완전 동일":exact_n,
            "내용 동일":content_n,
            "유사":similar_n,
            "읽기 실패":failed_n
        }
        for k,b in self.filter_buttons.items():
            b.configure(text=f"{k} {self.result_counts[k]}")
        self.active_filter="전체"
        self.apply_filter("전체")
        self.status["text"]=(f"완료 · 대상 {len(self.records)}개 · 결과 그룹 {len(self.groups)}개 · "
                             f"읽기 실패 {len(errs)}개")
        self.running=False;self.start["state"]="normal";self.save["state"]="normal"

    def apply_filter(self,key):
        self.active_filter=key
        for k,b in self.filter_buttons.items():
            if k==key:
                b.configure(background="#315b7b",foreground="white",relief="sunken")
            else:
                b.configure(background="#eef2f5",foreground="#344553",relief="flat")
        self.render_results()

    def render_results(self):
        self.tree.delete(*self.tree.get_children())
        shown=0
        for gi,(j,score,rs) in enumerate(self.groups,1):
            if self.active_filter not in ("전체",j):continue
            tag={"완전 동일":"exact","내용 동일":"content","유사":"similar"}.get(j,"")
            root=self.tree.insert("","end",text=str(gi),values=(j,group_score_text(j,score),"","","","",""),open=True,tags=(tag,))
            for r in rs:
                self.tree.insert(root,"end",values=("","",display_safe(r.name),r.kind,self.sz(r.size),
                    time.strftime("%Y-%m-%d %H:%M",time.localtime(r.mtime)),display_safe(r.path)))
            shown+=1
        if self.active_filter in ("전체","읽기 실패"):
            errs=[r for r in self.records if r.err]
            if errs:
                root=self.tree.insert("","end",text="오류",values=("읽기 실패","",f"{len(errs)}개","","","",""),open=True,tags=("error",))
                for r in errs:
                    self.tree.insert(root,"end",values=("","",r.name,r.ext,self.sz(r.size),"",display_safe(r.path)+" | "+display_safe(r.err)))
        if self.tree.get_children():
            self.tree.yview_moveto(0);self.tree.xview_moveto(0)

    def selected(self):
        d={r.path:r for r in self.records};out=[]
        for x in self.tree.selection():
            v=self.tree.item(x,"values")
            if len(v)>=7 and v[6] in d:out.append(d[v[6]])
        return out
    def diffwin(self):
        rs=self.selected()
        if len(rs)!=2:return messagebox.showinfo(APP,"파일 두 개를 Ctrl+클릭으로 선택해 주세요.")
        if str(self.pair_compare_btn["state"])=="disabled":
            return messagebox.showinfo(APP,"현재 파일 비교가 진행 중입니다. 완료된 뒤 다시 시도해 주세요.")
        a,b=rs
        if self.skip_diff_transfer_notice:
            return self.transfer_to_pair(a,b)
        self.diff_transfer_notice(a,b)

    def diff_transfer_notice(self,a,b):
        w=tk.Toplevel(self)
        w.title("파일 2개 비교")
        w.resizable(False,False)
        w.transient(self)
        w.grab_set()
        body=ttk.Frame(w,padding=(22,18,22,8));body.pack(fill="both",expand=True)
        ttk.Label(body,text="선택한 두 파일을 ‘파일 2개 비교’ 탭에서 비교합니다.",
                  font=("Malgun Gothic",10,"bold")).pack(anchor="w")
        ttk.Label(body,text="기존에 비교하던 파일과 결과가 있다면 새 파일로 교체됩니다.",
                  foreground="#5f6f7a").pack(anchor="w",pady=(7,14))
        skip=tk.BooleanVar(value=False)
        ttk.Checkbutton(body,text="이번 실행 중에는 다시 보지 않기",variable=skip).pack(anchor="w")
        buttons=ttk.Frame(w,padding=(22,8,22,18));buttons.pack(fill="x")
        def cancel():
            w.destroy()
        def proceed():
            if skip.get():self.skip_diff_transfer_notice=True
            w.destroy()
            self.transfer_to_pair(a,b)
        ttk.Button(buttons,text="취소",command=cancel).pack(side="right")
        ttk.Button(buttons,text="비교하기",command=proceed).pack(side="right",padx=(0,7))
        w.protocol("WM_DELETE_WINDOW",cancel)
        w.update_idletasks()
        x=self.winfo_rootx()+(self.winfo_width()-w.winfo_width())//2
        y=self.winfo_rooty()+(self.winfo_height()-w.winfo_height())//2
        w.geometry(f"+{max(0,x)}+{max(0,y)}")
        w.focus_set()

    def transfer_to_pair(self,a,b):
        # 폴더 검사 결과는 그대로 두고 비교 탭의 A/B와 결과만 새 선택으로 교체한다.
        self.reset_pair()
        self.pair_paths[0].set(a.path)
        self.pair_paths[1].set(b.path)
        self.pair_headers[0].set(self.short_path(a.path))
        self.pair_headers[1].set(self.short_path(b.path))
        self.tabs.select(self.pair_tab)
        self.update_idletasks()
        self.compare_pair()
    def csv(self):
        p=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv")],initialfile="문서_중복_유사본_결과.csv")
        if not p:return
        with open(p,"w",encoding="utf-8-sig",newline="") as f:
            w=csv.writer(f);w.writerow(["그룹","판정","유사도","파일명","형식","크기","수정일","경로"])
            for gi,(j,s,rs) in enumerate(self.groups,1):
                for r in rs:w.writerow([gi,j,f"{shown_group_score(j,s):.2f}%",r.name,r.kind,r.size,time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(r.mtime)),r.path])
        messagebox.showinfo(APP,"저장했습니다.")
    @staticmethod
    def sz(n):
        for u in ("B","KB","MB","GB"):
            if n<1024 or u=="GB":return f"{int(n)} B" if u=="B" else f"{n:.1f} {u}"
            n/=1024
if __name__=="__main__":App().mainloop()
