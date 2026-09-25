import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import syhwp

APP='HWP 5.x 읽기 테스트'
class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP); self.geometry('850x650'); self.minsize(650,450)
        top=ttk.Frame(self,padding=10); top.pack(fill='x')
        ttk.Label(top,text='HWP 5.x 읽기 테스트',font=('',14,'bold')).pack(side='left')
        ttk.Button(top,text='HWP 파일 선택',command=self.open_hwp).pack(side='right')
        self.info=ttk.Label(self,text='HWP 5.x 파일을 선택해 주세요.',padding=(10,0,10,8)); self.info.pack(fill='x')
        self.text=tk.Text(self,wrap='word',font=('',10)); self.text.pack(fill='both',expand=True,padx=10,pady=(0,10))
    def open_hwp(self):
        p=filedialog.askopenfilename(filetypes=[('한글 HWP','*.hwp'),('모든 파일','*.*')])
        if not p:return
        try:
            fmt=syhwp.detect_format(p)
            if fmt!='hwp5': raise ValueError(f'HWP 5.x로 인식되지 않습니다. 감지 형식: {fmt}')
            doc=syhwp.open(p)
            content=syhwp.extract_text(p)
            self.text.delete('1.0','end'); self.text.insert('1.0',content)
            ver=getattr(doc,'version','확인 불가')
            self.info.config(text=f'{Path(p).name}  |  형식: {fmt}  |  버전: {ver}  |  추출 문자: {len(content):,}자')
        except Exception as e:
            self.text.delete('1.0','end'); self.text.insert('1.0',f'읽기 실패\n\n{type(e).__name__}: {e}')
            self.info.config(text=Path(p).name+'  |  읽기 실패')
            messagebox.showerror(APP,f'파일을 읽지 못했습니다.\n\n{e}')
if __name__=='__main__': App().mainloop()
