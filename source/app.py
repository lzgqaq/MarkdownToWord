"""Windows desktop interface; all widget updates run on the Tk main thread."""
from pathlib import Path
import os
import queue
import threading
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from converter import convert

class App:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.result = None
        self.events = queue.Queue()
        root.title('Markdown 转 Word')
        root.geometry('740x510')
        root.minsize(640, 470)
        root.configure(bg='#f3f5f8')
        root.protocol('WM_DELETE_WINDOW', self.close)
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI', 10))
        style.configure('TFrame', background='#f3f5f8')
        style.configure('TLabel', background='#f3f5f8', foreground='#243247')
        style.configure('Title.TLabel', font=('Microsoft YaHei UI', 22, 'bold'))
        style.configure('TButton', padding=(13, 8))
        style.configure('Accent.TButton', background='#245cd6', foreground='white')
        style.map('Accent.TButton', background=[('active','#1949ac'),('disabled','#9daac2')])
        frame = ttk.Frame(root, padding=28)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text='Markdown 转 Word', style='Title.TLabel').grid(row=0,column=0,columnspan=2,sticky='w')
        ttk.Label(frame, text='选择文档，一键生成可编辑的 Word 文件。').grid(row=1,column=0,columnspan=2,sticky='w',pady=(6,22))
        self.source, self.output = tk.StringVar(), tk.StringVar()
        ttk.Label(frame, text='Markdown 文件').grid(row=2,column=0,sticky='w')
        self.source_entry = ttk.Entry(frame, textvariable=self.source)
        self.source_entry.grid(row=3,column=0,sticky='ew',pady=(7,17),ipady=6)
        self.pick_button = ttk.Button(frame,text='选择文件',command=self.pick)
        self.pick_button.grid(row=3,column=1,padx=(12,0),pady=(7,17))
        ttk.Label(frame,text='Word 保存位置').grid(row=4,column=0,sticky='w')
        self.output_entry = ttk.Entry(frame,textvariable=self.output)
        self.output_entry.grid(row=5,column=0,sticky='ew',pady=(7,20),ipady=6)
        self.save_button = ttk.Button(frame,text='另存为…',command=self.save)
        self.save_button.grid(row=5,column=1,padx=(12,0),pady=(7,20))
        self.convert_button = ttk.Button(frame,text='转换为 Word',style='Accent.TButton',command=self.start)
        self.convert_button.grid(row=6,column=0,columnspan=2,sticky='ew')
        self.progress = ttk.Progressbar(frame,mode='indeterminate')
        self.progress.grid(row=7,column=0,columnspan=2,sticky='ew',pady=(14,10))
        self.status = tk.StringVar(value='就绪 · 本地处理，无需联网')
        ttk.Label(frame,textvariable=self.status,wraplength=600).grid(row=8,column=0,columnspan=2,sticky='w')
        buttons = ttk.Frame(frame)
        buttons.grid(row=9,column=0,columnspan=2,sticky='w',pady=(16,0))
        self.open_button = ttk.Button(buttons,text='打开 Word',state='disabled',command=lambda:self.open(False))
        self.open_button.pack(side='left')
        self.folder_button = ttk.Button(buttons,text='打开文件夹',state='disabled',command=lambda:self.open(True))
        self.folder_button.pack(side='left',padx=10)
        self.locked = [self.source_entry,self.output_entry,self.pick_button,self.save_button,self.convert_button]
        root.after(100,self.poll)

    def pick(self):
        filename = filedialog.askopenfilename(title='选择 Markdown 文件',filetypes=[('Markdown 文件','*.md *.markdown'),('所有文件','*.*')])
        if filename:
            self.source.set(filename)
            self.output.set(str(Path(filename).with_suffix('.docx')))

    def save(self):
        filename = filedialog.asksaveasfilename(title='保存 Word 文件',defaultextension='.docx',initialfile=Path(self.output.get()).name or '文档.docx',filetypes=[('Word 文档','*.docx')])
        if filename: self.output.set(filename)

    def start(self):
        source, output = self.source.get().strip(), self.output.get().strip()
        if not source or not output:
            messagebox.showinfo('请选择文件','请选择 Markdown 文件并设置 Word 保存位置。')
            return
        overwrite = False
        if Path(output).exists():
            overwrite = messagebox.askyesno('确认覆盖',f'文件已存在：\n{output}\n\n是否覆盖？')
            if not overwrite: return
        self.busy = True
        self.result = None
        for widget in self.locked: widget.configure(state='disabled')
        self.open_button.configure(state='disabled')
        self.folder_button.configure(state='disabled')
        self.progress.start(12)
        self.status.set('正在转换，请稍候…')
        threading.Thread(target=self.worker,args=(source,output,overwrite),daemon=True).start()

    def worker(self, source, output, overwrite):
        try: self.events.put(('success',Path(output).resolve(),convert(source,output,overwrite)))
        except Exception as exc: self.events.put(('error',str(exc),None))

    def poll(self):
        try:
            kind, value, warnings = self.events.get_nowait()
        except queue.Empty: pass
        else:
            self.busy = False
            self.progress.stop()
            for widget in self.locked: widget.configure(state='normal')
            if kind == 'success':
                self.result = value
                self.status.set('转换完成' + ('，有内容提示' if warnings else '') + '。可打开 Word 查看。')
                self.open_button.configure(state='normal')
                self.folder_button.configure(state='normal')
                if warnings: messagebox.showwarning('已完成，存在提示','\n\n'.join(warnings)[:6000])
            else:
                self.status.set('转换失败，请检查文件后重试。')
                messagebox.showerror('转换失败',value)
        self.root.after(100,self.poll)

    def open(self, folder):
        try: os.startfile(str(self.result.parent if folder else self.result))
        except OSError as exc: messagebox.showerror('无法打开',str(exc))

    def close(self):
        if self.busy:
            messagebox.showinfo('正在转换','请等待转换结束后关闭窗口。')
        else: self.root.destroy()

def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    root = tk.Tk()
    app = App(root)
    if len(sys.argv) == 5 and sys.argv[1] == '--smoke-test':
        # Packaging QA: exercise the same Tk callbacks and worker as a user click.
        root.withdraw()
        report = Path(sys.argv[4])
        app.source.set(sys.argv[2])
        app.output.set(sys.argv[3])
        failures = []
        messagebox.showerror = lambda title, text: failures.append(text)
        messagebox.showwarning = lambda title, text: failures.append(text)
        def check():
            if app.busy:
                root.after(100,check)
            else:
                report.write_text(json.dumps({'success':bool(app.result), 'errors':failures,
                    'status':app.status.get()},ensure_ascii=False),encoding='utf-8')
                root.destroy()
        root.after(20,app.start)
        root.after(200,check)
        root.after(200000,root.destroy)
    root.mainloop()

if __name__ == '__main__': main()
