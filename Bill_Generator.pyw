from pathlib import Path
import webbrowser
import subprocess
import tempfile
import traceback
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

root = tk.Tk()
root.report_callback_exception = lambda kind, value, tb: messagebox.showerror('Bill Generator', ''.join(traceback.format_exception(kind,value,tb)),parent=root)
root.title('Senura | Bill Generator')
root.geometry('730x510')
root.minsize(640,490)
header=tk.Frame(root,bg='#18394a',padx=24,pady=16)
header.pack(fill='x')
tk.Label(header,text='Monthly hiring bills',font=('Segoe UI',20,'bold'),fg='white',bg='#18394a').pack(anchor='w')
tk.Label(header,text='Senura Civil Engineering (Pvt) Ltd',fg='#cce4e6',bg='#18394a').pack(anchor='w')
footer=ttk.Frame(root,padding=(22,8))
footer.pack(side='bottom',fill='x')
ttk.Label(footer,text='Developed by Sahan Kaushalya  ·  0763708504').pack(side='left')
ttk.Button(footer,text='GitHub · sksahan',command=lambda: webbrowser.open('https://github.com/sksahan')).pack(side='right')
destination=tk.StringVar()
def choose_destination():
    path=filedialog.askdirectory(title='Choose folder for individual bills',parent=root)
    if path:
        destination.set(path)
panel = ttk.Frame(root, padding=22)
panel.pack(fill='both', expand=True)
panel.columnconfigure(0, weight=1)
source = tk.StringVar()
status = tk.StringVar(value='Choose a site monthly revenue report.')
def browse():
    path = filedialog.askopenfilename(title='Select monthly report', filetypes=[('Excel workbook','*.xlsx')])
    if path:
        source.set(path)
        if not destination.get():
            destination.set(str(Path(path).parent))
def generate():
    report = Path(source.get())
    if not report.is_file() or report.suffix.lower() != '.xlsx':
        messagebox.showerror('Report required','Select a monthly Excel report.')
        return
    target = destination.get().strip()
    if not target or not Path(target).is_dir():
        choose_destination()
        target=destination.get().strip()
        if not target or not Path(target).is_dir():
            return
    try:
        for name in ('engine.ps1','template.xlsx'):
            if not Path(__file__).with_name(name).is_file():
                raise OSError('Missing '+name+'. Extract all files from the ZIP into one folder.')
        log = tempfile.NamedTemporaryFile(prefix='Senura_Bill_',suffix='.log',delete=False)
        with log:
            proc = subprocess.Popen(['powershell.exe','-NoProfile','-NonInteractive','-STA','-ExecutionPolicy','Bypass','-File',str(Path(__file__).with_name('engine.ps1')),str(report.resolve()),target], creationflags=subprocess.CREATE_NO_WINDOW,stdout=log,stderr=subprocess.STDOUT)
    except OSError as error:
        messagebox.showerror('Cannot start generator',str(error))
        return
    for widget in controls:
        widget.state(['disabled'])
    progress.start(12)
    status.set('Generating bills in Excel…')
    started=time.monotonic()
    def poll():
        code=proc.poll()
        lines=Path(log.name).read_bytes().decode('utf-8',errors='replace').strip().splitlines()
        last=lines[-1] if lines else 'Starting generator...'
        if code is None:
            status.set(f'{last} ({int(time.monotonic()-started)} s)')
            root.after(500,poll)
            return
        progress.stop()
        for widget in controls:
            widget.state(['!disabled'])
        status.set(last)
        if code:
            details='\n'.join(lines[-10:]) or f'Generator exited with code {code}.'
            messagebox.showerror('Generation failed',details+'\n\nLog: '+log.name,parent=root)
        else:
            messagebox.showinfo('Bills generated',last,parent=root)
    poll()
controls=[]
for row,label,var,action in [(0,'Monthly site report (.xlsx)',source,browse),(2,'Save bills in',destination,choose_destination)]:
    ttk.Label(panel,text=label,font=('Segoe UI',10,'bold')).grid(row=row,column=0,sticky='w',pady=(5,4))
    entry=ttk.Entry(panel,textvariable=var)
    entry.grid(row=row+1,column=0,sticky='ew',ipady=4)
    button=ttk.Button(panel,text='Browse…',command=action)
    button.grid(row=row+1,column=1,padx=(10,0))
    controls.extend((entry,button))
run=ttk.Button(panel,text='Generate individual bills',command=generate)
run.grid(row=4,column=0,sticky='w',pady=(18,12))
controls.append(run)
progress=ttk.Progressbar(panel,mode='indeterminate')
progress.grid(row=5,column=0,columnspan=2,sticky='ew')
ttk.Label(panel,textvariable=status,wraplength=570).grid(row=6,column=0,columnspan=2,sticky='w',pady=8)
root.mainloop()
