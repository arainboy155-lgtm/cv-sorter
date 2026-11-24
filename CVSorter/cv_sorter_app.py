"""
CV Sorter - Teacher CV Auto Sorting Desktop App
================================================
Chalane ke liye:   python cv_sorter_app.py
EXE banane ke liye: build_exe.bat chalayein (README dekhein)
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import cv_engine as eng

try:  # Drag & drop optional hai
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

APP_TITLE = "CV Sorter - Teacher CV Auto Sorting"
COMPANY_NAME = "Your Company Name"   # <-- apni company ka naam yahan likhein

PRIMARY = "#1F4E78"
ACCENT = "#2E86C1"
BG = "#F4F6F9"


def open_path(path):
    if not path or not os.path.exists(path):
        messagebox.showwarning("Not found", f"Ye file/folder nahi mila:\n{path}")
        return
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class CVSorterApp:
    def __init__(self, root):
        self.root = root
        self.cfg = eng.load_config()
        self.q = queue.Queue()
        self.busy = False
        self.watch_on = tk.BooleanVar(value=False)
        self.seen_in_watch = set()

        root.title(f"{APP_TITLE}  |  {COMPANY_NAME}")
        root.geometry("1180x720")
        root.minsize(950, 600)
        root.configure(bg=BG)
        self._style()
        self._build_ui()
        self.refresh_table()
        self.root.after(200, self._poll_queue)

    # ------------------------------------------------------------------ UI
    def _style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TFrame", background=BG)
        s.configure("TLabel", background=BG, font=("Segoe UI", 10))
        s.configure("Header.TLabel", background=PRIMARY, foreground="white",
                    font=("Segoe UI", 16, "bold"))
        s.configure("Sub.TLabel", background=PRIMARY, foreground="#D6EAF8", font=("Segoe UI", 9))
        s.configure("TButton", font=("Segoe UI", 10), padding=6)
        s.configure("Accent.TButton", background=ACCENT, foreground="white",
                    font=("Segoe UI", 10, "bold"), padding=8)
        s.map("Accent.TButton", background=[("active", PRIMARY)])
        s.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        s.configure("Stat.TLabel", background="white", font=("Segoe UI", 9), padding=6)

    def _build_ui(self):
        # Header
        header = tk.Frame(self.root, bg=PRIMARY, height=60)
        header.pack(fill="x")
        ttk.Label(header, text="  📄 CV Sorter", style="Header.TLabel").pack(side="left", pady=10)
        ttk.Label(header, text=f"{COMPANY_NAME}  •  Teacher CVs ko khud pehchan kar alag folders mein save karta hai",
                  style="Sub.TLabel").pack(side="left", padx=10)

        # Toolbar
        bar = ttk.Frame(self.root, padding=(10, 8))
        bar.pack(fill="x")
        ttk.Button(bar, text="➕ Add CV Files", style="Accent.TButton",
                   command=self.add_files).pack(side="left", padx=3)
        ttk.Button(bar, text="📁 Add Whole Folder", command=self.add_folder).pack(side="left", padx=3)
        ttk.Button(bar, text="📂 Open Sorted Folder",
                   command=lambda: open_path(self.cfg["output_folder"])).pack(side="left", padx=3)
        ttk.Button(bar, text="📊 Export Excel Report", command=self.export).pack(side="left", padx=3)
        ttk.Button(bar, text="⚙ Settings", command=self.open_settings).pack(side="left", padx=3)
        ttk.Checkbutton(bar, text="Auto-Watch Folder", variable=self.watch_on,
                        command=self.toggle_watch).pack(side="right", padx=6)

        # Drop zone
        drop_text = ("CV files yahan Drag & Drop karein, ya 'Add CV Files' dabayein  (PDF, DOCX, TXT)"
                     if HAS_DND else
                     "'Add CV Files' ya 'Add Whole Folder' dabayein  (PDF, DOCX, TXT)")
        self.drop = tk.Label(self.root, text=drop_text, bg="white", fg="#566573",
                             font=("Segoe UI", 11), relief="groove", bd=2, height=3)
        self.drop.pack(fill="x", padx=12, pady=(0, 6))
        if HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)

        # Progress
        pf = ttk.Frame(self.root, padding=(12, 0))
        pf.pack(fill="x")
        self.progress = ttk.Progressbar(pf, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.status = ttk.Label(pf, text="Ready", width=40)
        self.status.pack(side="left", padx=8)

        # Main split: table (left) + stats/log (right)
        main = ttk.PanedWindow(self.root, orient="horizontal")
        main.pack(fill="both", expand=True, padx=12, pady=8)

        left = ttk.Frame(main)
        main.add(left, weight=4)
        sf = ttk.Frame(left)
        sf.pack(fill="x", pady=(0, 4))
        ttk.Label(sf, text="🔍 Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_table())
        ttk.Entry(sf, textvariable=self.search_var, width=40).pack(side="left", padx=5)
        ttk.Label(sf, text="(naam, subject, email, phone)", foreground="#888").pack(side="left")
        ttk.Button(sf, text="Refresh", command=self.refresh_table).pack(side="right")

        cols = ("id", "date", "name", "subject", "conf", "email", "phone", "exp", "file")
        heads = ("#", "Date", "Candidate Name", "Subject", "Match %", "Email", "Phone",
                 "Exp (Yrs)", "Original File")
        widths = (40, 110, 160, 190, 70, 190, 110, 70, 200)
        tf = ttk.Frame(left)
        tf.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h, command=lambda c=c: self._sort_by(c, False))
            self.tree.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        self.tree.tag_configure("unsorted", background="#FDEBD0")
        self.tree.tag_configure("low", background="#FCF3CF")
        self.tree.bind("<Double-1>", lambda e: self.open_selected())
        self.tree.bind("<Button-3>", self._context_menu)

        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="📄 Open CV", command=self.open_selected)
        self.menu.add_command(label="📂 Open Containing Folder", command=self.open_selected_folder)
        self.menu.add_separator()
        self.menu.add_command(label="✏ Change Subject (galat detect hua?)", command=self.change_subject)
        self.menu.add_command(label="🗑 Remove from List", command=self.remove_selected)

        right = ttk.Frame(main)
        main.add(right, weight=1)
        ttk.Label(right, text="📈 Subject-wise Count", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.stats = tk.Text(right, height=14, width=34, font=("Segoe UI", 9), bd=0,
                             bg="white", state="disabled", padx=8, pady=6)
        self.stats.pack(fill="x", pady=(2, 10))
        ttk.Label(right, text="📝 Activity Log", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.log = tk.Text(right, height=14, width=34, font=("Consolas", 8), bd=0,
                           bg="#1C2833", fg="#D5DBDB", state="disabled", padx=6, pady=6)
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("ok", foreground="#58D68D")
        self.log.tag_configure("warn", foreground="#F5B041")
        self.log.tag_configure("err", foreground="#EC7063")

        ttk.Label(self.root, text="Tip: Kisi row par Right-Click karke subject badal sakte hain. "
                                  "Orange rows = manual review chahiye.",
                  foreground="#666").pack(anchor="w", padx=12, pady=(0, 6))

    # ------------------------------------------------------------ actions
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="CV files select karein",
            filetypes=[("CV Files", "*.pdf *.docx *.txt"), ("PDF", "*.pdf"),
                       ("Word", "*.docx"), ("Text", "*.txt")])
        if files:
            self.process(list(files))

    def add_folder(self):
        folder = filedialog.askdirectory(title="CVs wala folder select karein")
        if folder:
            self.process(self._collect(folder))

    def _collect(self, folder):
        out = []
        out_root = os.path.abspath(self.cfg["output_folder"])
        for dirpath, _, names in os.walk(folder):
            if os.path.abspath(dirpath).startswith(out_root):
                continue  # sorted folder ko dobara scan na karo
            for n in names:
                if n.lower().endswith(eng.SUPPORTED_EXT) and not n.startswith("~$"):
                    out.append(os.path.join(dirpath, n))
        return out

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        files = []
        for p in paths:
            if os.path.isdir(p):
                files += self._collect(p)
            else:
                files.append(p)
        self.process(files)

    def process(self, files):
        if self.busy:
            messagebox.showinfo("Busy", "Pehle wali CVs process ho rahi hain, thora intezar karein.")
            return
        if not files:
            messagebox.showinfo("No files", "Koi PDF/DOCX/TXT file nahi mili.")
            return
        self.busy = True
        self.progress.configure(maximum=len(files), value=0)
        self._log(f"--- {len(files)} CV(s) process ho rahi hain ---")
        threading.Thread(target=self._worker, args=(files,), daemon=True).start()

    def _worker(self, files):
        summary = {"sorted": 0, "unsorted": 0, "duplicate": 0, "error": 0}
        for i, f in enumerate(files, 1):
            r = eng.process_cv(f, self.cfg)
            summary[r["status"]] += 1
            self.q.put(("result", i, len(files), r))
        self.q.put(("done", summary))

    def _poll_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                if item[0] == "result":
                    _, i, n, r = item
                    self.progress.configure(value=i)
                    self.status.configure(text=f"{i}/{n}  {r['file'][:30]}")
                    if r["status"] == "sorted":
                        self._log(f"✔ {r['file']} → {r['subject']} ({r['confidence']}%)", "ok")
                    elif r["status"] == "unsorted":
                        self._log(f"? {r['file']} → Manual Review. {r['message'] if r['message'] != 'OK' else ''}", "warn")
                    elif r["status"] == "duplicate":
                        self._log(f"⧉ {r['file']} → Duplicate (pehle se save hai)", "warn")
                    else:
                        self._log(f"✖ {r['file']} → Error: {r['message']}", "err")
                elif item[0] == "done":
                    s = item[1]
                    self.busy = False
                    self.status.configure(text="Done ✔")
                    self.refresh_table()
                    msg = (f"Sorted: {s['sorted']}\nManual Review: {s['unsorted']}\n"
                           f"Duplicates: {s['duplicate']}\nErrors: {s['error']}")
                    self._log(msg.replace("\n", " | "))
                    if not self.watch_on.get():
                        messagebox.showinfo("Kaam mukammal", msg)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_queue)

    # -------------------------------------------------------- watch mode
    def toggle_watch(self):
        if self.watch_on.get():
            folder = self.cfg.get("watch_folder")
            if not folder or not os.path.isdir(folder):
                folder = filedialog.askdirectory(title="Kaunsa folder watch karna hai? (jahan naye CVs aate hain)")
                if not folder:
                    self.watch_on.set(False)
                    return
                self.cfg["watch_folder"] = folder
                eng.save_config(self.cfg)
            self.seen_in_watch = set()
            self._log(f"👁 Watching: {folder}  (har 10 second check hoga)")
            self._watch_tick()
        else:
            self._log("👁 Auto-Watch band")

    def _watch_tick(self):
        if not self.watch_on.get():
            return
        folder = self.cfg.get("watch_folder")
        if folder and os.path.isdir(folder) and not self.busy:
            new = [f for f in self._collect(folder) if f not in self.seen_in_watch]
            self.seen_in_watch.update(new)
            if new:
                self.process(new)
        self.root.after(10000, self._watch_tick)

    # -------------------------------------------------------- table
    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        self.paths = {}
        for r in eng.all_records(self.search_var.get().strip() if hasattr(self, "search_var") else ""):
            rid, date, name, subj, conf, email, phone, exp, orig, saved = r
            tag = "unsorted" if subj == eng.UNSORTED else ("low" if (conf or 0) < 50 else "")
            self.tree.insert("", "end", iid=str(rid), tags=(tag,),
                             values=(rid, date, name or "-", subj, f"{conf}%", email or "-",
                                     phone or "-", exp or "-", orig))
            self.paths[str(rid)] = saved
        self._update_stats()

    def _update_stats(self):
        counts = eng.subject_counts()
        total = sum(counts.values())
        self.stats.configure(state="normal")
        self.stats.delete("1.0", "end")
        self.stats.insert("end", f"Total CVs: {total}\n\n")
        for s, n in counts.items():
            self.stats.insert("end", f"{s:<30} {n}\n")
        self.stats.configure(state="disabled")

    def _sort_by(self, col, desc):
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        def key(x):
            v = x[0].rstrip("%")
            return (0, float(v)) if v.replace(".", "", 1).isdigit() else (1, v.lower())
        data.sort(key=key, reverse=desc)
        for i, (_, k) in enumerate(data):
            self.tree.move(k, "", i)
        self.tree.heading(col, command=lambda: self._sort_by(col, not desc))

    def _context_menu(self, e):
        row = self.tree.identify_row(e.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self.menu.tk_popup(e.x_root, e.y_root)

    def _selected(self):
        return list(self.tree.selection())

    def open_selected(self):
        for rid in self._selected()[:5]:
            open_path(self.paths.get(rid))

    def open_selected_folder(self):
        sel = self._selected()
        if sel:
            open_path(os.path.dirname(self.paths.get(sel[0], "")))

    def change_subject(self):
        sel = self._selected()
        if not sel:
            return
        subjects = list(self.cfg["subjects"].keys()) + [eng.UNSORTED]
        dlg = tk.Toplevel(self.root)
        dlg.title("Subject badlein")
        dlg.transient(self.root)
        dlg.grab_set()
        ttk.Label(dlg, text=f"{len(sel)} CV(s) ka sahi subject chunein:", padding=10).pack()
        var = tk.StringVar(value=subjects[0])
        ttk.Combobox(dlg, textvariable=var, values=subjects, width=35, state="readonly").pack(padx=15)

        def ok():
            for rid in sel:
                eng.update_subject(int(rid), var.get(), self.cfg["output_folder"])
            self._log(f"✏ {len(sel)} CV(s) → {var.get()} mein move ki gayi", "ok")
            dlg.destroy()
            self.refresh_table()
        ttk.Button(dlg, text="Save & Move", style="Accent.TButton", command=ok).pack(pady=12)

    def remove_selected(self):
        sel = self._selected()
        if sel and messagebox.askyesno("Remove", f"{len(sel)} record(s) list se hata dein?\n"
                                                 "(File folder mein rahegi, sirf list se hategi)"):
            for rid in sel:
                eng.delete_record(int(rid))
            self.refresh_table()

    def export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile="CV_Report.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if path:
            saved = eng.export_report(path)
            self._log(f"📊 Report saved: {saved}", "ok")
            if messagebox.askyesno("Report ready", f"Report save ho gayi:\n{saved}\n\nAbhi kholein?"):
                open_path(saved)

    def _log(self, msg, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # -------------------------------------------------------- settings
    def open_settings(self):
        SettingsWindow(self)


class SettingsWindow:
    def __init__(self, app):
        self.app = app
        self.cfg = app.cfg
        self.subjects = {k: dict(v) for k, v in self.cfg["subjects"].items()}
        w = self.win = tk.Toplevel(app.root)
        w.title("Settings")
        w.geometry("760x580")
        w.transient(app.root)
        w.grab_set()

        nb = ttk.Notebook(w)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # --- General tab
        g = ttk.Frame(nb, padding=14)
        nb.add(g, text="General")
        self.out_var = tk.StringVar(value=self.cfg["output_folder"])
        self.watch_var = tk.StringVar(value=self.cfg.get("watch_folder", ""))
        self.mode_var = tk.StringVar(value=self.cfg.get("mode", "copy"))
        self.min_var = tk.IntVar(value=self.cfg.get("min_score", 4))
        self.rename_var = tk.BooleanVar(value=self.cfg.get("rename_files", True))

        def row(r, label, var):
            ttk.Label(g, text=label).grid(row=r, column=0, sticky="w", pady=6)
            ttk.Entry(g, textvariable=var, width=55).grid(row=r, column=1, padx=6)
            ttk.Button(g, text="Browse", command=lambda: self._browse(var)).grid(row=r, column=2)
        row(0, "Sorted CVs kahan save hon:", self.out_var)
        row(1, "Auto-Watch folder (naye CVs):", self.watch_var)

        ttk.Label(g, text="Original file ka kya karein:").grid(row=2, column=0, sticky="w", pady=6)
        mf = ttk.Frame(g)
        mf.grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(mf, text="Copy (original wahin rahe)", value="copy",
                        variable=self.mode_var).pack(side="left")
        ttk.Radiobutton(mf, text="Move (original hata do)", value="move",
                        variable=self.mode_var).pack(side="left", padx=10)

        ttk.Label(g, text="Minimum match score:").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Spinbox(g, from_=1, to=30, textvariable=self.min_var, width=6).grid(row=3, column=1, sticky="w")
        ttk.Label(g, text="(is se kam score ho to CV 'Manual Review' folder mein jayegi)",
                  foreground="#777").grid(row=4, column=1, sticky="w")
        ttk.Checkbutton(g, text="File ka naam badal kar 'Candidate Name - Subject' rakho",
                        variable=self.rename_var).grid(row=5, column=1, sticky="w", pady=10)

        # --- Subjects tab
        s = ttk.Frame(nb, padding=10)
        nb.add(s, text="Subjects & Keywords")
        lf = ttk.Frame(s)
        lf.pack(side="left", fill="y")
        ttk.Label(lf, text="Subjects / Folders").pack(anchor="w")
        self.lb = tk.Listbox(lf, width=30, height=20, exportselection=False)
        self.lb.pack(fill="y", expand=True)
        self.lb.bind("<<ListboxSelect>>", self._on_select)
        bf = ttk.Frame(lf)
        bf.pack(fill="x", pady=4)
        ttk.Button(bf, text="+ Add", command=self._add_subject).pack(side="left")
        ttk.Button(bf, text="− Delete", command=self._del_subject).pack(side="left", padx=4)

        rf = ttk.Frame(s, padding=(12, 0))
        rf.pack(side="left", fill="both", expand=True)
        ttk.Label(rf, text="Strong keywords (har line par aik) — jaise 'english teacher', 'ma english'",
                  wraplength=380).pack(anchor="w")
        self.strong_txt = tk.Text(rf, height=9, font=("Consolas", 9))
        self.strong_txt.pack(fill="both", expand=True, pady=(2, 8))
        ttk.Label(rf, text="Weak keywords (madadgar alfaaz) — jaise 'grammar', 'literature'",
                  wraplength=380).pack(anchor="w")
        self.weak_txt = tk.Text(rf, height=9, font=("Consolas", 9))
        self.weak_txt.pack(fill="both", expand=True, pady=2)
        self.current = None
        self._fill_list()

        bb = ttk.Frame(w, padding=8)
        bb.pack(fill="x")
        ttk.Button(bb, text="Reset Subjects to Default", command=self._reset).pack(side="left")
        ttk.Button(bb, text="Cancel", command=w.destroy).pack(side="right")
        ttk.Button(bb, text="💾 Save", style="Accent.TButton", command=self._save).pack(side="right", padx=6)

    def _browse(self, var):
        d = filedialog.askdirectory()
        if d:
            var.set(d)

    def _fill_list(self):
        self.lb.delete(0, "end")
        for k in self.subjects:
            self.lb.insert("end", k)
        if self.subjects:
            self.lb.selection_set(0)
            self._on_select()

    def _store_current(self):
        if self.current and self.current in self.subjects:
            get = lambda t: [x.strip().lower() for x in t.get("1.0", "end").splitlines() if x.strip()]
            self.subjects[self.current] = {"strong": get(self.strong_txt), "weak": get(self.weak_txt)}

    def _on_select(self, _=None):
        sel = self.lb.curselection()
        if not sel:
            return
        self._store_current()
        self.current = self.lb.get(sel[0])
        d = self.subjects[self.current]
        for t, key in ((self.strong_txt, "strong"), (self.weak_txt, "weak")):
            t.delete("1.0", "end")
            t.insert("1.0", "\n".join(d.get(key, [])))

    def _add_subject(self):
        name = simpledialog.askstring("Naya Subject", "Subject / folder ka naam (jaise 'Arabic Teacher'):",
                                      parent=self.win)
        if name and name.strip() and name.strip() not in self.subjects:
            self._store_current()
            n = name.strip()
            first = n.lower().replace(" teacher", "")
            self.subjects[n] = {"strong": [f"{first} teacher", f"ma {first}", f"bs {first}"],
                                "weak": [first]}
            self.current = None
            self._fill_list()
            idx = list(self.subjects).index(n)
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(idx)
            self._on_select()

    def _del_subject(self):
        sel = self.lb.curselection()
        if sel:
            name = self.lb.get(sel[0])
            if messagebox.askyesno("Delete", f"'{name}' delete karein?", parent=self.win):
                self.subjects.pop(name, None)
                self.current = None
                self._fill_list()

    def _reset(self):
        if messagebox.askyesno("Reset", "Saare subjects default par wapis le aayein?", parent=self.win):
            self.subjects = {k: dict(v) for k, v in eng.DEFAULT_SUBJECTS.items()}
            self.current = None
            self._fill_list()

    def _save(self):
        self._store_current()
        self.cfg.update(output_folder=self.out_var.get().strip() or eng.DEFAULT_CONFIG["output_folder"],
                        watch_folder=self.watch_var.get().strip(),
                        mode=self.mode_var.get(), min_score=int(self.min_var.get()),
                        rename_files=bool(self.rename_var.get()), subjects=self.subjects)
        eng.save_config(self.cfg)
        self.app._log("⚙ Settings save ho gayi", "ok")
        self.win.destroy()


def main():
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    CVSorterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
