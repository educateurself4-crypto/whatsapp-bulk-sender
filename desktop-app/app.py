"""
WhatsApp Bulk Sender - desktop app (Tkinter).
Talks ONLY to your backend webhooks. The Meta access token never lives in this app.

Backend connection is configured once via File -> Settings (saved to ~/.wa_bulk_sender/settings.json).
Run:  python app.py
"""
from __future__ import annotations

import csv
import json
import os
import queue
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk
import sv_ttk

from core import (
    N8nClient,
    N8nError,
    Template,
    build_contacts,
    clean_param,
    make_campaign_id,
    normalize_phone,
    parse_template,
    read_table,
    run_campaign,
)

APP_DIR = os.path.join(os.path.expanduser("~"), ".wa_bulk_sender")
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
HISTORY_FILE = os.path.join(APP_DIR, "history.csv")
CHUNK_SIZE = 50


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        sv_ttk.set_theme("light")
        self.title("WhatsApp Bulk Sender")
        self.geometry("900x700")
        self.minsize(800, 600)

        self.q: queue.Queue = queue.Queue()  # worker threads -> UI
        self.cancel = threading.Event()
        self.worker: threading.Thread | None = None

        self.templates: list[Template] = []
        self.headers: list[str] = []
        self.rows: list[list[str]] = []
        self.var_boxes: list[ttk.Combobox] = []
        self.report = None  # last ValidationReport
        self.last_campaign_id = ""
        self.last_status: dict = {}

        # Backend settings (loaded from disk; never shown to the user as required fields)
        self._settings: dict = {}

        self._build_ui()
        self._load_settings()
        self.after(150, self._drain_queue)

        self._settings["url"] = "https://palegoldenrod-elk-207353.hostingersite.com"

        # Auto-load templates if settings are already configured
        if self._settings.get("url"):
            self.after(500, self.on_load_templates)
        else:
            self.log_line("No backend configured yet. Please configure the backend settings file manually.")

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # Menu bar
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)
        self.configure(menu=menubar)

        # 1. template
        f1 = ttk.LabelFrame(self, text="1. Message Template (approved templates only)")
        f1.pack(fill="x", padx=10, pady=(10, 4))
        f1.columnconfigure(0, weight=1)

        top_row = ttk.Frame(f1)
        top_row.grid(row=0, column=0, sticky="ew", padx=8, pady=4)
        top_row.columnconfigure(0, weight=1)
        self.tpl_combo = ttk.Combobox(top_row, state="readonly")
        self.tpl_combo.grid(row=0, column=0, sticky="ew")
        self.btn_refresh_tpl = ttk.Button(top_row, text="↺ Refresh", command=self.on_load_templates, width=10)
        self.btn_refresh_tpl.grid(row=0, column=1, padx=(6, 0))

        self.tpl_combo.bind("<<ComboboxSelected>>", self.on_select_template)
        self.tpl_preview = tk.Text(f1, height=4, wrap="word", state="disabled", background="#f5f5f5")
        self.tpl_preview.grid(row=1, column=0, sticky="ew", padx=8, pady=4)

        # 2. file
        f2 = ttk.LabelFrame(self, text="2. Contacts File (Excel or CSV)")
        f2.pack(fill="x", padx=10, pady=4)
        f2.columnconfigure(1, weight=1)
        ttk.Button(f2, text="Browse...", command=self.on_browse).grid(row=0, column=0, **pad)
        self.file_lbl = ttk.Label(f2, text="No file selected")
        self.file_lbl.grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f2, text="Phone column").grid(row=1, column=0, sticky="w", **pad)
        self.phone_combo = ttk.Combobox(f2, state="readonly")
        self.phone_combo.grid(row=1, column=1, sticky="w", **pad)
        self.phone_combo.bind("<<ComboboxSelected>>", lambda _e: self._invalidate())
        self.map_frame = ttk.Frame(f2)
        self.map_frame.grid(row=2, column=0, columnspan=2, sticky="ew", **pad)
        
        self.tree_frame = ttk.Frame(f2)
        self.tree_frame.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        self.tree = ttk.Treeview(self.tree_frame, height=6, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        # Contacts action row
        contacts_btn_row = ttk.Frame(f2)
        contacts_btn_row.grid(row=4, column=0, columnspan=2, sticky="w", **pad)
        
        self.btn_remove = ttk.Button(contacts_btn_row, text="Remove selected", command=self.on_remove_selected, state="disabled")
        self.btn_remove.pack(side="left", padx=(0, 4))
        
        ttk.Button(contacts_btn_row, text="Validate contacts", command=self.on_validate).pack(side="left", padx=(0, 4))
        self.btn_import_sheets = ttk.Button(
            contacts_btn_row, text="📊 Import to Google Sheets", command=self.on_import_sheets, state="disabled"
        )
        self.btn_import_sheets.pack(side="left", padx=4)
        self.btn_export_invalid = ttk.Button(
            contacts_btn_row, text="Export invalid rows", command=self.on_export_invalid, state="disabled"
        )
        self.btn_export_invalid.pack(side="left", padx=4)

        self.val_lbl = ttk.Label(f2, text="")
        self.val_lbl.grid(row=5, column=0, columnspan=2, sticky="w", **pad)

        # 3. send
        f3 = ttk.LabelFrame(self, text="3. Send")
        f3.pack(fill="both", expand=True, padx=10, pady=4)
        f3.columnconfigure(1, weight=1)
        self.consent_var = tk.BooleanVar()
        ttk.Checkbutton(
            f3,
            variable=self.consent_var,
            text="I confirm every contact in this file has opted in to receive WhatsApp marketing messages from us.",
        ).grid(row=0, column=0, columnspan=4, sticky="w", **pad)

        ttk.Label(f3, text="Max messages this run").grid(row=1, column=0, sticky="w", **pad)
        self.max_var = tk.StringVar(value="250")
        ttk.Entry(f3, textvariable=self.max_var, width=8).grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(
            f3, text="(keep at or below your WhatsApp Manager daily limit; new numbers start around 250/day)"
        ).grid(row=1, column=2, columnspan=2, sticky="w", **pad)

        ttk.Label(f3, text="Test number").grid(row=2, column=0, sticky="w", **pad)
        self.test_var = tk.StringVar()
        ttk.Entry(f3, textvariable=self.test_var, width=18).grid(row=2, column=1, sticky="w", **pad)
        self.btn_test = ttk.Button(f3, text="Send 1 test message", command=self.on_test)
        self.btn_test.grid(row=2, column=2, sticky="w", **pad)

        bar = ttk.Frame(f3)
        bar.grid(row=3, column=0, columnspan=4, sticky="ew", **pad)
        self.btn_start = ttk.Button(bar, text="Start sending", command=self.on_start, state="disabled")
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(bar, text="Stop", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        self.btn_refresh = ttk.Button(bar, text="Refresh delivery status", command=self.on_refresh, state="disabled")
        self.btn_refresh.pack(side="left", padx=4)
        self.btn_export = ttk.Button(bar, text="Export failures", command=self.on_export, state="disabled")
        self.btn_export.pack(side="left", padx=4)

        self.progress = ttk.Progressbar(f3, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=4, sticky="ew", **pad)
        self.stats_lbl = ttk.Label(f3, text="")
        self.stats_lbl.grid(row=5, column=0, columnspan=4, sticky="w", **pad)

        self.log = tk.Text(f3, height=9, wrap="word", state="disabled")
        self.log.grid(row=6, column=0, columnspan=4, sticky="nsew", **pad)
        f3.rowconfigure(6, weight=1)

    # ------------------------------------------------------- small helpers
    def log_line(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def client(self) -> N8nClient:
        url = self._settings.get("url", "").strip()
        # Fallback to a dummy key if none provided, in case the test webhook doesn't enforce auth
        key = self._settings.get("key", "").strip() or "test-key"
        if not url:
            raise N8nError("Backend not configured. Please configure the backend settings file manually.")
        return N8nClient(url, key)

    def cc(self) -> str:
        return self._settings.get("cc", "91").strip().lstrip("+") or "91"

    def run_bg(self, fn, done=None):
        """Run fn() in a thread; deliver ('done'|'error', ...) to the UI thread."""

        def target():
            try:
                res = fn()
                self.q.put(("call", done, res))
            except Exception as e:  # noqa: BLE001 - shown to the user
                self.q.put(("error", str(e)))

        threading.Thread(target=target, daemon=True).start()

    def _drain_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "call":
                    _, fn, res = item
                    if fn:
                        fn(res)
                elif kind == "error":
                    self.log_line("ERROR: " + item[1])
                    self._set_running(False)
                    messagebox.showerror("Error", item[1])
                elif kind == "progress":
                    _, msg, done, total = item
                    self.log_line(msg)
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = done
                elif kind == "finished":
                    self._set_running(False)
                    self.log_line(item[1])
                    self.on_refresh()
        except queue.Empty:
            pass
        self.after(150, self._drain_queue)

    def _set_running(self, running: bool):
        self.btn_start.configure(state="disabled" if running else ("normal" if self.report and self.report.valid else "disabled"))
        self.btn_stop.configure(state="normal" if running else "disabled")
        self.btn_test.configure(state="disabled" if running else "normal")

    # ---------------------------------------------------------- settings
    def _load_settings(self):
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                self._settings = json.load(f)
        except (OSError, ValueError):
            self._settings = {}

    def _save_settings(self, settings: dict):
        self._settings = settings
        os.makedirs(APP_DIR, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f)

    # ------------------------------------------------------- 1: templates
    def on_load_templates(self):
        try:
            c = self.client()
        except N8nError as e:
            messagebox.showerror("Backend not configured", str(e))
            return
        self.log_line("Loading approved templates…")

        def done(raw):
            self.templates = [parse_template(t) for t in raw if (t.get("status") or "").upper() == "APPROVED"]
            self.templates.sort(key=lambda t: (not t.supported, t.name))
            self.tpl_combo["values"] = [t.label for t in self.templates]
            self.log_line(f"Loaded {len(self.templates)} approved template(s).")
            if not self.templates:
                messagebox.showinfo("No templates", "No APPROVED templates were returned for this WhatsApp Business Account.")

        self.run_bg(c.list_templates, done)

    def selected_template(self) -> Template | None:
        i = self.tpl_combo.current()
        return self.templates[i] if 0 <= i < len(self.templates) else None

    def on_select_template(self, _evt=None):
        t = self.selected_template()
        if not t:
            return
        txt = t.body or "(no body)"
        if not t.supported:
            txt += f"\n\nNOT SUPPORTED by this app: {t.reason}"
        self.tpl_preview.configure(state="normal")
        self.tpl_preview.delete("1.0", "end")
        self.tpl_preview.insert("1.0", txt)
        self.tpl_preview.configure(state="disabled")
        self._rebuild_mapping()
        self._invalidate()

    def on_remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            self.tree.delete(item)
        # Rebuild self.rows from the treeview to keep them perfectly in sync
        self.rows = []
        for child in self.tree.get_children():
            # values come back as a list or tuple. 
            # We convert everything to strings to match CSV reading format.
            vals = [str(x) for x in self.tree.item(child)["values"]]
            self.rows.append(vals)
            
        self.file_lbl.configure(text=f"List edited manually  -  {len(self.rows)} rows remaining")
        self._invalidate()

    # ---------------------------------------------------------- 2: contacts
    def on_browse(self):
        path = filedialog.askopenfilename(filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv")])
        if not path:
            return
        try:
            self.headers, self.rows = read_table(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Cannot read file", str(e))
            return
        self.file_lbl.configure(text=f"{os.path.basename(path)}  -  {len(self.rows)} rows")
        self.phone_combo["values"] = self.headers
        guess = next((h for h in self.headers if any(k in h.lower() for k in ("phone", "mobile", "whatsapp", "number", "contact"))), self.headers[0])
        self.phone_combo.set(guess)
        self.tree["columns"] = self.headers
        for h in self.headers:
            self.tree.heading(h, text=h)
            self.tree.column(h, width=110, stretch=True)
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            self.tree.insert("", "end", values=r)
        self._rebuild_mapping()
        self._invalidate()
        # Enable import and remove buttons as soon as a file is loaded
        self.btn_import_sheets.configure(state="normal")
        self.btn_remove.configure(state="normal")
        self.log_line(f"Loaded {len(self.rows)} rows from {os.path.basename(path)}")

    def _rebuild_mapping(self):
        for w in self.map_frame.winfo_children():
            w.destroy()
        self.var_boxes = []
        t = self.selected_template()
        if not t or not self.headers:
            return
        for n in range(1, t.var_count + 1):
            ttk.Label(self.map_frame, text="{{%d}} <-" % n).grid(row=0, column=(n - 1) * 2, padx=(0, 2))
            cb = ttk.Combobox(self.map_frame, state="readonly", values=self.headers, width=18)
            if n < len(self.headers) + 1:
                others = [h for h in self.headers if h != self.phone_combo.get()]
                if n - 1 < len(others):
                    cb.set(others[n - 1])
            cb.grid(row=0, column=(n - 1) * 2 + 1, padx=(0, 10))
            cb.bind("<<ComboboxSelected>>", lambda _e: self._invalidate())
            self.var_boxes.append(cb)

    def _invalidate(self):
        self.report = None
        self.val_lbl.configure(text="")
        self.btn_start.configure(state="disabled")
        self.btn_export_invalid.configure(state="disabled")

    def on_validate(self):
        t = self.selected_template()
        if not t:
            return messagebox.showwarning("Template", "Choose a template first.")
        if not t.supported:
            return messagebox.showwarning("Template", f"This template cannot be used here: {t.reason}")
        if not self.rows:
            return messagebox.showwarning("File", "Choose a contacts file first.")
        if not self.phone_combo.get():
            return messagebox.showwarning("File", "Choose the phone column.")
        if any(not b.get() for b in self.var_boxes):
            return messagebox.showwarning("Mapping", "Map every {{n}} variable to a column.")
        try:
            c = self.client()
        except N8nError as e:
            return messagebox.showerror("Error", str(e))

        phone_idx = self.headers.index(self.phone_combo.get())
        var_idx = [self.headers.index(b.get()) for b in self.var_boxes]
        self.log_line("Checking the opt-out list and validating contacts...")

        def work():
            return c.optouts()

        def done(optouts):
            rep = build_contacts(self.rows, phone_idx, var_idx, self.cc(), optouts)
            self.report = rep
            self.val_lbl.configure(
                text=f"{len(rep.valid)} valid  |  {len(rep.invalid)} invalid  |  "
                f"{rep.duplicates} duplicates  |  {rep.opted_out} opted out (skipped)"
            )
            self.btn_export_invalid.configure(state="normal" if rep.invalid else "disabled")
            self.btn_start.configure(state="normal" if rep.valid else "disabled")
            self.log_line(self.val_lbl.cget("text"))

        self.run_bg(work, done)

    def on_export_invalid(self):
        if not self.report or not self.report.invalid:
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="invalid_rows.csv")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Excel row", "Phone as entered", "Problem"])
            w.writerows(self.report.invalid)
        self.log_line(f"Saved {path}")

    def on_import_sheets(self):
        """Push all loaded rows (as dicts keyed by header) to Google Sheets via the backend."""
        if not self.rows or not self.headers:
            return messagebox.showwarning("File", "Load a contacts file first.")
        try:
            c = self.client()
        except N8nError as e:
            return messagebox.showerror("Error", str(e))

        contacts = [{h: row[i] for i, h in enumerate(self.headers)} for row in self.rows]
        self.log_line(f"Importing {len(contacts)} rows to Google Sheets…")
        self.btn_import_sheets.configure(state="disabled")

        def work():
            return c.import_contacts_to_sheet(contacts, sheet_tab="Contacts")

        def done(count):
            self.log_line(f"✔ {count} rows imported to Google Sheets (Contacts tab).")
            self.btn_import_sheets.configure(state="normal")

        self.run_bg(work, done)

    # ------------------------------------------------------------- 3: send
    def on_test(self):
        t = self.selected_template()
        if not t or not t.supported:
            return messagebox.showwarning("Template", "Choose a supported template first.")
        phone, err = normalize_phone(self.test_var.get(), self.cc())
        if err:
            return messagebox.showwarning("Test number", err)
        if not self.rows or any(not b.get() for b in self.var_boxes):
            return messagebox.showwarning("File", "Load a file and map the variables first (row 1 of data is used for the values).")
        try:
            c = self.client()
        except N8nError as e:
            return messagebox.showerror("Error", str(e))
        var_idx = [self.headers.index(b.get()) for b in self.var_boxes]
        params = [clean_param(self.rows[0][i]) or "-" for i in var_idx]
        cid = "TEST_" + make_campaign_id()
        if not messagebox.askyesno("Send test", f"Send '{t.name}' to {phone} using the values from the first data row?"):
            return
        self.log_line(f"Sending test message to {phone}...")

        def work():
            c.send_chunk(cid, t, [{"phone": phone, "params": params}])
            return cid

        self.run_bg(work, lambda cid: self.log_line(f"Test message sent (id {cid}). It should arrive within seconds."))

    def on_start(self):
        t, rep = self.selected_template(), self.report
        if not (t and rep and rep.valid):
            return
        if not self.consent_var.get():
            return messagebox.showwarning("Opt-in", "Tick the opt-in confirmation first.")
        try:
            limit = int(self.max_var.get())
            if limit < 1:
                raise ValueError
        except ValueError:
            return messagebox.showwarning("Limit", "Max messages must be a positive whole number.")
        try:
            c = self.client()
        except N8nError as e:
            return messagebox.showerror("Error", str(e))

        contacts = rep.valid[:limit]
        skipped = len(rep.valid) - len(contacts)
        msg = f"Send template '{t.name}' to {len(contacts)} contacts?"
        if skipped:
            msg += f"\n\n{skipped} valid contacts are over your 'Max messages' limit and will NOT be sent in this run."
        msg += "\n\nMarketing messages are charged per delivered message by Meta."
        if not messagebox.askyesno("Confirm campaign", msg):
            return

        cid = make_campaign_id()
        self.last_campaign_id = cid
        self.cancel.clear()
        self._set_running(True)
        self.btn_refresh.configure(state="normal")
        self.btn_export.configure(state="normal")
        self.progress["value"] = 0
        self.log_line(f"Campaign {cid} started: {len(contacts)} contacts.")
        os.makedirs(APP_DIR, exist_ok=True)
        with open(HISTORY_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([cid, t.name, t.language, len(contacts), datetime.now().isoformat(timespec="seconds")])

        def work():
            def on_progress(m, done, total):
                self.q.put(("progress", m, done, total))

            n = run_campaign(c, cid, t, contacts, self.cancel, on_progress, chunk_size=CHUNK_SIZE)
            self.q.put(("finished", f"Finished. {n} contacts were submitted."))

        self.worker = threading.Thread(target=self._worker_wrapper, args=(work,), daemon=True)
        self.worker.start()

    def _worker_wrapper(self, fn):
        try:
            fn()
        except Exception as e:  # noqa: BLE001
            self.q.put(("error", str(e)))

    def on_stop(self):
        self.cancel.set()
        self.log_line("Stop requested - no new chunks will be sent.")

    def on_refresh(self):
        if not self.last_campaign_id:
            return
        try:
            c = self.client()
        except N8nError:
            return

        def done(st):
            self.last_status = st
            self.stats_lbl.configure(
                text=f"Accepted {st['accepted']}  |  Delivered {st['delivered']}  |  Read {st['read']}  |  "
                f"Failed at send {st['send_failed']}  |  Failed on delivery {st['delivery_failed']}"
            )

        self.run_bg(lambda: c.campaign_status(self.last_campaign_id), done)

    def on_export(self):
        fails = (self.last_status or {}).get("failures", [])
        if not fails:
            return messagebox.showinfo("Export", "No failures recorded (press 'Refresh delivery status' first).")
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=f"{self.last_campaign_id}_failures.csv")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Phone", "Stage", "Reason"])
            for x in fails:
                w.writerow([x.get("phone", ""), x.get("stage", ""), x.get("reason", "")])
        self.log_line(f"Saved {path}")


if __name__ == "__main__":
    App().mainloop()
