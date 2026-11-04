# Journey_Converter.pyw  —  double-click to run, no terminal needed.
# Converts MoEngage final_campaign_report into Akasa Journey Tracker format (Book 1).

import sys, os, traceback, logging, json

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

LOG_FILE      = os.path.join(BASE_DIR, "journey_converter_errors.log")
MAPPINGS_FILE = os.path.join(BASE_DIR, "journey_mappings.json")

logging.basicConfig(
    filename=LOG_FILE, level=logging.WARNING,
    format="%(asctime)s  %(levelname)s  %(message)s"
)
def log_exc(msg=""):
    logging.error(msg + "\n" + traceback.format_exc())
def log_warn(msg):
    logging.warning(msg)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from datetime import datetime
except ImportError as e:
    import tkinter as tk
    tk.Tk().withdraw()
    tk.messagebox.showerror("Missing package",
        f"A required package is not installed:\n\n{e}\n\nRun: pip install pandas openpyxl")
    sys.exit(1)


# ─── CONSTANTS ────────────────────────────────────────────────────────────────

CHANNEL_ORDER = ["Email", "App Push", "Web Push", "WhatsApp"]
COST_CHANNELS = {"Email", "WhatsApp"}

OUT_COLS = [
    "Journey Name", "Channels",
    "Trip Started", "Trip Engaged",
    "Sent", "Delivered/Impressions", "Delivery/Impression %",
    "Open/Read", "Clicks", "CTR",
    "VT Conv Users", "VT Total Conv users", "VT Revenue",
    "CT Conv Users", "CT Conv Events", "CT Revenue",
    "Avg Order Value", "Cost", "ROI",
]

COL_WIDTHS = {
    "Journey Name": 36, "Channels": 12,
    "Trip Started": 13, "Trip Engaged": 13,
    "Sent": 10, "Delivered/Impressions": 21, "Delivery/Impression %": 21,
    "Open/Read": 10, "Clicks": 8, "CTR": 8,
    "VT Conv Users": 14, "VT Total Conv users": 19, "VT Revenue": 15,
    "CT Conv Users": 14, "CT Conv Events": 15, "CT Revenue": 13,
    "Avg Order Value": 16, "Cost": 10, "ROI": 8,
}

INR_COLS = {"VT Revenue", "CT Revenue", "Avg Order Value", "Cost"}
INR_FMT  = "[$\u20b9-4009]#,##,##0;[$\u20b9-4009]-"

BG        = "#F0F4F8"
BLUE_DARK = "#1F4E79"
BLUE_MID  = "#2E75B6"
GREEN     = "#375623"

HEADER_FILL = PatternFill("solid", start_color="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=9)
BODY_FONT   = Font(name="Arial", size=9)
CENTER      = Alignment(horizontal="center", vertical="center")
LEFT        = Alignment(horizontal="left",   vertical="center")
ODD_FILL    = PatternFill("solid", start_color="DCE6F1")
EVEN_FILL   = PatternFill("solid", start_color="FFFFFF")
THIN        = Border(
    left=Side(style="thin",   color="B8CCE4"),
    right=Side(style="thin",  color="B8CCE4"),
    top=Side(style="thin",    color="B8CCE4"),
    bottom=Side(style="thin", color="B8CCE4"),
)


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def to_num(val):
    try:
        v = float(val)
        return 0.0 if (v != v) else v
    except Exception:
        return 0.0

def load_mappings():
    if os.path.exists(MAPPINGS_FILE):
        try:
            with open(MAPPINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            log_exc("load_mappings failed")
    return {}

def save_mappings_to_file(m):
    with open(MAPPINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(m, f, indent=2, ensure_ascii=False)


# ─── TRANSFORM ────────────────────────────────────────────────────────────────

def transform(df_raw, mappings, email_rate=0.007, wa_rate=0.007):
    df = df_raw.copy()
    df["campaign"] = df["campaign"].astype(str)
    df["Channel"]  = df["Channel"].replace({"Web": "Web Push"})
    df["jname"]    = df["campaign"].map(mappings).fillna(df["campaign"])

    df_ov = df[df["Channel"] == "overAll"]
    df_ch = df[df["Channel"] != "overAll"]

    # aggregate overAll per journey (sums across merged tags)
    overall = {}
    for jn, grp in df_ov.groupby("jname"):
        overall[jn] = {
            "ts":  to_num(grp["Trips Started Total users"].sum()),
            "te":  to_num(grp["Trips engaged"].sum()),
            "vtc": to_num(grp["Goal 1 View Through Total Conversion users (CVR)"].sum()),
            "vtr": to_num(grp["View Through Revenue (Primary Goal)"].sum()),
        }

    def ch_metrics(grp, ch):
        # clicks: uppercase col for push, lowercase for email/whatsapp
        if ch in ("App Push", "Web Push"):
            clicks = to_num(grp["Total Clicks"].sum())
        else:
            clicks = to_num(grp["Total clicks"].sum())

        # delivered / impressions
        if ch == "App Push":
            delivered = to_num(grp["Total Impressions"].sum())
        elif ch in ("Web Push", "WhatsApp"):
            delivered = to_num(grp["Total Delivered"].sum())
        else:
            delivered = 0.0

        # open / read
        if ch == "Email":
            open_read = to_num(grp["Total Open"].sum())
        elif ch == "WhatsApp":
            open_read = to_num(grp["Total Read"].sum())
        elif ch == "App Push":
            open_read = "-"
        else:
            open_read = None

        return {
            "sent":      to_num(grp["Total Sent"].sum()),
            "delivered": delivered,
            "open_read": open_read,
            "clicks":    clicks,
            "vt_cu":     to_num(grp["Goal 1 View Through Converted Users"].sum()),
            "ct_cu":     to_num(grp["Goal 1 Click Through Converted Users"].sum()),
            "ct_ce":     to_num(grp["Goal 1 Click Through Conversion Events"].sum()),
            "ct_tr":     to_num(grp["Goal 1 Click Through Total Revenue"].sum()),
        }

    rows     = []
    warnings = []

    # preserve first-appearance order of journeys
    seen = {}
    for jn in df_ch["jname"]:
        seen.setdefault(jn, True)

    for jn in seen:
        if not any(mappings.get(t) == jn for t in mappings) and jn not in mappings:
            warnings.append(f"Unmapped tag used as journey name: {jn}")
            log_warn(f"Unmapped tag: {jn}")

        jgrp  = df_ch[df_ch["jname"] == jn]
        ov    = overall.get(jn, {"ts": None, "te": None, "vtc": None, "vtr": None})
        first = True

        for ch in CHANNEL_ORDER:
            cg = jgrp[jgrp["Channel"] == ch]
            if cg.empty:
                continue

            m = ch_metrics(cg, ch)

            # skip empty channel rows
            if (m["sent"] == 0 and m["delivered"] == 0 and
                    m["clicks"] == 0 and m["vt_cu"] == 0 and m["ct_cu"] == 0):
                continue

            sent      = int(m["sent"])
            delivered = int(m["delivered"])
            clicks    = int(m["clicks"])
            vt_cu     = int(m["vt_cu"])
            ct_cu     = int(m["ct_cu"])
            ct_ce     = int(m["ct_ce"])
            ct_tr     = int(m["ct_tr"])
            del_pct   = round(delivered / sent, 2) if sent else 0
            ctr       = round(clicks    / sent, 2) if sent else 0
            aov       = round(ct_tr     / ct_ce, 0) if ct_ce else 0

            if isinstance(m["open_read"], float):
                open_read = int(m["open_read"])
            else:
                open_read = m["open_read"]

            rows.append({
                "Journey Name":          jn if first else None,
                "Channels":              ch,
                "Trip Started":          int(ov["ts"]) if first and ov["ts"] else None,
                "Trip Engaged":          int(ov["te"]) if first and ov["te"] else None,
                "Sent":                  sent,
                "Delivered/Impressions": delivered,
                "Delivery/Impression %": del_pct,
                "Open/Read":             open_read,
                "Clicks":                clicks,
                "CTR":                   ctr,
                "VT Conv Users":         vt_cu,
                "VT Total Conv users":   int(ov["vtc"]) if first and ov["vtc"] else None,
                "VT Revenue":            int(ov["vtr"]) if first and ov["vtr"] else None,
                "CT Conv Users":         ct_cu,
                "CT Conv Events":        ct_ce,
                "CT Revenue":            ct_tr,
                "Avg Order Value":       int(aov),
                "_ch":                   ch,
                "_er":                   email_rate,
                "_wr":                   wa_rate,
            })
            first = False

    return rows, warnings


# ─── WRITE ────────────────────────────────────────────────────────────────────

def write_xlsx(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Journey Report"

    for c, col in enumerate(OUT_COLS, 1):
        cell = ws.cell(row=1, column=c, value=col)
        cell.font, cell.fill, cell.alignment, cell.border = (
            HEADER_FONT, HEADER_FILL, CENTER, THIN)
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTHS.get(col, 12)

    sent_let  = get_column_letter(OUT_COLS.index("Sent") + 1)
    ctrev_let = get_column_letter(OUT_COLS.index("CT Revenue") + 1)
    cost_let  = get_column_letter(OUT_COLS.index("Cost") + 1)

    for r, row in enumerate(rows, 2):
        fill = ODD_FILL if r % 2 == 1 else EVEN_FILL
        ch   = row["_ch"]
        rate = row["_er"] if ch == "Email" else row["_wr"]

        for c, col in enumerate(OUT_COLS, 1):
            if col == "Cost":
                val = (f"=ROUND({rate}*{sent_let}{r},0)"
                       if ch in COST_CHANNELS else None)
            elif col == "ROI":
                val = (f"=IFERROR(ROUND({ctrev_let}{r}/{cost_let}{r},0),-1)"
                       if ch in COST_CHANNELS else None)
            else:
                val = row.get(col)

            cell = ws.cell(row=r, column=c, value=val)
            cell.font      = BODY_FONT
            cell.fill      = fill
            cell.border    = THIN
            cell.alignment = LEFT if col in ("Journey Name", "Channels") else CENTER
            if col in INR_COLS and isinstance(val, (int, float)):
                cell.number_format = INR_FMT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(out_path)


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Akasa  \u00b7  Journey Report Converter")
        self.resizable(False, False)
        self.configure(bg=BG)

        self._input_path   = tk.StringVar()
        self._email_rate   = tk.StringVar(value="0.007")
        self._wa_rate      = tk.StringVar(value="0.007")
        self._status_msg   = tk.StringVar(value="Ready.")
        self._mapping_vars = {}
        self._mappings     = load_mappings()
        self._unique_tags  = []

        self._build_ui()

    def _build_ui(self):
        # title bar
        hdr = tk.Frame(self, bg=BLUE_DARK)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Akasa  \u00b7  Journey Report Converter",
                 font=("Arial", 12, "bold"), bg=BLUE_DARK, fg="white",
                 anchor="w", padx=12, pady=8).pack(fill="x")
        tk.Label(hdr,
                 text="Converts MoEngage final_campaign_report into Journey Tracker\n"
                      "format (Book 1). Same journey name in both rows = auto-merged.",
                 font=("Arial", 9), bg=BLUE_DARK, fg="#B8D4F0",
                 anchor="w", padx=14, justify="left").pack(fill="x", pady=(0, 10))

        # step 1: file
        frm1 = tk.LabelFrame(self,
            text=" 1.  Select final_campaign_report.xlsx ",
            bg=BG, font=("Arial", 9, "bold"), fg=BLUE_DARK)
        frm1.pack(fill="x", padx=12, pady=(10, 4))

        r1 = tk.Frame(frm1, bg=BG)
        r1.pack(fill="x", padx=8, pady=8)
        tk.Entry(r1, textvariable=self._input_path, width=50,
                 font=("Arial", 9), relief="solid", bd=1
                 ).pack(side="left", padx=(0, 6))
        tk.Button(r1, text="Browse\u2026", command=self._browse,
                  bg=BLUE_MID, fg="white", font=("Arial", 9, "bold"),
                  relief="flat", padx=10).pack(side="left")

        self._file_lbl = tk.Label(frm1, text="",
                                   font=("Arial", 8), bg=BG, anchor="w", padx=10)
        self._file_lbl.pack(fill="x", pady=(0, 4))

        # step 2: mapping table
        frm2 = tk.LabelFrame(self,
            text=" 2.  Journey Name Mapping  "
                 "(same name = merged & summed) ",
            bg=BG, font=("Arial", 9, "bold"), fg=BLUE_DARK)
        frm2.pack(fill="x", padx=12, pady=4)

        hbar = tk.Frame(frm2, bg="#D6E4F0")
        hbar.pack(fill="x", padx=8, pady=(6, 0))
        tk.Label(hbar, text="MoEngage Tag  (from export)", width=36,
                 bg="#D6E4F0", font=("Arial", 8, "bold"), anchor="w"
                 ).pack(side="left", padx=4, pady=3)
        tk.Label(hbar, text="Journey Name  (in tracker)",
                 bg="#D6E4F0", font=("Arial", 8, "bold"), anchor="w"
                 ).pack(side="left")

        ctr = tk.Frame(frm2, bg=BG)
        ctr.pack(fill="x", padx=8, pady=4)

        self._canvas = tk.Canvas(ctr, bg=BG, height=150, highlightthickness=0)
        sb = tk.Scrollbar(ctr, orient="vertical", command=self._canvas.yview)
        self._map_frame = tk.Frame(self._canvas, bg=BG)
        self._map_frame.bind(
            "<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self._map_frame, anchor="nw")
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")

        tk.Label(self._map_frame,
                 text="Load a file first to see its tags here.",
                 bg=BG, font=("Arial", 8), fg="#888", padx=6, pady=8
                 ).pack(anchor="w")

        save_row = tk.Frame(frm2, bg=BG)
        save_row.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(save_row, text="Save Mappings",
                  command=self._save_mappings,
                  bg=GREEN, fg="white", font=("Arial", 8, "bold"),
                  relief="flat", padx=8, pady=3).pack(side="left")
        self._save_lbl = tk.Label(save_row, text="",
                                   bg=BG, font=("Arial", 8), fg=GREEN)
        self._save_lbl.pack(side="left", padx=8)

        # step 3: cost rates
        frm3 = tk.LabelFrame(self,
            text=" 3.  Cost Rate per Message Sent ",
            bg=BG, font=("Arial", 9, "bold"), fg=BLUE_DARK)
        frm3.pack(fill="x", padx=12, pady=4)

        rr = tk.Frame(frm3, bg=BG)
        rr.pack(fill="x", padx=10, pady=8)
        tk.Label(rr, text="Email:", bg=BG, font=("Arial", 9)).pack(side="left")
        tk.Entry(rr, textvariable=self._email_rate, width=8,
                 font=("Arial", 9), relief="solid", bd=1
                 ).pack(side="left", padx=(4, 20))
        tk.Label(rr, text="WhatsApp:", bg=BG, font=("Arial", 9)).pack(side="left")
        tk.Entry(rr, textvariable=self._wa_rate, width=8,
                 font=("Arial", 9), relief="solid", bd=1
                 ).pack(side="left", padx=(4, 0))
        tk.Label(rr, text="   Cost = rate \u00d7 Sent  (written as live Excel formula)",
                 bg=BG, font=("Arial", 8), fg="#888").pack(side="left", padx=8)

        # convert button
        tk.Button(self, text="\u25b6  Convert & Save As\u2026",
                  command=self._run_convert,
                  bg=BLUE_DARK, fg="white", font=("Arial", 10, "bold"),
                  relief="flat", padx=24, pady=8).pack(pady=(8, 10))

        # status bar
        tk.Label(self, textvariable=self._status_msg,
                 bg="#D6E4F0", fg=BLUE_DARK, font=("Arial", 9),
                 anchor="w", padx=10, pady=4).pack(fill="x", side="bottom")

    def _browse(self):
        p = filedialog.askopenfilename(
            title="Select final_campaign_report",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if not p:
            return
        self._input_path.set(p)
        try:
            df   = pd.read_excel(p, engine="openpyxl")
            tags = sorted(df["campaign"].astype(str).unique().tolist())
            self._unique_tags = tags
            self._file_lbl.config(
                text=f"\u2714  {len(df)} rows  |  {len(tags)} unique tags found.",
                fg=GREEN)
            self._populate_mapping_table(tags)
            self._status_msg.set(f"Loaded: {os.path.basename(p)}")
        except Exception as e:
            log_exc("browse failed")
            self._file_lbl.config(text=f"\u26a0  Could not read file: {e}",
                                   fg="#C0392B")

    def _populate_mapping_table(self, tags):
        for w in self._map_frame.winfo_children():
            w.destroy()
        self._mapping_vars = {}

        for i, tag in enumerate(tags):
            row_bg = "#FFFFFF" if i % 2 == 0 else "#EEF4FB"
            rf = tk.Frame(self._map_frame, bg=row_bg)
            rf.pack(fill="x")
            display = tag if len(tag) <= 40 else tag[:37] + "\u2026"
            tk.Label(rf, text=display, width=36,
                     bg=row_bg, font=("Arial", 8), anchor="w"
                     ).pack(side="left", padx=4, pady=3)
            var = tk.StringVar(value=self._mappings.get(tag, ""))
            tk.Entry(rf, textvariable=var, width=28,
                     font=("Arial", 8), relief="solid", bd=1
                     ).pack(side="left", padx=(2, 4), pady=3)
            self._mapping_vars[tag] = var

        self._canvas.update_idletasks()
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _save_mappings(self):
        new_m = {tag: var.get().strip()
                 for tag, var in self._mapping_vars.items()
                 if var.get().strip()}
        try:
            save_mappings_to_file(new_m)
            self._mappings = new_m
            self._save_lbl.config(text=f"\u2714  {len(new_m)} mapping(s) saved.")
        except Exception as e:
            log_exc("save_mappings failed")
            messagebox.showerror("Save error", str(e))

    def _run_convert(self):
        path = self._input_path.get().strip()
        if not path:
            messagebox.showerror("No file", "Browse and select a file first.")
            return
        try:
            er = float(self._email_rate.get())
            wr = float(self._wa_rate.get())
        except ValueError:
            messagebox.showerror("Invalid rate",
                "Cost rates must be decimal numbers e.g. 0.007")
            return

        live = {tag: var.get().strip()
                for tag, var in self._mapping_vars.items()
                if var.get().strip()}
        if not live:
            live = self._mappings

        try:
            df = pd.read_excel(path, engine="openpyxl")
        except Exception as e:
            log_exc("read failed")
            messagebox.showerror("Read error",
                f"{e}\n\nMake sure the file is not open in Excel.")
            return

        try:
            rows, warnings = transform(df, live, er, wr)
        except Exception as e:
            log_exc("transform failed")
            messagebox.showerror("Conversion error", str(e))
            return

        if not rows:
            messagebox.showwarning("No data",
                "No channel rows produced.\n"
                "Check that the file is a MoEngage journey report.")
            return

        out_path = filedialog.asksaveasfilename(
            title="Save Journey Report",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"Journey_Report_{datetime.today().strftime('%d%b%Y')}.xlsx",
            initialdir=BASE_DIR,
        )
        if not out_path:
            return

        try:
            write_xlsx(rows, out_path)
        except Exception as e:
            log_exc("write failed")
            messagebox.showerror("Write error",
                f"{e}\n\nMake sure the output file is not open in Excel.")
            return

        unmapped = [t for t in self._unique_tags if t not in live]
        status   = f"\u2714  {len(rows)} channel rows written."
        if unmapped:
            status += f"  \u26a0 {len(unmapped)} unmapped tag(s) — raw tag used."
        self._status_msg.set(status)

        detail = f"{len(rows)} rows saved to:\n{out_path}"
        if warnings:
            detail += (f"\n\n\u26a0  {len(warnings)} tag(s) had no mapping — "
                       f"raw tag used as journey name.\nSee: {LOG_FILE}")
        messagebox.showinfo("Done", detail)


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

try:
    app = App()
    app.mainloop()
except Exception:
    log_exc("Unhandled startup exception")
    try:
        import tkinter as tk
        from tkinter import messagebox
        tk.Tk().withdraw()
        messagebox.showerror("Startup error",
            f"Converter failed to start.\nSee: {LOG_FILE}")
    except Exception:
        pass
