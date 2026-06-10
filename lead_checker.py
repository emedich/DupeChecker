"""
Lead Checker - Desktop Tool  |  Calder Capital
Checks a leads CSV against a Keap contacts CSV for duplicates.
Results are written back into the original leads file.

Tiered Remove Duplicate Logic (Same-Row Required):
1. Email matches (strongest proof)
2. OR Name + Company match on same row
3. OR Name + Website match on same row
4. OR any 3+ of the 5 categories match on same row

Usage:  python lead_checker.py   (opens the GUI)
Requires:  pip install pandas
"""

import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

SUPPORT_EMAIL = "leads@caldergr.com"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_SUFFIX_PATTERN = re.compile(
    r'\b(?:inc|ltd|llc|corp|co|group|solutions|services|holding|holdings|'
    r'company|partners|partner|associates|venture|ventures|llp|plc|sa|kg|'
    r'gmbh|ag|management|consulting|capital|financial|investments|technology|'
    r'systems|media|marketing|digital|software|enterprise)\b',
    re.IGNORECASE,
)
_NON_ALNUM   = re.compile(r'[^a-z0-9\s]')
_MULTI_SPACE = re.compile(r'\s+')


def clean_company(name: str) -> str:
    if not isinstance(name, str) or not name.strip():
        return ''
    s = name.lower()
    s = _SUFFIX_PATTERN.sub('', s)
    s = _NON_ALNUM.sub('', s)
    s = _MULTI_SPACE.sub(' ', s).strip()
    return s


def clean_web(url: str) -> str:
    if not isinstance(url, str) or not url.strip():
        return ''
    s = re.sub(r'^(https?://)?(www\.)?', '', url.strip().lower())
    s = s.split('/')[0]
    return s.strip()


def col(df: pd.DataFrame, candidates: list):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def require_col(df: pd.DataFrame, candidates: list, label: str) -> str:
    c = col(df, candidates)
    if c is None:
        raise ValueError(
            f"Required column not found: {label}\n"
            f"Expected one of: {candidates}\n"
            f"Columns present: {list(df.columns)}"
        )
    return c


def read_csv_safe(path: str, **kwargs) -> pd.DataFrame:
    for enc in ('utf-8-sig', 'latin-1', 'cp1252'):
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False,
                               encoding=enc, **kwargs)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    return pd.read_csv(path, dtype=str, keep_default_na=False,
                       encoding='utf-8', encoding_errors='replace', **kwargs)


_HEADER_MARKERS = [
    'Company Name', 'First Name', 'Executive First Name',
    'Last Name', 'Executive Last Name', 'Email', 'Website',
]


def detect_skip_rows(path: str) -> int:
    for enc in ('utf-8-sig', 'latin-1', 'cp1252', 'utf-8'):
        try:
            with open(path, encoding=enc, errors='replace') as f:
                for i, line in enumerate(f):
                    if i > 10:
                        break
                    if any(m in line for m in _HEADER_MARKERS):
                        return i
            break
        except Exception:
            continue
    return 0


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def process(leads_path: str, keap_path: str,
            progress_callback=None) -> str:

    def _p(msg: str):
        if progress_callback:
            progress_callback(msg)

    # 1. Load files
    _p("Loading Keap contacts…")
    keap_df = read_csv_safe(keap_path)
    _p(f"  {len(keap_df):,} Keap contacts loaded.")

    _p("Loading leads…")
    skip = detect_skip_rows(leads_path)
    leads_df = read_csv_safe(leads_path,
                             skiprows=list(range(skip)) if skip > 0 else None)
    _p(f"  {len(leads_df):,} leads loaded.\n")

    # 2. Validate columns
    _p("Validating columns…")
    l_first = require_col(leads_df, ['Executive First Name', 'First Name'], 'Executive First Name (leads)')
    l_last  = require_col(leads_df, ['Executive Last Name',  'Last Name'], 'Executive Last Name (leads)')
    l_email = require_col(leads_df, ['Executive Email', 'Email'], 'Executive Email (leads)')
    l_web   = require_col(leads_df, ['Website'], 'Website (leads)')
    l_comp  = require_col(leads_df, ['Company Name', 'Company'], 'Company Name (leads)')

    k_first  = require_col(keap_df, ['First Name', 'Executive First Name'], 'First Name (Keap)')
    k_last   = require_col(keap_df, ['Last Name',  'Executive Last Name'], 'Last Name (Keap)')
    k_email1 = require_col(keap_df, ['Email'], 'Email (Keap)')
    k_email2 = col(keap_df, ['Email Address 2'])
    k_email3 = col(keap_df, ['Email Address 3'])
    k_web    = require_col(keap_df, ['Website'], 'Website (Keap)')
    k_comp   = require_col(keap_df, ['Company Name', 'Company'], 'Company Name (Keap)')
    _p("  All required columns found.\n")

    # 3. Build lookup structures (vectorised)
    _p("Building lookup structures…")
    k_full_s  = (keap_df[k_first].str.strip().str.lower() + ' ' +
                 keap_df[k_last].str.strip().str.lower()).str.strip()
    k_em1_s   = keap_df[k_email1].str.strip().str.lower()
    k_em2_s   = keap_df[k_email2].str.strip().str.lower() if k_email2 else pd.Series([''] * len(keap_df), dtype=str)
    k_em3_s   = keap_df[k_email3].str.strip().str.lower() if k_email3 else pd.Series([''] * len(keap_df), dtype=str)
    k_web_s   = keap_df[k_web].str.strip().apply(clean_web)
    k_comp_s  = keap_df[k_comp].str.strip()
    k_comp_lo = k_comp_s.str.lower()
    k_comp_tr = k_comp_s.apply(clean_company)

    # Sets for individual column matches (any row)
    set_name    = set(k_full_s[k_full_s != ''])
    set_email   = set(k_em1_s[k_em1_s != '']) | set(k_em2_s[k_em2_s != '']) | set(k_em3_s[k_em3_s != ''])
    set_website = set(k_web_s[k_web_s != ''])
    set_company = set(k_comp_lo[k_comp_lo != ''])
    set_trunc   = set(k_comp_tr[k_comp_tr != ''])

    # Compound sets for tiered Remove Duplicate (must be same row)
    # We store tuples of (name, email, web, comp, trunc) for each row
    keap_rows = list(zip(k_full_s, k_em1_s, k_em2_s, k_em3_s, k_web_s, k_comp_lo, k_comp_tr))
    
    # For speed, we'll use dictionaries to find rows by name/email
    from collections import defaultdict
    rows_by_name = defaultdict(list)
    for r in keap_rows:
        if r[0]: rows_by_name[r[0]].append(r)
    
    _p("  Lookup structures ready.\n")

    # 4. Process leads
    _p("Processing leads…")
    l_full_s  = (leads_df[l_first].str.strip().str.lower() + ' ' +
                 leads_df[l_last].str.strip().str.lower()).str.strip()
    l_email_s = leads_df[l_email].str.strip().str.lower()
    l_web_s   = leads_df[l_web].str.strip().apply(clean_web)
    l_comp_raw = leads_df[l_comp].str.strip()

    def check_split_match(raw_name, target_set, cleaner=None):
        if not raw_name: return False
        parts = [p.strip() for p in raw_name.split(';')]
        for p in parts:
            val = cleaner(p) if cleaner else p.lower()
            if val and val in target_set: return True
        return False

    match_name    = l_full_s.isin(set_name) & (l_full_s != '')
    match_email   = l_email_s.isin(set_email) & (l_email_s != '')
    match_website = l_web_s.isin(set_website) & (l_web_s != '')
    match_company = l_comp_raw.apply(lambda x: check_split_match(x, set_company))
    match_trunc   = l_comp_raw.apply(lambda x: check_split_match(x, set_trunc, clean_company))
    l_comp_tr_display = l_comp_raw.apply(lambda x: clean_company(x.split(';')[0]))

    # 5. Tiered Remove Duplicate (Same-Row)
    def is_rd(idx):
        em = l_email_s.iat[idx]
        if em and em in set_email: return True # Rule 1: Email match (any row)
        
        fn = l_full_s.iat[idx]
        if not fn: return False
        
        web = l_web_s.iat[idx]
        raw_comp = l_comp_raw.iat[idx]
        comp_parts = [p.strip().lower() for p in raw_comp.split(';')]
        trunc_parts = [clean_company(p) for p in raw_comp.split(';')]
        
        # Check all Keap rows that have this name
        for kr in rows_by_name[fn]:
            # kr = (name, em1, em2, em3, web, comp, trunc)
            k_em1, k_em2, k_em3, k_web, k_comp, k_trunc = kr[1], kr[2], kr[3], kr[4], kr[5], kr[6]
            
            # Rule 2: Name + Company
            if any(p == k_comp for p in comp_parts if p): return True
            # Rule 3: Name + Website
            if web and web == k_web: return True
            
            # Rule 4: 3+ matches on same row
            m_count = 1 # Name already matches
            if em and em in (k_em1, k_em2, k_em3): m_count += 1
            if web and web == k_web: m_count += 1
            if any(p == k_comp for p in comp_parts if p): m_count += 1
            if any(p == k_trunc for p in trunc_parts if p): m_count += 1
            if m_count >= 3: return True
            
        return False

    remove_dup = pd.Series([is_rd(i) for i in range(len(leads_df))], index=leads_df.index)

    # 6. Insert columns at the LEFT
    yes = 'YES'
    new_cols = {
        'Remove Duplicate':           remove_dup.apply(lambda x: yes if x else 'NO'),
        'Match - Truncated Company':  match_trunc.apply(lambda x: yes if x else ''),
        'Match - Company Name':       match_company.apply(lambda x: yes if x else ''),
        'Match - Website':            match_website.apply(lambda x: yes if x else ''),
        'Match - Email':              match_email.apply(lambda x: yes if x else ''),
        'Match - Full Name':          match_name.apply(lambda x: yes if x else ''),
        'Truncated Company Name':     l_comp_tr_display
    }
    for col_name, col_data in reversed(list(new_cols.items())):
        if col_name in leads_df.columns: leads_df.drop(columns=[col_name], inplace=True)
        leads_df.insert(0, col_name, col_data)

    _p(f"Saving results back to:\n  {leads_path}\n")
    leads_df.to_csv(leads_path, index=False)
    total, removes = len(leads_df), int(remove_dup.sum())
    summary = f"Complete!\n\n  Leads processed  : {total:,}\n  Remove Duplicate : {removes:,}\n\nResults saved to original file."
    _p(summary)
    return summary

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

BG, PANEL, ACCENT, TEXT, SUBTEXT, SUCCESS = "#1f2937", "#111827", "#3b82f6", "#f9fafb", "#9ca3af", "#10b981"
FONT_UI, FONT_BOLD, FONT_TITLE, FONT_LOG = ("Segoe UI", 10), ("Segoe UI", 10, "bold"), ("Segoe UI", 16, "bold"), ("Consolas", 9)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Lead Checker  |  Calder Capital")
        self.configure(bg=BG)
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg=ACCENT, pady=14); header.pack(fill='x')
        tk.Label(header, text="Lead Checker", font=FONT_TITLE, bg=ACCENT, fg=TEXT).pack()
        tk.Label(header, text="Calder Capital  •  Duplicate Detection Tool", font=("Segoe UI", 9), bg=ACCENT, fg="#bfdbfe").pack()
        body = tk.Frame(self, bg=BG, padx=24, pady=20); body.pack(fill='both')
        self._leads_var, self._keap_var = tk.StringVar(), tk.StringVar()
        self._file_row(body, "Leads CSV", "(the file to check)", self._leads_var, self._browse_leads)
        self._file_row(body, "Keap Contacts CSV", "(your master CRM export)", self._keap_var, self._browse_keap)
        btn_frame = tk.Frame(self, bg=BG, pady=4); btn_frame.pack(fill='x', padx=24)
        self._run_btn = tk.Button(btn_frame, text="▶   Run Check", font=FONT_BOLD, bg=ACCENT, fg=TEXT, relief='flat', padx=20, pady=8, cursor='hand2', command=self._run)
        self._run_btn.pack(fill='x')
        prog_frame = tk.Frame(self, bg=BG, padx=24, pady=6); prog_frame.pack(fill='x')
        self._progress = ttk.Progressbar(prog_frame, mode='indeterminate', length=600); self._progress.pack(fill='x')
        log_frame = tk.Frame(self, bg=PANEL, padx=24, pady=12); log_frame.pack(fill='both')
        self._log = tk.Text(log_frame, height=12, width=74, state='disabled', font=FONT_LOG, bg=PANEL, fg=TEXT, relief='flat', wrap='word', padx=8, pady=6)
        self._log.pack(side='left', fill='both', expand=True)
        sb = tk.Scrollbar(log_frame, command=self._log.yview); sb.pack(side='right', fill='y'); self._log['yscrollcommand'] = sb.set
        self._log.tag_config('ok', foreground=SUCCESS); self._log.tag_config('err', foreground="#f87171"); self._log.tag_config('muted', foreground=SUBTEXT)
        footer = tk.Frame(self, bg=BG, pady=8); footer.pack(fill='x')
        tk.Label(footer, text=f"Support: {SUPPORT_EMAIL}", font=("Segoe UI", 8), bg=BG, fg=SUBTEXT).pack()

    def _file_row(self, parent, label, sublabel, var, cmd):
        frame = tk.Frame(parent, bg=BG, pady=6); frame.pack(fill='x')
        tk.Label(frame, text=label, font=FONT_BOLD, bg=BG, fg=TEXT).pack(anchor='w')
        tk.Label(frame, text=sublabel, font=("Segoe UI", 8), bg=BG, fg=SUBTEXT).pack(anchor='w')
        row = tk.Frame(frame, bg=BG); row.pack(fill='x', pady=(4, 0))
        tk.Entry(row, textvariable=var, width=58, font=FONT_UI, bg="#374151", fg=TEXT, relief='flat', highlightthickness=1, highlightbackground="#4b5563", highlightcolor=ACCENT).pack(side='left', ipady=5, padx=(0, 8))
        tk.Button(row, text="Browse…", font=FONT_UI, bg="#374151", fg=TEXT, relief='flat', padx=10, pady=4, cursor='hand2', command=cmd).pack(side='left')

    def _browse_leads(self):
        p = filedialog.askopenfilename(title="Select Leads CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p: self._leads_var.set(p)

    def _browse_keap(self):
        p = filedialog.askopenfilename(title="Select Keap Contacts CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p: self._keap_var.set(p)

    def _log_msg(self, msg: str, tag: str = ''):
        self._log.configure(state='normal'); self._log.insert('end', msg + '\n', tag); self._log.see('end'); self._log.configure(state='disabled')

    def _run(self):
        leads, keap = self._leads_var.get().strip(), self._keap_var.get().strip()
        if not leads or not keap:
            messagebox.showwarning("Missing files", "Please select both files.")
            return
        self._run_btn.configure(state='disabled', text="Running…")
        self._log.configure(state='normal'); self._log.delete('1.0', 'end'); self._log.configure(state='disabled')
        self._log_msg("Starting…\n", 'muted'); self._progress.start(12)
        def worker():
            try:
                process(leads, keap, progress_callback=lambda m: self.after(0, self._log_msg, m))
                self.after(0, self._log_msg, "\nAll done!", 'ok')
            except ValueError as ve:
                self.after(0, self._log_msg, f"\nCOLUMN ERROR:\n{ve}", 'err')
                self.after(0, messagebox.showerror, "Column Not Found", str(ve))
            except Exception:
                msg = f"An unexpected error occurred.\n\nPlease contact support:\n{SUPPORT_EMAIL}"
                self.after(0, self._log_msg, f"\n{msg}", 'err')
                self.after(0, messagebox.showerror, "Unexpected Error", msg)
            finally:
                self.after(0, self._progress.stop)
                self.after(0, self._run_btn.configure, {'state': 'normal', 'text': '▶   Run Check'})
        threading.Thread(target=worker, daemon=True).start()

if __name__ == '__main__':
    App().mainloop()
