"""
Lead Checker - Desktop Tool  |  Calder Capital
Checks a leads CSV against a Keap contacts CSV for duplicates.
Results are written back into the original leads file.

Tiered Remove Duplicate Logic (Same-Row Required):
1. YES: Email matches (any row)
2. YES: Name (exact/nickname) + Company match on same row
3. YES: Name (exact/nickname) + Website match on same row
4. YES: Any 3+ of the 5 categories match on same row
5. Check Contact Name: Website + Company match on same row (but name doesn't)

Usage:  python lead_checker.py   (opens the GUI)
Requires:  pip install pandas
"""

import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from collections import defaultdict

import pandas as pd

SUPPORT_EMAIL = "leads@caldergr.com"

# ---------------------------------------------------------------------------
# Nickname Dictionary & Root Name Mapping
# ---------------------------------------------------------------------------
NICKNAMES = {
    'robert': ['rob', 'bob', 'bobby', 'bert'],
    'william': ['bill', 'billy', 'will', 'willy', 'liam'],
    'richard': ['dick', 'rick', 'rich', 'ritchie'],
    'james': ['jim', 'jimmy', 'jamie'],
    'joseph': ['joe', 'joey'],
    'christopher': ['chris', 'topher'],
    'matthew': ['matt', 'matty'],
    'michael': ['mike', 'mikey'],
    'thomas': ['tom', 'tommy'],
    'david': ['dave', 'davy'],
    'anthony': ['tony'],
    'kenneth': ['ken', 'kenny'],
    'steven': ['steve', 'stephan', 'stephen'],
    'andrew': ['drew', 'andy'],
    'gregory': ['greg'],
    'joshua': ['josh'],
    'timothy': ['tim', 'timmy'],
    'ronald': ['ron', 'ronnie'],
    'jeffrey': ['jeff', 'geoff'],
    'ryan': ['ry'],
    'nicholas': ['nick', 'nicky'],
    'jonathan': ['jon', 'john', 'johnny'],
    'charles': ['charlie', 'chuck'],
    'edward': ['ed', 'eddie', 'ted', 'teddy'],
    'elizabeth': ['liz', 'lizzie', 'beth', 'betsy', 'eliza'],
    'katherine': ['kate', 'katie', 'kathy', 'kat'],
    'catherine': ['cat', 'cathy', 'katie'],
    'margaret': ['maggie', 'peggy', 'marge'],
    'susan': ['sue', 'susie'],
    'dorothy': ['dot', 'dottie'],
    'rebecca': ['becca', 'becky'],
    'deborah': ['deb', 'debbie'],
    'patricia': ['pat', 'patty', 'trish', 'tricia'],
    'jennifer': ['jen', 'jenny'],
    'kimberly': ['kim'],
    'alexandra': ['alex', 'ali', 'lexi'],
    'alexander': ['alex', 'xander'],
    'samuel': ['sam', 'sammy'],
    'benjamin': ['ben', 'benny'],
    'daniel': ['dan', 'danny'],
    'phillip': ['phil'],
    'douglas': ['doug'],
    'patrick': ['pat', 'ricky'],
    'raymond': ['ray'],
    'gerald': ['jerry'],
    'lawrence': ['larry'],
    'terrence': ['terry'],
    'bradley': ['brad'],
}

# Map every name to its "Root" (e.g., rob -> robert)
ROOT_MAP = {}
for root, nicks in NICKNAMES.items():
    ROOT_MAP[root] = root
    for n in nicks:
        ROOT_MAP[n] = root

def get_root(name):
    if not name: return ""
    name = name.lower().strip()
    return ROOT_MAP.get(name, name)

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
        if c in df.columns: return c
    return None

def require_col(df: pd.DataFrame, candidates: list, label: str) -> str:
    c = col(df, candidates)
    if c is None:
        raise ValueError(f"Required column not found: {label}\nExpected one of: {candidates}")
    return c

def read_csv_safe(path: str, **kwargs) -> pd.DataFrame:
    for enc in ('utf-8-sig', 'latin-1', 'cp1252'):
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc, **kwargs)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    return pd.read_csv(path, dtype=str, keep_default_na=False, encoding='utf-8', encoding_errors='replace', **kwargs)

def detect_skip_rows(path: str) -> int:
    markers = ['Company Name', 'First Name', 'Executive First Name', 'Last Name', 'Executive Last Name', 'Email', 'Website']
    for enc in ('utf-8-sig', 'latin-1', 'cp1252', 'utf-8'):
        try:
            with open(path, encoding=enc, errors='replace') as f:
                for i, line in enumerate(f):
                    if i > 10: break
                    if any(m in line for m in markers): return i
            break
        except Exception: continue
    return 0

# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def process(leads_path: str, keap_path: str, progress_callback=None) -> str:
    def _p(msg: str):
        if progress_callback: progress_callback(msg)

    # 1. Load files
    _p("Loading Keap contacts…")
    keap_df = read_csv_safe(keap_path)
    _p(f"  {len(keap_df):,} Keap contacts loaded.")

    _p("Loading leads…")
    skip = detect_skip_rows(leads_path)
    leads_df = read_csv_safe(leads_path, skiprows=list(range(skip)) if skip > 0 else None)
    _p(f"  {len(leads_df):,} leads loaded.\n")

    # 2. Validate columns
    l_first = require_col(leads_df, ['Executive First Name', 'First Name'], 'Executive First Name')
    l_last  = require_col(leads_df, ['Executive Last Name',  'Last Name'], 'Executive Last Name')
    l_email = require_col(leads_df, ['Executive Email', 'Email'], 'Executive Email')
    l_web   = require_col(leads_df, ['Website'], 'Website')
    l_comp  = require_col(leads_df, ['Company Name', 'Company'], 'Company Name')

    k_first  = require_col(keap_df, ['First Name', 'Executive First Name'], 'First Name (Keap)')
    k_last   = require_col(keap_df, ['Last Name',  'Executive Last Name'], 'Last Name (Keap)')
    k_email1 = require_col(keap_df, ['Email'], 'Email (Keap)')
    k_email2 = col(keap_df, ['Email Address 2'])
    k_email3 = col(keap_df, ['Email Address 3'])
    k_web    = require_col(keap_df, ['Website'], 'Website (Keap)')
    k_comp   = require_col(keap_df, ['Company Name', 'Company'], 'Company Name (Keap)')

    # 3. Pre-calculate Keap data (Vectorized)
    _p("Building lookup structures…")
    k_roots = (keap_df[k_first].apply(get_root) + " " + keap_df[k_last].str.strip().str.lower()).values
    k_em1_vals   = keap_df[k_email1].str.strip().str.lower().values
    k_em2_vals   = keap_df[k_email2].str.strip().str.lower().values if k_email2 else [""] * len(keap_df)
    k_em3_vals   = keap_df[k_email3].str.strip().str.lower().values if k_email3 else [""] * len(keap_df)
    k_web_vals   = keap_df[k_web].str.strip().apply(clean_web).values
    k_comp_vals  = keap_df[k_comp].str.strip().str.lower().values
    k_trunc_vals = keap_df[k_comp].str.strip().apply(clean_company).values

    # Global sets for fast membership checks
    set_email = set(k_em1_vals) | set(k_em2_vals) | set(k_em3_vals)
    set_email.discard("")
    set_name = set(k_roots)
    set_name.discard("")
    set_web = set(k_web_vals)
    set_web.discard("")
    set_comp = set(k_comp_vals)
    set_comp.discard("")
    set_trunc = set(k_trunc_vals)
    set_trunc.discard("")

    # Signature sets for Same-Row matching
    set_rd_name_comp = set(zip(k_roots, k_comp_vals)) | set(zip(k_roots, k_trunc_vals))
    set_rd_name_web  = set(zip(k_roots, k_web_vals))
    set_rd_web_comp  = set(zip(k_web_vals, k_comp_vals)) | set(zip(k_web_vals, k_trunc_vals))

    # For 3+ matches, use a row-indexed dictionary only for candidates
    rows_by_name = defaultdict(list)
    for i, root in enumerate(k_roots):
        if root: rows_by_name[root].append(i)

    _p("  Lookup structures ready.\n")

    # 4. Process leads (Optimized Loop)
    _p("Processing leads…")
    l_roots = (leads_df[l_first].apply(get_root) + " " + leads_df[l_last].str.strip().str.lower()).values
    l_emails = leads_df[l_email].str.strip().str.lower().values
    l_webs   = leads_df[l_web].str.strip().apply(clean_web).values
    l_comps_raw = leads_df[l_comp].str.strip().values
    
    # Prepare display columns
    l_comp_tr_display = [clean_company(c.split(';')[0]) for c in l_comps_raw]
    
    match_name_list, match_email_list, match_web_list, match_comp_list, match_tr_list = [], [], [], [], []
    remove_dup_list = []

    for i in range(len(leads_df)):
        root = l_roots[i]
        em   = l_emails[i]
        web  = l_webs[i]
        raw_comp = l_comps_raw[i]
        
        comp_parts = [p.strip().lower() for p in raw_comp.split(';')]
        trunc_parts = [clean_company(p) for p in raw_comp.split(';')]
        
        # Individual column matches
        m_name = root in set_name if root else False
        m_em   = em in set_email if em else False
        m_web  = web in set_web if web else False
        m_comp = any(p in set_comp for p in comp_parts if p)
        m_tr   = any(p in set_trunc for p in trunc_parts if p)
        
        match_name_list.append("YES" if m_name else "")
        match_email_list.append("YES" if m_em else "")
        match_web_list.append("YES" if m_web else "")
        match_comp_list.append("YES" if m_comp else "")
        match_tr_list.append("YES" if m_tr else "")

        # Tiered Remove Duplicate
        status = "NO"
        if m_em: 
            status = "YES"
        elif root:
            # Check Name + Company
            if any((root, p) in set_rd_name_comp for p in comp_parts if p): status = "YES"
            # Check Name + Website
            elif web and (root, web) in set_rd_name_web: status = "YES"
            # Check 3+ Matches on same row
            else:
                for idx in rows_by_name[root]:
                    cnt = 1 # Name match
                    if em and em in (k_em1_vals[idx], k_em2_vals[idx], k_em3_vals[idx]): cnt += 1
                    if web and web == k_web_vals[idx]: cnt += 1
                    if any(p == k_comp_vals[idx] for p in comp_parts if p): cnt += 1
                    if any(p == k_trunc_vals[idx] for p in trunc_parts if p): cnt += 1
                    if cnt >= 3:
                        status = "YES"
                        break
        
        # If not YES, check for "Check Contact Name"
        if status == "NO" and web:
            if any((web, p) in set_rd_web_comp for p in comp_parts if p):
                status = "Check Contact Name"
        
        remove_dup_list.append(status)

    # 5. Insert columns at the LEFT
    new_cols = {
        'Remove Duplicate':           remove_dup_list,
        'Match - Truncated Company':  match_tr_list,
        'Match - Company Name':       match_comp_list,
        'Match - Website':            match_web_list,
        'Match - Email':              match_email_list,
        'Match - Full Name':          match_name_list,
        'Truncated Company Name':     l_comp_tr_display
    }
    for col_name, col_data in reversed(list(new_cols.items())):
        if col_name in leads_df.columns: leads_df.drop(columns=[col_name], inplace=True)
        leads_df.insert(0, col_name, col_data)

    _p(f"Saving results back to:\n  {leads_path}\n")
    leads_df.to_csv(leads_path, index=False)
    total = len(leads_df)
    removes = remove_dup_list.count("YES")
    checks = remove_dup_list.count("Check Contact Name")
    summary = f"Complete!\n\n  Leads processed  : {total:,}\n  Remove Duplicate : {removes:,}\n  Check Contact    : {checks:,}\n\nResults saved to original file."
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
