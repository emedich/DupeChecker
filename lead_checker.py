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
# Nickname Dictionary
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

# Reverse mapping for faster lookup
NICKNAME_MAP = {}
for real_name, nicks in NICKNAMES.items():
    NICKNAME_MAP[real_name] = real_name
    for n in nicks:
        NICKNAME_MAP[n] = real_name

def get_root_name(name):
    return NICKNAME_MAP.get(name.lower(), name.lower())

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

def process(leads_path: str, keap_path: str, progress_callback=None) -> str:
    def _p(msg: str):
        if progress_callback:
            progress_callback(msg)

    # 1. Load files
    _p("Loading Keap contacts…")
    keap_df = read_csv_safe(keap_path)
    _p(f"  {len(keap_df):,} Keap contacts loaded.")

    _p("Loading leads…")
    skip = detect_skip_rows(leads_path)
    leads_df = read_csv_safe(leads_path, skiprows=list(range(skip)) if skip > 0 else None)
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

    # 3. Build lookup structures
    _p("Building lookup structures…")
    
    def get_name_variants(first, last):
        f = first.strip().lower()
        l = last.strip().lower()
        if not f or not l: return set()
        root = get_root_name(f)
        variants = {f"{f} {l}"}
        # Add the root name version if different
        if root != f:
            variants.add(f"{root} {l}")
        # Add all other nicknames for that root
        if root in NICKNAMES:
            for nick in NICKNAMES[root]:
                variants.add(f"{nick} {l}")
        return variants

    # We'll pre-process Keap data for matching
    keap_rows = []
    set_email = set()
    set_website = set()
    set_company = set()
    set_trunc = set()
    set_name = set()
    
    rows_by_name = defaultdict(list)
    rows_by_web = defaultdict(list)
    rows_by_comp = defaultdict(list)

    for i in range(len(keap_df)):
        f = keap_df.at[i, k_first]
        l = keap_df.at[i, k_last]
        em1 = keap_df.at[i, k_email1].strip().lower()
        em2 = keap_df.at[i, k_email2].strip().lower() if k_email2 else ""
        em3 = keap_df.at[i, k_email3].strip().lower() if k_email3 else ""
        web = clean_web(keap_df.at[i, k_web])
        comp = keap_df.at[i, k_comp].strip().lower()
        trunc = clean_company(keap_df.at[i, k_comp])
        
        name_variants = get_name_variants(f, l)
        
        row_data = {
            'names': name_variants,
            'emails': {e for e in [em1, em2, em3] if e},
            'web': web,
            'comp': comp,
            'trunc': trunc
        }
        
        keap_rows.append(row_data)
        
        # Add to global sets for individual column matches
        set_name.update(name_variants)
        set_email.update(row_data['emails'])
        if web: set_website.add(web)
        if comp: set_company.add(comp)
        if trunc: set_trunc.add(trunc)
        
        # Add to lookup dicts for same-row matching
        for v in name_variants:
            rows_by_name[v].append(row_data)
        if web:
            rows_by_web[web].append(row_data)
        if comp:
            rows_by_comp[comp].append(row_data)
        if trunc:
            rows_by_comp[trunc].append(row_data)

    _p("  Lookup structures ready.\n")

    # 4. Process leads
    _p("Processing leads…")
    
    l_first_s = leads_df[l_first].str.strip()
    l_last_s  = leads_df[l_last].str.strip()
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

    def is_rd_tiered(idx):
        em = l_email_s.iat[idx]
        if em and em in set_email: return "YES" # Rule 1: Email match (any row)
        
        f = l_first_s.iat[idx]
        l = l_last_s.iat[idx]
        name_variants = get_name_variants(f, l)
        web = l_web_s.iat[idx]
        raw_comp = l_comp_raw.iat[idx]
        comp_parts = [p.strip().lower() for p in raw_comp.split(';')]
        trunc_parts = [clean_company(p) for p in raw_comp.split(';')]
        
        # Potential candidates for same-row matching
        candidates = []
        for v in name_variants:
            candidates.extend(rows_by_name[v])
        if web:
            candidates.extend(rows_by_web[web])
        for p in comp_parts:
            if p: candidates.extend(rows_by_comp[p])
        for p in trunc_parts:
            if p: candidates.extend(rows_by_comp[p])
            
        # Unique candidates by object ID to avoid double checking
        seen_ids = set()
        unique_candidates = []
        for c in candidates:
            if id(c) not in seen_ids:
                seen_ids.add(id(c))
                unique_candidates.append(c)
        
        check_contact = False
        
        for cr in unique_candidates:
            # Match flags for this specific Keap row
            m_name = any(v in cr['names'] for v in name_variants)
            m_em   = em in cr['emails'] if em else False
            m_web  = (web == cr['web']) if web else False
            m_comp = any(p == cr['comp'] for p in comp_parts if p)
            m_tr   = any(p == cr['trunc'] for p in trunc_parts if p)
            
            # Rule 2: Name + Company
            if m_name and m_comp: return "YES"
            # Rule 3: Name + Website
            if m_name and m_web: return "YES"
            # Rule 4: 3+ matches on same row
            if (int(m_name) + int(m_em) + int(m_web) + int(m_comp) + int(m_tr)) >= 3:
                return "YES"
            
            # Rule 5: Website + Company (Check Contact Name)
            if m_web and m_comp:
                check_contact = True
        
        return "Check Contact Name" if check_contact else "NO"

    # Individual column matches (for display columns)
    match_name = pd.Series([any(v in set_name for v in get_name_variants(f, l)) 
                           for f, l in zip(l_first_s, l_last_s)], index=leads_df.index)
    match_email   = l_email_s.isin(set_email) & (l_email_s != '')
    match_website = l_web_s.isin(set_website) & (l_web_s   != '')
    match_company = l_comp_raw.apply(lambda x: check_split_match(x, set_company))
    match_trunc   = l_comp_raw.apply(lambda x: check_split_match(x, set_trunc, clean_company))
    l_comp_tr_display = l_comp_raw.apply(lambda x: clean_company(x.split(';')[0]))

    remove_dup = pd.Series([is_rd_tiered(i) for i in range(len(leads_df))], index=leads_df.index)

    # 6. Insert columns at the LEFT
    yes = 'YES'
    new_cols = {
        'Remove Duplicate':           remove_dup,
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
    total = len(leads_df)
    removes = int((remove_dup == "YES").sum())
    checks = int((remove_dup == "Check Contact Name").sum())
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

