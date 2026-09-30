#!/usr/bin/env python3
"""Simplify Finance Dashboard Generator — FY 2026-27 Full Redesign"""

import io, os, json, hashlib, base64, calendar, tempfile, re
import pandas as pd, requests
from datetime import datetime, date
from zoneinfo import ZoneInfo
SYDNEY = ZoneInfo('Australia/Sydney')
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────────
SCRIPT_DIR       = Path(__file__).parent
GITHUB_CFG       = SCRIPT_DIR / 'github_config.json'
GITHUB_PAGES_CFG = SCRIPT_DIR / 'github_pages_config.json'
HTML_HASH_FILE   = SCRIPT_DIR / 'last_html_hash.txt'

EXCEL_CANDIDATES = [
    Path('/Users/alanhemmings/Library/CloudStorage/OneDrive-SimplifyFinance/Communication site - Operations/Dashboards/Business Data.xlsx'),
    Path('/Users/alanhemmings/Library/CloudStorage/OneDrive-SimplifyFinanceGroupPtyLtd/Documents/Simplify Finance/Dashboards/Business Data.xlsx'),
    SCRIPT_DIR / 'Business Data.xlsx',
]

MONTHLY_TARGETS = {
    'Jul': 22_900_000, 'Aug': 27_800_000, 'Sep': 32_100_000,
    'Oct': 30_900_000, 'Nov': 28_800_000, 'Dec': 31_400_000,
    'Jan': 20_600_000, 'Feb': 19_700_000, 'Mar': 30_200_000,
    'Apr': 24_200_000, 'May': 28_700_000, 'Jun': 32_900_000,
}
ANNUAL_TARGET = 330_000_000
FY_MONTHS = ['Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun']
# Variants used in CreditTeam sheet
MONTH_ALIASES = {
    'jul':'Jul','july':'Jul','aug':'Aug','august':'Aug',
    'sep':'Sep','sept':'Sep','september':'Sep',
    'oct':'Oct','october':'Oct','nov':'Nov','november':'Nov',
    'dec':'Dec','december':'Dec','jan':'Jan','january':'Jan',
    'feb':'Feb','february':'Feb','mar':'Mar','march':'Mar',
    'apr':'Apr','april':'Apr','may':'May','jun':'Jun','june':'Jun',
}

def find_excel():
    for p in EXCEL_CANDIDATES:
        if p.exists():
            return p
    raise FileNotFoundError('Cannot find Business Data.xlsx')

def is_cloud():
    """True when running in GitHub Actions with Azure credentials available."""
    return bool(os.environ.get('AZURE_CLIENT_SECRET'))

def get_graph_token():
    """Obtain a Microsoft Graph API token via client credentials flow."""
    r = requests.post(
        f'https://login.microsoftonline.com/{os.environ["AZURE_TENANT_ID"]}/oauth2/v2.0/token',
        data={
            'grant_type':    'client_credentials',
            'client_id':     os.environ['AZURE_CLIENT_ID'],
            'client_secret': os.environ['AZURE_CLIENT_SECRET'],
            'scope':         'https://graph.microsoft.com/.default',
        }, timeout=30)
    r.raise_for_status()
    return r.json()['access_token']

def download_excel_cloud():
    """Download Business Data.xlsx from SharePoint via Graph API, return local temp path.
    Uses file metadata to get a fresh direct download URL (bypasses SharePoint caching).
    Logs lastModifiedDateTime so we can confirm the file version is current."""
    token = get_graph_token()
    h = {
        'Authorization': f'Bearer {token}',
        'Cache-Control': 'no-cache, no-store',
        'Pragma': 'no-cache',
    }
    base = 'https://graph.microsoft.com/v1.0/sites/simplifyfin.sharepoint.com'

    REQUIRED_SHEETS = {'Year On Year Stats', 'Leave', 'CreditTeam'}

    def _try_download(url, use_auth=True):
        """Download from url, return (path, sheet_names) or (None, None)."""
        try:
            if use_auth:
                headers = h
            else:
                # Pre-signed URL — no auth header, but still ask for no-cache
                headers = {'Cache-Control': 'no-cache, no-store', 'Pragma': 'no-cache'}
            r = requests.get(url, headers=headers, timeout=60)
            if r.status_code != 200:
                return None, None
            tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
            tmp.write(r.content); tmp.close()
            p = Path(tmp.name)
            try:
                xl = pd.ExcelFile(p)
                return p, xl.sheet_names
            except Exception:
                p.unlink(missing_ok=True)
                return None, None
        except Exception:
            return None, None

    # Candidate metadata paths — fetch metadata first to get a fresh download URL
    # and to confirm the file's last-modified time (detects stale cached copies)
    meta_paths = [
        f'{base}/drive/root:/Operations/Dashboards/Business Data.xlsx',
        f'{base}/lists/Operations/drive/root:/Dashboards/Business Data.xlsx',
        f'{base}/lists/Documents/drive/root:/Operations/Dashboards/Business Data.xlsx',
    ]

    best_path, best_sheets, best_modified = None, [], 'unknown'

    for meta_url in meta_paths:
        try:
            r = requests.get(meta_url, headers=h, timeout=30)
            if r.status_code != 200:
                continue
            item = r.json()
            modified = item.get('lastModifiedDateTime', 'unknown')
            short = meta_url.replace(base, '')
            print(f'  Found file at: {short}')
            print(f'  Last modified on SharePoint: {modified}')
            # Use /content via Graph API with Bearer token — bypasses SharePoint CDN cache.
            # @microsoft.graph.downloadUrl points to Azure Blob Storage which caches
            # aggressively and ignores Cache-Control headers. The /content endpoint
            # routes through Microsoft's API servers and respects our no-cache headers.
            item_id  = item.get('id')
            drive_id = item.get('parentReference', {}).get('driveId')
            if item_id and drive_id:
                content_url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content'
                p, sheets = _try_download(content_url, use_auth=True)
            else:
                # Fallback to pre-signed URL if we can't build the content URL
                dl_url = item.get('@microsoft.graph.downloadUrl')
                if dl_url:
                    p, sheets = _try_download(dl_url, use_auth=False)
                else:
                    p, sheets = None, None
            if p is None:
                continue
            print(f'  Sheets ({len(sheets)}): {sheets}')
            if REQUIRED_SHEETS.issubset(set(sheets)):
                if best_path and best_path != p:
                    best_path.unlink(missing_ok=True)
                best_path, best_sheets, best_modified = p, sheets, modified
                break
            else:
                missing = sorted(REQUIRED_SHEETS - set(sheets))
                print(f'  ✗ Missing sheets: {missing} — trying next path')
                if best_path is None:
                    best_path, best_sheets = p, sheets
                else:
                    p.unlink(missing_ok=True)
        except Exception as e:
            print(f'  Error checking {meta_url.replace(base, "")}: {e}')
            continue

    if best_path:
        if not REQUIRED_SHEETS.issubset(set(best_sheets)):
            print(f'  ⚠ Warning: best file missing sheets {sorted(REQUIRED_SHEETS - set(best_sheets))}')
        return best_path, best_modified

    # Last resort: search across SharePoint
    print('  Searching SharePoint for Business Data.xlsx ...')
    r = requests.get(
        f"{base}/drive/root/search(q='Business Data')",
        headers=h, timeout=30)
    if r.status_code == 200:
        for item in r.json().get('value', []):
            name = item.get('name', '')
            if 'Business Data' in name and name.endswith('.xlsx'):
                modified = item.get('lastModifiedDateTime', 'unknown')
                pref = item.get('parentReference', {}).get('path', '')
                print(f'    Found: {name} at {pref} (modified: {modified})')
                item_id2  = item.get('id')
                drive_id2 = item.get('parentReference', {}).get('driveId')
                if item_id2 and drive_id2:
                    content_url2 = f'https://graph.microsoft.com/v1.0/drives/{drive_id2}/items/{item_id2}/content'
                    p, sheets = _try_download(content_url2, use_auth=True)
                else:
                    dl_url = item.get('@microsoft.graph.downloadUrl')
                    p, sheets = (_try_download(dl_url, use_auth=False) if dl_url else (None, None))
                if p:
                    print(f'    Sheets: {sheets}')
                    return p, modified

    raise FileNotFoundError('Cannot find Business Data.xlsx on SharePoint — check the file path and permissions.')

def safe_num(v):
    try:
        f = float(v)
        return f if (f == f) and f > 0 else 0.0
    except (TypeError, ValueError):
        return 0.0

def norm_month(s):
    return MONTH_ALIASES.get(str(s).strip().lower(), '')

# ── Current year data ──────────────────────────────────────────────────────────
def read_current_year(df):
    latest_row, latest_num = None, 0
    for i, row in df.iterrows():
        cell = str(row.iloc[0])
        if cell.startswith('Year ') and 'Book' not in cell and 'TOTAL' not in cell:
            parts = cell.split()
            if len(parts) >= 2:
                try:
                    n = int(parts[1])
                    if n > latest_num:
                        latest_num, latest_row = n, i
                except ValueError:
                    pass
    if latest_row is None:
        raise ValueError('No Year section found')
    print(f'  Current year: {df.iloc[latest_row, 0]} (row {latest_row+1})')
    month_data = {}
    for offset in range(2, 15):
        idx = latest_row + offset
        if idx >= len(df): break
        row = df.iloc[idx]
        name = norm_month(row.iloc[0])
        if not name: continue
        month_data[name] = {
            'lodgements':    safe_num(row.iloc[2]),
            'settlements':   safe_num(row.iloc[6]),
            'deals_lodged':  int(safe_num(row.iloc[1])),
            'deals_settled': int(safe_num(row.iloc[4])),
        }
    return month_data

# ── All-time total + 5-year history ───────────────────────────────────────────
def read_history(df):
    """Returns (all_time_total, list of last 5 year dicts for chart)"""
    year_sections = []  # list of (year_num, label, start_row, is_book2)

    for i, row in df.iterrows():
        cell = str(row.iloc[0]).strip()
        if cell.startswith('Year ') and 'TOTAL' not in cell:
            parts = cell.split()
            if len(parts) >= 2:
                try:
                    n = int(parts[1])
                    is_b2 = 'Book' in cell
                    # Extract FY label e.g. "2025 - 2026" → "FY26"
                    years_part = [p for p in parts if p.isdigit() and len(p)==4]
                    label = f'FY{years_part[1][2:]}' if len(years_part) >= 2 else f'FY?{n}'
                    year_sections.append({'num': n, 'label': label, 'row': i, 'book2': is_b2})
                except (ValueError, IndexError):
                    pass

    # Compute all-time settlements
    all_time = 0.0
    for sec in year_sections:
        # Find Total row
        for offset in range(12, 20):
            idx = sec['row'] + offset
            if idx >= len(df): break
            if str(df.iloc[idx, 0]).strip() == 'Total':
                s = safe_num(df.iloc[idx, 6])
                all_time += s
                break

    # Build 5-year chart data (last 5 distinct year nums, combine book variants)
    # Group by year num
    from collections import defaultdict
    year_monthly = defaultdict(lambda: [0.0]*12)
    year_labels  = {}
    for sec in year_sections:
        n = sec['num']
        year_labels[n] = sec['label']
        for offset in range(2, 14):
            idx = sec['row'] + offset
            if idx >= len(df): break
            row = df.iloc[idx]
            m = norm_month(row.iloc[0])
            if m in FY_MONTHS:
                mi = FY_MONTHS.index(m)
                year_monthly[n][mi] += safe_num(row.iloc[6])

    distinct_years = sorted(year_monthly.keys())
    last5 = distinct_years[-5:] if len(distinct_years) >= 5 else distinct_years
    history = [{'label': year_labels[n], 'months': year_monthly[n]} for n in last5]

    return all_time, history

# ── CreditTeam BCs and LOs ─────────────────────────────────────────────────────
def read_credit_team(path, cur_short):
    try:
        df = pd.read_excel(path, sheet_name='CreditTeam', header=None)
    except Exception:
        return 0, 0, 0
    total_bc = total_lo = total_cp = 0
    in_month = False
    for i in range(1, len(df)):  # skip row 0 (title row)
        row = df.iloc[i]
        cell = str(row.iloc[0]).strip()
        # Skip header and separator rows only
        if cell.lower() in ('month', 'total', ''):
            continue
        # Update month tracking (only for rows that have a valid month name)
        nm = norm_month(cell)
        if nm:
            in_month = (nm == cur_short)
        # Count BC/LO/CP for all rows in the current month section
        if in_month:
            total_bc += int(safe_num(row.iloc[2]))
            total_lo += int(safe_num(row.iloc[3]))
            total_cp += int(safe_num(row.iloc[4]) if len(row) > 4 else 0)
    return total_bc, total_lo, total_cp

# ── Leave data ─────────────────────────────────────────────────────────────────
def read_leave(path, cur_month_full):
    """Read Leave sheet. Structure: row 0 = Names | <date>, rows 1+ = Name | Dates.
       Only returns staff who have actual leave dates entered."""
    try:
        df = pd.read_excel(path, sheet_name='Leave', header=None)
    except Exception:
        return [], cur_month_full + ' Leave'
    if len(df) == 0:
        return [], cur_month_full + ' Leave'

    # Row 0: col 0 = label ("Names"), col 1 = date object → derive month title
    title = cur_month_full + ' Leave'
    try:
        date_cell = df.iloc[0, 1]
        if hasattr(date_cell, 'strftime'):
            title = date_cell.strftime('%B') + ' Leave'
        elif str(date_cell) not in ('nan', ''):
            title = pd.to_datetime(date_cell).strftime('%B') + ' Leave'
    except Exception:
        pass

    # Rows 1+: Name | Dates — only include staff with actual dates
    # Also filter out leave that has already ended (date in the past)
    today = datetime.now(SYDNEY).date()

    def latest_day_in_string(s):
        """Extract the highest day number from a date string like '28th', '3-5 Oct', etc."""
        # First try ordinal suffixes: 28th, 3rd, 1st, 2nd
        ordinals = re.findall(r'(\d{1,2})(?:st|nd|rd|th)', str(s))
        if ordinals:
            valid = [int(n) for n in ordinals if 1 <= int(n) <= 31]
            if valid:
                return max(valid)
        # Fall back: any isolated 1-2 digit number (not part of a longer number)
        nums = re.findall(r'(?<!\d)(\d{1,2})(?!\d)', str(s))
        valid = [int(n) for n in nums if 1 <= int(n) <= 31]
        return max(valid) if valid else None

    entries = []
    for i in range(1, len(df)):
        row = df.iloc[i]
        name = str(row.iloc[0]).strip()
        if not name or name.lower() == 'nan':
            continue
        dates = ''
        if len(row) > 1:
            d = str(row.iloc[1]).strip()
            if d and d.lower() != 'nan':
                dates = d
        if dates:  # only show staff who have leave entered
            # Check if leave has already ended — if so, skip it
            last_day = latest_day_in_string(dates)
            if last_day is not None:
                try:
                    leave_end = date(today.year, today.month, last_day)
                    if leave_end < today:
                        print(f'    ↳ Skipping {name} ({dates}) — leave ended {leave_end} (today={today})')
                        continue  # Leave already ended — exclude from display
                except ValueError:
                    pass  # Invalid date (e.g. day 31 in a short month) — include to be safe
            entries.append({'name': name, 'dates': dates})
    return entries, title

def read_lender_mix(path):
    """Read 'Lender Mix' sheet. Two datasets sit side by side:
      - FY26 Settlements: col A = Lender, col C = $ Settled  (rows 3+)
      - Loan Book 30/6/26: col E = Lender, col G = $ amount  (rows 3+)
    Normalises name capitalisation, merges 'Commercial' variants into their
    parent bank, and groups lenders below 3% of total into 'Other'.
    Returns (fy26_list, book_list) — each a list of {name, amount} dicts."""
    try:
        df = pd.read_excel(path, sheet_name='Lender Mix', header=None)
    except Exception:
        return [], []

    # Commercial variants → parent bank name
    COMMERCIAL = {
        'anz commercial':       'ANZ',
        'cba commercial':       'CBA',
        'nab commercial':       'NAB',
        'st george commercial': 'St George',
        'la trobe commercial':  'La Trobe',
        'liberty commercial':   'Liberty',
    }
    # Capitalisation fixes (handles mangled entries like 'BankweSt', 'FirStmac')
    NAME_FIX = {
        'amp': 'AMP',               'anz': 'ANZ',
        'auswide ltd': 'Auswide',   'bank australia': 'Bank Australia',
        'bankwest': 'Bankwest',     'bendigo': 'Bendigo',
        'bluestone': 'Bluestone',   'bom': 'BOM',
        'boq': 'BOQ',               'citibank': 'Citibank',
        'cba': 'CBA',               'firstmac': 'Firstmac',
        'ing': 'ING',               'la trobe': 'La Trobe',
        'liberty': 'Liberty',       'ma money': 'MA Money',
        'macquarie': 'Macquarie',   'me bank': 'ME Bank',
        'mortgage ezy': 'Mortgage Ezy', 'move': 'MOVE',
        'nab': 'NAB',               'orde': 'Orde',
        'people first': 'People First', 'pepper': 'Pepper',
        'resimac': 'Resimac',       'st george': 'St George',
        'suncorp': 'Suncorp',       'think tank': 'Think Tank',
        'ubank': 'uBank',           'westpac': 'Westpac',
        'wlth': 'WLTH',
    }
    SKIP = {'lender', '# settled', '$ settled', 'grand total', 'total', '', 'nan'}

    def canonical(raw):
        low = str(raw).strip().lower()
        if low in COMMERCIAL:
            return COMMERCIAL[low]           # merge commercial into parent
        return NAME_FIX.get(low, str(raw).strip())

    def read_col(col_lender, col_amount):
        """Aggregate $ amounts by canonical lender name."""
        totals = {}
        for i in range(2, len(df)):          # rows 0–1 are header rows
            row = df.iloc[i]
            raw = str(row.iloc[col_lender]).strip()
            if not raw or raw.lower() in SKIP:
                continue
            amt = safe_num(row.iloc[col_amount]) if len(row) > col_amount else 0.0
            if amt <= 0:
                continue
            name = canonical(raw)
            totals[name] = totals.get(name, 0.0) + amt
        return totals

    def bucket(totals, pct_threshold=3.0):
        """Group lenders below threshold into 'Other'. Return sorted list."""
        grand = sum(totals.values())
        if grand <= 0:
            return []
        other = 0.0
        out = []
        for name, amt in totals.items():
            if amt / grand * 100 >= pct_threshold:
                out.append({'name': name, 'amount': round(amt)})
            else:
                other += amt
        if other > 0:
            out.append({'name': 'Other', 'amount': round(other)})
        out.sort(key=lambda x: x['amount'], reverse=True)
        return out

    # FY26 Settlements: col 0 (A) = Lender, col 2 (C) = $ Settled
    # Loan Book:        col 4 (E) = Lender, col 6 (G) = $ amount
    fy26 = bucket(read_col(0, 2))
    book = bucket(read_col(4, 6))
    return fy26, book

# ── Brokers (Lead Generation) ─────────────────────────────────────────────────
def read_brokers(path):
    """Read Brokers sheet. R11/S11/T11 (row 10, cols 17/18/19) = total Calls/Connections/Leads."""
    try:
        df = pd.read_excel(path, sheet_name='Brokers', header=None)
        calls       = int(safe_num(df.iloc[10, 17]))
        connections = int(safe_num(df.iloc[10, 18]))
        leads       = int(safe_num(df.iloc[10, 19]))
        return calls, connections, leads
    except Exception:
        return 0, 0, 0

# ── Lead Pipeline ─────────────────────────────────────────────────────────────
def read_lead_pipeline(path):
    """Read Lead Pipeline sheet. Row 0 = headers, rows 1+ = weekly data.
       Columns: Week Commencing, BC, Waiting Customer, LO, Compliance, Total Leads"""
    try:
        df = pd.read_excel(path, sheet_name='Lead Pipeline', header=None)
    except Exception:
        return []
    rows = []
    for i in range(1, len(df)):
        row = df.iloc[i]
        week = row.iloc[0]
        if pd.isna(week): continue
        bc   = int(safe_num(row.iloc[1]))
        wait = int(safe_num(row.iloc[2]))
        lo   = int(safe_num(row.iloc[3]))
        comp = int(safe_num(row.iloc[4]))
        tot  = int(safe_num(row.iloc[5]))
        if tot == 0 and bc == 0 and lo == 0: continue
        try:
            label = pd.to_datetime(week, dayfirst=True).strftime('%-d %b')
        except Exception:
            label = str(week)
        rows.append({'week': label, 'bc': bc, 'waiting': wait, 'lo': lo, 'compliance': comp, 'total': tot})
    return rows

# ── Lodgement Pipeline ─────────────────────────────────────────────────────────
def read_lodgement_pipeline(path):
    """Read Lodgement Pipeline sheet. Row 0 = headers, rows 1+ = weekly data.
       Columns: Week Commencing, Lodged, OA, Pre-Approved, Formal, Contracts, Settlement Booked, Total"""
    try:
        df = pd.read_excel(path, sheet_name='Lodgement Pipeline', header=None)
    except Exception:
        return []
    rows = []
    for i in range(1, len(df)):
        row = df.iloc[i]
        week = row.iloc[0]
        if pd.isna(week): continue
        lodged  = round(safe_num(row.iloc[1]))
        oa      = round(safe_num(row.iloc[2]))
        preapp  = round(safe_num(row.iloc[3]))
        formal  = round(safe_num(row.iloc[4]))
        contr   = round(safe_num(row.iloc[5]))
        setbook = round(safe_num(row.iloc[6]))
        total   = round(safe_num(row.iloc[7]))
        if total == 0: continue
        try:
            label = pd.to_datetime(week, dayfirst=True).strftime('%-d %b')
        except Exception:
            label = str(week)
        rows.append({'week': label, 'lodged': lodged, 'oa': oa, 'pre_approved': preapp,
                     'formal': formal, 'contracts': contr, 'settlement_booked': setbook, 'total': total})
    return rows

# ── Build payload ──────────────────────────────────────────────────────────────
def build_data(month_data, bc, lo, cp, broker_calls, broker_connections, broker_leads,
               leave, leave_title, all_time, history,
               lender_mix=None, lender_mix_book=None,
               lead_pipeline=None, lodgement_pipeline=None):
    now            = datetime.now(SYDNEY)
    cur_month      = now.strftime('%b')
    days_in_month  = calendar.monthrange(now.year, now.month)[1]
    day_of_month   = now.day
    days_remaining = days_in_month - day_of_month

    # Business days (Mon–Fri) in month and elapsed so far
    biz_total   = sum(1 for dd in range(1, days_in_month + 1)
                      if datetime(now.year, now.month, dd).weekday() < 5)
    biz_elapsed = sum(1 for dd in range(1, day_of_month + 1)
                      if datetime(now.year, now.month, dd).weekday() < 5)

    months = []
    ytd_setts = ytd_lodge = completed_setts = completed_months = 0

    for name in FY_MONTHS:
        md      = month_data.get(name, {})
        setts   = md.get('settlements', 0)
        lodge   = md.get('lodgements',  0)
        target  = MONTHLY_TARGETS.get(name, 0)
        is_cur  = (name == cur_month)
        is_past = (not is_cur and setts > 0)
        is_empty= (setts == 0 and lodge == 0)
        if not is_empty:
            ytd_setts += setts; ytd_lodge += lodge
        if is_past:
            completed_months += 1; completed_setts += setts
        pct = round(setts / target * 100, 1) if target > 0 and setts > 0 else 0
        months.append({
            'name': name, 'lodgements': lodge, 'settlements': setts,
            'deals_lodged': md.get('deals_lodged', 0),
            'deals_settled': md.get('deals_settled', 0),
            'target': target, 'pct_of_target': pct,
            'is_current': is_cur, 'is_past': is_past, 'is_empty': is_empty,
        })

    cur_setts  = month_data.get(cur_month, {}).get('settlements', 0)
    cur_lodge  = month_data.get(cur_month, {}).get('lodgements',  0)
    cur_ld_n   = month_data.get(cur_month, {}).get('deals_lodged', 0)
    cur_st_n   = month_data.get(cur_month, {}).get('deals_settled', 0)
    cur_target = MONTHLY_TARGETS.get(cur_month, 0)

    # MTD target based on business days (daily target × business days elapsed)
    mtd_target = round((cur_target / biz_total) * biz_elapsed) if biz_total > 0 else 0

    # Pace against MTD target
    if biz_elapsed > 0 and mtd_target > 0:
        pace_pct = round(cur_setts / mtd_target * 100, 1)
        pace = 'On Track' if pace_pct >= 95 else ('Slightly Behind' if pace_pct >= 80 else 'Behind')
    else:
        pace_pct, pace = 0.0, 'Early Days'

    # Year-end projection: pace-extrapolated current month + completed + remaining targets
    cur_month_idx = FY_MONTHS.index(cur_month) if cur_month in FY_MONTHS else 0
    if cur_setts > 0 and biz_elapsed > 0 and biz_total > 0:
        proj_this_month = round((cur_setts / biz_elapsed) * biz_total)
        remaining_targets = sum(MONTHLY_TARGETS.get(m, 0) for m in FY_MONTHS[cur_month_idx + 1:])
        projection = round(completed_setts + proj_this_month + remaining_targets)
    else:
        projection = None

    return {
        'annual_target': ANNUAL_TARGET, 'ytd_settlements': ytd_setts, 'ytd_lodgements': ytd_lodge,
        'current_month': cur_month, 'current_month_full': now.strftime('%B'),
        'day_of_month': day_of_month, 'days_in_month': days_in_month, 'days_remaining': days_remaining,
        'biz_days_total': biz_total, 'biz_days_elapsed': biz_elapsed, 'mtd_target': mtd_target,
        'months': months, 'projection': projection, 'completed_months': completed_months,
        'current_month_settlements': cur_setts, 'current_month_lodgements': cur_lodge,
        'current_month_deals_lodged': cur_ld_n, 'current_month_deals_settled': cur_st_n,
        'current_month_target': cur_target,
        'pace_pct': pace_pct, 'pace_status': pace,
        'bc_total': bc, 'lo_total': lo, 'cp_total': cp,
        'broker_calls': broker_calls, 'broker_connections': broker_connections, 'broker_leads': broker_leads,
        'leave': leave, 'leave_title': leave_title,
        'all_time_settlements': all_time,
        'history': history,
        'lender_mix': lender_mix or [],
        'lender_mix_book': lender_mix_book or [],
        'lead_pipeline': lead_pipeline or [],
        'lodgement_pipeline': lodgement_pipeline or [],
        'last_updated': now.strftime('%d %b %Y %-I:%M %p'),
    }

# ── Gist ────────────────────────────────────────────────────────────────────────
def push_gist(data, cfg):
    h = {'Authorization':f'token {cfg["token"]}','Accept':'application/vnd.github.v3+json'}

    # ── Staleness guard ────────────────────────────────────────────────────────
    # Compare this file's SharePoint timestamp against the one stored in the
    # current Gist. If we're about to push data from an OLDER file, skip the
    # update so the Gist (and TV) always hold the most recent data seen.
    new_ts = data.get('_sp_modified', '')
    if new_ts and new_ts != 'unknown':
        try:
            r_cur = requests.get(f'https://api.github.com/gists/{cfg["gist_id"]}',
                                 headers=h, timeout=30)
            if r_cur.status_code == 200:
                cur_file = r_cur.json().get('files', {}).get('dashboard_data.json', {})
                cur_content = cur_file.get('content', '')
                if cur_content:
                    cur_data = json.loads(cur_content)
                    cur_ts = cur_data.get('_sp_modified', '')
                    if cur_ts and cur_ts != 'unknown' and new_ts < cur_ts:
                        print(f'  ⚠ Stale SharePoint file detected!')
                        print(f'    Current Gist was from: {cur_ts}')
                        print(f'    This file is from:     {new_ts}')
                        print(f'    Skipping update — TV keeps the current (newer) data.')
                        return
                    print(f'  ✓ Freshness check passed ({new_ts} >= {cur_ts or "none"})')
        except Exception as e:
            print(f'  ⚠ Staleness check failed ({e}) — proceeding with update')

    r = requests.patch(f'https://api.github.com/gists/{cfg["gist_id"]}', headers=h,
        json={'files':{'dashboard_data.json':{'content':json.dumps(data,indent=2)}}},timeout=30)
    r.raise_for_status(); print('  ✓ Gist updated')

# ── GitHub Pages ────────────────────────────────────────────────────────────────
def deploy_html(html, cfg):
    hsh = hashlib.md5(html.encode()).hexdigest()
    if HTML_HASH_FILE.exists() and HTML_HASH_FILE.read_text().strip()==hsh:
        print('  ✓ HTML unchanged'); return
    h = {'Authorization':f'token {cfg["token"]}','Accept':'application/vnd.github.v3+json'}
    url = f'https://api.github.com/repos/{cfg["repo"]}/contents/index.html'
    resp = requests.get(url,headers=h,timeout=30)
    sha = resp.json().get('sha') if resp.status_code==200 else None
    pl = {'message':f'Dashboard FY27 {datetime.now(SYDNEY).strftime("%Y-%m-%d %H:%M")}',
          'content':base64.b64encode(html.encode()).decode()}
    if sha: pl['sha']=sha
    r = requests.put(url,headers=h,json=pl,timeout=30)
    r.raise_for_status(); HTML_HASH_FILE.write_text(hsh); print('  ✓ HTML deployed')


# ── HTML ────────────────────────────────────────────────────────────────────────
def build_html(gist_url):
    return r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><title>Simplify Finance | FY 2026-27</title>
<!-- Force browser to never cache this page — TV must always load fresh HTML from GitHub Pages -->
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
*{box-sizing:border-box;margin:0;padding:0}
html{font-size:20px}
body{background:#0d1117;color:#e6edf3;font-family:'Segoe UI',system-ui,sans-serif;height:100vh;overflow:hidden}
.db{display:grid;grid-template-columns:1.4fr 1.3fr 1fr;grid-template-rows:2.5fr 1.5fr 1.5fr;
    grid-template-areas:"perf focus leave""ytd rot rot""vis rot rot";gap:9px;padding:9px;height:100vh}
.pnl{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:14px;
     overflow:hidden;display:flex;flex-direction:column}
#pp{grid-area:perf}#pf{grid-area:focus}#pl{grid-area:leave}
#py{grid-area:ytd}#pr{grid-area:rot;position:relative}
/* Vision panel — bottom-left, sits below YTD */
#pv{grid-area:vis;align-items:center;justify-content:center}
/* Panel title — small label, clearly secondary */
.ptitle{font-size:.62rem;color:#7d8590;text-transform:uppercase;letter-spacing:.12em;
        margin-bottom:8px;border-bottom:1px solid #21262d;padding-bottom:5px;flex-shrink:0}
/* ── Table ───────────────────────────────────────────────── */
table{width:100%;border-collapse:collapse}
th{color:#7d8590;font-size:.6rem;text-transform:uppercase;letter-spacing:.05em;
   padding:4px 7px;text-align:right;border-bottom:1px solid #30363d;white-space:nowrap}
th:first-child{text-align:left}
th.grp{text-align:center;color:#7d8590;font-size:.58rem;border-bottom:1px solid #30363d;padding-bottom:2px}
/* Data rows — bright and bold so they read from distance */
td{padding:4px 7px;text-align:right;font-size:.88rem;font-weight:600;
   color:#c9d1d9;border-bottom:1px solid #1a1f26;white-space:nowrap}
td:first-child{text-align:left;color:#7d8590;font-size:.82rem;font-weight:400}
/* Current month — teal highlight, slightly larger */
tr.cur td{background:rgba(0,232,196,.1);color:#00e8c4;font-weight:700;font-size:.92rem}
tr.cur td:first-child{color:#00e8c4;font-weight:700}
/* Future months — very dim and compact, just placeholder rows */
tr.fut td{color:#2d333b;padding:1px 7px;font-size:.7rem;line-height:1.1}
tr.fut td:first-child{color:#2d333b;font-size:.7rem}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:4px;vertical-align:middle}
.g{background:#3fb950}.a{background:#d29922}.r{background:#f85149}.n{background:#2d333b}
.lupd{font-size:.5rem;color:#3a4150;text-align:right;margin-top:auto;padding-top:3px}
.rp{display:inline-block;width:5px;height:5px;border-radius:50%;background:#3fb950;margin-left:3px;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
/* ── Month in Focus ───────────────────────────────────────── */
.f-month{font-size:1.1rem;font-weight:800;color:#00e8c4;margin-bottom:1px;margin-top:0}
/* Section dividers — very dim */
.f-section{font-size:.52rem;color:#7d8590;text-transform:uppercase;letter-spacing:.1em;
           margin-top:1px;margin-bottom:0px;border-top:1px solid #21262d;padding-top:1px}
.f-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:0px}
/* Labels readable but clearly secondary to data values */
.f-label{font-size:.65rem;color:#8b949e}
.f-val{font-size:.82rem;color:#f0f6fc;font-weight:700}
.f-grid{display:grid;grid-template-columns:1fr auto auto;gap:0px 14px;align-items:baseline;margin-bottom:0px}
.f-col-hdr{font-size:.52rem;color:#4e5866;text-transform:uppercase;text-align:right}
/* Big % number — hero element */
.f-pct{font-size:1.45rem;font-weight:900;line-height:1;margin:1px 0}
.f-bar{height:5px;background:#21262d;border-radius:4px;margin:2px 0}
.f-bar-fill{height:100%;border-radius:4px;transition:width .6s}
.f-days{font-size:.58rem;color:#8b949e;margin-bottom:1px}
.f-pace{font-size:.7rem;font-weight:700;padding:2px 8px;border-radius:12px;display:inline-block}
.p-on{background:rgba(63,185,80,.18);color:#3fb950;border:1px solid #3fb950}
.p-sl{background:rgba(210,153,34,.18);color:#d29922;border:1px solid #d29922}
.p-be{background:rgba(248,81,73,.18);color:#f85149;border:1px solid #f85149}
.p-ea{background:rgba(125,133,144,.15);color:#7d8590;border:1px solid #484f58}
/* ── Leave ────────────────────────────────────────────────── */
.leave-entry{display:flex;justify-content:space-between;align-items:baseline;
             margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid #1a1f26}
.leave-entry:last-child{border-bottom:none}
.leave-name{font-size:.88rem;color:#f0f6fc;font-weight:700}
.leave-dates{font-size:.78rem;color:#7d8590;text-align:right}
.no-leave{font-size:.82rem;color:#3a4150;font-style:italic;margin-top:8px}
/* ── YTD ──────────────────────────────────────────────────── */
.ytd-h{font-size:.6rem;color:#7d8590;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px}
.ytd-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px}
/* Hero number — large and teal */
.ytd-num{font-size:2rem;font-weight:800;color:#00e8c4}
.ytd-of{font-size:.72rem;color:#4e5866}
.ytd-bar{height:14px;background:#21262d;border-radius:7px;overflow:hidden;margin-bottom:8px}
.ytd-bar-fill{height:100%;background:linear-gradient(90deg,#00e8c4,#00b4d8);border-radius:7px;transition:width 1s}
.ytd-stats{display:flex;gap:20px}
/* Stat labels dim, stat values bright */
.ytd-stat label{font-size:.58rem;color:#7d8590;text-transform:uppercase;letter-spacing:.08em;display:block;margin-bottom:2px}
.ytd-stat span{font-size:.92rem;color:#f0f6fc;font-weight:700}
/* ── Rotating panel ───────────────────────────────────────── */
.rot-view{position:absolute;inset:0;padding:14px;display:flex;flex-direction:column;
          transition:opacity .8s;opacity:0;pointer-events:none}
.rot-view.active{opacity:1;pointer-events:auto}
.rot-title{font-size:.62rem;color:#7d8590;text-transform:uppercase;letter-spacing:.12em;
           margin-bottom:6px;border-bottom:1px solid #21262d;padding-bottom:5px;flex-shrink:0}
canvas{flex:1;width:100%;min-height:0;display:block}
.rot-indicator{display:flex;gap:7px;justify-content:center;position:absolute;bottom:8px;left:0;right:0}
.rot-dot{width:7px;height:7px;border-radius:50%;background:#2d333b;transition:background .3s}
.rot-dot.active{background:#00e8c4}
/* ── Vision ───────────────────────────────────────────────── */
.vis-text{font-size:1rem;color:#8b949e;text-align:center;line-height:1.7;font-style:italic}
.vis-text b{color:#00e8c4;font-style:normal;font-weight:700;font-size:1.05rem}
</style></head><body>
<div class="db">

  <!-- Performance Table -->
  <div class="pnl" id="pp">
    <div class="ptitle" id="perf-title">Financial Year 2027 Performance</div>
    <table><thead>
      <tr>
        <th rowspan="2">Month</th>
        <th colspan="2" class="grp">Lodgements</th>
        <th colspan="2" class="grp">Settlements</th>
        <th rowspan="2" style="text-align:center">Settlement<br>to Target</th>
      </tr>
      <tr><th>#</th><th>$</th><th>#</th><th>$</th></tr>
    </thead><tbody id="tb"></tbody></table>
    <div class="lupd" id="lu">&mdash;</div>
  </div>

  <!-- Month in Focus -->
  <div class="pnl" id="pf">
    <div class="ptitle" style="margin-bottom:3px;padding-bottom:3px">Month in Focus</div>
    <div class="f-month" id="f-month">&mdash;</div>
    <div class="f-section">Lead Generation</div>
    <div class="f-row"><span class="f-label">Calls</span><span class="f-val" id="f-calls">&mdash;</span></div>
    <div class="f-row"><span class="f-label">Connections</span><span class="f-val" id="f-connections">&mdash;</span></div>
    <div class="f-row"><span class="f-label">Leads</span><span class="f-val" id="f-leads">&mdash;</span></div>
    <div class="f-section">Pipeline</div>
    <div class="f-row"><span class="f-label">Borrowing Capacities</span><span class="f-val" id="f-bc">&mdash;</span></div>
    <div class="f-row"><span class="f-label">Loan Options</span><span class="f-val" id="f-lo">&mdash;</span></div>
    <div class="f-row"><span class="f-label">Compliance Prep</span><span class="f-val" id="f-cp">&mdash;</span></div>
    <div class="f-section">This Month</div>
    <div class="f-grid">
      <span></span><span class="f-col-hdr">#</span><span class="f-col-hdr">$</span>
      <span class="f-label">Lodged</span><span class="f-val" id="f-ld-n">&mdash;</span><span class="f-val" id="f-ld-d">&mdash;</span>
      <span class="f-label">Settled</span><span class="f-val" id="f-st-n">&mdash;</span><span class="f-val" id="f-st-d">&mdash;</span>
    </div>
    <div class="f-section">Settlement Pace</div>
    <div class="f-pct" id="f-pct">&mdash;</div>
    <div class="f-bar" style="margin:1px 0 2px"><div class="f-bar-fill" id="f-bar" style="width:0"></div></div>
    <span class="f-pace p-ea" id="f-pace">&mdash;</span>
  </div>

  <!-- Leave -->
  <div class="pnl" id="pl">
    <div class="ptitle" id="leave-title">Leave</div>
    <div id="leave-list"></div>
  </div>

  <!-- YTD Progress -->
  <div class="pnl" id="py">
    <div class="ytd-h">FY 2026&ndash;27 &mdash; Annual Target: $330M</div>
    <div class="ytd-row">
      <div class="ytd-num" id="y-num">$0</div>
      <div class="ytd-of" id="y-pct">0% of target</div>
    </div>
    <div class="ytd-bar"><div class="ytd-bar-fill" id="y-bar" style="width:0%"></div></div>
    <div class="ytd-stats">
      <div class="ytd-stat"><label>Projected Year-End</label><span id="y-proj">&mdash;</span></div>
      <div class="ytd-stat"><label>Remaining</label><span id="y-rem">$330M</span></div>
    </div>
  </div>

  <!-- Rotating: Gauge + 5yr Chart + Lender Mix (x2) + Pipelines -->
  <div class="pnl" id="pr" style="padding:0;overflow:visible">
    <div class="rot-view active" id="rv0">
      <div class="rot-title">All Time Settlements</div>
      <canvas id="gauge-canvas"></canvas>
    </div>
    <div class="rot-view" id="rv1">
      <div class="rot-title">5 Year Settlement History</div>
      <canvas id="chart-canvas"></canvas>
    </div>
    <div class="rot-view" id="rv2">
      <div class="rot-title">FY26 Settlement Mix</div>
      <canvas id="donut-canvas"></canvas>
    </div>
    <div class="rot-view" id="rv3">
      <div class="rot-title">Lead Pipeline</div>
      <canvas id="lead-canvas"></canvas>
    </div>
    <div class="rot-view" id="rv4">
      <div class="rot-title">Lodgement Pipeline</div>
      <canvas id="lodge-canvas"></canvas>
    </div>
    <div class="rot-view" id="rv5">
      <div class="rot-title">Loan Book Mix &mdash; Jun 2026</div>
      <canvas id="donut-canvas-book"></canvas>
    </div>
    <div class="rot-indicator">
      <div class="rot-dot active" id="rd0"></div>
      <div class="rot-dot" id="rd1"></div>
      <div class="rot-dot" id="rd2"></div>
      <div class="rot-dot" id="rd3"></div>
      <div class="rot-dot" id="rd4"></div>
      <div class="rot-dot" id="rd5"></div>
    </div>
  </div>

  <!-- Vision -->
  <div class="pnl" id="pv">
    <p class="vis-text">
      <b>By making loans simple, efficient, and stress-free,</b><br>
      we empower our clients to achieve their property goals<br>
      and build long-term financial security.
    </p>
  </div>

</div>
<script>
var GIST_ID='GIST_URL_PLACEHOLDER';
var ANN=330000000;
var rotIdx=0, D=null;

function fm(n){if(n>=1e9)return'$'+(n/1e9).toFixed(2)+'B';if(n>=1e6)return'$'+(n/1e6).toFixed(1)+'M';if(n>=1e3)return'$'+(n/1e3).toFixed(0)+'K';return'$'+Math.round(n);}
function col(p){if(p<=0)return'#484f58';if(p>=100)return'#3fb950';if(p>=75)return'#d29922';return'#f85149';}
function dcls(p){if(p<=0)return'n';if(p>=100)return'g';if(p>=75)return'a';return'r';}

function cvSize(cv){
  // Size canvas from parent panel, not from getBoundingClientRect (avoids stale/zero values)
  var pv=cv.parentElement;
  var titleEl=pv.querySelector('.rot-title');
  var titleH=titleEl?titleEl.offsetHeight+10:36;
  var cs=window.getComputedStyle(pv);
  var pH=parseFloat(cs.paddingTop)+parseFloat(cs.paddingBottom);
  var pW=parseFloat(cs.paddingLeft)+parseFloat(cs.paddingRight);
  var w=Math.max(pv.clientWidth-pW,100);
  var h=Math.max(pv.clientHeight-pH-titleH,100);
  cv.width=w; cv.height=h;
  cv.style.width=w+'px'; cv.style.height=h+'px';
}
function drawGauge(allTime){
  var cv=document.getElementById('gauge-canvas');
  cvSize(cv);
  var ctx=cv.getContext('2d');
  ctx.clearRect(0,0,cv.width,cv.height);
  var cx=cv.width/2,cy=cv.height*0.72,r=Math.min(cx,cy)*0.8;
  var maxVal=Math.ceil(allTime/500000000)*500000000+500000000;
  var pct=Math.min(allTime/maxVal,1);
  // Background
  ctx.beginPath();ctx.arc(cx,cy,r,Math.PI,2*Math.PI);
  ctx.strokeStyle='#21262d';ctx.lineWidth=r*0.18;ctx.stroke();
  // Value arc
  ctx.beginPath();ctx.arc(cx,cy,r,Math.PI,Math.PI+pct*Math.PI);
  ctx.strokeStyle='#00e8c4';ctx.lineWidth=r*0.18;ctx.stroke();
  // Inner glow ring
  ctx.beginPath();ctx.arc(cx,cy,r,Math.PI,Math.PI+pct*Math.PI);
  ctx.strokeStyle='rgba(0,232,196,0.15)';ctx.lineWidth=r*0.3;ctx.stroke();
  // Value text
  ctx.fillStyle='#e6edf3';ctx.textAlign='center';ctx.textBaseline='middle';
  ctx.font='bold '+Math.round(r*0.26)+'px Segoe UI,sans-serif';
  ctx.fillText(fm(allTime),cx,cy-r*0.08);
  ctx.fillStyle='#7d8590';ctx.font=Math.round(r*0.11)+'px Segoe UI,sans-serif';
  ctx.fillText('Total Settlements',cx,cy+r*0.18);
  // Min/max labels
  ctx.fillStyle='#484f58';ctx.font=Math.round(r*0.1)+'px Segoe UI,sans-serif';
  ctx.textAlign='left';ctx.fillText('$0',cx-r-4,cy+4);
  ctx.textAlign='right';ctx.fillText(fm(maxVal),cx+r+4,cy+4);
}

function drawChart(history){
  var cv=document.getElementById('chart-canvas');
  cvSize(cv);
  var ctx=cv.getContext('2d');
  ctx.clearRect(0,0,cv.width,cv.height);
  if(!history||!history.length)return;
  // Scale fonts from width (more reliable than height across screen sizes)
  var fs=Math.max(12,Math.round(cv.width*0.022));
  var pL=fs*5,pR=12,pT=12,pB=Math.round(fs*6);
  var cW=cv.width-pL-pR,cH=cv.height-pT-pB;
  var maxV=0;
  history.forEach(function(yr){yr.months.forEach(function(v){if(v>maxV)maxV=v;});});
  if(maxV===0)return;
  maxV=Math.ceil(maxV/5000000)*5000000;
  var colors=['#58a6ff','#7ee787','#d29922','#f78166','#00e8c4'];
  var months=D?D.months.map(function(m){return m.name;}):['Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun'];
  // Grid lines + Y axis labels
  ctx.strokeStyle='#21262d';ctx.lineWidth=0.5;
  for(var g=0;g<=4;g++){
    var gy=pT+cH-(g/4)*cH;
    ctx.beginPath();ctx.moveTo(pL,gy);ctx.lineTo(pL+cW,gy);ctx.stroke();
    ctx.fillStyle='#7d8590';ctx.font=fs+'px Segoe UI';ctx.textAlign='right';
    ctx.fillText(fm(maxV*g/4),pL-4,gy+fs*0.35);
  }
  // Lines — current year thicker and brightest
  history.forEach(function(yr,yi){
    ctx.beginPath();
    ctx.strokeStyle=colors[yi%colors.length];
    ctx.lineWidth=yi===history.length-1?7:3.5;
    ctx.globalAlpha=yi===history.length-1?1:0.65;
    var started=false;
    yr.months.forEach(function(v,mi){
      var x=pL+(mi/11)*cW,y=pT+cH-(v/maxV)*cH;
      if(v===0){started=false;return;}
      if(!started){ctx.moveTo(x,y);started=true;}else{ctx.lineTo(x,y);}
    });
    ctx.stroke();
    ctx.globalAlpha=1;
  });
  // X axis labels
  months.forEach(function(m,i){
    var x=pL+(i/11)*cW;
    ctx.fillStyle='#8b949e';ctx.font=fs+'px Segoe UI';ctx.textAlign='center';
    ctx.fillText(m,x,pT+cH+fs*1.4);
  });
  // Legend — centred, spaced evenly
  ctx.font=(fs+1)+'px Segoe UI';
  var swW=fs*1.4,swH=fs*0.5,gap=fs*0.6;
  var totalW=0;
  history.forEach(function(yr){totalW+=ctx.measureText(yr.label).width+swW+gap+fs*1.5;});
  var lx=(cv.width-totalW)/2,ly=cv.height-fs*0.6;
  history.forEach(function(yr,yi){
    ctx.fillStyle=colors[yi%colors.length];
    ctx.fillRect(lx,ly-swH-2,swW,swH+2);
    ctx.fillStyle='#c9d1d9';ctx.textAlign='left';
    ctx.fillText(yr.label,lx+swW+gap,ly);
    lx+=ctx.measureText(yr.label).width+swW+gap+fs*1.5;
  });
}

function drawDonut(canvasId, lenders, centreLabel){
  var cv=document.getElementById(canvasId);
  if(!cv||!lenders||!lenders.length)return;
  cvSize(cv);
  var ctx=cv.getContext('2d');
  ctx.clearRect(0,0,cv.width,cv.height);
  var colors=['#00e8c4','#00b4d8','#3fb950','#d29922','#f85149','#a371f7','#fb8f44','#79c0ff','#56d364','#e3b341','#ff7b72','#ffa657'];
  var total=lenders.reduce(function(s,l){return s+l.amount;},0);
  if(total<=0)return;
  // Layout: donut on left (~62%), legend on right (~38%)
  var fs=Math.max(10,Math.round(cv.width*0.022));
  var legW=Math.round(cv.width*0.38);
  var donutW=cv.width-legW;
  var cx=donutW/2;
  var r=Math.min(cx*0.9,cv.height/2*0.9);
  var cy=cv.height/2;
  var startAngle=-Math.PI/2;
  // Segments
  lenders.forEach(function(l,i){
    var sweep=l.amount/total*Math.PI*2;
    ctx.beginPath();
    ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,r,startAngle,startAngle+sweep);
    ctx.closePath();
    ctx.fillStyle=colors[i%colors.length];
    ctx.fill();
    startAngle+=sweep;
  });
  // Inner cutout
  ctx.beginPath();
  ctx.arc(cx,cy,r*0.52,0,Math.PI*2);
  ctx.fillStyle='#161b22';
  ctx.fill();
  // Centre: label + total
  ctx.textAlign='center';
  ctx.fillStyle='#7d8590';
  ctx.font=Math.round(fs*0.8)+'px Inter,sans-serif';
  ctx.fillText(centreLabel||'Total',cx,cy-Math.round(fs*0.7));
  ctx.fillStyle='#f0f6fc';
  ctx.font='bold '+Math.round(fs*1.3)+'px Inter,sans-serif';
  ctx.fillText(fm(total),cx,cy+Math.round(fs*0.75));
  // Legend — single column on right, vertically centred
  var swW=Math.round(fs*0.7),swH=Math.round(fs*0.6),gap=4;
  var rowH=Math.round(fs*1.5);
  var legX=donutW+6;
  var startY=Math.max(4,Math.round((cv.height-lenders.length*rowH)/2))+Math.round(fs*0.6);
  ctx.font=Math.round(fs*0.82)+'px Inter,sans-serif';
  lenders.forEach(function(l,i){
    var lly=startY+i*rowH;
    var pct=(l.amount/total*100).toFixed(1)+'%';
    ctx.fillStyle=colors[i%colors.length];
    ctx.fillRect(legX,lly-swH+2,swW,swH);
    ctx.fillStyle='#c9d1d9';
    ctx.textAlign='left';
    ctx.fillText(l.name+'  '+pct,legX+swW+gap,lly);
  });
}

function drawStackedBar(cv,rows,keys,labels,colors,fmVal,stepSize){
  cvSize(cv);
  var ctx=cv.getContext('2d');
  ctx.clearRect(0,0,cv.width,cv.height);
  var fs=Math.max(9,Math.round(cv.width*0.027));
  var sw=Math.round(fs*0.7),sh=Math.round(fs*0.6),gap2=4;
  // Legend at top — wrap into rows as needed
  ctx.font=Math.round(fs*0.82)+'px Inter,sans-serif';
  var legItems=labels.map(function(lb,i){return{lb:lb,w:sw+gap2+ctx.measureText(lb).width+10};});
  var legRowH=sh+5,legRows=[[]],rowW=0;
  legItems.forEach(function(it){
    if(rowW+it.w>cv.width-8&&legRows[legRows.length-1].length>0){legRows.push([]);rowW=0;}
    legRows[legRows.length-1].push(it);rowW+=it.w;
  });
  var legTotalH=legRows.length*legRowH+6;
  var padL=Math.round(cv.width*0.11),padR=8,padT=legTotalH,padB=Math.round(fs*1.6);
  var W=cv.width-padL-padR,H=cv.height-padT-padB;
  // Compute axis max as next step multiple above data max
  var dataMax=Math.max.apply(null,rows.map(function(r){
    return keys.reduce(function(s,k){return s+(r[k]||0);},0);
  }));
  if(dataMax<=0)return;
  var axisMax=Math.ceil(dataMax/stepSize)*stepSize;
  var numSteps=axisMax/stepSize;
  var slot=W/rows.length,bw=Math.round(slot*0.58),bgap=Math.round((slot-bw)/2);
  // Draw legend
  var ly=sh+3;
  legRows.forEach(function(row){
    var lx=padL;
    row.forEach(function(it){
      var gi=labels.indexOf(it.lb);
      ctx.fillStyle=colors[gi%colors.length];ctx.fillRect(lx,ly-sh+1,sw,sh);
      ctx.fillStyle='#c9d1d9';ctx.textAlign='left';ctx.font=Math.round(fs*0.82)+'px Inter,sans-serif';
      ctx.fillText(it.lb,lx+sw+gap2,ly);lx+=it.w;
    });
    ly+=legRowH;
  });
  // Y-axis grid lines at each step
  ctx.font=Math.round(fs*0.72)+'px Inter,sans-serif';ctx.textAlign='right';
  for(var si=0;si<=numSteps;si++){
    var yv=si*stepSize;
    var y=padT+H-Math.round(yv/axisMax*H);
    ctx.fillStyle='#8b949e';ctx.fillText(fmVal(yv),padL-3,y+fs*0.3);
    ctx.strokeStyle=si===0?'#30363d':'#21262d';ctx.lineWidth=1;
    ctx.beginPath();ctx.moveTo(padL,y);ctx.lineTo(padL+W,y);ctx.stroke();
  }
  // Stacked bars with in-segment labels
  var segFs=Math.max(8,Math.round(fs*0.75));
  rows.forEach(function(r,i){
    var x=padL+i*slot+bgap;
    var base=padT+H;
    keys.forEach(function(k,ki){
      var val=r[k]||0;
      if(val<=0)return;
      var bh=Math.round(val/axisMax*H);
      ctx.fillStyle=colors[ki%colors.length];
      ctx.fillRect(x,base-bh,bw,bh);
      // Show value inside segment if tall enough
      var minH=segFs+4;
      if(bh>=minH){
        ctx.fillStyle='rgba(0,0,0,0.6)';ctx.font='bold '+segFs+'px Inter,sans-serif';
        ctx.textAlign='center';
        ctx.fillText(fmVal(val),x+bw/2,base-bh/2+segFs*0.35);
      }
      base-=bh;
    });
    // Week label below
    ctx.fillStyle='#8b949e';ctx.font=Math.round(fs*0.78)+'px Inter,sans-serif';ctx.textAlign='center';
    ctx.fillText(r.week,x+bw/2,padT+H+fs*1.2);
  });
}

function drawLeadPipeline(rows){
  var cv=document.getElementById('lead-canvas');
  if(!cv||!rows||!rows.length)return;
  drawStackedBar(cv,rows,
    ['bc','lo','compliance','waiting'],
    ['BC','LO','Compliance','Waiting'],
    ['#00e8c4','#00b4d8','#d29922','#a371f7'],
    function(v){return String(v);},
    10
  );
}

function drawLodgementPipeline(rows){
  var cv=document.getElementById('lodge-canvas');
  if(!cv||!rows||!rows.length)return;
  function fmM(v){return v>=1000000?'$'+(v/1000000).toFixed(0)+'M':v>=1000?'$'+(v/1000).toFixed(0)+'K':'$'+v;}
  drawStackedBar(cv,rows,
    ['lodged','oa','pre_approved','formal','contracts','settlement_booked'],
    ['Lodged','OA','Pre-Appr','Formal','Contracts','Sett Bkd'],
    ['#00e8c4','#00b4d8','#3fb950','#d29922','#f85149','#a371f7'],
    fmM,
    10000000
  );
}

function rotate(){
  rotIdx=(rotIdx+1)%6;
  for(var i=0;i<6;i++){
    document.getElementById('rv'+i).classList.toggle('active',rotIdx===i);
    document.getElementById('rd'+i).classList.toggle('active',rotIdx===i);
  }
  if(rotIdx===0&&D)setTimeout(function(){drawGauge(D.all_time_settlements);},900);
  if(rotIdx===1&&D)setTimeout(function(){drawChart(D.history);},900);
  if(rotIdx===2&&D)setTimeout(function(){drawDonut('donut-canvas',D.lender_mix,'FY26 Setts');},900);
  if(rotIdx===3&&D)setTimeout(function(){drawLeadPipeline(D.lead_pipeline);},900);
  if(rotIdx===4&&D)setTimeout(function(){drawLodgementPipeline(D.lodgement_pipeline);},900);
  if(rotIdx===5&&D)setTimeout(function(){drawDonut('donut-canvas-book',D.lender_mix_book,'Loan Book');},900);
}

function update(d){
  D=d;
  document.getElementById('lu').innerHTML='Updated: '+d.last_updated+'<span class="rp"></span>';

  // Performance table — 6 columns: Month | Lodg.# | Lodg.$ | Sett.# | Sett.$ | Settlement to Target
  var tb=document.getElementById('tb');tb.innerHTML='';
  d.months.forEach(function(m){
    var tr=document.createElement('tr');
    if(m.is_current)tr.className='cur';
    else if(m.is_empty&&!m.is_past)tr.className='fut';
    var ldN=m.deals_lodged>0?m.deals_lodged:'&mdash;';
    var ldD=m.lodgements>0?fm(m.lodgements):'&mdash;';
    var stN=m.deals_settled>0?m.deals_settled:'&mdash;';
    var stD=m.settlements>0?fm(m.settlements):'&mdash;';
    var vs='&mdash;';
    if(m.is_past&&m.settlements>0&&m.target>0)
      vs='<span class="dot '+dcls(m.pct_of_target)+'"></span>'+m.pct_of_target.toFixed(1)+'%';
    tr.innerHTML='<td>'+m.name+'</td><td>'+ldN+'</td><td>'+ldD+'</td><td>'+stN+'</td><td>'+stD+'</td><td style="text-align:center">'+vs+'</td>';
    tb.appendChild(tr);
  });

  // Month in Focus
  document.getElementById('f-month').textContent=d.current_month_full||d.current_month;
  document.getElementById('f-calls').textContent=d.broker_calls||0;
  document.getElementById('f-connections').textContent=d.broker_connections||0;
  document.getElementById('f-leads').textContent=d.broker_leads||0;
  document.getElementById('f-bc').textContent=d.bc_total||0;
  document.getElementById('f-lo').textContent=d.lo_total||0;
  document.getElementById('f-cp').textContent=d.cp_total||0;
  // Lodged/Settled split into # and $
  document.getElementById('f-ld-n').textContent=d.current_month_deals_lodged||0;
  document.getElementById('f-ld-d').textContent=d.current_month_lodgements>0?fm(d.current_month_lodgements):'—';
  document.getElementById('f-st-n').textContent=d.current_month_deals_settled||0;
  document.getElementById('f-st-d').textContent=d.current_month_settlements>0?fm(d.current_month_settlements):'—';
  // Settlement pace — business-days MTD basis
  var pp=d.pace_pct||0;
  var pe=document.getElementById('f-pct');pe.textContent=pp.toFixed(1)+'%';pe.style.color=col(pp);
  var bf=document.getElementById('f-bar');bf.style.width=Math.min(pp,100)+'%';bf.style.background=col(pp);
  var fdEl=document.getElementById('f-days');
  if(fdEl){var bizTxt=(d.biz_days_elapsed||0)+' of '+(d.biz_days_total||0)+' business days — MTD target: '+fm(d.mtd_target||0);fdEl.textContent=bizTxt;}
  var pac=document.getElementById('f-pace');pac.textContent=d.pace_status;pac.className='f-pace';
  if(d.pace_status==='On Track')pac.classList.add('p-on');
  else if(d.pace_status==='Slightly Behind')pac.classList.add('p-sl');
  else if(d.pace_status==='Behind')pac.classList.add('p-be');
  else pac.classList.add('p-ea');

  // Leave
  document.getElementById('leave-title').textContent=d.leave_title||'Leave';
  var ll=document.getElementById('leave-list');ll.innerHTML='';
  if(d.leave&&d.leave.length){
    d.leave.forEach(function(e){
      var div=document.createElement('div');div.className='leave-entry';
      div.innerHTML='<span class="leave-name">'+e.name+'</span><span class="leave-dates">'+e.dates+'</span>';
      ll.appendChild(div);
    });
  }else{
    ll.innerHTML='<div class="no-leave">No leave scheduled</div>';
  }

  // YTD
  var yp=d.ytd_settlements/ANN*100;
  document.getElementById('y-num').textContent=fm(d.ytd_settlements);
  document.getElementById('y-pct').textContent=yp.toFixed(1)+'% of $330M target';
  document.getElementById('y-bar').style.width=Math.min(yp,100)+'%';
  document.getElementById('y-proj').innerHTML=d.projection?fm(d.projection):'&mdash;';
  document.getElementById('y-rem').textContent=fm(Math.max(0,ANN-d.ytd_settlements));

  // Rotating views
  if(rotIdx===0)drawGauge(d.all_time_settlements);
  else if(rotIdx===1)drawChart(d.history);
  else if(rotIdx===2)drawDonut('donut-canvas',d.lender_mix,'FY26 Setts');
  else if(rotIdx===3)drawLeadPipeline(d.lead_pipeline);
  else if(rotIdx===4)drawLodgementPipeline(d.lodgement_pipeline);
  else drawDonut('donut-canvas-book',d.lender_mix_book,'Loan Book');
}

function go(){
  // Fetch via GitHub API (not raw CDN) — API is never cached by GitHub's CDN,
  // so we always get the latest Gist content without stale-cache issues.
  fetch('https://api.github.com/gists/'+GIST_ID,{cache:'no-store',headers:{'Accept':'application/vnd.github.v3+json'}})
    .then(function(r){
      if(r.status===403||r.status===429){console.warn('GitHub API rate limited — will retry next interval');return null;}
      if(!r.ok){console.warn('Gist fetch error: '+r.status);return null;}
      return r.json();
    })
    .then(function(g){
      if(!g)return;
      var f=g.files&&g.files['dashboard_data.json'];
      if(!f){console.warn('dashboard_data.json not found in gist');return;}
      var content=f.content;
      if(!content&&f.truncated){
        // Content too large — fall back to raw_url (commit-specific, so no CDN stale issue)
        fetch(f.raw_url,{cache:'no-store'}).then(function(r){return r.json();}).then(update).catch(console.error);
        return;
      }
      try{update(JSON.parse(content));}catch(e){console.error(e);}
    })
    .catch(function(e){console.error('Gist fetch failed:',e);});
}

window.addEventListener('resize',function(){
  if(!D)return;
  if(rotIdx===0)drawGauge(D.all_time_settlements);
  else if(rotIdx===1)drawChart(D.history);
  else if(rotIdx===2)drawDonut('donut-canvas',D.lender_mix,'FY26 Setts');
  else if(rotIdx===3)drawLeadPipeline(D.lead_pipeline);
  else if(rotIdx===4)drawLodgementPipeline(D.lodgement_pipeline);
  else drawDonut('donut-canvas-book',D.lender_mix_book,'Loan Book');
});

// Fullscreen — click, Enter, Space, F, or F11 on TV remote
function toggleFS(){
  if(!document.fullscreenElement){
    (document.documentElement.requestFullscreen||document.documentElement.webkitRequestFullscreen||function(){}).call(document.documentElement);
  }else{
    (document.exitFullscreen||document.webkitExitFullscreen||function(){}).call(document);
  }
}
document.addEventListener('click',toggleFS);
document.addEventListener('keydown',function(e){
  if(['F11','f','F','Enter',' '].indexOf(e.key)>-1){toggleFS();e.preventDefault();}
});

go();
// Data refreshes every 5 min — GitHub API unauthenticated limit is 60 req/hr;
// 60-second polling hits that ceiling. Data only changes every 15 min anyway.
setInterval(go,300000);
setInterval(rotate,60000);

// Hard-reload every 10 minutes so the TV always runs the latest HTML from GitHub Pages.
// A soft reload can serve a browser-cached old copy — we force a brand-new page load
// by changing the URL with a timestamp query param (browser treats it as a different URL).
setInterval(function(){
  var base=window.location.href.split('?')[0];
  window.location.replace(base+'?r='+Date.now());
},600000);
</script></body></html>""".replace('GIST_URL_PLACEHOLDER', gist_url)  # gist_url is now the gist ID

# ── Main ────────────────────────────────────────────────────────────────────────
def main():
    print(f'[{datetime.now(SYDNEY).strftime("%Y-%m-%d %H:%M:%S")} AEST] Simplify Finance Dashboard — FY2026-27')
    cloud = is_cloud()

    # ── Config ──────────────────────────────────────────────────────────────────
    if cloud:
        print('  ℹ Cloud mode — reading from SharePoint')
        gh_token  = os.environ['GH_TOKEN']
        gist_id   = os.environ['GIST_ID']
        pages_repo = os.environ['PAGES_REPO']
        gist_cfg  = {'token': gh_token, 'gist_id': gist_id,
                     'raw_url': f'https://gist.githubusercontent.com/AlanSimplifyFinance/{gist_id}/raw/dashboard_data.json'}
        pages_cfg = {'token': gh_token, 'repo': pages_repo}
    else:
        try:
            gist_cfg  = json.loads(GITHUB_CFG.read_text())
            pages_cfg = json.loads(GITHUB_PAGES_CFG.read_text())
        except FileNotFoundError as e:
            print(f'  ✗ Config: {e}'); return 1

    # ── Excel source ────────────────────────────────────────────────────────────
    if cloud:
        try:
            path, sp_modified = download_excel_cloud()
            file_hash = hashlib.md5(Path(path).read_bytes()).hexdigest()[:8]
            file_size = Path(path).stat().st_size
            print(f'  ✓ Downloaded: {file_size:,} bytes | hash={file_hash} | modified={sp_modified}')
        except Exception as e:
            print(f'  ✗ SharePoint download: {e}'); return 1
    else:
        try:
            path = find_excel()
            print(f'  ✓ Local Excel: {path.name}')
        except FileNotFoundError as e:
            print(f'  ✗ Excel: {e}'); return 1

    df_stats = pd.read_excel(path, sheet_name='Year On Year Stats', header=None)

    try:
        month_data = read_current_year(df_stats)
        print(f'  ✓ {len(month_data)} months read')
    except Exception as e:
        print(f'  ✗ Year data: {e}'); return 1

    try:
        all_time, history = read_history(df_stats)
        print(f'  ✓ All-time: {all_time:,.0f} | History: {len(history)} years')
    except Exception as e:
        print(f'  ✗ History: {e}'); all_time, history = 0, []

    now = datetime.now(SYDNEY)
    cur_short = now.strftime('%b')
    cur_full  = now.strftime('%B')

    bc, lo, cp = read_credit_team(path, cur_short)
    print(f'  ✓ BCs: {bc} | LOs: {lo} | CPs: {cp}')

    broker_calls, broker_connections, broker_leads = read_brokers(path)
    print(f'  ✓ Lead gen: {broker_calls} calls | {broker_connections} connections | {broker_leads} leads')

    leave, leave_title = read_leave(path, cur_full)
    if leave:
        print(f'  ✓ Leave entries: {len(leave)} ({leave_title}): {[e["name"] for e in leave]}')
    else:
        print(f'  ✓ Leave entries: 0 — none current ({leave_title})')

    lender_mix, lender_mix_book = read_lender_mix(path)
    print(f'  ✓ Lender mix (FY26 settlements): {len(lender_mix)} entries')
    print(f'  ✓ Lender mix (Loan Book):         {len(lender_mix_book)} entries')

    lead_pipeline = read_lead_pipeline(path)
    print(f'  ✓ Lead pipeline: {len(lead_pipeline)} weeks')

    lodgement_pipeline = read_lodgement_pipeline(path)
    print(f'  ✓ Lodgement pipeline: {len(lodgement_pipeline)} weeks')

    data = build_data(month_data, bc, lo, cp, broker_calls, broker_connections, broker_leads,
                      leave, leave_title, all_time, history,
                      lender_mix, lender_mix_book, lead_pipeline, lodgement_pipeline)
    print(f'  ✓ YTD: ${data["ytd_settlements"]:,.0f} | Pace: {data["pace_status"]}')
    # Tag the data with the SharePoint file timestamp — used as staleness guard in push_gist
    if cloud:
        data['_sp_modified'] = sp_modified

    try:
        push_gist(data, gist_cfg)
    except Exception as e:
        print(f'  ✗ Gist: {e}'); return 1

    html = build_html(gist_cfg['gist_id'])  # pass gist ID so HTML uses API (not cached CDN)
    try:
        deploy_html(html, pages_cfg)
    except Exception as e:
        print(f'  ✗ Deploy: {e}'); return 1

    print('  ✓ Done')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
