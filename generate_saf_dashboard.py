#!/usr/bin/env python3
"""Simplify Asset Finance Dashboard Generator"""

import io, os, json, hashlib, base64, tempfile
import pandas as pd, requests
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

SYDNEY = ZoneInfo('Australia/Sydney')
SCRIPT_DIR = Path(__file__).parent

# ── Config ──────────────────────────────────────────────────────────────────
EXCEL_CANDIDATES = [
    SCRIPT_DIR / 'SAF - Application Report.xlsx',
    Path.home() / 'Library/CloudStorage/OneDrive-SimplifyFinance/Communication site - Operations/Dashboards/SAF - Application Report.xlsx',
]

# ── Cloud helpers ────────────────────────────────────────────────────────────
def is_cloud():
    return bool(os.environ.get('AZURE_CLIENT_SECRET'))

def get_graph_token():
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
    token = get_graph_token()
    headers = {'Authorization': f'Bearer {token}'}
    FILE = 'SAF - Application Report.xlsx'
    # Try most likely SharePoint paths (same root drive as Business Data.xlsx)
    paths = [
        f'https://graph.microsoft.com/v1.0/sites/simplifyfin.sharepoint.com/drive/root:/Operations/Dashboards/{FILE}:/content',
        f'https://graph.microsoft.com/v1.0/sites/simplifyfin.sharepoint.com/drive/root:/Operations/Asset Finance/{FILE}:/content',
        f'https://graph.microsoft.com/v1.0/sites/simplifyfin.sharepoint.com/drive/root:/Shared Documents/Asset Finance/{FILE}:/content',
        f'https://graph.microsoft.com/v1.0/sites/simplifyfin.sharepoint.com/drive/root:/Asset Finance/{FILE}:/content',
        f'https://graph.microsoft.com/v1.0/sites/simplifyfin.sharepoint.com/drive/root:/Shared Documents/{FILE}:/content',
    ]
    for url in paths:
        r = requests.get(url, headers=headers, timeout=60)
        if r.status_code == 200:
            tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
            tmp.write(r.content); tmp.close()
            print(f'  ✓ Downloaded from: {url}')
            return Path(tmp.name)
        print(f'  ✗ {r.status_code} at {url}')
    raise FileNotFoundError(f'Could not find {FILE} in SharePoint — confirm the path and update EXCEL_CANDIDATES')

def find_excel():
    for p in EXCEL_CANDIDATES:
        if p.exists():
            return p
    raise FileNotFoundError('SAF - Application Report.xlsx not found locally')

# ── Data helpers ─────────────────────────────────────────────────────────────
def safe_num(v):
    try: return float(v) if not pd.isna(v) else 0.0
    except: return 0.0

SAB_PARTNERS = {'beach box capital pty ltd', 'simplify finance - jaden head'}

def classify_channel(row):
    biz = str(row['Business']).strip().lower()
    if 'roam' in biz:
        return 'Roam Finance'
    ref = str(row['Referral Partner']).strip().lower()
    if ref in SAB_PARTNERS:
        return 'Simplify Asset Broker'
    return 'Simplify Asset Finance'

CHANNELS = ['Simplify Asset Finance', 'Simplify Asset Broker', 'Roam Finance']
CH_SHORT  = {'Simplify Asset Finance': 'SAF', 'Simplify Asset Broker': 'SAB', 'Roam Finance': 'Roam'}
CH_COLOR  = {'Simplify Asset Finance': '#00e8c4', 'Simplify Asset Broker': '#00b4d8', 'Roam Finance': '#3fb950'}
FY_MONTHS = ['Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun']

def read_applications(path):
    df = pd.read_excel(path, sheet_name='Applications', header=0)
    df.columns = df.columns.str.strip()
    df['Loan Amount']    = pd.to_numeric(df['Loan Amount'],    errors='coerce').fillna(0)
    df['Enquiry Date']   = pd.to_datetime(df['Enquiry Date'],   errors='coerce')
    df['Submission Date']= pd.to_datetime(df['Submission Date'],errors='coerce')
    df['Settlement Date']= pd.to_datetime(df['Settlement Date'],errors='coerce')
    df['Business']       = df['Business'].astype(str).str.strip()
    df['Channel']        = df.apply(classify_channel, axis=1)
    df['Status']         = df['Status'].astype(str).str.strip()
    return df

def metrics_for(df, enq_month=None, sub_month=None, sett_month=None):
    """Return (enq_n, sub_n, sub_val, sett_n, sett_val) for a given period."""
    enq  = df[df['Enquiry Date'].dt.to_period('M') == enq_month] if enq_month else df.iloc[0:0]
    sub  = df[df['Submission Date'].dt.to_period('M') == sub_month] if sub_month else df.iloc[0:0]
    sett = df[(df['Settlement Date'].dt.to_period('M') == sett_month) &
              (df['Status'] == 'Settled')] if sett_month else df.iloc[0:0]
    return {
        'enq_n':    int(len(enq)),
        'sub_n':    int(len(sub)),
        'sub_val':  round(float(sub['Loan Amount'].sum())),
        'sett_n':   int(len(sett)),
        'sett_val': round(float(sett['Loan Amount'].sum())),
    }

def build_data(df):
    now = datetime.now(SYDNEY)
    cur_period = now.to_period('M') if hasattr(now, 'to_period') else pd.Period(now, freq='M')
    cur_period = pd.Period(now.strftime('%Y-%m'), freq='M')
    cur_month_full = now.strftime('%B %Y')

    # ── Current month by channel ──────────────────────────────────────────
    current_channels = []
    tot_cur = {'enq_n':0,'sub_n':0,'sub_val':0,'sett_n':0,'sett_val':0}
    for ch in CHANNELS:
        ch_df = df[df['Channel'] == ch]
        m = metrics_for(ch_df, cur_period, cur_period, cur_period)
        m['channel'] = ch
        m['short']   = CH_SHORT[ch]
        m['color']   = CH_COLOR[ch]
        current_channels.append(m)
        for k in tot_cur: tot_cur[k] += m[k]

    # ── Full FY breakdown ─────────────────────────────────────────────────
    # Determine current FY start
    fy_start_year = now.year if now.month >= 7 else now.year - 1
    fy_start = pd.Period(f'{fy_start_year}-07', freq='M')

    monthly = []
    for i, mon_name in enumerate(FY_MONTHS):
        period = fy_start + i
        mon_channels = []
        tot_mon = {'enq_n':0,'sub_n':0,'sub_val':0,'sett_n':0,'sett_val':0}
        is_future = period > cur_period
        for ch in CHANNELS:
            ch_df = df[df['Channel'] == ch]
            m = metrics_for(ch_df, period, period, period)
            m['channel'] = ch
            m['short']   = CH_SHORT[ch]
            m['color']   = CH_COLOR[ch]
            mon_channels.append(m)
            for k in tot_mon: tot_mon[k] += m[k]
        is_current = (period == cur_period)
        monthly.append({
            'month': mon_name,
            'period': str(period),
            'is_current': is_current,
            'is_future': is_future,
            'channels': mon_channels,
            'total': tot_mon,
        })

    # ── FY settlement history (total $ per month, one series per FY) ────
    # Data starts Aug 2023 (FY24). Include all FYs up to current.
    data_start_fy  = 2023   # Jul 2023 = FY24 start
    cur_fy_start   = now.year if now.month >= 7 else now.year - 1
    fy_palette     = ['#484f58', '#3fb950', '#00b4d8', '#00e8c4']  # oldest → newest
    fy_years       = list(range(data_start_fy, cur_fy_start + 1))
    fy_series      = []
    for idx, fy_yr in enumerate(fy_years):
        fy_label = f'FY{(fy_yr + 1) % 100:02d}'
        color    = fy_palette[min(idx, len(fy_palette) - 1)]
        # Always use the last colour for the current FY regardless of count
        if fy_yr == cur_fy_start:
            color = fy_palette[-1]
        values, saf_values = [], []
        saf_df = df[df['Channel'] == 'Simplify Asset Finance']
        for i in range(12):          # 0=Jul … 5=Dec, 6=Jan … 11=Jun
            mo = i + 7 if i < 6 else i - 5
            yr = fy_yr if i < 6 else fy_yr + 1
            p  = pd.Period(f'{yr}-{mo:02d}', freq='M')
            # For current FY: null out current month and beyond (only show completed months)
            if fy_yr == cur_fy_start and p >= cur_period:
                values.append(None)
                saf_values.append(None)
            else:
                sett = df[(df['Settlement Date'].dt.to_period('M') == p) &
                          (df['Status'] == 'Settled')]
                values.append(round(float(sett['Loan Amount'].sum())))
                saf_sett = saf_df[(saf_df['Settlement Date'].dt.to_period('M') == p) &
                                  (saf_df['Status'] == 'Settled')]
                saf_values.append(round(float(saf_sett['Loan Amount'].sum())))
        fy_series.append({'fy': fy_label, 'color': color,
                          'is_current': fy_yr == cur_fy_start,
                          'values': values, 'saf_values': saf_values})
    history = {'months': FY_MONTHS, 'series': fy_series}

    return {
        'current_month': cur_month_full,
        'current_channels': current_channels,
        'current_total': tot_cur,
        'monthly': monthly,
        'history': history,
        'last_updated': now.strftime('%d %b %Y %-I:%M %p'),
    }

# ── Gist push ────────────────────────────────────────────────────────────────
def push_gist(data, cfg):
    payload = json.dumps(data, default=str)
    h = hashlib.md5(payload.encode()).hexdigest()
    r = requests.patch(
        f'https://api.github.com/gists/{cfg["gist_id"]}',
        headers={'Authorization': f'token {cfg["token"]}',
                 'Accept': 'application/vnd.github.v3+json'},
        json={'files': {'saf_dashboard_data.json': {'content': payload}}},
        timeout=30)
    r.raise_for_status()
    print(f'  ✓ Gist updated ({h[:8]})')

# ── GitHub Pages deploy ──────────────────────────────────────────────────────
def deploy_pages(html, cfg):
    token, repo = cfg['token'], cfg['repo']
    headers = {'Authorization': f'token {token}',
                'Accept': 'application/vnd.github.v3+json'}
    h = hashlib.md5(html.encode()).hexdigest()
    encoded = base64.b64encode(html.encode()).decode()
    url = f'https://api.github.com/repos/{repo}/contents/saf-dashboard.html'
    r = requests.get(url, headers=headers, timeout=30)
    sha = r.json().get('sha') if r.status_code == 200 else None
    payload = {'message': f'SAF Dashboard {datetime.now(SYDNEY).strftime("%Y-%m-%d %H:%M")}',
               'content': encoded}
    if sha: payload['sha'] = sha
    r = requests.put(url, headers=headers, json=payload, timeout=30)
    r.raise_for_status()
    print(f'  ✓ HTML deployed ({h[:8]})')

# ── HTML template ─────────────────────────────────────────────────────────────
def build_html(gist_url):
    return """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<title>Simplify Asset Finance | Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0d1117;color:#e6edf3;font-family:'Segoe UI',system-ui,sans-serif;
     height:100vh;overflow:hidden;display:flex;flex-direction:column}
/* Header */
.hdr{display:flex;align-items:center;justify-content:space-between;
     padding:10px 20px;border-bottom:1px solid #30363d;flex-shrink:0}
.hdr-title{font-size:1.1rem;font-weight:700;color:#00e8c4;letter-spacing:.04em}
.hdr-month{font-size:.85rem;color:#8b949e}
.hdr-updated{font-size:.72rem;color:#484f58}
/* Main grid */
.main{display:grid;grid-template-columns:1fr 1.7fr;grid-template-rows:1fr 1fr;
      gap:12px;padding:12px;flex:1;min-height:0}
.pnl{background:#161b22;border:1px solid #30363d;border-radius:10px;
     padding:10px;overflow:hidden;display:flex;flex-direction:column}
.ptitle{font-size:.6rem;color:#7d8590;text-transform:uppercase;letter-spacing:.12em;margin-bottom:4px}
/* Tables */
table{width:100%;border-collapse:collapse;font-size:.63rem}
th{font-size:.52rem;color:#7d8590;text-transform:uppercase;letter-spacing:.06em;
   padding:1px 4px;text-align:right;border-bottom:1px solid #30363d}
th:first-child{text-align:left}
td{padding:0px 4px;line-height:1.55;text-align:right;color:#c9d1d9;border-bottom:1px solid #1a1f26;white-space:nowrap}
td:first-child{text-align:left}
tr.month-sep td{border-top:2px solid #21262d}
tr.cur-month td{background:rgba(0,232,196,.06)}
tr.future td{color:#484f58}
tr.ch-total td{color:#00e8c4;font-weight:800;font-size:.66rem;background:#1a2230;border-top:1px solid #21262d}
tr.grand-total td{color:#00e8c4;font-weight:700;border-top:1px solid #30363d;background:#1c2128}
.ch-dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px}
.num{color:#f0f6fc;font-weight:700}
.val{color:#58a6ff}
.dim{color:#484f58}
/* Panel placement */
#pnl-year{grid-column:1;grid-row:1/3}
canvas{width:100%;height:100%}
/* Chart rotation */
.chart-hdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;flex-shrink:0}
.rdots{display:flex;gap:5px}
.rdot{width:6px;height:6px;border-radius:50%;background:#30363d;cursor:pointer}
.rdot.active{background:#00e8c4}
.chart-view{flex:1;min-height:0;display:none}
.chart-view.active{display:flex;flex-direction:column}
</style>
</head><body>
<div class="hdr">
  <div class="hdr-title">Simplify Asset Finance &mdash; Pipeline Dashboard</div>
  <div class="hdr-month" id="hdr-month">&mdash;</div>
  <div class="hdr-updated" id="hdr-updated">Loading&hellip;</div>
</div>
<div class="main">
  <!-- Full year table — left column, full height -->
  <div class="pnl" id="pnl-year">
    <div class="ptitle">FY2026&ndash;27 Monthly Breakdown</div>
    <table>
      <thead><tr>
        <th>Month</th><th>Channel</th>
        <th>Enq</th>
        <th>Sub #</th><th>Sub $</th>
        <th>Sett #</th><th>Sett $</th>
      </tr></thead>
      <tbody id="year-body"></tbody>
    </table>
  </div>
  <!-- Current month — right column, top -->
  <div class="pnl" id="pnl-cur">
    <div class="ptitle" id="cur-title">Current Month</div>
    <table>
      <thead><tr>
        <th>Channel</th>
        <th>Enq</th>
        <th>Sub #</th><th>Sub $</th>
        <th>Sett #</th><th>Sett $</th>
      </tr></thead>
      <tbody id="cur-body"></tbody>
    </table>
  </div>
  <!-- Chart — right column, bottom (rotating) -->
  <div class="pnl" id="chart-pnl">
    <div class="chart-hdr">
      <div class="ptitle" id="chart-title">Total Settlements by Month</div>
      <div class="rdots"><span class="rdot active" id="cd0"></span><span class="rdot" id="cd1"></span></div>
    </div>
    <div class="chart-view active" id="cv0"><canvas id="total-canvas"></canvas></div>
    <div class="chart-view" id="cv1"><canvas id="saf-canvas"></canvas></div>
  </div>
</div>

<script>
var GIST='GIST_URL_PLACEHOLDER';
var D=null;

function fm(v){
  if(v>=1000000)return'$'+(v/1000000).toFixed(1)+'M';
  if(v>=1000)return'$'+(v/1000).toFixed(0)+'K';
  return'$'+v;
}
function fn(v){return v>0?v:'—';}

function update(d){
  D=d;
  document.getElementById('hdr-month').textContent=d.current_month||'';
  document.getElementById('hdr-updated').textContent='Updated: '+d.last_updated;

  // Current month table
  var cb=document.getElementById('cur-body');cb.innerHTML='';
  d.current_channels.forEach(function(ch){
    var tr=document.createElement('tr');
    tr.innerHTML='<td><span class="ch-dot" style="background:'+ch.color+'"></span>'+ch.short+'</td>'+
      '<td class="num">'+fn(ch.enq_n)+'</td>'+
      '<td class="num">'+fn(ch.sub_n)+'</td><td class="val">'+(ch.sub_val>0?fm(ch.sub_val):'—')+'</td>'+
      '<td class="num">'+fn(ch.sett_n)+'</td><td class="val">'+(ch.sett_val>0?fm(ch.sett_val):'—')+'</td>';
    cb.appendChild(tr);
  });
  var t=d.current_total;
  var tot=document.createElement('tr');tot.className='grand-total';
  tot.innerHTML='<td>Total</td>'+
    '<td class="num">'+fn(t.enq_n)+'</td>'+
    '<td class="num">'+fn(t.sub_n)+'</td><td class="val">'+(t.sub_val>0?fm(t.sub_val):'—')+'</td>'+
    '<td class="num">'+fn(t.sett_n)+'</td><td class="val">'+(t.sett_val>0?fm(t.sett_val):'—')+'</td>';
  cb.appendChild(tot);

  // Full year table
  var yb=document.getElementById('year-body');yb.innerHTML='';
  d.monthly.forEach(function(mon,mi){
    var baseCls=mon.is_future?'future':(mon.is_current?'cur-month':'');
    mon.channels.forEach(function(ch,ci){
      var tr=document.createElement('tr');
      var cls=baseCls;
      if(ci===0&&mi>0)cls+=' month-sep';
      tr.className=cls.trim();
      var monthCell=ci===0?'<td rowspan="4" style="vertical-align:middle;color:'+(mon.is_current?'#00e8c4':'#c9d1d9')+';font-weight:'+(mon.is_current?700:400)+'">'+mon.month+'</td>':'';
      tr.innerHTML=monthCell+
        '<td style="text-align:left"><span class="ch-dot" style="background:'+ch.color+'"></span>'+ch.short+'</td>'+
        '<td>'+(ch.enq_n>0?ch.enq_n:'<span class="dim">—</span>')+'</td>'+
        '<td>'+(ch.sub_n>0?ch.sub_n:'<span class="dim">—</span>')+'</td>'+
        '<td class="val">'+(ch.sub_val>0?fm(ch.sub_val):'<span class="dim">—</span>')+'</td>'+
        '<td>'+(ch.sett_n>0?ch.sett_n:'<span class="dim">—</span>')+'</td>'+
        '<td class="val">'+(ch.sett_val>0?fm(ch.sett_val):'<span class="dim">—</span>')+'</td>';
      yb.appendChild(tr);
    });
    // Month total row
    var t=mon.total;
    var tot=document.createElement('tr');
    tot.className='ch-total'+(mon.is_future?' future':'');
    tot.innerHTML=
      '<td style="color:#8b949e;font-size:.65rem;text-transform:uppercase;letter-spacing:.06em">Total</td>'+
      '<td class="num">'+(t.enq_n>0?t.enq_n:'—')+'</td>'+
      '<td class="num">'+(t.sub_n>0?t.sub_n:'—')+'</td>'+
      '<td class="val">'+(t.sub_val>0?fm(t.sub_val):'—')+'</td>'+
      '<td class="num">'+(t.sett_n>0?t.sett_n:'—')+'</td>'+
      '<td class="val">'+(t.sett_val>0?fm(t.sett_val):'—')+'</td>';
    yb.appendChild(tot);
  });

  drawChart(d.history,'total-canvas','values');
  drawChart(d.history,'saf-canvas','saf_values');
}

var CHART_TITLES=['Total Settlements by Month','SAF Settlements by Month'];
var chartIdx=0;
function rotateChart(){
  chartIdx=(chartIdx+1)%2;
  document.getElementById('cv0').className='chart-view'+(chartIdx===0?' active':'');
  document.getElementById('cv1').className='chart-view'+(chartIdx===1?' active':'');
  document.getElementById('cd0').className='rdot'+(chartIdx===0?' active':'');
  document.getElementById('cd1').className='rdot'+(chartIdx===1?' active':'');
  document.getElementById('chart-title').textContent=CHART_TITLES[chartIdx];
  if(D){
    setTimeout(function(){
      if(chartIdx===0)drawChart(D.history,'total-canvas','values');
      else drawChart(D.history,'saf-canvas','saf_values');
    },200);
  }
}
setInterval(rotateChart,10000);

function drawChart(history,canvasId,valKey){
  valKey=valKey||'values';
  var cv=document.getElementById(canvasId||'total-canvas');
  if(!cv||!history||!history.series||!history.series.length)return;
  var dpr=window.devicePixelRatio||1;
  var rect=cv.getBoundingClientRect();
  cv.width=Math.round(rect.width*dpr);
  cv.height=Math.round(rect.height*dpr);
  var ctx=cv.getContext('2d');
  ctx.scale(dpr,dpr);
  var W=rect.width,H=rect.height;
  ctx.clearRect(0,0,W,H);

  var months=history.months;
  var series=history.series;
  var fs=Math.max(9,Math.round(W*0.022));
  var legH=Math.round(fs*1.5);
  var padL=Math.round(W*0.1),padR=10,padT=8,padB=Math.round(fs*1.4+legH);
  var cW=W-padL-padR,cH=H-padT-padB;

  // Max across all series (use selected value key)
  var maxVal=0;
  series.forEach(function(s){(s[valKey]||s.values).forEach(function(v){if(v&&v>maxVal)maxVal=v;});});
  if(maxVal<=0)return;
  var step=Math.pow(10,Math.floor(Math.log10(maxVal)));
  if(maxVal/step>5)step*=2;
  var axMax=Math.ceil(maxVal/step)*step;

  // Y-axis gridlines and labels
  ctx.font=Math.round(fs*0.78)+'px Inter,sans-serif';ctx.textAlign='right';ctx.fillStyle='#8b949e';
  for(var v=0;v<=axMax;v+=step){
    var y=padT+cH-Math.round(v/axMax*cH);
    ctx.fillText(fm(v),padL-3,y+fs*0.3);
    ctx.strokeStyle=v===0?'#30363d':'#1a1f26';ctx.lineWidth=1;
    ctx.beginPath();ctx.moveTo(padL,y);ctx.lineTo(padL+cW,y);ctx.stroke();
  }

  // X-axis month labels (Jul through Jun)
  var slotW=cW/months.length;
  ctx.fillStyle='#8b949e';ctx.textAlign='center';ctx.font=Math.round(fs*0.78)+'px Inter,sans-serif';
  months.forEach(function(m,i){
    ctx.fillText(m,padL+i*slotW+slotW/2,padT+cH+fs*1.2);
  });

  // One line per FY — older years thinner/dimmer, nulls break the line
  series.forEach(function(s){
    var isCur=s.is_current;
    var vals=s[valKey]||s.values;
    ctx.beginPath();
    ctx.strokeStyle=s.color;
    ctx.lineWidth=isCur?3.5:2.5;
    ctx.lineJoin='round';
    ctx.globalAlpha=isCur?1:0.65;
    var started=false;
    vals.forEach(function(v,i){
      if(v===null||v===undefined){started=false;return;}
      var x=padL+i*slotW+slotW/2;
      var y=padT+cH-Math.round(v/axMax*cH);
      if(!started){ctx.moveTo(x,y);started=true;}else{ctx.lineTo(x,y);}
    });
    ctx.stroke();
    ctx.globalAlpha=1;
  });

  // Legend
  ctx.font=Math.round(fs*0.82)+'px Inter,sans-serif';ctx.textAlign='left';
  var lx=padL,ly=H-Math.round(fs*0.3);
  series.forEach(function(s){
    ctx.fillStyle=s.color;
    ctx.fillRect(lx,ly-Math.round(fs*0.75),Math.round(fs*0.75),Math.round(fs*0.6));
    ctx.fillStyle='#c9d1d9';
    ctx.fillText(s.fy,lx+Math.round(fs*0.75)+4,ly);
    lx+=ctx.measureText(s.fy).width+Math.round(fs*1.8)+8;
  });
}

function go(){
  var x=new XMLHttpRequest();x.open('GET',GIST+'?t='+Date.now(),true);
  x.onload=function(){if(x.status===200){try{update(JSON.parse(x.responseText))}catch(e){console.error(e)}}};
  x.send();
}

var inFullscreen=false;
function goFullScreen(){
  var el=document.documentElement;
  if     (el.requestFullscreen)       el.requestFullscreen();
  else if(el.webkitRequestFullscreen) el.webkitRequestFullscreen();
  else if(el.mozRequestFullScreen)    el.mozRequestFullScreen();
  else if(el.msRequestFullscreen)     el.msRequestFullscreen();
  inFullscreen=true;
}
function onFsChange(){
  var isFs=!!(document.fullscreenElement||document.webkitFullscreenElement||
              document.mozFullScreenElement||document.msFullscreenElement);
  if(inFullscreen&&!isFs){setTimeout(goFullScreen,800);}
}
document.addEventListener('fullscreenchange',onFsChange);
document.addEventListener('webkitfullscreenchange',onFsChange);
document.addEventListener('mozfullscreenchange',onFsChange);
document.addEventListener('MSFullscreenChange',onFsChange);
window.addEventListener('load',function(){setTimeout(goFullScreen,500);});
document.addEventListener('click',goFullScreen);
window.addEventListener('resize',function(){
  if(D){
    drawChart(D.history,'total-canvas','values');
    drawChart(D.history,'saf-canvas','saf_values');
  }
});
go();
setInterval(go,60000);
</script>
</body></html>""".replace('GIST_URL_PLACEHOLDER', gist_url)

# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    print(f'[{datetime.now(SYDNEY).strftime("%Y-%m-%d %H:%M:%S")} AEST] SAF Dashboard Generator')
    cloud = is_cloud()

    if cloud:
        print('  ℹ Cloud mode — reading from SharePoint')
        gh_token   = os.environ['GH_TOKEN']
        gist_id    = os.environ['SAF_GIST_ID']
        pages_repo = os.environ['PAGES_REPO']
        gist_cfg   = {
            'token':   gh_token,
            'gist_id': gist_id,
            'raw_url': f'https://gist.githubusercontent.com/AlanSimplifyFinance/{gist_id}/raw/saf_dashboard_data.json'
        }
        pages_cfg  = {'token': gh_token, 'repo': pages_repo}
        path = download_excel_cloud()
    else:
        gist_cfg  = json.loads((SCRIPT_DIR / 'github_config.json').read_text())
        pages_cfg = json.loads((SCRIPT_DIR / 'github_pages_config.json').read_text())
        # Override gist_id for SAF if a separate config exists
        saf_cfg = SCRIPT_DIR / 'github_saf_config.json'
        if saf_cfg.exists():
            gist_cfg = json.loads(saf_cfg.read_text())
        path = find_excel()

    print(f'  ✓ Reading: {path.name}')
    df = read_applications(path)
    print(f'  ✓ {len(df)} applications loaded')
    print(f'  ✓ Channels: {df["Channel"].value_counts().to_dict()}')

    data = build_data(df)
    print(f'  ✓ Current month: {data["current_month"]} | '
          f'Sett: {data["current_total"]["sett_n"]} (${data["current_total"]["sett_val"]:,.0f})')

    gist_url = gist_cfg['raw_url']
    html = build_html(gist_url)

    try:
        push_gist(data, gist_cfg)
    except Exception as e:
        print(f'  ✗ Gist push failed: {e}')

    try:
        deploy_pages(html, pages_cfg)
    except Exception as e:
        print(f'  ✗ Pages deploy failed: {e}')

    if cloud and path.name.startswith('tmp'):
        path.unlink(missing_ok=True)

    print('  ✓ Done')

if __name__ == '__main__':
    main()
