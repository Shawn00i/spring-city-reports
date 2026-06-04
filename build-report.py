#!/usr/bin/env python3
"""
📊 Spring City Golf Monthly Sales Dashboard Generator
Reads Excel reports → generates a polished, interactive, shareable HTML dashboard.
Usage:  python3 build-report.py
"""

import os, openpyxl, json
from datetime import datetime

def n(v): return v if isinstance(v, (int, float)) else 0
def fmt(n): return f'¥{n:,.0f}'
DESKTOP = os.path.expanduser('/Users/shawn/Work/高尔夫/05_球会报告与报表/Monthly Report Gen')
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(WORKSPACE, 'monthly-report-dashboard.html')

def find_latest_file(pattern):
    import glob
    # Search in DESKTOP and any month subdirectories
    search_dirs = [DESKTOP] + [os.path.join(DESKTOP, d) for d in os.listdir(DESKTOP) if os.path.isdir(os.path.join(DESKTOP, d))]
    all_files = []
    for sd in search_dirs:
        for p in [pattern, pattern.replace('*', '*(1)')]:
            files = glob.glob(os.path.join(sd, p))
            all_files.extend(files)
    if all_files:
        return max(all_files, key=os.path.getmtime)
    print(f'  ⚠️  No file: {pattern}')
    return None

MON = datetime.now().strftime('%b %d, %Y')
month_key_map = {'jan':'Jan','feb':'Feb','mar':'Mar','apr':'Apr','may':'May','jun':'Jun','jul':'Jul','aug':'Aug','sep':'Sep','oct':'Oct','nov':'Nov','dec':'Dec'}
month_keys_ordered = list(month_key_map.keys())
month_titles_ordered = list(month_key_map.values())

print(f'📊 Generating sales dashboard — {MON}')
print('📁 Checking data files...')
for p in ['Key Accounts-0525-2026*','2025*高尔夫轮次收入对比*','Production Review-0525-2026*','Groups & Event OTB*']:
    f = find_latest_file(p)
    if f: print(f'  ✅ {os.path.basename(f)}')

ka_file = find_latest_file('Key Accounts-0525-2026*')
rounds_file = find_latest_file('2025*高尔夫轮次收入对比*')
prod_file = find_latest_file('Production Review-0525-2026*')
otb_file = find_latest_file('Groups & Event OTB*')

# ═══════════════════════════════════════════
# 1. KEY ACCOUNTS
# ═══════════════════════════════════════════
accounts = []
if ka_file:
    wb = openpyxl.load_workbook(ka_file, data_only=True)
    rows = list(wb['TOP A'].iter_rows(min_row=1, values_only=True))
    def is_company_name(name):
        if not name: return False
        name = str(name).strip()
        if len(name) < 4: return False
        if name == 'Year': return False
        if name[0].isdigit(): return False
        if ' VS ' in name.upper() or name.upper().startswith('VS') or 'vs2' in name.lower(): return False
        return True
    i = 0
    while i < len(rows):
        name = str(rows[i][1]).strip() if rows[i][1] else ''
        if is_company_name(name):
            d4 = d5 = d6 = None
            for j in range(i+1, min(i+12, len(rows))):
                yr = str(rows[j][1]).strip() if rows[j][1] else ''
                if yr == '2024' and not d4: d4 = rows[j]
                if yr == '2025' and not d5: d5 = rows[j]
                if yr == '2026' and not d6: d6 = rows[j]
            if d4 or d5 or d6:
                y26 = sum(n(d6[c]) for c in range(2,7)) if d6 else 0
                y25 = sum(n(d5[c]) for c in range(2,7)) if d5 else 0
                if y26 > 0 or y25 > 0:
                    accounts.append({'name':name,'ft':n(d4[14])if d4 else 0,'y25':y25,'y26':y26,'m26':[n(d6[c])if d6 else 0 for c in range(2,7)],'m25':[n(d5[c])if d5 else 0 for c in range(2,7)]})
        i += 1
    print(f'  Accounts: {len(accounts)}')

# ═══════════════════════════════════════════
# 2. REVENUE SEGMENTS
# ═══════════════════════════════════════════
segs = []
if rounds_file:
    wb2 = openpyxl.load_workbook(rounds_file, data_only=True)
    for row in wb2['2025VS2026'].iter_rows(min_row=3, values_only=True):
        s = str(row[1] or row[2] or row[3]).strip() if (row[1] or row[2] or row[3]) else ''
        if s and len(s)>2 and s not in ['Golf','Golf (Percentage)'] and 'Sub-Total' not in s and '#' not in s:
            a,b,c2,d2 = n(row[4]), n(row[7]), n(row[17]), n(row[19])
            if b>0 or d2>0:
                segs.append({'n': s, 'r25': a, 'v25': b, 'r26': c2, 'v26': d2})
    print(f'  Segments: {len(segs)}')

# ═══════════════════════════════════════════
# 3. SALES TEAM
# ═══════════════════════════════════════════
team = []
if prod_file:
    wb3 = openpyxl.load_workbook(prod_file, data_only=True)
    for row in wb3['Sales-2026'].iter_rows(min_row=3, values_only=True):
        ns = str(row[1]).strip() if row[1] else ''
        if ns and 'Total' not in ns and 'Dead' not in ns and 'New' not in ns and len(ns)<20:
            team.append({'n': ns, 'y26': n(row[2]), 'y25': n(row[3])})
    print(f'  Sellers: {len(team)}')

# ═══════════════════════════════════════════
# 4. GA REVENUE
# ═══════════════════════════════════════════
gd26 = gi26 = gd25 = gi25 = None
if prod_file:
    rows_g = list(wb3['2026-GA'].iter_rows(min_row=3, values_only=True))
    for rw in rows_g:
        cc = str(rw[7]).strip() if rw[7] else ''
        yy = str(rw[8]).strip() if rw[8] else ''
        def mk_ga(r): return {'jan':n(r[9]),'feb':n(r[10]),'mar':n(r[11]),'apr':n(r[12]),'may':n(r[13]),'ytd':n(r[14])} if len(r) > 14 else None
        if cc == 'Domestic' and yy == '2026': gd26 = mk_ga(rw)
        if cc == 'Domestic' and yy == '2025': gd25 = mk_ga(rw)
        if cc == 'International' and yy == '2026': gi26 = mk_ga(rw)
        if cc == 'International' and yy == '2025': gi25 = mk_ga(rw)

# ═══════════════════════════════════════════
# 5. ACCOUNTS REVIEW
# ═══════════════════════════════════════════
oc = nc2 = 0
if prod_file:
    for row in wb3['Accounts Review-2025'].iter_rows(min_row=4, values_only=True):
        ss = str(row[2]).strip() if row[2] else ''
        if ss == 'Old': oc += 1
        elif ss == 'New': nc2 += 1

# ═══════════════════════════════════════════
# 6. REGIONAL BREAKDOWN (exclude Excel "Total" row)
# ═══════════════════════════════════════════
regs_raw = {}
if prod_file:
    wsr = wb3['2026-YTD']
    for row in wsr.iter_rows(min_row=4, values_only=True):
        rr = str(row[1]).strip() if row[1] else ''
        if rr and len(rr) > 2:
            # Skip total/汇总 rows from Excel
            if rr.strip().lower() in ['total', 'totals', '合计', '汇总', '总计']:
                continue
            k = rr[:12]
            if k not in regs_raw:
                regs_raw[k] = {'label': rr, 'ytd26':0,'ytd25':0,'may26':0,'may25':0,'jan26':0,'jan25':0,'feb26':0,'feb25':0,'mar26':0,'mar25':0,'apr26':0,'apr25':0}
            rd = regs_raw[k]
            rd['ytd26'] += n(row[6]); rd['ytd25'] += n(row[9])
            rd['jan26'] += n(row[16]); rd['jan25'] += n(row[19])
            rd['feb26'] += n(row[23]); rd['feb25'] += n(row[26])
            rd['mar26'] += n(row[30]); rd['mar25'] += n(row[33])
            rd['apr26'] += n(row[37]); rd['apr25'] += n(row[40])
            rd['may26'] += n(row[48]); rd['may25'] += n(row[51])
    print(f'  Regions: {len(regs_raw)}')

# ═══════════════════════════════════════════
# 7. EVENT OTB
# ═══════════════════════════════════════════
otb_data = {}
if otb_file:
    try:
        wb_otb = openpyxl.load_workbook(otb_file, data_only=True)
        if 'Summary' in wb_otb.sheetnames:
            ws = wb_otb['Summary']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    label = str(row[0]).strip()
                    otb_data[label] = {
                        'total': n(row[1]) if len(row) > 1 else 0,
                        'confirmed': n(row[2]) if len(row) > 2 else 0,
                        'count_confirmed': int(n(row[3])) if len(row) > 3 else 0,
                        'count_lost': int(n(row[4])) if len(row) > 4 else 0,
                    }
    except Exception as e:
        print(f'  ⚠️  OTB read error: {e}')

# ═══════════════════════════════════════════
# CALCULATIONS
# ═══════════════════════════════════════════

# Detect how many months of data we actually have from the GA data
months_with_data = 0
if gd26:
    for mk in month_keys_ordered:
        if n(gd26.get(mk, 0)) > 0:
            months_with_data += 1
if months_with_data == 0:
    months_with_data = 5  # fallback to Jan-May
available_months = month_titles_ordered[:months_with_data]
ytd_label = f'{available_months[0]}–{available_months[-1]}'

ka26 = sum(a['y26'] for a in accounts)
ka25 = sum(a['y25'] for a in accounts)
ck = ((ka26/ka25)*100-100) if ka25 else 0
ck_s = f'+{ck:.1f}%' if ck >= 0 else f'{ck:.1f}%'

t3 = sorted(accounts, key=lambda a: a['y26'], reverse=True)[:3]

st26 = sum(s['y26'] for s in team)
st25 = sum(s['y25'] for s in team)
sc = ((st26/st25)*100-100) if st25 else 0
sc_s = f'+{sc:.1f}%' if sc >= 0 else f'{sc:.1f}%'

sv26 = sum(s['v26'] for s in segs)
sv25 = sum(s['v25'] for s in segs)
sr26 = sum(s['r26'] for s in segs)
sr25 = sum(s['r25'] for s in segs)
av26 = sv26/sr26 if sr26 else 0
av25 = sv25/sr25 if sr25 else 0

combined = ka26 + st26
combined_25 = ka25 + st25
combined_chg = ((combined/combined_25)*100-100) if combined_25 else 0

ts2 = sum(s['v26'] for s in segs)
ts3 = sorted(segs, key=lambda s: s['v26'], reverse=True)[:10]
cols = ['#d4af37','#2d6b4f','#1a4a38','#0f2e24','#4a8a6e','#b8960f','#6b4f2d','#8a6e4a','#3a7a5e','#a08040']
rs2 = sorted(regs_raw.items(), key=lambda x: x[1]['ytd26'], reverse=True)[:15]

# Monthly trend data for chart
monthly_trend = []
if gd26 and gi26:
    for mk, mt in month_key_map.items():
        dv = gd26.get(mk)
        iv = gi26.get(mk)
        if dv is not None and iv is not None and (dv > 0 or iv > 0):
            monthly_trend.append({'month': mt, 'domestic': dv, 'international': iv, 'total': dv + iv})

# Total for regions (from ALL regions, not just top 15)
mr26 = sum(d['may26'] for d in regs_raw.values())
mr25 = sum(d['may25'] for d in regs_raw.values())
yr26_total = sum(d['ytd26'] for d in regs_raw.values())
yr25_total = sum(d['ytd25'] for d in regs_raw.values())

# ═══════════════════════════════════════════
# EXECUTIVE SUMMARY
# ═══════════════════════════════════════════
def gen_summary():
    lines = []
    chg_dir = '增长' if combined_chg >= 0 else '下降'
    lines.append(f'截至{available_months[-1]}，2026年YTD总营收{fmt(combined)}，同比{chg_dir}{abs(combined_chg):.1f}%。')
    if accounts:
        ka_dir = '增长' if ck >= 0 else '下降'
        lines.append(f'大客户板块营收{fmt(ka26)}，同比{ka_dir}{abs(ck):.1f}%。')
        if t3:
            top = t3[0]
            top_chg = ((top['y26']/top['y25'])*100-100) if top['y25'] else None
            top_dir = '增长' if top_chg and top_chg >= 0 else '下降'
            top_chg_str = f'{abs(top_chg):.1f}%' if top_chg is not None else '新客户'
            lines.append(f'头部客户{top["name"][:10]}贡献{fmt(top["y26"])}，同比{top_dir}{top_chg_str}，占大客户板块{top["y26"]/ka26*100:.1f}%。')
    if team:
        team_dir = '增长' if sc >= 0 else '下降'
        best = max(team, key=lambda t: t['y26'])
        lines.append(f'销售团队YTD营收{fmt(st26)}，同比{team_dir}{abs(sc):.1f}%。{best["n"]}以{fmt(best["y26"])}领跑团队。')
    if segs and ts3:
        top_seg = ts3[0]
        seg_chg = ((top_seg['v26']/top_seg['v25'])*100-100) if top_seg['v25'] else 0
        seg_dir = '增长' if seg_chg >= 0 else '下降'
        lines.append(f'最大收入板块"{top_seg["n"]}"营收{fmt(top_seg["v26"])}，同比{seg_dir}{abs(seg_chg):.1f}%。')
    if gd26 and gi26:
        ds = gd26['ytd']/(gd26['ytd']+gi26['ytd'])*100
        lines.append(f'国内GA YTD营收{fmt(gd26["ytd"])}，国际GA{fmt(gi26["ytd"])}，国内占比{ds:.1f}%。')
    if regs_raw:
        reg_chg = ((yr26_total/yr25_total)*100-100) if yr25_total else 0
        reg_dir = '增长' if reg_chg >= 0 else '下降'
        lines.append(f'区域市场YTD同比{reg_dir}{abs(reg_chg):.1f}%。')
    return ''.join(lines)

exec_summary = gen_summary()

# ═══════════════════════════════════════════
# HTML GENERATION
# ═══════════════════════════════════════════
H = []

# ─── HEAD ───
H.append(f'''<!DOCTYPE html>
<html lang="zh-CN" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Spring City Sales Dashboard · {ytd_label} 2026</title>
<meta name="description" content="Spring City Golf Monthly Sales Dashboard — Auto-generated {MON}">
<meta property="og:title" content="Spring City Sales Dashboard">
<meta property="og:description" content="{ytd_label} 2026 · Data from monthly reports">
<meta property="og:type" content="website">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
:root {{
  --bg: #f4f3ef;
  --surface: #ffffff;
  --text: #1a1a18;
  --text2: #7a7a74;
  --text3: #a5a59e;
  --border: #e6e3db;
  --accent: #d4af37;
  --accent2: #b8960f;
  --green: #2d6b4f;
  --green-bg: #e8f3ed;
  --red: #b33a3a;
  --red-bg: #f7e8e8;
  --dark-bg: #0f2e24;
  --dark-text: #f5f0e8;
  --shadow: 0 1px 3px rgba(0,0,0,.04);
  --shadow-hover: 0 4px 12px rgba(0,0,0,.06);
  --radius: 10px;
  --chart-bg: #fff;
}}
[data-theme="dark"] {{
  --bg: #141412;
  --surface: #1e1e1b;
  --text: #e8e4dc;
  --text2: #99958c;
  --text3: #6b685e;
  --border: #2a2824;
  --accent: #d4af37;
  --accent2: #c4a030;
  --green: #4aaf7a;
  --green-bg: #1a2e24;
  --red: #e05555;
  --red-bg: #2e1a1a;
  --dark-bg: #1a1a16;
  --dark-text: #e8e4dc;
  --shadow: 0 1px 3px rgba(0,0,0,.2);
  --shadow-hover: 0 4px 12px rgba(0,0,0,.3);
  --chart-bg: #1e1e1b;
}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Inter','PingFang SC','Microsoft YaHei',sans-serif;background:var(--bg);color:var(--text);padding:24px 16px;transition:background .3s,color .3s}}
.c{{max-width:1160px;margin:0 auto}}
h1{{font-size:24px;font-weight:700;letter-spacing:-0.5px;line-height:1.2}}
h1 span{{color:var(--accent)}}
.st{{color:var(--text2);font-size:12px;margin-top:2px}}
.hdr{{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;margin-bottom:24px}}
.hdr-r{{display:flex;gap:8px;align-items:center;flex-wrap:wrap}}
.bd{{background:var(--dark-bg);color:var(--dark-text);padding:5px 12px;border-radius:6px;font-size:11px;font-weight:500;white-space:nowrap}}
.gr{{display:grid;gap:10px;margin-bottom:20px}}
.g4{{grid-template-columns:repeat(auto-fit,minmax(170px,1fr))}}
.g3{{grid-template-columns:repeat(auto-fit,minmax(195px,1fr))}}
.g2{{grid-template-columns:1fr 1fr}}
@media(max-width:760px){{.g2,.gf{{grid-template-columns:1fr}} .hdr{{flex-direction:column}}}}
.kc{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);transition:box-shadow .2s,transform .2s,background .3s,border .3s;animation:fU .4s backwards}}
.kc:hover{{box-shadow:var(--shadow-hover);transform:translateY(-2px)}}
.kt{{font-size:10px;font-weight:500;color:var(--text2);text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}}
.kv{{font-size:21px;font-weight:700;letter-spacing:-.3px}}
.kv-sm{{font-size:18px;font-weight:700;letter-spacing:-.2px}}
.kd{{font-size:11px;font-weight:500;margin-top:3px}}
.gn{{color:var(--green)}}.rd{{color:var(--red)}}
.sec{{margin-bottom:24px;animation:fU .4s backwards}}
.sh{{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap}}
.sh .l{{font-size:14px;font-weight:600}}
.sh .t{{font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;background:var(--accent);color:#0f2e24;padding:2px 8px;border-radius:4px}}
.ca{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;transition:background .3s,border .3s}}
.cb{{padding:12px 14px;transition:background .3s}}
.tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{text-align:left;padding:7px 9px;border-bottom:2px solid var(--border);font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);white-space:nowrap;cursor:pointer;user-select:none}}
th:hover{{color:var(--accent2)}}
th.srt::after{{content:" ▲";font-size:8px}}th.srt-d::after{{content:" ▼";font-size:8px}}
td{{padding:7px 9px;border-bottom:1px solid var(--border);white-space:nowrap;transition:background .3s}}
tr:last-child td{{border-bottom:none}}tr:hover td{{background:var(--bg)}}
.ar{{text-align:right;font-weight:500}}.pr{{text-align:right;font-size:11px;font-weight:500}}
.bar{{display:flex;height:6px;border-radius:3px;overflow:hidden;margin:4px 0}}
.spark{{display:flex;align-items:flex-end;gap:3px;height:28px;margin:4px 0}}
.sb-w{{width:10px;border-radius:2px 2px 0 0;min-height:3px;transition:background .3s}}
.prg{{height:5px;background:var(--border);border-radius:3px;overflow:hidden;margin:3px 0;transition:background .3s}}
.pf{{height:100%;border-radius:3px;min-width:4px}}
.ft{{text-align:center;padding:20px 0;font-size:11px;color:var(--text3);border-top:1px solid var(--border);margin-top:6px;transition:color .3s,border .3s}}
.lk{{display:inline-block;background:var(--accent);color:#0f2e24;padding:4px 14px;border-radius:6px;font-size:11px;font-weight:600;text-decoration:none;letter-spacing:.5px;transition:opacity .2s;white-space:nowrap}}
.lk:hover{{opacity:.8}}
.dmt{{background:none;border:1px solid var(--border);border-radius:6px;padding:5px 10px;cursor:pointer;font-size:13px;color:var(--text);transition:all .2s;background:var(--surface)}}
.dmt:hover{{border-color:var(--accent)}}
/* Summary box */
.sb-b{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 18px;box-shadow:var(--shadow);margin-bottom:20px;line-height:1.7;font-size:13.5px;color:var(--text);border-left:3px solid var(--accent);transition:background .3s,border .3s,color .3s}}
.sb-b strong{{color:var(--accent2)}}
/* Search */
.sbar{{border:1px solid var(--border);border-radius:6px;padding:5px 10px;font-size:12px;font-family:inherit;background:var(--surface);color:var(--text);width:200px;transition:border .3s;outline:none;margin-bottom:8px}}
.sbar:focus{{border-color:var(--accent)}}
/* Chart */
.ch{{padding:12px 14px;height:200px;position:relative}}
/* Print */
@media print{{body{{padding:0;background:#fff}} .dmt,.sbar{{display:none}} .kc{{break-inside:avoid;box-shadow:none;border:1px solid #ddd}} .ca{{box-shadow:none}} @page{{margin:1cm}}}}
@keyframes fU{{from{{opacity:0;transform:translateY(10px)}}to{{opacity:1;transform:translateY(0)}}}}
</style>
</head>
<body>
<div class="c">''')

# ─── HEADER ───
H.append(f'''<div class="hdr">
<div>
<h1><span>Spring City</span> Sales Dashboard</h1>
<div class="st">⛳ YTD {ytd_label} {datetime.now().year} · Generated {MON}</div>
</div>
<div class="hdr-r">
<span class="bd">📅 {MON}</span>
<button class="dmt" onclick="toggleTheme()" title="Dark/Light mode">🌓</button>
</div>
</div>''')

# ─── EXECUTIVE SUMMARY ───
H.append(f'<div class="sb-b">📋 <strong>月度简报</strong> · {exec_summary}</div>')

# ─── KEY METRICS ───
ka_dir_cls = 'gn' if ck >= 0 else 'rd'
sc_dir_cls = 'gn' if sc >= 0 else 'rd'
av_dir_cls = 'gn' if av26 >= av25 else 'rd'
comb_dir_cls = 'gn' if combined_chg >= 0 else 'rd'
av_arrow = '↑' if av26 >= av25 else '↓'

H.append(f'''<div class="gr g4">
<div class="kc"><div class="kt">Combined Revenue YTD</div><div class="kv">{fmt(combined)}</div><div class="kd {comb_dir_cls}">{("+" if combined_chg>=0 else "")}{combined_chg:.1f}% vs 2025</div></div>
<div class="kc"><div class="kt">Key Accounts YTD</div><div class="kv">{fmt(ka26)}</div><div class="kd {ka_dir_cls}">{ck_s}</div></div>
<div class="kc"><div class="kt">Sales Team YTD</div><div class="kv">{fmt(st26)}</div><div class="kd {sc_dir_cls}">{sc_s}</div></div>
<div class="kc"><div class="kt">Avg Revenue / Round</div><div class="kv kv-sm">{fmt(av26)}</div><div class="kd {av_dir_cls}">{av_arrow} vs 2025 ({fmt(av25)})</div></div>
</div>''')

# ─── TOP 3 ACCOUNTS ───
H.append('<div class="gr g3">')
for i, a in enumerate(t3):
    g = ((a['y26']/a['y25'])*100-100) if a['y25'] else None
    gs = f'+{g:.1f}%' if g is not None and g >= 0 else (f'{g:.1f}%' if g is not None else 'NEW')
    dd = 'gn' if g is not None and g >= 0 else ('rd' if g is not None else 'gn')
    mx = max(a['m26']) if max(a['m26']) else 1
    colors = ['#d4af37','#2d6b4f','#0f2e24']
    bs = ''.join(f'<div class="sb-w" style="height:{max(8,(m/mx)*24)}px;background:{colors[i]}"></div>' for m in a['m26'])
    H.append(f'''<div class="kc" style="animation-delay:{.25+i*.05}s">
<div class="kt">#{i+1} {a["name"][:18]}</div>
<div class="kv kv-sm">{fmt(a["y26"])}</div>
<div class="kd {dd}">{gs}</div>
<div style="font-size:9px;color:var(--text3);margin-top:4px">{' · '.join(available_months[:5])}</div>
<div class="spark">{bs}</div>
</div>''')
H.append('</div>')

# ─── KEY ACCOUNTS TABLE ───
ac = len(accounts)
H.append(f'''<div class="sec">
<div class="sh"><span class="l">🏆 Key Accounts</span><span class="t">{ac} accounts · YTD {ytd_label}</span><input class="sbar" type="text" placeholder="🔍 Filter accounts..." oninput="filterTable(this,'tbl-ka')" style="margin-left:auto"></div>
<div class="ca"><div class="tw">
<table id="tbl-ka"><thead><tr>
<th onclick="sortTable('tbl-ka',0)">Account</th>
<th class="ar" onclick="sortTable('tbl-ka',1)">2024 Full</th>
<th class="ar" onclick="sortTable('tbl-ka',2)">2025 YTD</th>
<th class="ar" onclick="sortTable('tbl-ka',3)">2026 YTD</th>
<th class="pr" onclick="sortTable('tbl-ka',4)" title="Year-over-year change">Change</th>
<th class="pr" onclick="sortTable('tbl-ka',5)">Share</th>
</tr></thead><tbody>''')
for a in accounts:
    g = ((a['y26']/a['y25'])*100-100) if a['y25'] else None
    if g is None:
        gs, dd = '—', 'gn'  # new account with no prior year
    else:
        gs = f'+{g:.1f}%' if g >= 0 else f'{g:.1f}%'
        dd = 'gn' if g >= 0 else 'rd'
    sh = a['y26']/ka26*100 if ka26 else 0
    H.append(f'<tr><td>{a["name"][:35]}</td><td class="ar">{fmt(a["ft"])}</td><td class="ar">{fmt(a["y25"])}</td><td class="ar">{fmt(a["y26"])}</td><td class="pr {dd}">{gs}</td><td class="pr">{sh:.1f}%</td></tr>')
H.append('</tbody></table></div></div></div>')

# ─── REVENUE SEGMENTS + SALES TEAM ───
H.append('<div class="gr g2">')

# Revenue Segments
H.append(f'''<div class="sec">
<div class="sh"><span class="l">⛳ Revenue Segments</span><span class="t">Top 10</span></div>
<div class="ca"><div class="cb">
<div class="bar">''')
for i, s in enumerate(ts3):
    p = s['v26']/ts2*100 if ts2 else 0
    H.append(f'<div style="width:{p}%;background:{cols[i%len(cols)]};height:6px" title="{s["n"][:20]}: {p:.1f}%"></div>')
H.append('</div><table id="tbl-seg"><thead><tr><th>Segment</th><th class="ar">2026 YTD</th><th class="pr">Share</th><th class="pr">Change</th></tr></thead><tbody>')
for i, s in enumerate(ts3):
    c = ((s['v26']/s['v25'])*100-100) if s['v25'] else 0
    cs = f'+{c:.1f}%' if c >= 0 else f'{c:.1f}%'
    dd = 'gn' if c >= 0 else 'rd'
    sh = s['v26']/ts2*100 if ts2 else 0
    H.append(f'<tr><td><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:{cols[i%len(cols)]};margin-right:5px"></span>{s["n"][:30]}</td><td class="ar">{fmt(s["v26"])}</td><td class="pr">{sh:.1f}%</td><td class="pr {dd}">{cs}</td></tr>')
H.append('</tbody></table></div></div></div>')

# Sales Team
H.append(f'''<div class="sec">
<div class="sh"><span class="l">👥 Sales Team</span><span class="t">YTD · {len(team)} sellers</span></div>
<div class="ca"><div class="cb">''')
st_max = max(s['y26'] for s in team) if team else 1
for s in team:
    gs = ((s['y26']/s['y25'])*100-100) if s['y25'] else None
    if gs is None:
        gss, dd = '—', 'gn'
    else:
        gss = f'+{gs:.1f}%' if gs >= 0 else f'{gs:.1f}%'
        dd = 'gn' if gs >= 0 else 'rd'
    sh = s['y26']/st26*100 if st26 else 0
    bw = max(6, (s['y26']/st_max)*100) if st_max else 6
    H.append(f'''<div style="margin-bottom:9px">
<div style="display:flex;justify-content:space-between;font-size:12px">
<span style="font-weight:500">{s["n"]}</span>
<span style="font-weight:600">{fmt(s["y26"])}</span>
</div>
<div class="prg"><div class="pf" style="width:{bw}%;background:#d4af37"></div></div>
<div style="display:flex;justify-content:space-between;font-size:10px;color:var(--text2)">
<span>{sh:.1f}%</span>
<span class="{dd}">{gss}</span>
</div>
</div>''')

# Shawn section
for s in team:
    if s['n'] == 'Shawn':
        hp = s['y25'] and s['y25'] > 0
        sz = f'+{int((s["y26"]/s["y25"]-1)*100)}%' if hp else 'New accounts / N/A'
        d2 = 'gn' if hp else 'rd'
        H.append(f'''<div style="border-top:1px solid var(--border);padding-top:9px;margin-top:9px">
<div style="font-size:10px;font-weight:500;color:var(--text2);margin-bottom:2px">Your Performance</div>
<div class="kv kv-sm">{fmt(s["y26"])}</div>
<div class="kd {d2}">{sz}</div>
<div style="font-size:11px;color:var(--text2);margin-top:3px">New accounts: {nc2} · Total: {oc+nc2}</div>
</div>''')
H.append('</div></div></div></div>')

# ─── GA REVENUE TREND (section containing both chart and table) ───
if gd26 or gi26:
    H.append(f'<div class="sec"><div class="sh"><span class="l">📈 GA Revenue Trend</span><span class="t">Monthly · {ytd_label}</span></div>')
    # KPI cards
    H.append('<div class="gr g3">')
    if gd26:
        H.append(f'<div class="kc"><div class="kt">Domestic GA YTD</div><div class="kv kv-sm">{fmt(gd26["ytd"])}</div></div>')
    if gi26:
        H.append(f'<div class="kc"><div class="kt">International GA YTD</div><div class="kv kv-sm">{fmt(gi26["ytd"])}</div></div>')
    if gd26 and gi26:
        ds = gd26['ytd']/(gd26['ytd']+gi26['ytd'])*100
        H.append(f'<div class="kc"><div class="kt">Domestic Share</div><div class="kv kv-sm">{ds:.1f}%</div></div>')
    H.append('</div>')
    # Chart
    H.append(f'<div class="ca"><div class="ch"><canvas id="gaChart"></canvas></div></div>')
    # Table (inside same sec)
    H.append('<div class="ca" style="margin-top:10px"><div class="tw"><table><thead><tr><th>Month</th><th class="ar">Domestic 2026</th>')
    if gd25:
        H.append('<th class="ar">Domestic 2025</th><th class="pr">YoY</th>')
    H.append('<th class="ar">International 2026</th><th class="ar">Total</th></tr></thead><tbody>')
    active_months_displayed = 0
    for mk, mn in month_key_map.items():
        d_26 = n(gd26.get(mk, 0)) if gd26 else 0
        d_25 = n(gd25.get(mk, 0)) if gd25 else 0
        i_26 = n(gi26.get(mk, 0)) if gi26 else 0
        if d_26 == 0 and i_26 == 0 and active_months_displayed >= months_with_data:
            continue  # skip empty months
        if d_26 == 0 and i_26 == 0:
            continue  # skip months with no data at all
        active_months_displayed += 1
        tt = d_26 + i_26
        H.append(f'<tr><td>{mn}</td><td class="ar">{fmt(d_26)}</td>')
        if gd25:
            cd = ((d_26/d_25)*100-100) if d_25 else 0
            dd = 'gn' if cd >= 0 else 'rd'
            H.append(f'<td class="ar">{fmt(d_25)}</td><td class="pr {dd}">{"↑" if cd>=0 else "↓"} {abs(cd):.1f}%</td>')
        H.append(f'<td class="ar">{fmt(i_26)}</td><td class="ar">{fmt(tt)}</td></tr>')
    H.append('</tbody></table></div></div></div>')

# ─── REGIONAL BREAKDOWN ───
H.append(f'''<div class="sec">
<div class="sh"><span class="l">🌍 Regional Breakdown</span><span class="t">Top 15 · May + YTD</span><input class="sbar" type="text" placeholder="🔍 Filter..." oninput="filterTable(this,'tbl-reg')" style="margin-left:auto"></div>
<div class="ca"><div class="tw">
<table id="tbl-reg"><thead><tr>
<th onclick="sortTable('tbl-reg',0)">Region</th>
<th class="ar" onclick="sortTable('tbl-reg',1)">May 2026</th>
<th class="ar" onclick="sortTable('tbl-reg',2)">May 2025</th>
<th class="pr" onclick="sortTable('tbl-reg',3)">Chg</th>
<th class="ar" onclick="sortTable('tbl-reg',4)">YTD 2026</th>
<th class="ar" onclick="sortTable('tbl-reg',5)">YTD 2025</th>
<th class="pr" onclick="sortTable('tbl-reg',6)">Chg</th>
</tr></thead><tbody>''')
for reg, d in rs2:
    mc = ((d['may26']/max(d['may25'],1))*100-100) if d['may25'] else 0
    yc = ((d['ytd26']/max(d['ytd25'],1))*100-100) if d['ytd25'] else 0
    md = 'gn' if mc >= 0 else 'rd'
    yd = 'gn' if yc >= 0 else 'rd'
    H.append(f'<tr><td>{reg}</td><td class="ar">{fmt(d["may26"])}</td><td class="ar">{fmt(d["may25"])}</td><td class="pr {md}">{"↑" if mc>=0 else "↓"} {abs(mc):.1f}%</td><td class="ar">{fmt(d["ytd26"])}</td><td class="ar">{fmt(d["ytd25"])}</td><td class="pr {yd}">{"↑" if yc>=0 else "↓"} {abs(yc):.1f}%</td></tr>')
tmc = ((mr26/mr25)*100-100) if mr25 else 0
tyc = ((yr26_total/yr25_total)*100-100) if yr25_total else 0
tmd = 'gn' if tmc >= 0 else 'rd'
tyd = 'gn' if tyc >= 0 else 'rd'
H.append(f'<tr style="font-weight:600;border-top:2px solid var(--border)"><td>Total</td><td class="ar">{fmt(mr26)}</td><td class="ar">{fmt(mr25)}</td><td class="pr {tmd}">{"↑" if tmc>=0 else "↓"} {abs(tmc):.1f}%</td><td class="ar">{fmt(yr26_total)}</td><td class="ar">{fmt(yr25_total)}</td><td class="pr {tyd}">{"↑" if tyc>=0 else "↓"} {abs(tyc):.1f}%</td></tr>')
H.append('</tbody></table></div></div></div>')

# ─── FOOTER ───
H.append(f'''<div class="ft">
<div>Spring City Golf · YTD {ytd_label} {datetime.now().year} · {MON}</div>
<div style="margin-top:4px;font-size:10px">Auto-generated from monthly Excel reports</div>
</div>
</div>

<script>
function toggleTheme() {{
  const html = document.documentElement;
  html.dataset.theme = html.dataset.theme === 'dark' ? 'light' : 'dark';
  localStorage.setItem('sc-theme', html.dataset.theme);
}}
if (localStorage.getItem('sc-theme') === 'dark') document.documentElement.dataset.theme = 'dark';

function sortTable(tblId, col) {{
  const table = document.getElementById(tblId);
  if (!table) return;
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const header = table.querySelectorAll('th')[col];
  const isNum = header.classList.contains('ar') || header.classList.contains('pr');
  const asc = !header.classList.contains('srt');
  table.querySelectorAll('th').forEach(th => {{ th.classList.remove('srt','srt-d'); }});
  rows.sort((a, b) => {{
    let va = a.cells[col].textContent.trim().replace(/[¥,↑↓+\s]/g,'') || '0';
    let vb = b.cells[col].textContent.trim().replace(/[¥,↑↓+\s]/g,'') || '0';
    if (isNum) {{
      return asc ? (parseFloat(va)||0) - (parseFloat(vb)||0) : (parseFloat(vb)||0) - (parseFloat(va)||0);
    }}
    return asc ? va.localeCompare(vb) : vb.localeCompare(va);
  }});
  rows.forEach(r => tbody.appendChild(r));
  header.classList.add(asc ? 'srt' : 'srt-d');
}}

function filterTable(input, tblId) {{
  const q = input.value.toLowerCase();
  const table = document.getElementById(tblId);
  if (!table) return;
  table.querySelectorAll('tbody tr').forEach(r => {{
    r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none';
  }});
}}

const gaCanvas = document.getElementById('gaChart');
if (gaCanvas) {{
  const months = {json.dumps([m['month'] for m in monthly_trend])};
  const domestic = {json.dumps([m['domestic'] for m in monthly_trend])};
  const international = {json.dumps([m['international'] for m in monthly_trend])};
  new Chart(gaCanvas, {{
    type: 'bar',
    data: {{
      labels: months,
      datasets: [
        {{ label: 'Domestic', data: domestic, backgroundColor: '#2d6b4f', borderRadius: 4 }},
        {{ label: 'International', data: international, backgroundColor: '#d4af37', borderRadius: 4 }}
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: 'top', labels: {{ font: {{ size: 11 }}, usePointStyle: true }} }},
        tooltip: {{ callbacks: {{ label: ctx => '¥' + ctx.parsed.y.toLocaleString() }} }}
      }},
      scales: {{
        x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }},
        y: {{ beginAtZero: true, ticks: {{ callback: v => '¥' + (v/10000).toFixed(0) + 'w', font: {{ size: 10 }} }} }}
      }}
    }}
  }});
}}
</script>
</body>
</html>''')

# ─── WRITE ───
content = '\n'.join(H)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(content)

size_kb = os.path.getsize(OUT) / 1024
print(f'\n✅ Dashboard: {OUT}')
print(f'   Size: {size_kb:.0f} KB · {len(accounts)} accounts · {len(segs)} segments · {len(team)} sellers · {len(regs_raw)} regions')
