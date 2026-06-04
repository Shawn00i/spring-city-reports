#!/usr/bin/env python3
"""
📊 Spring City Golf Sales Dashboard v2
Primary source: Rolling Forecast (budget vs forecast by segment)
Supplemented by: Key Accounts, Sales Team, Production Review

Usage:  python3 build-report.py
"""

import os, openpyxl, json
from datetime import datetime

def n(v): return v if isinstance(v, (int, float)) else 0
def fmt(n): return f'¥{n:,.0f}'

DATA_DIR = os.path.expanduser('~/Desktop/Monthly Report Gen')
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(WORKSPACE, 'monthly-report-dashboard.html')

def find_file(pattern, subdirs=True):
    import glob
    dirs = [DATA_DIR]
    if subdirs:
        dirs += [os.path.join(DATA_DIR, d) for d in os.listdir(DATA_DIR) if os.path.isdir(os.path.join(DATA_DIR, d))]
    for d in dirs:
        for p in [pattern, pattern.replace('*', '*(1)')]:
            files = glob.glob(os.path.join(d, p))
            if files: return max(files, key=os.path.getmtime)
    return None

MON = datetime.now().strftime('%b %d, %Y')
print(f'📊 v2 Dashboard — {MON}')

# ─── Locate data files ───
files = {
    'forecast': find_file('2026 Rolling Forecast*'),
    'accounts': find_file('Key Accounts-0525-2026*'),
    'rounds': find_file('2025*高尔夫轮次收入对比*'),
    'prod': find_file('Production Review-0525-2026*'),
}
for k, v in files.items():
    if v: print(f'  ✅ {k}: {os.path.basename(v)}')

# ═══════════════════════════════════════════
# 1. ROLLING FORECAST — new primary data
# ═══════════════════════════════════════════
months_lbl = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
categories = []  # {sn, name, rep, aob, rolling, yoy, monthly_aob: [...], monthly_rolling: [...], ytd_actual}
month_col_start = 7  # Col 7 = Jan data

if files['forecast']:
    wb = openpyxl.load_workbook(files['forecast'], data_only=True)
    ws = wb['2026 Rolling Forecast']

    current_sn = None
    current_name = ''
    current_rep = ''

    for row in ws.iter_rows(min_row=3, max_row=379, values_only=True):
        sn = row[0]
        name = str(row[1] or '').strip()
        rep = str(row[3] or '').strip() if row[3] else ''
        rtype = str(row[4] or '').strip()

        # New category starts when sn has a number in this column
        if sn is not None and isinstance(sn, (int, float)):
            current_sn = int(sn)
            current_name = name
            current_rep = rep

        if current_sn is None:
            continue

        def read_monthly(r):
            return [n(r[month_col_start + j]) if month_col_start + j < len(r) else 0 for j in range(12)]

        if rtype == '2025 Actual' and sn is not None and isinstance(sn, (int, float)):
            cat = {'sn': current_sn, 'name': current_name, 'rep': current_rep,
                   'aob': 0, 'rolling': 0, 'y25_actual': n(row[5]), 'yoy': 0,
                   'monthly_aob': [0]*12, 'monthly_rolling': [0]*12, 'monthly_y25': read_monthly(row) if row[5] else [0]*12}
            categories.append(cat)

        elif rtype == '2026 AOB' and categories:
            categories[-1]['aob'] = n(row[5])
            if n(row[5]):
                categories[-1]['monthly_aob'] = read_monthly(row)

        elif rtype == 'Rolling Forecast' and categories:
            categories[-1]['rolling'] = n(row[5])
            if n(row[5]):
                categories[-1]['monthly_rolling'] = read_monthly(row)

        elif rtype == 'yoy' and categories:
            categories[-1]['yoy'] = n(row[5])

# Filter to real categories (exclude subtotals, totals, grand totals)
real_cats = [c for c in categories if c['name'] and 'Sub-Total' not in c['name']
             and 'Total' not in c['name'] and 'Grand' not in c['name']
             and c['sn'] not in [24, 42]]  # skip empty-name/zero rows

print(f'  Categories: {len(real_cats)}')

# Compute totals
total_rolling = sum(c['rolling'] for c in real_cats)
total_aob = sum(c['aob'] for c in real_cats)
total_y25 = sum(c['y25_actual'] for c in real_cats)
total_var = total_rolling - total_aob
total_var_pct = (total_var / total_aob * 100) if total_aob else 0
total_yoy_pct = ((total_rolling / total_y25) * 100 - 100) if total_y25 else 0

# Group by responsible person
people = {}
for c in real_cats:
    rep = c['rep'] or 'Other'
    if rep not in people: people[rep] = {'rolling': 0, 'aob': 0, 'y25': 0, 'cats': []}
    people[rep]['rolling'] += c['rolling']
    people[rep]['aob'] += c['aob']
    people[rep]['y25'] += c['y25_actual']
    people[rep]['cats'].append(c)

# Sort people by rolling revenue
people_sorted = sorted(people.items(), key=lambda x: x[1]['rolling'], reverse=True)
print(f'  People: {len(people_sorted)}')

# YTD actual from rolling forecast (Jan-May actual months)
ytd_months = 5  # Jan-May available
monthly_rolling_ytd = {}
for c in real_cats:
    for j in range(min(ytd_months, len(c.get('monthly_rolling', [])))):
        if c['monthly_rolling'] and j < len(c['monthly_rolling']):
            monthly_rolling_ytd[j] = monthly_rolling_ytd.get(j, 0) + c['monthly_rolling'][j]
ytd_rolling = sum(v for v in monthly_rolling_ytd.values())
monthly_y25_ytd = {}
for c in real_cats:
    for j in range(min(ytd_months, len(c.get('monthly_y25', [])))):
        if c['monthly_y25'] and j < len(c['monthly_y25']):
            monthly_y25_ytd[j] = monthly_y25_ytd.get(j, 0) + c['monthly_y25'][j]
ytd_y25 = sum(v for v in monthly_y25_ytd.values())
ytd_yoy = ((ytd_rolling / ytd_y25) * 100 - 100) if ytd_y25 else 0

top_cats = sorted(real_cats, key=lambda c: c['rolling'], reverse=True)[:15]

# ═══════════════════════════════════════════
# 2. KEY ACCOUNTS (from old file)
# ═══════════════════════════════════════════
accounts = []
if files['accounts']:
    wb2 = openpyxl.load_workbook(files['accounts'], data_only=True)
    rows = list(wb2['TOP A'].iter_rows(min_row=1, values_only=True))
    def is_company_name(name):
        if not name: return False
        name = str(name).strip()
        if len(name) < 4: return False; 
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

ka26 = sum(a['y26'] for a in accounts)
ka25 = sum(a['y25'] for a in accounts)
ka_chg = ((ka26/ka25)*100-100) if ka25 else 0

# ═══════════════════════════════════════════
# 3. SALES TEAM (from Production Review)
# ═══════════════════════════════════════════
team = []
ga_dom = ga_int = None
if files['prod']:
    wb3 = openpyxl.load_workbook(files['prod'], data_only=True)
    for row in wb3['Sales-2026'].iter_rows(min_row=3, values_only=True):
        ns = str(row[1]).strip() if row[1] else ''
        if ns and 'Total' not in ns and 'Dead' not in ns and 'New' not in ns and len(ns)<20:
            team.append({'n': ns, 'y26': n(row[2]), 'y25': n(row[3])})
    # GA from production review
    rows_g = list(wb3['2026-GA'].iter_rows(min_row=3, values_only=True))
    for rw in rows_g:
        cc = str(rw[7]).strip() if rw[7] else ''
        yy = str(rw[8]).strip() if rw[8] else ''
        def mk_ga(r): return {months_lbl[j].lower(): float(n(r[9+j])) if 9+j < len(r) else 0 for j in range(5)} if len(r) > 14 else None
        if cc == 'Domestic' and yy == '2026': ga_dom = mk_ga(rw)
        if cc == 'International' and yy == '2026': ga_int = mk_ga(rw)

st26 = sum(s['y26'] for s in team)
st25 = sum(s['y25'] for s in team)
st_chg = ((st26/st25)*100-100) if st25 else 0

# ═══════════════════════════════════════════
# 4. EXECUTIVE SUMMARY
# ═══════════════════════════════════════════
def gen_summary():
    lines = []
    lines.append(f'2026年滚动预测（Rolling Forecast）总营收{fmt(total_rolling)}，相比预算（AOB {fmt(total_aob)}）{"增长" if total_var >= 0 else "减少"}{abs(total_var_pct):.1f}%。')
    lines.append(f'YTD（Jan-May）实际预测营收约{fmt(ytd_rolling)}，同比2025年同期{"增长" if ytd_yoy >= 0 else "下降"}{abs(ytd_yoy):.1f}%。')
    # Top person
    if people_sorted:
        top_p = people_sorted[0]
        lines.append(f'按负责人划分，{top_p[0]}的板块预测营收最高，为{fmt(top_p[1]["rolling"])}。')
    # Top category
    if top_cats:
        tc = top_cats[0]
        var = tc['rolling'] - tc['aob']
        lines.append(f'最大收入板块"{tc["name"]}"滚动预测{fmt(tc["rolling"])}，{"超预算" if var >= 0 else "低于预算"}{fmt(abs(var))}。')
    # Key accounts
    ka_dir = '增长' if ka_chg >= 0 else '下降'
    lines.append(f'大客户板块YTD营收{fmt(ka26)}，同比{ka_dir}{abs(ka_chg):.1f}%。')
    # Sales team
    st_dir = '增长' if st_chg >= 0 else '下降'
    lines.append(f'销售团队YTD营收{fmt(st26)}，同比{st_dir}{abs(st_chg):.1f}%。')
    return ''.join(lines)

exec_summary = gen_summary()

# ═══════════════════════════════════════════
# HTML GENERATION
# ═══════════════════════════════════════════
H = []
H.append(f'''<!DOCTYPE html>
<html lang="zh-CN" data-theme="light">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Spring City Sales Dashboard · {MON}</title>
<meta property="og:title" content="Spring City Sales Dashboard">
<meta property="og:description" content="Rolling Forecast {MON}">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
:root {{--bg:#f4f3ef;--surface:#fff;--text:#1a1a18;--text2:#7a7a74;--text3:#a5a59e;--border:#e6e3db;--accent:#d4af37;--accent2:#b8960f;--green:#2d6b4f;--green-bg:#e8f3ed;--red:#b33a3a;--red-bg:#f7e8e8;--dark:#0f2e24;--dt:#f5f0e8;--shadow:0 1px 3px rgba(0,0,0,.04);--sh-h:0 4px 12px rgba(0,0,0,.06);--r:10px}}
[data-theme="dark"] {{--bg:#141412;--surface:#1e1e1b;--text:#e8e4dc;--text2:#99958c;--text3:#6b685e;--border:#2a2824;--green:#4aaf7a;--red:#e05555}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Inter','PingFang SC','Microsoft YaHei',sans-serif;background:var(--bg);color:var(--text);padding:24px 16px;transition:background .3s,color .3s}}
.c{{max-width:1160px;margin:0 auto}}
h1{{font-size:24px;font-weight:700;letter-spacing:-0.5px}} h1 span{{color:var(--accent)}}
.st{{color:var(--text2);font-size:12px;margin-top:2px}}
.hdr{{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;margin-bottom:24px}}
.hdr-r{{display:flex;gap:8px;align-items:center;flex-wrap:wrap}}
.bd{{background:var(--dark);color:var(--dt);padding:5px 12px;border-radius:6px;font-size:11px;font-weight:500}}
.gr{{display:grid;gap:10px;margin-bottom:20px}}
.g4{{grid-template-columns:repeat(auto-fit,minmax(170px,1fr))}}
.g3{{grid-template-columns:repeat(auto-fit,minmax(195px,1fr))}}
.g2{{grid-template-columns:1fr 1fr}}
@media(max-width:760px){{.g2{{grid-template-columns:1fr}}.hdr{{flex-direction:column}}}}
.kc{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:14px 16px;box-shadow:var(--shadow);transition:box-shadow .2s,transform .2s,background .3s;animation:fU .4s backwards}}
.kc:hover{{box-shadow:var(--sh-h);transform:translateY(-2px)}}
.kt{{font-size:10px;font-weight:500;color:var(--text2);text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}}
.kv{{font-size:21px;font-weight:700;letter-spacing:-.3px}}
.kv-sm{{font-size:18px;font-weight:700;letter-spacing:-.2px}}
.kd{{font-size:11px;font-weight:500;margin-top:3px}}
.gn{{color:var(--green)}}.rd{{color:var(--red)}}
.sec{{margin-bottom:24px;animation:fU .4s backwards}}
.sh{{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap}}
.sh .l{{font-size:14px;font-weight:600}} .sh .t{{font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;background:var(--accent);color:#0f2e24;padding:2px 8px;border-radius:4px}}
.ca{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--shadow);overflow:hidden;transition:background .3s}}
.cb{{padding:12px 14px}} .tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{text-align:left;padding:7px 9px;border-bottom:2px solid var(--border);font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);cursor:pointer;user-select:none}}
th:hover{{color:var(--accent2)}} th.srt::after{{content:" \\u25b2";font-size:8px}}th.srt-d::after{{content:" \\u25bc";font-size:8px}}
td{{padding:7px 9px;border-bottom:1px solid var(--border);white-space:nowrap;transition:background .3s}}
tr:last-child td{{border-bottom:none}}tr:hover td{{background:var(--bg)}}
.ar{{text-align:right;font-weight:500}}.pr{{text-align:right;font-size:11px;font-weight:500}}
.prg{{height:5px;background:var(--border);border-radius:3px;overflow:hidden;margin:3px 0;transition:background .3s}}
.pf{{height:100%;border-radius:3px;min-width:4px}}
.ft{{text-align:center;padding:20px 0;font-size:11px;color:var(--text3);border-top:1px solid var(--border)}}
.dmt{{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:5px 10px;cursor:pointer;font-size:13px;color:var(--text);transition:all .2s}}
.dmt:hover{{border-color:var(--accent)}}
.sb-b{{background:var(--surface);border-left:3px solid var(--accent);border-radius:var(--r);padding:14px 18px;box-shadow:var(--shadow);margin-bottom:20px;line-height:1.7;font-size:13.5px;color:var(--text);transition:background .3s,color .3s}}
.sb-b strong{{color:var(--accent2)}}
.sbar{{border:1px solid var(--border);border-radius:6px;padding:5px 10px;font-size:12px;font-family:inherit;background:var(--surface);color:var(--text);width:200px;outline:none;margin-bottom:8px;transition:border .3s}}
.sbar:focus{{border-color:var(--accent)}}
.ch{{padding:12px 14px;height:220px;position:relative}}
@media print{{body{{padding:0;background:#fff}}.dmt,.sbar{{display:none}}.kc{{break-inside:avoid;box-shadow:none;border:1px solid #ddd}}.ca{{box-shadow:none}}.ch{{height:250px}}@page{{margin:1cm}}}}
@keyframes fU{{from{{opacity:0;transform:translateY(10px)}}to{{opacity:1;transform:translateY(0)}}}}
</style></head><body><div class="c">''')

# ─── HEADER ───
H.append(f'''<div class="hdr"><div><h1><span>Spring City</span> Sales Dashboard</h1><div class="st">Rolling Forecast · {MON}</div></div>
<div class="hdr-r"><span class="bd">📅 {MON}</span><button class="dmt" onclick="toggleTheme()">🌓</button></div></div>''')

# ─── EXECUTIVE SUMMARY ───
H.append(f'<div class="sb-b">📋 <strong>月度简报</strong> · {exec_summary}</div>')

# ─── KPI CARDS ───
var_cls = 'gn' if total_var >= 0 else 'rd'
ytd_cls = 'gn' if ytd_yoy >= 0 else 'rd'
ka_cls = 'gn' if ka_chg >= 0 else 'rd'
st_cls = 'gn' if st_chg >= 0 else 'rd'
H.append(f'''<div class="gr g4">
<div class="kc"><div class="kt">Rolling Forecast FY 2026</div><div class="kv">{fmt(total_rolling)}</div><div class="kd {var_cls}">vs AOB {fmt(total_aob)} · {("+" if total_var_pct>=0 else "")}{total_var_pct:.1f}%</div></div>
<div class="kc"><div class="kt">YTD Rolling (Jan–May)</div><div class="kv">{fmt(ytd_rolling)}</div><div class="kd {ytd_cls}">{"+" if ytd_yoy>=0 else ""}{ytd_yoy:.1f}% vs 2025 YTD</div></div>
<div class="kc"><div class="kt">Key Accounts YTD</div><div class="kv">{fmt(ka26)}</div><div class="kd {ka_cls}">{"+" if ka_chg>=0 else ""}{ka_chg:.1f}% vs 2025</div></div>
<div class="kc"><div class="kt">Sales Team YTD</div><div class="kv">{fmt(st26)}</div><div class="kd {st_cls}">{"+" if st_chg>=0 else ""}{st_chg:.1f}% vs 2025</div></div>
</div>''')

# ─── BY RESPONSIBLE PERSON ───
person_colors = ['#d4af37','#2d6b4f','#1a4a38','#0f2e24','#4a8a6e']
H.append('<div class="gr g3">')
for pi, (name, pd) in enumerate(people_sorted):
    var = pd['rolling'] - pd['aob']
    vcls = 'gn' if var >= 0 else 'rd'
    if pd['y25']:
        yoy = ((pd['rolling']/pd['y25'])*100-100)
        ycls = 'gn' if yoy >= 0 else 'rd'
        yoy_s = f'{"+" if yoy>=0 else ""}{yoy:.1f}% vs 2025'
    else:
        ycls = 'gn'
        yoy_s = '—'
    cat_count = len(pd['cats'])
    H.append(f'<div class="kc" style="animation-delay:{.2+pi*.08}s;border-top:3px solid {person_colors[pi%len(person_colors)]}"><div class="kt">{name}</div><div class="kv kv-sm">{fmt(pd["rolling"])}</div><div class="kd {vcls}">{"+" if var>=0 else ""}{fmt(abs(var))} vs AOB · {cat_count} segments</div><div class="kd {ycls}" style="margin-top:2px">{yoy_s}</div></div>')
H.append('</div>')

# ─── MONTHLY TREND CHART (AOB vs Rolling) ───
monthly_aob_total = []
monthly_rolling_total = []
for j in range(12):
    ma = sum(c.get('monthly_aob',[0]*12)[j] for c in real_cats if c.get('monthly_aob') and len(c['monthly_aob'])>j)
    mr = sum(c.get('monthly_rolling',[0]*12)[j] for c in real_cats if c.get('monthly_rolling') and len(c['monthly_rolling'])>j)
    monthly_aob_total.append(ma)
    monthly_rolling_total.append(mr)

H.append(f'''<div class="sec"><div class="sh"><span class="l">\U0001f4c8 Monthly Revenue · AOB vs Rolling Forecast</span><span class="t">Full Year 2026 (\u00a5w = \u00a510,000)</span></div>
<div class="ca"><div class="ch"><canvas id="monthlyChart"></canvas></div></div></div>''')

# ─── TOP 15 SEGMENTS TABLE ───
H.append(f'''<div class="sec"><div class="sh"><span class="l">🏷️ Top Revenue Segments</span><span class="t">Rolling Forecast · {len(real_cats)} categories</span><input class="sbar" type="text" placeholder="🔍 Filter..." oninput="filterTable(this,'tbl-cat')" style="margin-left:auto"></div>
<div class="ca"><div class="tw"><table id="tbl-cat"><thead><tr>
<th onclick="sortTable('tbl-cat',0)">Segment</th><th onclick="sortTable('tbl-cat',1)">Rep</th><th class="ar" onclick="sortTable('tbl-cat',2)">2025 Actual</th>
<th class="ar" onclick="sortTable('tbl-cat',3)">2026 AOB</th><th class="ar" onclick="sortTable('tbl-cat',4)">Rolling Fcst</th><th class="pr" onclick="sortTable('tbl-cat',5)">Var vs AOB</th><th class="pr" onclick="sortTable('tbl-cat',6)">Var % vs AOB</th>
<th class="pr" onclick="sortTable('tbl-cat',7)">YoY vs 2025</th></tr></thead><tbody>''')
for cat in top_cats:
    var = cat['rolling'] - cat['aob']
    vpct = (var/cat['aob']*100) if cat['aob'] else 0
    yoy_pct = ((cat['rolling']/cat['y25_actual'])*100-100) if cat['y25_actual'] else 0
    vcls = 'gn' if var >= 0 else 'rd'
    ycls = 'gn' if yoy_pct >= 0 else 'rd'
    H.append(f'<tr><td>{cat["name"][:30]}</td><td>{cat["rep"][:18] or "—"}</td><td class="ar">{fmt(cat["y25_actual"]) if cat["y25_actual"] else "—"}</td><td class="ar">{fmt(cat["aob"])}</td><td class="ar">{fmt(cat["rolling"])}</td><td class="pr {vcls}">{"+" if var>=0 else ""}{fmt(abs(var))}</td><td class="pr {vcls}">{"+" if vpct>=0 else ""}{vpct:.1f}%</td><td class="pr {ycls}">{"+" if yoy_pct>=0 else ""}{yoy_pct:.1f}%</td></tr>')
H.append('</tbody></table></div></div></div>')

# ─── GA REVENUE (from Rolling Forecast) ───
# Find GA-Domestic and GA-International in forecast categories
ga_dom_cat = next((c for c in real_cats if 'GA-Domestic' in c['name']), None)
ga_int_cat = next((c for c in real_cats if 'GA-International' in c['name']), None)

if ga_dom_cat or ga_int_cat:
    H.append(f'<div class="sec"><div class="sh"><span class="l">📡 GA Revenue (Rolling Forecast)</span><span class="t">Domestic + International</span></div>')
    H.append('<div class="gr g3">')
    if ga_dom_cat:
        dom_var = ga_dom_cat['rolling'] - ga_dom_cat['aob']
        dom_vc = 'gn' if dom_var >= 0 else 'rd'
        dom_var_pct = (abs(dom_var)/ga_dom_cat['aob']*100) if ga_dom_cat and ga_dom_cat['aob'] else 0
        dom_var_sgn = '+' if dom_var >= 0 else ''
        H.append(f'<div class="kc"><div class="kt">GA-Domestic Rolling</div><div class="kv kv-sm">{fmt(ga_dom_cat["rolling"])}</div><div class="kd {dom_vc}">vs AOB {fmt(ga_dom_cat["aob"])} · {dom_var_sgn}{dom_var_pct:.1f}%</div></div>')
    if ga_int_cat:
        int_var = ga_int_cat['rolling'] - ga_int_cat['aob']
        int_vc = 'gn' if int_var >= 0 else 'rd'
        int_var_pct = (abs(int_var)/ga_int_cat['aob']*100) if ga_int_cat['aob'] else 0
        int_var_sgn = '+' if int_var >= 0 else ''
        H.append(f'<div class="kc"><div class="kt">GA-International Rolling</div><div class="kv kv-sm">{fmt(ga_int_cat["rolling"])}</div><div class="kd {int_vc}">vs AOB {fmt(ga_int_cat["aob"])} · {int_var_sgn}{int_var_pct:.1f}%</div></div>')
    ga_total = (ga_dom_cat['rolling'] if ga_dom_cat else 0) + (ga_int_cat['rolling'] if ga_int_cat else 0)
    ga_aob = (ga_dom_cat['aob'] if ga_dom_cat else 0) + (ga_int_cat['aob'] if ga_int_cat else 0)
    ga_var_pct = ((ga_total/ga_aob)*100-100) if ga_aob else 0
    H.append(f'<div class="kc"><div class="kt">GA Total Rolling</div><div class="kv kv-sm">{fmt(ga_total)}</div><div class="kd gn">{"+" if ga_var_pct>=0 else ""}{ga_var_pct:.1f}% vs AOB</div></div>')
    H.append('</div>')

    # GA monthly table
    ga_months = {m: [] for m in months_lbl}
    for j, m in enumerate(months_lbl):
        dv = ga_dom_cat.get('monthly_rolling', [0]*12)[j] if ga_dom_cat and ga_dom_cat.get('monthly_rolling') and len(ga_dom_cat['monthly_rolling']) > j else 0
        da = ga_dom_cat.get('monthly_aob', [0]*12)[j] if ga_dom_cat and ga_dom_cat.get('monthly_aob') and len(ga_dom_cat['monthly_aob']) > j else 0
        iv = ga_int_cat.get('monthly_rolling', [0]*12)[j] if ga_int_cat and ga_int_cat.get('monthly_rolling') and len(ga_int_cat['monthly_rolling']) > j else 0
        ga_months[m] = {'dom': dv, 'int': iv, 'aob': da}
    H.append('<div class="ca" style="margin-top:10px"><div class="tw"><table><thead><tr><th>Month</th><th class="ar">GA-Domestic (Rolling)</th><th class="ar">GA-International (Rolling)</th><th class="ar">Total</th></tr></thead><tbody>')
    for m in months_lbl:
        d = ga_months[m]
        if d['dom'] == 0 and d['int'] == 0: continue
        H.append(f'<tr><td>{m}</td><td class="ar">{fmt(d["dom"])}</td><td class="ar">{fmt(d["int"])}</td><td class="ar">{fmt(d["dom"]+d["int"])}</td></tr>')
    H.append('</tbody></table></div></div></div>')

# ─── KEY ACCOUNTS ───
ac = len(accounts)
t3 = sorted(accounts, key=lambda a: a['y26'], reverse=True)[:3]
H.append('<div class="gr g3">')
for i, a in enumerate(t3):
    g = ((a['y26']/a['y25'])*100-100) if a['y25'] else None
    gs = f'+{g:.1f}%' if g is not None and g >= 0 else (f'{g:.1f}%' if g is not None else 'NEW')
    dd = 'gn' if g is not None and g >= 0 else ('rd' if g is not None else 'gn')
    mx = max(a['m26']) if max(a['m26']) else 1
    colors = ['#d4af37','#2d6b4f','#0f2e24']
    bs = ''.join(f'<div class="sb-w" style="height:{max(8,(m/mx)*24)}px;background:{colors[i]}"></div>' for m in a['m26'])
    H.append(f'<div class="kc" style="border-top:3px solid {colors[i]}"><div class="kt">#{i+1} {a["name"][:18]}</div><div class="kv kv-sm">{fmt(a["y26"])}</div><div class="kd {dd}">{gs}</div><div class="spark" style="margin-top:4px">{bs}</div></div>')
H.append('</div>')

H.append(f'''<div class="sec"><div class="sh"><span class="l">🏆 Key Accounts</span><span class="t">{ac} accounts · Jan–May YTD</span><input class="sbar" type="text" placeholder="Filter..." oninput="filterTable(this,'tbl-ka')" style="margin-left:auto"></div>
<div class="ca"><div class="tw"><table id="tbl-ka"><thead><tr><th onclick="sortTable('tbl-ka',0)">Account</th><th class="ar" onclick="sortTable('tbl-ka',1)">2024 Full</th><th class="ar" onclick="sortTable('tbl-ka',2)">2025 YTD</th><th class="ar" onclick="sortTable('tbl-ka',3)">2026 YTD</th><th class="pr" onclick="sortTable('tbl-ka',4)">Change vs 2025</th><th class="pr" onclick="sortTable('tbl-ka',5)">Share</th></tr></thead><tbody>''')
for a in accounts:
    g = ((a['y26']/a['y25'])*100-100) if a['y25'] else None
    gs = f'+{g:.1f}%' if g is not None and g >= 0 else (f'{g:.1f}%' if g is not None else 'N/A')
    dd = 'gn' if g is not None and g >= 0 else ('rd' if g is not None else 'gn')
    sh = a['y26']/ka26*100 if ka26 else 0
    H.append(f'<tr><td>{a["name"][:35]}</td><td class="ar">{fmt(a["ft"])}</td><td class="ar">{fmt(a["y25"])}</td><td class="ar">{fmt(a["y26"])}</td><td class="pr {dd}">{gs}</td><td class="pr">{sh:.1f}%</td></tr>')
H.append('</tbody></table></div></div></div>')

# ─── SALES TEAM ───
team_html_parts = []
for s in team:
    pct_team = s['y26']/st26*100 if st26 else 0
    bw = max(6, s['y26']/max(st26,1)*100)
    if s['y25'] and s['y25'] > 0:
        chg = ((s['y26']/s['y25'])-1)*100
        chg_cls = 'gn' if chg >= 0 else 'rd'
        chg_txt = f'+{int(chg)}%' if chg >= 0 else f'{int(chg)}%'
    else:
        chg_cls = 'gn'
        chg_txt = 'N/A'
    team_html_parts.append(f'<div style="margin-bottom:8px"><div style="display:flex;justify-content:space-between;font-size:12px"><span style="font-weight:500">{s["n"]}</span><span style="font-weight:600">{fmt(s["y26"])}</span></div><div class="prg"><div class="pf" style="width:{bw}%;background:#d4af37"></div></div><div style="display:flex;justify-content:space-between;font-size:10px;color:var(--text2)"><span>{pct_team:.1f}% of team</span><span class="{chg_cls}">{chg_txt}</span></div></div>')

H.append(f'''<div class="sec"><div class="sh"><span class="l">👥 Sales Team</span><span class="t">Jan–May YTD · {len(team)} sellers</span></div>
<div class="ca"><div class="cb">
{''.join(team_html_parts)}
</div></div></div>''')

# ─── FOOTER ───
H.append(f'''<div class="ft">Spring City Golf · {MON} · Data: Rolling Forecast + Monthly Reports</div>
</div>

<script>
function toggleTheme(){{const h=document.documentElement;h.dataset.theme=h.dataset.theme==='dark'?'light':'dark';localStorage.setItem('sc-theme',h.dataset.theme)}}
if(localStorage.getItem('sc-theme')==='dark')document.documentElement.dataset.theme='dark';

function sortTable(t,i){{const t2=document.getElementById(t);if(!t2)return;const b=t2.querySelector('tbody'),r=Array.from(b.querySelectorAll('tr')),h=t2.querySelectorAll('th')[i],n=h.classList.contains('ar')||h.classList.contains('pr'),a=!h.classList.contains('srt');t2.querySelectorAll('th').forEach(x=>x.classList.remove('srt','srt-d'));r.sort((x,y)=>{{let va=x.cells[i].textContent.trim().replace(/[¥,↑↓+\s]/g,'')||'0',vb=y.cells[i].textContent.trim().replace(/[¥,↑↓+\s]/g,'')||'0';return n?(a?1:-1)*((parseFloat(va)||0)-(parseFloat(vb)||0)):(a?1:-1)*va.localeCompare(vb)}});r.forEach(x=>b.appendChild(x));h.classList.add(a?'srt':'srt-d')}}

function filterTable(i,t){{const q=i.value.toLowerCase(),t2=document.getElementById(t);if(!t2)return;t2.querySelectorAll('tbody tr').forEach(r=>r.style.display=r.textContent.toLowerCase().includes(q)?'':'none')}}

const mc=document.getElementById('monthlyChart');
if(mc){{
const m={json.dumps(months_lbl)};
const aob={json.dumps([round(v/10000) for v in monthly_aob_total])};
const roll={json.dumps([round(v/10000) for v in monthly_rolling_total])};
new Chart(mc,{{
type:'bar',data:{{labels:m,datasets:[
{{label:'AOB (Budget)',data:aob,backgroundColor:'rgba(212,175,55,0.5)',borderColor:'#d4af37',borderWidth:1,borderRadius:3}},
{{label:'Rolling Forecast',data:roll,backgroundColor:'rgba(45,107,79,0.7)',borderColor:'#2d6b4f',borderWidth:1,borderRadius:3}}
]}},
options:{{
responsive:true,maintainAspectRatio:false,
plugins:{{legend:{{position:'top',labels:{{font:{{size:11}},usePointStyle:true}}}},
tooltip:{{callbacks:{{label:ctx=>'¥'+(ctx.parsed.y*10000).toLocaleString()}}}}}},
scales:{{x:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}},
y:{{beginAtZero:true,ticks:{{callback:v=>'¥'+(v*10000/10000).toLocaleString()+'w',font:{{size:10}}}}}}}}
}})}};
</script>
</body></html>''')

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(H))

size = os.path.getsize(OUT)/1024
print(f'\n✅ Dashboard: {OUT} ({size:.0f} KB) · {len(real_cats)} categories · {len(accounts)} accounts · {len(team)} sellers')
