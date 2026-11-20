#!/usr/bin/env python3
"""
property_market_report.py — Monthly Central Coast property market report generator
Usage: python3 property_market_report.py <input.xlsx> <MONTH> <YEAR> [--publish]
"""

import sys, re, statistics, pathlib, datetime, subprocess, argparse, json
import openpyxl

REPORT_DIR   = pathlib.Path("/home/azureuser/property-reports")
COUNCIL_HTML = pathlib.Path("/home/azureuser/council-work/councildata-ai/index.html")
COUNCIL_REPO = pathlib.Path("/home/azureuser/council-work/councildata-ai")
GITHUB_PAGES = "https://pspedding.github.io/reports"

# ── helpers ──────────────────────────────────────────────────────────────────
def load_data(path):
    w = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sh = list(w)[0]
    rows = list(sh.iter_rows(values_only=True))
    headers = [str(h) for h in rows[0]]
    data = []
    for row in rows[1:]:
        if row[0] is None: continue
        data.append({headers[i]: row[i] for i in range(len(headers))})
    return data

def safe_float(v):
    if v is None: return None
    try: return float(v)
    except (ValueError, TypeError): return None

def fv(d, key):
    return safe_float(d.get(key))

def fmt_price(v):
    if v is None: return "—"
    return "$" + f"{int(v):,}"

def fmt_pct(v, dp=1):
    if v is None: return "—"
    sign = "+" if float(v) >= 0 else ""
    return f"{sign}{float(v)*100:.{dp}f}%"

# ── analysis ─────────────────────────────────────────────────────────────────
def analyse(data, month, year):
    hp  = [float(d["House Price (Current)"]) for d in data if d.get("House Price (Current)")]
    up  = [float(d["Unit Price (Current)"]) for d in data if d.get("Unit Price (Current)")]
    vis = [float(d["Listing Visits per Property"]) for d in data if d.get("Listing Visits per Property")]
    yoy = [fv(d,"House Price YoY (%)") for d in data if fv(d,"House Price YoY (%)") is not None]

    near_1m = sorted([d for d in data if d.get("House Price (Current)") and
                      960_000 <= float(d["House Price (Current)"]) < 1_000_000],
                     key=lambda d: -float(d["House Price (Current)"]))

    no_unit = sorted([d for d in data if not d.get("Unit Price (Current)")],
                     key=lambda d: -float(d.get("House Price (Current)") or 0))

    big_mom_h = sorted([d for d in data if fv(d,"MoM Change (%)") is not None
                        and abs(fv(d,"MoM Change (%)") or 0) > 0.10],
                       key=lambda d: -(fv(d,"MoM Change (%)") or 0))

    big_yoy_h = sorted([d for d in data if fv(d,"House Price YoY (%)") is not None
                        and abs(fv(d,"House Price YoY (%)") or 0) > 0.10],
                       key=lambda d: -(fv(d,"House Price YoY (%)") or 0))

    big_mom_u = sorted([d for d in data if fv(d,"Unit MoM Change (%)") is not None
                        and abs(fv(d,"Unit MoM Change (%)") or 0) > 0.10],
                       key=lambda d: -(fv(d,"Unit MoM Change (%)") or 0))

    visits_sorted = sorted([d for d in data if d.get("Listing Visits per Property")],
                           key=lambda d: -float(d["Listing Visits per Property"]))

    # Price bands
    band_defs = [
        ("Under $900k",  0,         900_000),
        ("$900k-$1M",    900_000,   1_000_000),
        ("$1M-$1.25M",   1_000_000, 1_250_000),
        ("$1.25M-$1.5M", 1_250_000, 1_500_000),
        ("$1.5M-$2M",    1_500_000, 2_000_000),
        ("$2M+",         2_000_000, 99_999_999),
    ]
    band_counts = [(lbl, sum(1 for p in hp if lo <= p < hi)) for lbl, lo, hi in band_defs]

    n_above_1m = sum(1 for p in hp if p >= 1_000_000)

    return dict(
        month=month, year=year, n=len(data),
        median_house=statistics.median(hp),
        median_unit=statistics.median(up) if up else None,
        mean_yoy=statistics.mean(yoy) if yoy else 0,
        median_visits=round(statistics.median(vis)) if vis else 0,
        n_above_1m=n_above_1m,
        n_no_unit=len(no_unit),
        near_1m=near_1m, no_unit=no_unit,
        big_mom_h=big_mom_h, big_yoy_h=big_yoy_h,
        big_mom_u=big_mom_u,
        visits_sorted=visits_sorted, band_counts=band_counts,
        data=data
    )

# ── HTML render ──────────────────────────────────────────────────────────────
def render_html(a):
    m, y = a["month"], a["year"]
    gen  = datetime.date.today().strftime("%-d %B %Y")

    # ── chart data preparation ──
    # YoY gainers: top 10 by YoY growth (ascending for horizontal bar)
    yoy_pos = [d for d in a["big_yoy_h"] if (fv(d,"House Price YoY (%)") or 0) > 0][:10]
    yoy_neg = [d for d in a["big_yoy_h"] if (fv(d,"House Price YoY (%)") or 0) < 0]
    yoy_pos_labels = json.dumps([d["Suburb"] for d in reversed(yoy_pos)])
    yoy_pos_data   = json.dumps([round((fv(d,"House Price YoY (%)") or 0)*100,1) for d in reversed(yoy_pos)])
    yoy_neg_labels = json.dumps([d["Suburb"] for d in reversed(yoy_neg)])
    yoy_neg_data   = json.dumps([round((fv(d,"House Price YoY (%)") or 0)*100,1) for d in reversed(yoy_neg)])

    # Interested buyers chart: top 15
    top15 = a["visits_sorted"][:15]
    vis_labels = json.dumps([d["Suburb"] for d in reversed(top15)])
    vis_data   = json.dumps([float(d["Listing Visits per Property"]) for d in reversed(top15)])
    vis_prices = json.dumps([fmt_price(d.get("House Price (Current)")) for d in reversed(top15)])

    # Price band chart
    band_labels = json.dumps([b[0] for b in a["band_counts"]])
    band_data   = json.dumps([b[1] for b in a["band_counts"]])

    # MoM big movers chart (house)
    mom_h_labels = json.dumps([d["Suburb"] for d in reversed(a["big_mom_h"])])
    mom_h_data   = json.dumps([round((fv(d,"MoM Change (%)") or 0)*100,1) for d in reversed(a["big_mom_h"])])
    mom_h_colors = json.dumps(["rgba(22,163,74,0.75)" if (fv(d,"MoM Change (%)") or 0)>0 else "rgba(220,38,38,0.75)" for d in reversed(a["big_mom_h"])])

    # ── table helpers ──
    def pct_cell(v):
        if v is None: return "<td>—</td>"
        cls = "pos" if v >= 0 else "neg"
        sign = "+" if v >= 0 else ""
        return f'<td class="{cls}">{sign}{v*100:.1f}%</td>'

    def outlook(d):
        mom = fv(d,"MoM Change (%)") or 0
        yoy = fv(d,"House Price YoY (%)") or 0
        if mom > 0.02 and yoy > 0.05: return "🟢 Strong"
        if mom > 0 and yoy > 0: return "🟢 Gaining"
        if mom < -0.02 or yoy < 0: return "🔴 Retreating"
        return "🟡 Steady"

    # near-$1M table rows
    near_rows = ""
    for d in a["near_1m"]:
        hp  = float(d["House Price (Current)"])
        gap = 1_000_000 - hp
        mom = fv(d,"MoM Change (%)")
        yoy = fv(d,"House Price YoY (%)")
        vis = d.get("Listing Visits per Property")
        up  = d.get("Unit Price (Current)")
        ol  = outlook(d)
        near_rows += f'<tr><td><strong>{d["Suburb"]}</strong></td>'
        near_rows += f'<td>{fmt_price(hp)}</td><td>${gap:,.0f}</td>'
        near_rows += pct_cell(mom)
        near_rows += pct_cell(yoy)
        near_rows += f'<td>{fmt_price(up) if up else "—"}</td>'
        near_rows += f'<td>{int(float(vis)):,}</td><td>{ol}</td></tr>'  if vis else f'<td>—</td><td>{ol}</td></tr>'

    # no-unit table rows
    nounit_rows = ""
    for d in a["no_unit"]:
        hp   = d.get("House Price (Current)")
        yoy  = fv(d,"House Price YoY (%)")
        rent = d.get("House Rent (P/W)")
        vis  = d.get("Listing Visits per Property")
        flag = d.get("$1M Median Suburb?","")
        fc   = "tag-green" if flag in ("Y","New") else "tag-amber"
        nounit_rows += f'<tr><td><strong>{d["Suburb"]}</strong></td>'
        nounit_rows += f'<td>{fmt_price(hp)}</td>'
        nounit_rows += pct_cell(yoy)
        nounit_rows += f'<td>{"$"+str(int(float(rent)))+"/wk" if rent else "—"}</td>'
        nounit_rows += f'<td>{int(float(vis)):,}</td>'  if vis else f'<td>—</td>'
        nounit_rows += f'<td class="{fc}">{flag if flag else "—"}</td></tr>'

    # large YoY table rows (all >±10%)
    yoy_rows = ""
    for d in a["big_yoy_h"]:
        hp  = d.get("House Price (Current)")
        yoy = fv(d,"House Price YoY (%)")
        mom = fv(d,"MoM Change (%)")
        caveat = " ⚠️" if abs(yoy or 0) > 0.25 else ""
        yoy_rows += f'<tr><td><strong>{d["Suburb"]}{caveat}</strong></td>'
        yoy_rows += f'<td>{fmt_price(hp)}</td>'
        yoy_rows += pct_cell(yoy)
        yoy_rows += pct_cell(mom)
        yoy_rows += '</tr>'

    # unit MoM big rows
    umom_rows = ""
    for d in a["big_mom_u"]:
        up   = d.get("Unit Price (Current)")
        umom = fv(d,"Unit MoM Change (%)")
        umom_rows += f'<tr><td><strong>{d["Suburb"]}</strong></td>'
        umom_rows += f'<td>{fmt_price(up) if up else "—"}</td>'
        umom_rows += pct_cell(umom)
        umom_rows += '</tr>'

    # ── change cards (Sep→Oct context) ──
    change_cards = """
    <div class="change-card up">
      <div class="suburb-name">Noraville &mdash; New $1M suburb</div>
      <div class="price">$1,000,000</div>
      <div class="badges">
        <span class="badge green">Sub-$1M → $1M</span>
        <span class="badge green">+5.6% from Sep $947,500</span>
        <span class="badge blue">+11.4% YoY</span>
      </div>
    </div>
    <div class="change-card up">
      <div class="suburb-name">Wadalba &mdash; Confirmed $1M+</div>
      <div class="price">$1,010,000</div>
      <div class="badges">
        <span class="badge blue">Sep crossing confirmed</span>
        <span class="badge green">+1.0% MoM steady</span>
        <span class="badge green">+9.8% YoY</span>
      </div>
    </div>
    <div class="change-card warn">
      <div class="suburb-name">Gosford &mdash; Extreme MoM surge</div>
      <div class="price">$1,555,000</div>
      <div class="badges">
        <span class="badge amber">+17.7% MoM +$233k</span>
        <span class="badge red">+50.4% YoY — may reflect data effects</span>
      </div>
    </div>
    <div class="change-card warn">
      <div class="suburb-name">West Gosford &mdash; Crossed $1M</div>
      <div class="price">$1,010,099</div>
      <div class="badges">
        <span class="badge amber">+28.7% MoM from $785k Sep</span>
        <span class="badge red">+7.0% YoY — 133 visits only</span>
      </div>
    </div>
    <div class="change-card down">
      <div class="suburb-name">Empire Bay unit &mdash; Sharp drop</div>
      <div class="price">Unit $450,000</div>
      <div class="badges">
        <span class="badge red">-29.1% unit MoM</span>
        <span class="badge amber">From $635k Sep</span>
        <span class="badge gray">House stable $1.29M</span>
      </div>
    </div>
    <div class="change-card down">
      <div class="suburb-name">Koolewong &mdash; Continued decline</div>
      <div class="price">$885,000</div>
      <div class="badges">
        <span class="badge red">-3.8% MoM Oct</span>
        <span class="badge red">-19.9% YoY</span>
        <span class="badge amber">From $975k near-$1M in Aug</span>
      </div>
    </div>
"""

    # ── pre-compute counts used in f-string ──
    n_yoy_h = len(a["big_yoy_h"])
    n_umom  = len(a["big_mom_u"])

    # ── KPI color helpers ──
    yoy_cls = "green" if a["mean_yoy"] >= 0 else "red"
    mom_delta = a["median_house"] - 1_145_500  # vs Sep
    mom_cls = "green" if mom_delta >= 0 else "amber"
    mom_str = ("+" if mom_delta >= 0 else "") + f"${abs(int(mom_delta)):,} vs Sep"

    # ── inline insight boxes ──
    insight_html = """
    <div class="insight-grid">
      <div class="insight-card amber">
        <div class="ic-head"><span class="ico">📉</span>↓ Median softening</div>
        <div class="ic-body">Median house price eased $15,250 from September's $1,145,500. YoY growth also moderated to +2.81% from +3.65%. October reflects a mild correction after two months of gains, though the market remains broadly above $1M median.</div>
      </div>
      <div class="insight-card green">
        <div class="ic-head"><span class="ico">✅</span>✓ 51 suburbs at or above $1M</div>
        <div class="ic-body">Net gain of 2 suburbs from September's 49. Noraville crossed the threshold at exactly $1,000,000. Wadalba confirmed its September crossing at $1,010,000. West Gosford also jumped above $1M this month (+28.7% MoM) but warrants independent verification given the magnitude of the move.</div>
      </div>
      <div class="insight-card red">
        <div class="ic-head"><span class="ico">⚠️</span>! Gosford &amp; West Gosford anomalies</div>
        <div class="ic-body">Gosford shows +50.4% YoY and +17.7% MoM at $1,555,000. West Gosford surged +28.7% MoM to $1,010,099. These extraordinary movements — in suburbs with notably low visit counts (251 and 133 respectively) — may reflect property-mix changes within the median calculation or other data effects. Without sales count data this cannot be verified. Do not use to represent LGA-wide appreciation.</div>
      </div>
      <div class="insight-card green">
        <div class="ic-head"><span class="ico">🏆</span>★ Demand: Wamberal leads</div>
        <div class="ic-body">Wamberal (3,454 visits/property) overtakes September's leader Terrigal (now 3,014). Both are coastal lifestyle suburbs above $1.6M. Bateau Bay (2,773) remains highest-engagement suburb in the accessible $1M–$1.3M band — sustained buyer engagement (listing visits proxy, not unique buyers) continues despite median softening.</div>
      </div>
    </div>
"""

    # ── assemble full HTML ──
    median_house_fmt = fmt_price(a["median_house"])
    median_unit_fmt  = fmt_price(a["median_unit"])
    mean_yoy_fmt     = fmt_pct(a["mean_yoy"])
    n_near1m = len(a["near_1m"])
    n_no_unit = a["n_no_unit"]
    n_above   = a["n_above_1m"]
    n_subs    = a["n"]
    median_vis = str(a["median_visits"])

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Central Coast Property Market \u2014 {m} {y}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet"/>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Inter",sans-serif;background:#f1f5f9;color:#1e293b;padding:24px 16px}}
.container{{max-width:960px;margin:0 auto}}
header{{text-align:center;margin-bottom:32px}}
header h1{{font-size:1.75rem;font-weight:700;color:#1e293b}}
header p{{color:#64748b;margin-top:6px;font-size:.95rem}}
.card{{background:#fff;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:24px;margin-bottom:24px}}
.card h2{{font-size:1.1rem;font-weight:600;color:#1e293b;margin-bottom:16px;padding-bottom:10px;border-bottom:2px solid #f1f5f9}}
.kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}}
@media(max-width:640px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
.kpi{{background:#f8fafc;border-radius:10px;padding:16px;text-align:center;border:1px solid #e2e8f0}}
.kpi .val{{font-size:1.4rem;font-weight:700;color:#2563eb}}
.kpi .val.green{{color:#16a34a}}.kpi .val.amber{{color:#d97706}}.kpi .val.red{{color:#dc2626}}
.kpi .lbl{{font-size:.78rem;color:#64748b;margin-top:4px;font-weight:500}}
.kpi .sub{{font-size:.72rem;color:#94a3b8;margin-top:2px}}
.change-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}}
.change-card{{border-radius:10px;padding:18px;border:1px solid #e2e8f0}}
.change-card.up{{background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-color:#86efac}}
.change-card.down{{background:linear-gradient(135deg,#fff7ed,#ffedd5);border-color:#fdba74}}
.change-card.warn{{background:linear-gradient(135deg,#fef2f2,#fee2e2);border-color:#fca5a5}}
.change-card .suburb-name{{font-size:.95rem;font-weight:700;color:#1d4ed8;margin-bottom:4px}}
.change-card.up .suburb-name{{color:#15803d}}.change-card.down .suburb-name{{color:#b45309}}.change-card.warn .suburb-name{{color:#b91c1c}}
.change-card .price{{font-size:1.2rem;font-weight:700;color:#1e293b;margin:4px 0 8px}}
.badges{{display:flex;gap:6px;flex-wrap:wrap}}
.badge{{padding:3px 9px;border-radius:20px;font-size:.73rem;font-weight:600}}
.badge.green{{background:#dcfce7;color:#15803d}}.badge.blue{{background:#dbeafe;color:#1d4ed8}}
.badge.amber{{background:#fef9c3;color:#a16207}}.badge.red{{background:#fee2e2;color:#b91c1c}}
.badge.gray{{background:#f1f5f9;color:#475569}}
.chart-wrap{{position:relative;height:340px}}
.chart-wrap.tall{{height:430px}}.chart-wrap.sm{{height:260px}}
table{{border-collapse:collapse;width:100%;font-size:.875rem}}
th{{text-align:left;padding:9px 12px;background:#f8fafc;font-weight:600;color:#475569;border-bottom:2px solid #e2e8f0}}
td{{padding:8px 12px;border-bottom:1px solid #f1f5f9}}
tr:nth-child(even) td{{background:#fafafa}}
.pos{{color:#16a34a;font-weight:600}}.neg{{color:#dc2626;font-weight:600}}
.tag-green{{color:#16a34a;font-weight:600}}.tag-amber{{color:#d97706;font-weight:600}}.tag-red{{color:#dc2626;font-weight:600}}
.note{{font-size:.82rem;color:#64748b;margin-top:12px;font-style:italic}}
.scroll{{overflow-x:auto}}
.insight-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
@media(max-width:600px){{.insight-grid{{grid-template-columns:1fr}}}}
.insight-card{{background:#f8fafc;border-radius:10px;padding:16px;border-left:4px solid #2563eb}}
.insight-card.green{{border-left-color:#16a34a}}.insight-card.amber{{border-left-color:#d97706}}.insight-card.red{{border-left-color:#dc2626}}
.ic-head{{font-weight:700;font-size:.9rem;color:#1e293b;margin-bottom:6px}}
.ic-body{{font-size:.82rem;color:#475569;line-height:1.5}}
.warn-box{{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px 14px;font-size:.8rem;color:#92400e;margin-top:14px}}
.info-box{{background:#eff6ff;border:1px solid #93c5fd;border-radius:8px;padding:12px 14px;font-size:.8rem;color:#1e40af;margin-top:14px}}
footer{{text-align:center;font-size:.78rem;color:#94a3b8;margin-top:8px;padding:16px}}
@media print{{.card{{break-inside:avoid;page-break-inside:avoid}}.insight-grid{{break-inside:avoid;page-break-inside:avoid}}.insight-card{{break-inside:avoid;page-break-inside:avoid}}.change-grid{{break-inside:avoid;page-break-inside:avoid}}.change-card{{break-inside:avoid;page-break-inside:avoid}}.kpi-grid{{break-inside:avoid;page-break-inside:avoid}}.warn-box{{break-inside:avoid;page-break-inside:avoid}}.card h2 span{{display:none}}.ico{{display:none}}}}
</style>
</head>
<body>
<div class="container">

<header>
  <h1>Central Coast Property Market \u2014 {m} {y}</h1>
  <p>{n_subs} suburbs tracked &nbsp;|&nbsp; Source: REA Group &nbsp;|&nbsp; Analysis: See-Change Solutions &nbsp;|&nbsp; Generated {gen}</p>
</header>

<!-- MARKET OVERVIEW KPIs -->
<div class="card">
  <h2><span>&#128202;</span> Market Overview &mdash; {m} {y}</h2>
  <div class="kpi-grid">
    <div class="kpi"><div class="val">{n_subs}</div><div class="lbl">Suburbs Tracked</div></div>
    <div class="kpi"><div class="val {mom_cls}">{median_house_fmt}</div><div class="lbl">Median House Price</div><div class="sub">{mom_str}</div></div>
    <div class="kpi"><div class="val {yoy_cls}">{mean_yoy_fmt}</div><div class="lbl">Mean House YoY Growth</div></div>
    <div class="kpi"><div class="val">{median_unit_fmt}</div><div class="lbl">Median Unit Price</div></div>
    <div class="kpi"><div class="val">{n_above}</div><div class="lbl">Suburbs at/above $1M (house)</div></div>
    <div class="kpi"><div class="val amber">{n_near1m}</div><div class="lbl">Approaching $1M (&ge;$960k)</div></div>
    <div class="kpi"><div class="val amber">{n_no_unit}</div><div class="lbl">No unit median reported</div></div>
    <div class="kpi"><div class="val">{median_vis}</div><div class="lbl">Median interested buyers/property (engagement proxy)</div></div>
  </div>
</div>

<!-- KEY INSIGHTS -->
<div class="card">
  <h2><span>&#128161;</span> Key Insights &mdash; {m} {y}</h2>
  {insight_html}
</div>

<!-- STATUS CHANGES -->
<div class="card">
  <h2><span>&#128260;</span> Suburb Status Changes &mdash; September to October 2026</h2>
  <div class="change-grid">
  {change_cards}
  </div>
  <p class="note">Gosford CBD and West Gosford show extraordinary single-month jumps. Both have notably low visit counts (251 and 133 respectively); sales count data would be needed to assess what drove these changes. Noraville ($1M, 986 visits) and Wadalba ($1,010k, 865 visits) crossed the $1M threshold with more typical visit activity for their price range.</p>
</div>

<!-- YOY GAINERS CHART -->
<div class="card">
  <h2><span>&#128200;</span> Top House Price Gainers &mdash; Year on Year (&gt;10%)</h2>
  <div class="chart-wrap tall"><canvas id="gainersChart"></canvas></div>
  <p class="note">[!]#9888; Gosford +50.4% YoY is highlighted; treat with caution as the cause cannot be confirmed without sales count data. Suburbs marked [!]#9888; have YoY changes &gt;25% which may reflect property-mix or other data effects.</p>
</div>

<!-- YOY DECLINERS CHART -->
<div class="card">
  <h2><span>&#128201;</span> House Price Decliners &mdash; Year on Year</h2>
  <div class="chart-wrap"><canvas id="losersChart"></canvas></div>
  <p class="note">Wagstaffe -39.5% YoY persists from Sep. Koolewong continues multi-month retreat. Large negative movements may reflect property-mix or other data effects — sales count data would be needed to assess causes.</p>
</div>

<!-- MOM BIG MOVERS CHART -->
<div class="card">
  <h2><span>&#128257;</span> House MoM Movements &gt;10% &mdash; October 2026</h2>
  <div class="chart-wrap sm"><canvas id="momChart"></canvas></div>
  <p class="note">Month-on-month changes above 10% magnitude. West Gosford (+28.7%) and Gosford (+17.7%) are outliers likely driven by sample composition changes.</p>
</div>

<!-- APPROACHING $1M -->
<div class="card">
  <h2><span>&#128280;</span> Approaching $1M &mdash; Suburbs &ge;$960k</h2>
  <div class="scroll">
  <table>
    <thead><tr><th>Suburb</th><th>House Price</th><th>Gap to $1M</th><th>MoM</th><th>YoY</th><th>Unit Price</th><th>Visits/prop</th><th>Outlook</th></tr></thead>
    <tbody>{near_rows}</tbody>
  </table>
  </div>
  <p class="note">Killarney Vale ($987,750, +1.1% MoM) is closest to the $1M threshold. Wyoming has held near $960k for three months with minimal momentum. Wyongah shows modest YoY growth (+18.8%) but retreated -1.3% MoM. All four remain genuine candidates if momentum returns.</p>
</div>

<!-- INTERESTED BUYERS CHART -->
<div class="card">
  <h2><span>&#128101;</span> Interested Buyers &mdash; Top 15 by Listing Visits per Property</h2>
  <div class="chart-wrap tall"><canvas id="visChart"></canvas></div>
  <div class="info-box">&#128161; &ldquo;Listing visits per property&rdquo; is a buyer engagement proxy (REA Group metric), not a count of unique buyers or settled transactions. Top 15 suburbs shown. Wamberal (3,454) and Terrigal (3,014) continue to dominate, suggesting strong aspirational demand for coastal lifestyle properties.</div>
</div>

<!-- NO UNIT SUBURBS -->
<div class="card">
  <h2><span>&#127968;</span> Suburbs with No Unit Price Data</h2>
  <div class="scroll">
  <table>
    <thead><tr><th>Suburb</th><th>House Price</th><th>House YoY</th><th>House Rent</th><th>Visits/prop</th><th>$1M Flag</th></tr></thead>
    <tbody>{nounit_rows}</tbody>
  </table>
  </div>
  <div class="warn-box">[!]#9888; Absence of a unit median indicates insufficient REA data for a reliable estimate — it does <em>not</em> confirm that no unit stock exists. Erina Heights (house $2,307,500), Pretty Beach ($1,982,500) and Killcare Heights ($1,820,000) are premium lifestyle areas where units may simply not trade frequently enough for a reliable REA median.</div>
</div>

<!-- HOUSE YOY FULL TABLE -->
<div class="card">
  <h2><span>&#128196;</span> All House YoY Changes &gt;&pm;10% ({n_yoy_h} suburbs)</h2>
  <div class="scroll">
  <table>
    <thead><tr><th>Suburb</th><th>House Price</th><th>YoY</th><th>MoM</th></tr></thead>
    <tbody>{yoy_rows}</tbody>
  </table>
  </div>
  <p class="note">[!]#9888; marks suburbs with YoY change exceeding &pm;25% — elevated potential for data or property-mix effects; treat these figures with extra caution.</p>
</div>

<!-- UNIT MOM CHANGES -->
<div class="card">
  <h2><span>&#127970;</span> Unit MoM Changes &gt;&pm;10% ({n_umom} suburbs)</h2>
  <div class="scroll">
  <table>
    <thead><tr><th>Suburb</th><th>Unit Price</th><th>Unit MoM</th></tr></thead>
    <tbody>{umom_rows}</tbody>
  </table>
  </div>
  <p class="note">Empire Bay unit dropped -29.1% MoM to $450,000 (from $635k in Sep). No corresponding house movement suggests unit market is separately driven — could reflect a single distressed sale in a thin market. Copacabana unit +18.1% MoM to $1,600,000.</p>
</div>

<!-- PRICE BAND DISTRIBUTION -->
<div class="card">
  <h2><span>&#128205;</span> House Price Band Distribution &mdash; {n_subs} Suburbs</h2>
  <div class="chart-wrap sm"><canvas id="bandChart"></canvas></div>
  <p class="note">Count of suburbs by house price band. Not weighted by listing volume or sales count. The $1M-$1.25M band is the most densely populated.</p>
</div>

<div class="warn-box">
  [!]#9888; <strong>Data limitations:</strong> All prices are suburb-level REA Group medians, not transaction-weighted LGA medians. YoY and MoM figures use prior-period REA data. &ldquo;Listing visits per property&rdquo; is an engagement proxy, not unique-buyer or sales data. No sales counts, zoning or feasibility data included. Missing unit median = insufficient REA data, not confirmed zero stock. Changes &gt;25% — especially in suburbs with &lt;500 visits/property — may reflect property-mix or other data effects rather than genuine market moves; sales count data would be needed to confirm. This report is for information only and is not investment advice.
</div>

<footer>
  <p>Source: REA Group &nbsp;|&nbsp; Central Coast LGA &nbsp;|&nbsp; {m} {y} &nbsp;|&nbsp; Generated {gen}</p>
  <p>See-Change Solutions &nbsp;|&nbsp; <a href="https://councildata.ai">councildata.ai</a></p>
</footer>

</div><!-- /container -->

<script>
const gainersLabels = {yoy_pos_labels};
const gainersData = {yoy_pos_data};
new Chart(document.getElementById("gainersChart"), {{
  type: "bar",
  data: {{ labels: gainersLabels, datasets: [{{
    label: "YoY Change (%)",
    data: gainersData,
    backgroundColor: gainersData.map(v => v > 25 ? "rgba(220,38,38,0.8)" : "rgba(22,163,74,0.75)"),
    borderColor: gainersData.map(v => v > 25 ? "rgba(185,28,28,1)" : "rgba(15,118,54,1)"),
    borderWidth: 1
  }}] }},
  options: {{ indexAxis: "y", responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => " +" + ctx.parsed.x + "%" }} }} }},
    scales: {{ x: {{ title: {{ display: true, text: "YoY Change (%)" }}, grid: {{ color: "#f1f5f9" }} }},
               y: {{ grid: {{ display: false }} }} }} }} }});

const losersLabels = {yoy_neg_labels};
const losersData = {yoy_neg_data};
new Chart(document.getElementById("losersChart"), {{
  type: "bar",
  data: {{ labels: losersLabels, datasets: [{{
    label: "YoY Change (%)",
    data: losersData,
    backgroundColor: "rgba(220,38,38,0.75)",
    borderColor: "rgba(185,28,28,1)", borderWidth: 1
  }}] }},
  options: {{ indexAxis: "y", responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => " " + ctx.parsed.x + "%" }} }} }},
    scales: {{ x: {{ title: {{ display: true, text: "YoY Change (%)" }}, grid: {{ color: "#f1f5f9" }} }},
               y: {{ grid: {{ display: false }} }} }} }} }});

const momLabels = {mom_h_labels};
const momData = {mom_h_data};
const momColors = {mom_h_colors};
new Chart(document.getElementById("momChart"), {{
  type: "bar",
  data: {{ labels: momLabels, datasets: [{{
    label: "MoM Change (%)", data: momData,
    backgroundColor: momColors, borderWidth: 1
  }}] }},
  options: {{ indexAxis: "y", responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }} }},
    scales: {{ x: {{ title: {{ display: true, text: "MoM Change (%)" }}, grid: {{ color: "#f1f5f9" }} }},
               y: {{ grid: {{ display: false }} }} }} }} }});

const visLabels = {vis_labels};
const visData = {vis_data};
const visPrices = {vis_prices};
new Chart(document.getElementById("visChart"), {{
  type: "bar",
  data: {{ labels: visLabels, datasets: [{{
    label: "Visits per property", data: visData,
    backgroundColor: "rgba(37,99,235,0.72)",
    borderColor: "rgba(29,78,216,1)", borderWidth: 1
  }}] }},
  options: {{ indexAxis: "y", responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }},
      tooltip: {{ callbacks: {{ afterLabel: ctx => "House: " + visPrices[ctx.dataIndex] }} }} }},
    scales: {{ x: {{ title: {{ display: true, text: "Listing visits per property" }}, grid: {{ color: "#f1f5f9" }} }},
               y: {{ grid: {{ display: false }} }} }} }} }});

const bandLabels = {band_labels};
const bandData = {band_data};
new Chart(document.getElementById("bandChart"), {{
  type: "bar",
  data: {{ labels: bandLabels, datasets: [{{
    label: "Number of suburbs", data: bandData,
    backgroundColor: ["rgba(100,116,139,0.7)","rgba(37,99,235,0.7)","rgba(22,163,74,0.7)","rgba(234,179,8,0.7)","rgba(234,88,12,0.7)","rgba(185,28,28,0.7)"],
    borderWidth: 1
  }}] }},
  options: {{ responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => " " + ctx.parsed.y + " suburbs" }} }} }},
    scales: {{ y: {{ title: {{ display: true, text: "Number of suburbs" }}, grid: {{ color: "#f1f5f9" }} }},
               x: {{ grid: {{ display: false }} }} }} }} }});
</script>
</body>
</html>"""

    n_yoy_h = len(a["big_yoy_h"])
    n_umom  = len(a["big_mom_u"])
    html = html.replace("{near_rows}", near_rows)
    html = html.replace("{nounit_rows}", nounit_rows)
    html = html.replace("{yoy_rows}", yoy_rows)
    html = html.replace("{umom_rows}", umom_rows)
    html = html.replace("{change_cards}", change_cards)
    html = html.replace("{insight_html}", insight_html)
    html = html.replace("{yoy_pos_labels}", yoy_pos_labels)
    html = html.replace("{yoy_pos_data}", yoy_pos_data)
    html = html.replace("{yoy_neg_labels}", yoy_neg_labels)
    html = html.replace("{yoy_neg_data}", yoy_neg_data)
    html = html.replace("{vis_labels}", vis_labels)
    html = html.replace("{vis_data}", vis_data)
    html = html.replace("{vis_prices}", vis_prices)
    html = html.replace("{band_labels}", band_labels)
    html = html.replace("{band_data}", band_data)
    html = html.replace("{mom_h_labels}", mom_h_labels)
    html = html.replace("{mom_h_data}", mom_h_data)
    html = html.replace("{mom_h_colors}", mom_h_colors)
    html = html.replace("{n_yoy_h}", str(n_yoy_h))
    html = html.replace("{n_umom}", str(n_umom))
    return html

# ── PDF ──────────────────────────────────────────────────────────────────────
def render_pdf(html_path, pdf_path):
    try:
        import weasyprint
        weasyprint.HTML(filename=str(html_path)).write_pdf(str(pdf_path))
        print("  PDF: weasyprint")
        return True
    except Exception as e:
        print(f"  weasyprint: {e}", file=sys.stderr)
    for browser in ["chromium-browser", "google-chrome", "chromium"]:
        try:
            r = subprocess.run([browser,"--headless","--no-sandbox",
                                f"--print-to-pdf={pdf_path}", str(html_path)],
                               capture_output=True, timeout=60)
            if r.returncode == 0:
                print(f"  PDF: {browser}")
                return True
        except Exception: pass
    print("WARNING: no PDF renderer found", file=sys.stderr)
    return False

# ── council links ─────────────────────────────────────────────────────────────
def update_council_links(html_name, pdf_name):
    src = COUNCIL_HTML.read_text()
    new = re.sub(
        r'href="https://pspedding\.github\.io/reports/Property_Market_[^"]*_Dashboard\.html"',
        f'href="https://pspedding.github.io/reports/{html_name}"',
        src)
    new = re.sub(
        r'href="https://pspedding\.github\.io/reports/Property_Market_[^"]*_Insights\.pdf"',
        f'href="https://pspedding.github.io/reports/{pdf_name}"',
        new)
    if new == src:
        print("WARNING: council link pattern not found", file=sys.stderr)
        return False
    COUNCIL_HTML.write_text(new)
    return True

# ── git push ──────────────────────────────────────────────────────────────────
def git_push(repo, message):
    subprocess.run(["git","-C",str(repo),"add","-A"], check=True)
    subprocess.run(["git","-C",str(repo),"commit","-m",message], check=True)
    subprocess.run(["git","-C",str(repo),"push"], check=True)

# ── main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Monthly property market report")
    ap.add_argument("xlsx", help="Source .xlsx from REA")
    ap.add_argument("month", help="e.g. Oct")
    ap.add_argument("year",  help="e.g. 2026")
    ap.add_argument("--publish", action="store_true",
                    help="Push to GitHub Pages and update councildata.ai")
    args = ap.parse_args()

    stem     = f"Property_Market_{args.month}_{args.year}"
    html_name= f"{stem}_Dashboard.html"
    pdf_name = f"{stem}_Insights.pdf"
    html_path= REPORT_DIR / html_name
    pdf_path = REPORT_DIR / pdf_name

    print(f"Loading {args.xlsx} ...")
    data = load_data(args.xlsx)
    print(f"  {len(data)} suburb rows")

    a = analyse(data, args.month, args.year)
    print(f"  Median house: {fmt_price(a['median_house'])} | >=1M (by price): {a['n_above_1m']} | No unit: {a['n_no_unit']}")

    print(f"Writing {html_path}")
    html_path.write_text(render_html(a))

    print(f"Rendering {pdf_path}")
    render_pdf(html_path, pdf_path)

    if args.publish:
        print("Pushing reports to GitHub Pages ...")
        git_push(REPORT_DIR, f"Add {args.month} {args.year} property market report")
        print("Updating councildata.ai ...")
        if update_council_links(html_name, pdf_name):
            git_push(COUNCIL_REPO, f"Update Property Market Intelligence to {args.month} {args.year}")
            print("  councildata.ai updated and pushed.")
        else:
            print("  WARNING: manual link update needed.")

    print("Done.")
    print(f"  HTML: {GITHUB_PAGES}/{html_name}")
    print(f"  PDF:  {GITHUB_PAGES}/{pdf_name}")

if __name__ == "__main__":
    main()