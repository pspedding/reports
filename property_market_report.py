#!/usr/bin/env python3
"""
property_market_report.py — Monthly Central Coast property market report generator

Usage:
  python3 property_market_report.py <input.xlsx> <MONTH> <YEAR> [--publish]

Output:
  /home/azureuser/property-reports/Property_Market_{MONTH}_{YEAR}_Dashboard.html
  /home/azureuser/property-reports/Property_Market_{MONTH}_{YEAR}_Insights.pdf

With --publish:
  Pushes both files to pspedding/reports GitHub Pages repo
  Updates councildata.ai index.html Property Market Intelligence links
  Pushes councildata.ai repo
"""

import sys, re, statistics, pathlib, datetime, subprocess, argparse
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
    return f"${int(v):,}"

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
                     key=lambda d: float(d.get("House Price (Current)") or 0))

    big_mom_h = sorted([d for d in data if fv(d,"MoM Change (%)") is not None
                        and abs(fv(d,"MoM Change (%)") or 0) > 0.10],
                       key=lambda d: fv(d,"MoM Change (%)") or 0)

    big_yoy_h = sorted([d for d in data if fv(d,"House Price YoY (%)") is not None
                        and abs(fv(d,"House Price YoY (%)") or 0) > 0.10],
                       key=lambda d: fv(d,"House Price YoY (%)") or 0)

    big_mom_u = sorted([d for d in data if fv(d,"Unit MoM Change (%)") is not None
                        and abs(fv(d,"Unit MoM Change (%)") or 0) > 0.10],
                       key=lambda d: fv(d,"Unit MoM Change (%)") or 0)

    big_yoy_u = sorted([d for d in data if fv(d,"Unit Price Change (%)") is not None
                        and abs(fv(d,"Unit Price Change (%)") or 0) > 0.10],
                       key=lambda d: fv(d,"Unit Price Change (%)") or 0)

    visits_sorted = sorted([d for d in data if d.get("Listing Visits per Property")],
                           key=lambda d: -float(d["Listing Visits per Property"]))

    bands = [
        ("Under $700k",  0,         700_000),
        ("$700k–$800k",  700_000,   800_000),
        ("$800k–$1M",    800_000,   1_000_000),
        ("$1M–$1.25M",   1_000_000, 1_250_000),
        ("$1.25M–$1.5M", 1_250_000, 1_500_000),
        ("$1.5M–$2M",    1_500_000, 2_000_000),
        ("$2M+",         2_000_000, 99_999_999),
    ]
    band_counts = [(lbl, sum(1 for p in hp if lo <= p < hi)) for lbl, lo, hi in bands]

    return dict(
        month=month, year=year, n=len(data),
        median_house=statistics.median(hp),
        median_unit=statistics.median(up) if up else None,
        mean_yoy=statistics.mean(yoy) if yoy else 0,
        total_visits=int(sum(vis)),
        n_above_1m=sum(1 for p in hp if p >= 1_000_000),
        n_no_unit=len(no_unit),
        near_1m=near_1m, no_unit=no_unit,
        big_mom_h=big_mom_h, big_yoy_h=big_yoy_h,
        big_mom_u=big_mom_u, big_yoy_u=big_yoy_u,
        visits_sorted=visits_sorted, band_counts=band_counts,
        data=data
    )

# ── HTML render ──────────────────────────────────────────────────────────────
def render_html(a):
    m, y = a["month"], a["year"]
    gen  = datetime.date.today().strftime("%-d %B %Y")

    def table_rows_house(rows, cols):
        out = []
        for d in rows:
            tds = []
            for col in cols:
                if col == "suburb":   tds.append(f'<td>{d["Suburb"]}</td>')
                elif col == "hp":     tds.append(f'<td>{fmt_price(d.get("House Price (Current)"))}</td>')
                elif col == "up":     tds.append(f'<td>{fmt_price(d.get("Unit Price (Current)"))}</td>')
                elif col == "gap":    tds.append(f'<td>${1_000_000-int(float(d["House Price (Current)"])):,}</td>')
                elif col == "mom":
                    v = d.get("MoM Change (%)")
                    cls = "pos" if v and float(v) >= 0 else "neg"
                    tds.append(f'<td class="{cls}">{fmt_pct(v)}</td>')
                elif col == "yoy":
                    v = d.get("House Price YoY (%)")
                    cls = "pos" if v and float(v) >= 0 else "neg"
                    tds.append(f'<td class="{cls}">{fmt_pct(v)}</td>')
                elif col == "umom":
                    v = d.get("Unit MoM Change (%)")
                    cls = "pos" if v and float(v) >= 0 else "neg"
                    tds.append(f'<td class="{cls}">{fmt_pct(v)}</td>')
                elif col == "uyoy":
                    v = d.get("Unit Price Change (%)")
                    cls = "pos" if v and float(v) >= 0 else "neg"
                    tds.append(f'<td class="{cls}">{fmt_pct(v)}</td>')
                elif col == "flag":   tds.append(f'<td>{d.get("$1M Median Suburb?","")}</td>')
                elif col == "vis":
                    v = d.get("Listing Visits per Property")
                    tds.append(f'<td>{int(float(v)):,}</td>' if v else "<td>—</td>")
                elif col == "rent":
                    v = d.get("House Rent (P/W)")
                    tds.append(f'<td>${float(v):.0f}/wk</td>' if v else "<td>—</td>")
            out.append("<tr>" + "".join(tds) + "</tr>")
        return "".join(out) or '<tr><td colspan="99">None</td></tr>'

    near_rows  = table_rows_house(a["near_1m"],    ["suburb","hp","gap","mom","yoy","flag","vis"])
    nounit_rows= table_rows_house(a["no_unit"],    ["suburb","hp","yoy","rent","vis"])
    bmomh_rows = table_rows_house(a["big_mom_h"],  ["suburb","hp","mom"])
    byoyh_rows = table_rows_house(a["big_yoy_h"],  ["suburb","hp","yoy"])
    bmomu_rows = table_rows_house(a["big_mom_u"],  ["suburb","up","umom"])
    byoyu_rows = table_rows_house(a["big_yoy_u"],  ["suburb","up","uyoy"])
    vis_rows   = table_rows_house(a["visits_sorted"][:15], ["suburb","vis","hp"])

    max_band = max((c for _,c in a["band_counts"]), default=1)
    bands_html = "".join(
        f'<div class="band"><span>{lbl}</span><div class="bar"><i style="width:{100*cnt/max_band:.0f}%"></i></div><b>{cnt}</b></div>'
        for lbl,cnt in a["band_counts"])

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Central Coast Property Market \u2014 {m} {y}</title>
<style>
*{{box-sizing:border-box}}body{{font-family:Inter,Arial,sans-serif;background:#f1f5f9;color:#1e293b;margin:0;padding:24px 16px;font-size:14px;line-height:1.55}}
.container{{max-width:960px;margin:0 auto}}header{{text-align:center;margin:0 0 28px}}
h1{{font-size:28px;margin:0}}header p{{color:#64748b;margin:4px}}
.card{{background:#fff;border-radius:12px;box-shadow:0 1px 4px #0001;padding:24px;margin:0 0 22px}}
.card h2{{font-size:18px;margin:0 0 14px;padding:0 0 9px;border-bottom:2px solid #f1f5f9}}
.card h3{{font-size:14px;margin:18px 0 7px;color:#1d4ed8}}
.kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}
.kpi{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:14px 8px;text-align:center}}
.val{{color:#2563eb;font-size:21px;font-weight:700}}.lbl{{font-size:11px;color:#64748b;line-height:1.3}}
.scroll{{overflow-x:auto}}table{{border-collapse:collapse;width:100%;font-size:12px}}
th{{background:#f8fafc;border-bottom:2px solid #e2e8f0;color:#475569;text-align:left;padding:7px 8px}}
td{{padding:7px 8px;border-bottom:1px solid #e2e8f0}}tbody tr:nth-child(even){{background:#fafafa}}
.pos{{color:#16a34a;font-weight:600}}.neg{{color:#dc2626;font-weight:600}}
.note{{font-size:12px;color:#64748b;margin:8px 0}}
.caution{{font-size:12px;background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px;color:#92400e;margin:14px 0}}
.band{{display:grid;grid-template-columns:130px 1fr 30px;gap:10px;align-items:center;margin:8px 0}}
.bar{{height:18px;background:#eff6ff;border-radius:5px}}.bar i{{display:block;height:100%;background:#3b82f6;border-radius:5px}}
footer{{text-align:center;color:#64748b;font-size:11px;padding:16px}}
@media(max-width:680px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
</style></head><body><main class="container">
<header><h1>Central Coast Property Market \u2014 {m} {y}</h1>
<p>{a["n"]} suburbs &middot; Source: REA Group &middot; Generated {gen}</p></header>

<section class="card"><h2>\U0001f4ca Market overview</h2><div class="kpi-grid">
<div class="kpi"><div class="val">{a["n"]}</div><div class="lbl">Suburbs tracked</div></div>
<div class="kpi"><div class="val">{fmt_price(a["median_house"])}</div><div class="lbl">Median suburb house price</div></div>
<div class="kpi"><div class="val">{fmt_pct(a["mean_yoy"])}</div><div class="lbl">Mean house YoY</div></div>
<div class="kpi"><div class="val">{fmt_price(a["median_unit"])}</div><div class="lbl">Median suburb unit price</div></div>
<div class="kpi"><div class="val">{a["n_above_1m"]}</div><div class="lbl">Suburbs at/above $1M (house)</div></div>
<div class="kpi"><div class="val">{len(a["near_1m"])}</div><div class="lbl">Near $1M (\u2265$960k, below $1M)</div></div>
<div class="kpi"><div class="val">{a["n_no_unit"]}</div><div class="lbl">No unit median reported</div></div>
<div class="kpi"><div class="val">{a["total_visits"]:,}</div><div class="lbl">Total visits/property</div></div>
</div></section>

<section class="card"><h2>\U0001f534 Near-$1M suburbs (\u2265$960k)</h2>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>House price</th><th>Gap to $1M</th><th>MoM</th><th>YoY</th><th>Flag</th><th>Visits/prop</th></tr></thead>
<tbody>{near_rows}</tbody></table></div>
<p class="note">Suburbs with house medians \u2265$960k but below $1M. "Flag" = source $1M field. Approaching-$1M label is a price screen, not a prediction.</p>
</section>

<section class="card"><h2>\U0001f3e0 Suburbs with no unit price data</h2>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>House price</th><th>House YoY</th><th>House rent</th><th>Visits/prop</th></tr></thead>
<tbody>{nounit_rows}</tbody></table></div>
<p class="note">Absence of a unit median = insufficient REA data, not confirmed absence of unit stock. Lower-priced suburbs here may merit investigation for housing mix and affordability policy.</p>
</section>

<section class="card"><h2>\U0001f4c8 Price changes &gt;10%</h2>
<h3>House MoM &gt;\u00b110% ({len(a["big_mom_h"])} suburbs)</h3>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>House price</th><th>MoM</th></tr></thead><tbody>{bmomh_rows}</tbody></table></div>
<h3>House YoY &gt;\u00b110% ({len(a["big_yoy_h"])} suburbs)</h3>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>House price</th><th>YoY</th></tr></thead><tbody>{byoyh_rows}</tbody></table></div>
<h3>Unit MoM &gt;\u00b110% ({len(a["big_mom_u"])} suburbs)</h3>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>Unit price</th><th>MoM</th></tr></thead><tbody>{bmomu_rows}</tbody></table></div>
<h3>Unit YoY &gt;\u00b110% ({len(a["big_yoy_u"])} suburbs)</h3>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>Unit price</th><th>YoY</th></tr></thead><tbody>{byoyu_rows}</tbody></table></div>
<p class="note">Strictly &gt;10% magnitude. Large swings often reflect thin samples or property-mix changes \u2014 verify before citing specific causes.</p>
</section>

<section class="card"><h2>\U0001f465 Interested buyers (listing visits per property)</h2>
<div class="scroll"><table><thead><tr><th>Suburb</th><th>Visits/property</th><th>House price</th></tr></thead>
<tbody>{vis_rows}</tbody></table></div>
<p class="note">Engagement proxy, not a count of unique buyers or completed transactions. Top 15 shown.</p>
</section>

<section class="card"><h2>\U0001f4cd House price clusters</h2>
<div>{bands_html}</div>
<p class="note">Count of suburbs by house price band. Not weighted by listing volume or sales count.</p>
</section>

<div class="caution">\u26a0\ufe0f <strong>Data limitations:</strong> Prices are suburb-level REA medians, not transaction-weighted LGA medians. YoY/MoM use prior-period REA figures. "Listing visits per property" is an engagement proxy. No sales counts, zoning or feasibility data included. Missing unit median = insufficient data, not zero stock. Large % changes may reflect thin samples.</div>

<footer>Source: REA Group &middot; Central Coast &middot; {m} {y} &middot; Generated {gen} &middot; Not for investment advice</footer>
</main></body></html>"""

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
    ap.add_argument("month", help="e.g. Sep")
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
    print(f"  {len(data)} suburb rows, period {data[0]['Date']}")

    a = analyse(data, args.month, args.year)
    print(f"  Median house: {fmt_price(a['median_house'])} | >=1M: {a['n_above_1m']} | No unit: {a['n_no_unit']}")

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