#!/usr/bin/env python3
"""Render an existing property dashboard PDF with print-only SVG charts.

Usage: python3 render_property_pdf.py Dashboard.html Source.xlsx Insights.pdf
The dashboard is read but never written. Chart datasets must match the REA workbook.
"""
import argparse
import html
import json
import pathlib
import re
import tempfile

import weasyprint
from property_market_report import analyse, fmt_price, fv, load_data

CHARTS = {
    "gainersChart": ("gainersLabels", "gainersData", "YoY change (%)"),
    "losersChart": ("losersLabels", "losersData", "YoY change (%)"),
    "momChart": ("momLabels", "momData", "MoM change (%)"),
    "visChart": ("visLabels", "visData", "Listing visits per property"),
    "bandChart": ("bandLabels", "bandData", "Number of suburbs"),
}


def arrays_from_dashboard(source):
    names = [item for pair in CHARTS.values() for item in pair[:2]] + ["visPrices"]
    arrays = {}
    for name in names:
        matches = re.findall(r"\bconst\s+" + name + r"\s*=\s*(\[[^\n;]*\]);", source)
        if len(matches) != 1:
            raise ValueError(f"Expected exactly one dashboard array: {name}")
        arrays[name] = json.loads(matches[0])
    return arrays


def verify_workbook(arrays, workbook):
    data = load_data(workbook)
    a = analyse(data, "", "")
    pos = [d for d in a["big_yoy_h"] if (fv(d, "House Price YoY (%)") or 0) > 0][:10]
    neg = [d for d in a["big_yoy_h"] if (fv(d, "House Price YoY (%)") or 0) < 0]
    mom = a["big_mom_h"]
    visits = a["visits_sorted"][:15]
    expected = {
        "gainersLabels": [d["Suburb"] for d in reversed(pos)],
        "gainersData": [round(fv(d, "House Price YoY (%)") * 100, 1) for d in reversed(pos)],
        "losersLabels": [d["Suburb"] for d in reversed(neg)],
        "losersData": [round(fv(d, "House Price YoY (%)") * 100, 1) for d in reversed(neg)],
        "momLabels": [d["Suburb"] for d in reversed(mom)],
        "momData": [round(fv(d, "MoM Change (%)") * 100, 1) for d in reversed(mom)],
        "visLabels": [d["Suburb"] for d in reversed(visits)],
        "visData": [float(d["Listing Visits per Property"]) for d in reversed(visits)],
        "visPrices": [fmt_price(d.get("House Price (Current)")) for d in reversed(visits)],
        "bandLabels": [label for label, count in a["band_counts"]],
        "bandData": [count for label, count in a["band_counts"]],
    }
    for name, values in expected.items():
        if arrays[name] != values:
            raise ValueError(f"Dashboard {name} does not match the REA workbook")
    print(f"Verified all five chart datasets against {len(data)} workbook suburbs")


def svg_chart(labels, values, axis_title, chart_id):
    if not labels or len(labels) != len(values):
        raise ValueError(f"Empty/misaligned chart {chart_id}")
    width, left, right, row_height, top = 780, 260, 55, 28, 37
    bottom = top + row_height * len(labels)
    height = bottom + 37
    plot_width = width - left - right
    negative = chart_id == "losersChart"
    ceiling = max(abs(float(v)) for v in values) * 1.12
    def esc(value):
        return html.escape(str(value), quote=True)
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" role="img" '
             f'aria-label="{esc(axis_title)} chart" viewBox="0 0 {width} {height}" '
             f'width="{width}" height="{height}" '
             f'style="font-family:DejaVu Sans,sans-serif">',
             f'<rect width="{width}" height="{height}" fill="#fff"/>']
    for i in range(6):
        x = left + plot_width * i / 5
        amount = ceiling * (5 - i if negative else i) / 5
        tick = (f"-{amount:.0f}" if negative and amount else f"{amount:.0f}")
        parts.append(f'<line x1="{x:.1f}" y1="{top-9}" x2="{x:.1f}" y2="{bottom}" stroke="#e2e8f0"/>')
        parts.append(f'<text x="{x:.1f}" y="{bottom+15}" text-anchor="middle" fill="#64748b" font-size="11">{tick}</text>')
    palette = ["#64748b", "#2563eb", "#16a34a", "#d9a306", "#ea580c", "#b91c1c"]
    for i, (label, value) in enumerate(zip(labels, values)):
        value = float(value)
        length = plot_width * abs(value) / ceiling
        y = top + i * row_height
        start = left + (plot_width - length if negative else 0)
        color = (palette[i] if chart_id == "bandChart" else
                 "#dc2626" if negative or (chart_id == "gainersChart" and value > 25) else
                 "#2563eb" if chart_id == "visChart" else "#16a34a")
        shown = (f"{value:+.1f}%" if chart_id in ("gainersChart", "losersChart", "momChart") else
                 f"{int(value):,}")
        parts.append(f'<text x="{left-11}" y="{y+16}" text-anchor="end" fill="#334155" font-size="12">{esc(label)}</text>')
        parts.append(f'<rect x="{start:.1f}" y="{y+3}" width="{length:.1f}" height="19" rx="3" fill="{color}"/>')
        text_x = max(left + 2, start - 7) if negative else min(width - 3, start + length + 6)
        anchor = "end" if negative else ("end" if text_x == width - 3 else "start")
        parts.append(f'<text x="{text_x:.1f}" y="{y+17}" text-anchor="{anchor}" fill="#1e293b" font-size="11" font-weight="bold">{esc(shown)}</text>')
    parts.append(f'<text x="{left+plot_width/2:.1f}" y="{height-4}" text-anchor="middle" fill="#475569" font-size="11">{esc(axis_title)}</text>')
    parts.append('</svg>')
    return ''.join(parts)


def render(source, arrays):
    for chart_id, (labels_key, data_key, axis_title) in CHARTS.items():
        pattern = rf'<canvas id="{chart_id}"></canvas>'
        svg = svg_chart(arrays[labels_key], arrays[data_key], axis_title, chart_id)
        source, count = re.subn(pattern, lambda _: svg, source)
        if count != 1:
            raise ValueError(f"Expected exactly one canvas: {chart_id}; found {count}")
    print_css = '''<style media="print">
    @page { size: A4; margin: 12mm; }
    body { padding: 0; }
    .card { break-inside: avoid; page-break-inside: avoid; padding: 16px; margin-bottom: 14px; }
    footer { padding: 4px; margin-top: 0; }
    .chart-wrap, .chart-wrap.tall, .chart-wrap.sm { height: auto; }
    .chart-wrap svg { width: 100%; height: auto; display: block; }
    </style>'''
    return source.replace('</head>', print_css + '</head>', 1)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("dashboard", type=pathlib.Path)
    parser.add_argument("workbook", type=pathlib.Path)
    parser.add_argument("pdf", type=pathlib.Path)
    args = parser.parse_args()
    original = args.dashboard.read_text(encoding="utf-8")
    arrays = arrays_from_dashboard(original)
    verify_workbook(arrays, args.workbook)
    pdf_html = render(original, arrays)
    with tempfile.NamedTemporaryFile(prefix="property-chart-pdf-", suffix=".pdf",
                                     dir=args.pdf.parent, delete=False) as temporary:
        tmp = pathlib.Path(temporary.name)
    try:
        weasyprint.HTML(string=pdf_html, base_url=str(args.dashboard.parent)).write_pdf(str(tmp))
        if tmp.stat().st_size < 30_000:
            raise ValueError("Rendered PDF is unexpectedly small")
        tmp.replace(args.pdf)
    finally:
        tmp.unlink(missing_ok=True)
    print(f"Rendered {args.pdf} ({args.pdf.stat().st_size:,} bytes); dashboard not written")


if __name__ == "__main__":
    main()
