#!/usr/bin/env python3
"""Render an existing property dashboard PDF with print-only SVG charts.

Usage: python3 render_property_pdf.py Dashboard.html Source.xlsx Insights.pdf
The dashboard is read but never written. Chart datasets must match the REA workbook.
"""
import argparse
import html
import json
import math
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
    # Leave ample room for long suburb names inside the SVG viewBox.
    width, left, right, row_height, top = 820, 300, 55, 28, 37
    bottom = top + row_height * len(labels)
    height = bottom + 37
    plot_width = width - left - right
    negative = chart_id == "losersChart"
    ceiling = max(abs(float(v)) for v in values) * 1.12
    raw_step = ceiling / 5
    step = (math.ceil(raw_step / 100) * 100 if raw_step >= 100 else
            math.ceil(raw_step / 2) * 2 if raw_step >= 8 else math.ceil(raw_step))
    ceiling = step * 5
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
                 "#d97706" if chart_id == "gainersChart" and value > 25 else
                 "#dc2626" if negative else
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


def pdf_safe_text(source):
    """Repair presentation defects in the PDF copy; never touch the dashboard file."""
    def once(old, new):
        nonlocal source
        # Month-specific copy may be absent from other reports using the same renderer.
        if source.count(old) > 1:
            raise ValueError(f"Repeated PDF-only text pattern: {old[:75]!r}")
        source = source.replace(old, new, 1)

    once('<p class="note">[!]#9888; Gosford +50.4% YoY is highlighted; treat with caution as the cause cannot be confirmed without sales count data. Suburbs marked [!]#9888; have YoY changes &gt;25% which may reflect property-mix or other data effects.</p>',
         '<p class="note">Gosford +50.4% YoY is shown in amber; treat with caution as the cause cannot be confirmed without sales count data. YoY changes &gt;25% may reflect property-mix or other data effects.</p>')
    once('<p class="note">[!]#9888; marks suburbs with YoY change exceeding &pm;25% — elevated potential for data or property-mix effects; treat these figures with extra caution.</p>',
         '<p class="note">Gosford and Wagstaffe have YoY changes exceeding &pm;25% — elevated potential for data or property-mix effects; treat these figures with extra caution.</p>')
    once('[!]#9888; <strong>Data limitations:</strong>', '<strong>Data limitations:</strong>')
    source = source.replace('[!]#9888;', '<strong>Caution:</strong>')
    if '[!]#9888;' in source:
        raise ValueError('Malformed warning marker remained in PDF')

    # Inline emojis and the numeric light-bulb entity render as tofu in the PDF font.
    source = re.sub(r'<span class="ico">[^<]+</span>', '', source)
    source = re.sub(r'(<h2>)<span>&#\d+;</span>\s*', r'\1', source)
    once('🟢 Gaining', '<span class="pdf-gaining">Gaining</span>')
    source = source.replace('🟡 Steady', '<span class="pdf-steady">Steady</span>')
    once('&#128161; &ldquo;Listing visits per property&rdquo;',
         '<strong>Note:</strong> &ldquo;Listing visits per property&rdquo;')
    once('Gosford ⚠️', 'Gosford (caution)')
    once('Wagstaffe ⚠️', 'Wagstaffe (caution)')

    # A signed zero is not a loss. Restrict correction to Wyoming's table row and note.
    once('<strong>Wyoming</strong></td><td>$960,000</td><td>$40,000</td><td class="neg">-0.0%</td>',
         '<strong>Wyoming</strong></td><td>$960,000</td><td>$40,000</td><td class="pdf-neutral">0.0%</td>')
    once('Wyoming ($960,000, -0.0% MoM, +1.5% YoY) (others).',
         'Wyoming ($960,000, 0.0% MoM, +1.5% YoY).')
    for unsupported in ('🟢', '🟡', '💡', '⚠️', '&#128161;'):
        if unsupported in source:
            raise ValueError(f'Unsupported PDF marker remained: {unsupported}')
    return source


def render(source, arrays):
    for chart_id, (labels_key, data_key, axis_title) in CHARTS.items():
        pattern = rf'<canvas id="{chart_id}"></canvas>'
        svg = svg_chart(arrays[labels_key], arrays[data_key], axis_title, chart_id)
        source, count = re.subn(pattern, lambda _: svg, source)
        if count != 1:
            raise ValueError(f"Expected exactly one canvas: {chart_id}; found {count}")
    source = pdf_safe_text(source)
    print_css = '''<style media="print">
    @page { size: A4; margin: 12mm; }
    body { padding: 0; }
    .card { break-inside: avoid; page-break-inside: avoid; padding: 16px; margin-bottom: 14px; }
    footer { padding: 4px; margin-top: 0; }
    .pdf-gaining { color: #15803d; font-weight: 600; }
    .pdf-steady { color: #a16207; font-weight: 600; }
    .pdf-neutral { color: #475569; font-weight: 600; }
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
