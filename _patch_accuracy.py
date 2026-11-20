import pathlib, re
p = pathlib.Path('/home/azureuser/property-reports/property_market_report.py')
src = p.read_text()
changes = []

# Fix 1: mom_str sign — 'else ""' -> 'else "-"' on the mom_str line
if 'else "") + f' in src:
    src = src.replace('else "") + f', 'else "-") + f', 1)
    changes.append('Fix1: mom_str sign')
else:
    changes.append('Fix1: NOT FOUND')

# Fix 2: Noraville badge — +5.6% -> dynamic fmt_mom call placeholder
# We'll resolve it via post-string replace since change_cards is a literal string
old_badge = '        <span class="badge green">+5.6% from Sep $947,500</span>'
new_badge = '        <span class="badge green">{_noraville_mom_badge}</span>'
if old_badge in src:
    src = src.replace(old_badge, new_badge, 1)
    changes.append('Fix2: Noraville badge placeholder')
else:
    changes.append('Fix2: NOT FOUND')

# Fix 3: near_1m note — replace hardcoded with placeholder
old_note = '  <p class="note">Killarney Vale ($987,750, +1.1% MoM) is closest to the $1M threshold. Wyoming has held near $960k for three months with minimal momentum. Wyongah shows modest YoY growth (+18.8%) but retreated -1.3% MoM. All four remain genuine candidates if momentum returns.</p>'
new_note = '  <p class="note">{near_1m_note}</p>'
if old_note in src:
    src = src.replace(old_note, new_note, 1)
    changes.append('Fix3: near_1m_note placeholder')
else:
    changes.append('Fix3: NOT FOUND')

# Fix 4: add dynamic lookup block + near_1m_note builder before change_cards
# Find the change_cards marker and insert before it
marker = '    # \u2500\u2500 change cards (Sep\u2192Oct context) \u2500\u2500'
insert = '''    # \u2500\u2500 dynamic suburb lookups for change cards / near_1m_note \u2500\u2500
    _smap = {d["Suburb"]: d for d in a["data"]}
    def _fmt_mom(sub):
        v = fv(_smap.get(sub, {}), "MoM Change (%)")
        if v is None: return "N/A"
        return ("+" if v >= 0 else "") + f"{v*100:.1f}%"
    def _fmt_yoy(sub):
        v = fv(_smap.get(sub, {}), "House Price YoY (%)")
        if v is None: return "N/A"
        return ("+" if v >= 0 else "") + f"{v*100:.1f}%"
    _noraville_mom_badge = _fmt_mom("Noraville") + " MoM from Sep $947,500"
    if a["near_1m"]:
        _top = a["near_1m"][0]
        near_1m_note = (
            _top["Suburb"] + " (" + fmt_price(_top.get("House Price (Current)")) + ", " + _fmt_mom(_top["Suburb"]) + " MoM) is closest to the $1M threshold. "
            + "; ".join(
                nd["Suburb"] + " (" + fmt_price(nd.get("House Price (Current)")) + ", " + _fmt_mom(nd["Suburb"]) + " MoM, " + _fmt_yoy(nd["Suburb"]) + " YoY)"
                for nd in a["near_1m"][1:]
            ) + (" (others)." if len(a["near_1m"]) > 1 else ".")
            + " All remain candidates if momentum returns."
        )
    else:
        near_1m_note = "No suburbs currently in the \u2265$960k band."
'''
if marker in src:
    src = src.replace(marker, insert + marker, 1)
    changes.append('Fix4: dynamic lookup block inserted')
else:
    changes.append('Fix4: marker NOT FOUND')

# Fix 5: subtitle print CSS
old_css = 'header p{{color:#64748b;margin-top:6px;font-size:.95rem}}'
new_css = old_css + chr(10) + '@media print{{header p{{font-size:.78rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}}}'
if old_css in src and '@media print{{header p' not in src:
    src = src.replace(old_css, new_css, 1)
    changes.append('Fix5: header p print CSS')
elif '@media print{{header p' in src:
    changes.append('Fix5: already present')
else:
    changes.append('Fix5: NOT FOUND')

p.write_text(src)
for c in changes: print(c)
print('Done.')