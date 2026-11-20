p = '/home/azureuser/property-reports/property_market_report.py'
src = open(p).read()
old = '    html = html.replace("{change_cards}", change_cards)'
new = old + chr(10) + '    html = html.replace("{_noraville_mom_badge}", _noraville_mom_badge)'
if old in src:
    src = src.replace(old, new, 1)
    open(p,'w').write(src)
    print('replace call added')
else:
    print('NOT FOUND')