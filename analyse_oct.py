import openpyxl, statistics, sys
path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
headers = [cell.value for cell in ws[1]]
rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None: continue
    rows.append(dict(zip(headers, row)))
print("rows:", len(rows))
def sf(v):
    try:
        if v is None or v == "": return None
        return float(v)
    except: return None
hp_all = [sf(r["House Price (Current)"]) for r in rows if sf(r["House Price (Current)"]) is not None]
up_all = [sf(r["Unit Price (Current)"]) for r in rows if sf(r["Unit Price (Current)"]) is not None]
yoy_all = [sf(r["House Price YoY (%)"]) for r in rows if sf(r["House Price YoY (%)"]) is not None]
print("med_house:", statistics.median(hp_all))
print("med_unit:", statistics.median(up_all))
print("mean_yoy:", round(statistics.mean(yoy_all)*100,2))
print("unit_count:", len(up_all), "unit_missing:", len(rows)-len(up_all))
over1m = [r for r in rows if r.get("$1M Median Suburb?") == "Y"]
print("over1m:", len(over1m))
near1m = [r for r in rows if sf(r["House Price (Current)"]) is not None and 960000 <= sf(r["House Price (Current)"]) < 1000000]
print("near1m:", len(near1m))
for r in near1m:
    hp=sf(r["House Price (Current)"]); mom=sf(r["MoM Change (%)"]); yoy=sf(r["House Price YoY (%)"])
    vis=sf(r["Listing Visits per Property"]); up=sf(r["Unit Price (Current)"])
    print("NEAR1M", r["Suburb"], int(hp), int(1000000-hp), round(mom*100 if mom else 0,2), round(yoy*100 if yoy else 0,2), vis, int(up) if up else "NA")
no_unit = [r for r in rows if sf(r["Unit Price (Current)"]) is None]
print("no_unit:", len(no_unit))
for r in no_unit:
    hp=sf(r["House Price (Current)"]); flag=r.get("$1M Median Suburb?","")
    print("NOUNIT", r["Suburb"], int(hp) if hp else 0, flag)
print("HMOM_BIG:")
for r in sorted(rows, key=lambda x: abs(sf(x["MoM Change (%)"]) or 0), reverse=True):
    v=sf(r["MoM Change (%)"])
    if v is not None and abs(v)>=0.10:
        print("HMOM", r["Suburb"], round(v*100,1), int(sf(r["House Price (Current)"])))
print("HYOY_BIG:")
for r in sorted(rows, key=lambda x: abs(sf(x["House Price YoY (%)"]) or 0), reverse=True):
    v=sf(r["House Price YoY (%)"])
    if v is not None and abs(v)>=0.10:
        print("HYOY", r["Suburb"], round(v*100,1), int(sf(r["House Price (Current)"])))
print("UMOM_BIG:")
for r in sorted(rows, key=lambda x: abs(sf(x["Unit MoM Change (%)"]) or 0), reverse=True):
    v=sf(r["Unit MoM Change (%)"])
    if v is not None and abs(v)>=0.10:
        up=sf(r["Unit Price (Current)"])
        print("UMOM", r["Suburb"], round(v*100,1), int(up) if up else "NA")
print("TOP15VIS:")
for r in sorted(rows, key=lambda x: sf(x["Listing Visits per Property"]) or 0, reverse=True)[:15]:
    v=sf(r["Listing Visits per Property"]); hp=sf(r["House Price (Current)"])
    up=sf(r["Unit Price (Current)"]); flag=r.get("$1M Median Suburb?","")
    print("VIS", r["Suburb"], v, int(hp), flag, int(up) if up else "NA")
print("BANDS:")
bm={}
for r in rows:
    hp=sf(r["House Price (Current)"])
    if hp is None: continue
    b=int(hp//700000)*700
    bm[b]=bm.get(b,0)+1
for b in sorted(bm): print("BAND", b, bm[b])
print("ALLSUB:")
for r in sorted(rows, key=lambda x: sf(x["House Price (Current)"]) or 0, reverse=True):
    hp=sf(r["House Price (Current)"]); up=sf(r["Unit Price (Current)"])
    yoy=sf(r["House Price YoY (%)"]); mom=sf(r["MoM Change (%)"])
    vis=sf(r["Listing Visits per Property"]); flag=r.get("$1M Median Suburb?","")
    rh=sf(r.get("House Rent (P/W)")); ru=sf(r.get("Unit Rent (P/W)"))
    ibt=sf(r.get("Interested Buyers Trend")); seg=r.get("Top Segment","")
    print("SUB", r["Suburb"], int(hp), round(mom*100 if mom else 0,2), round(yoy*100 if yoy else 0,2),
          int(up) if up else "NA", vis, flag, int(rh) if rh else 0, int(ru) if ru else 0, round(ibt,4) if ibt else "NA", str(seg).replace(" ","_"))