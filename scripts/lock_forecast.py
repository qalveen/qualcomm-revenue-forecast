"""Snapshot FY26 Q4 forecast (Base/Bull/Bear) from the model into forecasts/."""
import csv, datetime
from openpyxl import load_workbook

wb = load_workbook("model/qcom_revenue_forecast_model.xlsx", data_only=True)
a, act, g = wb["Assumptions"], wb["Actuals"], wb["Guidance"]

prior = {s: act[f"J{r}"].value for s, r in [("Handsets", 9), ("Automotive", 10), ("IoT", 11), ("QTL", 13)]}  # FY25 Q4
drv = {"Handsets": 8, "Automotive": 10, "IoT": 12, "QTL": 14}
today = datetime.date.today().isoformat()
out = f"forecasts/FY26Q4_forecast_locked_{today}.csv"

rows = []
for sc, c in [("Base", "C"), ("Bull", "D"), ("Bear", "E")]:
    f = {s: prior[s] * (1 + a[f"{c}{drv[s]}"].value) for s in prior}
    qct = f["Handsets"] + f["Automotive"] + f["IoT"]
    rows.append({"scenario": sc, "handsets": round(f["Handsets"]), "automotive": round(f["Automotive"]),
                 "iot": round(f["IoT"]), "qct_revenue": round(qct), "qtl_revenue": round(f["QTL"]),
                 "qct_ebt": round(qct * a[f"{c}16"].value), "qtl_ebt": round(f["QTL"] * a[f"{c}18"].value),
                 "qct_guide_low": g["N14"].value, "qct_guide_high": g["N15"].value,
                 "qtl_guide_low": g["N19"].value, "qtl_guide_high": g["N20"].value, "locked_on": today})

with open(out, "w", newline="") as fh:
    w = csv.DictWriter(fh, fieldnames=rows[0].keys()); w.writeheader(); w.writerows(rows)
for r in rows:
    print(f"{r['scenario']:5}  QCT {r['qct_revenue']:>6,}  QTL {r['qtl_revenue']:>5,}  "
          f"Handsets {r['handsets']:>6,}  Auto {r['automotive']:>5,}  IoT {r['iot']:>5,}")
print("saved", out)
