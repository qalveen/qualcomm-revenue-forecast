import pandas as pd

a = pd.read_csv("data/processed/segment_actuals.csv")
g = pd.read_csv("data/processed/guidance.csv")

counts = a.groupby(["fiscal_year", "fiscal_quarter"]).size()
bad = counts[counts != 5]
print("Quarters missing rows:", "none" if bad.empty else f"\n{bad}")

qct = a[a.segment == "QCT"].pivot_table(
    index=["fiscal_year", "fiscal_quarter"], columns="sub_segment", values="revenue_musd")
qct["streams_sum"] = qct[["Handsets", "Automotive", "IoT"]].sum(axis=1)
qct["gap_pct"] = (qct["Total"] - qct["streams_sum"]) / qct["Total"] * 100
print("\nQCT stream tie-out:\n", qct[["Total", "streams_sum", "gap_pct"]].round(1))

g["calc_mid"] = (g.low_musd + g.high_musd) / 2
off = g[(g.calc_mid - g.midpoint_musd).abs() > 0.5]
print("\nGuidance midpoint errors:", "none" if off.empty else f"\n{off}")
print(f"\nRows: actuals={len(a)} (expect 55), guidance={len(g)} (expect 33)")
