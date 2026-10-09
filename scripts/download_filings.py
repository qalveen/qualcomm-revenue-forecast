"""Download Financial_Report.xlsx for Qualcomm 10-Q/10-K filings FY24 Q1 - FY26 Q3."""
import time, requests, pathlib

UA = {"User-Agent": "Calvin Magsie calvin.magsie@gmail.com"}  # SEC requires contact info
CIK = "0000804328"
OUT = pathlib.Path("data/raw")
OUT.mkdir(parents=True, exist_ok=True)

subs = requests.get(f"https://data.sec.gov/submissions/CIK{CIK}.json", headers=UA).json()
r = subs["filings"]["recent"]

for form, acc, period in zip(r["form"], r["accessionNumber"], r["reportDate"]):
    if form not in ("10-Q", "10-K") or not ("2023-12-01" <= period <= "2026-07-31"):
        continue
    acc_nd = acc.replace("-", "")
    url = f"https://www.sec.gov/Archives/edgar/data/804328/{acc_nd}/Financial_Report.xlsx"
    fname = OUT / f"QCOM_{form}_{period}.xlsx"
    resp = requests.get(url, headers=UA)
    if resp.ok:
        fname.write_bytes(resp.content)
        print(f"OK   {fname.name}")
    else:
        print(f"FAIL {form} {period} ({resp.status_code}) -> {url}")
    time.sleep(0.5)  # stay under SEC rate limit
