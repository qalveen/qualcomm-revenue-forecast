"""Download Qualcomm earnings release PDFs, FY23 Q4 - FY26 Q3."""
import time, requests, pathlib

UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}
OUT = pathlib.Path("data/raw/earnings_releases")
OUT.mkdir(parents=True, exist_ok=True)
ORD = {1: "1st", 2: "2nd", 3: "3rd", 4: "4th"}

quarters = [(2023, 4)] + [(fy, q) for fy in (2024, 2025, 2026) for q in (1, 2, 3, 4)
                          if not (fy == 2026 and q == 4)]

for fy, q in quarters:
    url = (f"https://s204.q4cdn.com/645488518/files/doc_financials/{fy}/q{q}/"
           f"FY{fy}-{ORD[q]}-Quarter-Earnings-Release.pdf")
    fname = OUT / f"FY{fy}-Q{q}-earnings-release.pdf"
    resp = requests.get(url, headers=UA)
    if resp.ok and resp.content[:4] == b"%PDF":
        fname.write_bytes(resp.content)
        print(f"OK   {fname.name}")
    else:
        print(f"FAIL FY{fy} Q{q} ({resp.status_code}) -> {url}")
    time.sleep(0.5)
