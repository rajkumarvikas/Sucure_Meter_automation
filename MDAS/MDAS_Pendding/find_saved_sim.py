"""
Adds a "Saved Sim" column next to each meter in MDAS_result.xlsx.

Run:  python find_saved_sim.py
Reads MDAS_result.xlsx, writes MDAS_result_with_sim.xlsx (the original is not touched).

How the saved sim is found
  1. If the row has a "Sim Typed On Site" / "Sim Used" value or a "Sim used: <digits>" note in Message -> use it.
  2. Otherwise (older result files) it is worked out from the failed-sim list in Message:
     the meter's own sims are tried in order (Sim No, then Optional Sim). The first own sim that is
     NOT listed as failed is the one that was saved.
  3. If both own sims failed, a sim from another row (pool) was saved. Older result files did not
     record which one, so it is marked "UNKNOWN (pool sim)". Check that meter on the MDAS site,
     or look at the new "Sim Typed On Site" column produced by the latest mdas_auto.py.
"""
import re
import pandas as pd

SRC, DST = "MDAS_result.xlsx", "MDAS_result_with_sim.xlsx"


def digits(x):
    return re.sub(r"\D", "", str(x)) if pd.notna(x) else ""


def failed_sims(msg):
    return set(re.findall(r"(\d{15,20}):", str(msg) if pd.notna(msg) else ""))


def find_saved(row):
    if not str(row.get("Status", "")).startswith("SAVED"):
        return "", ""
    for col in ("Sim Typed On Site", "Sim Used"):
        v = digits(str(row.get(col, "")).split(" ")[0]) if col in row and pd.notna(row.get(col)) else ""
        if len(v) >= 15:
            return v, f"recorded in '{col}'"
    m = re.search(r"Sim used:\s*(\d{15,20})", str(row.get("Message", "")))
    if m:
        return m.group(1), "recorded in Message"
    failed = failed_sims(row.get("Message", ""))
    own = [digits(row.get("Sim No (Excel)")), digits(row.get("Optional Sim (Excel)"))]
    for s in own:
        if len(s) >= 15 and s not in failed:
            return s, "own Excel sim (not in the failed list)"
    return "UNKNOWN (pool sim)", "both own sims failed; the pool sim was not recorded"


df = pd.read_excel(SRC, dtype=str)
res = df.apply(find_saved, axis=1, result_type="expand")
df.insert(2, "Saved Sim", res[0])
df.insert(3, "Saved Sim Basis", res[1])
df.to_excel(DST, index=False)
saved = df[df["Status"].fillna("").str.startswith("SAVED")]
print(f"{len(saved)} saved meters")
print(f"  sim identified      : {(~saved['Saved Sim'].str.startswith('UNKNOWN')).sum()}")
print(f"  UNKNOWN (pool sim)  : {saved['Saved Sim'].str.startswith('UNKNOWN').sum()}")
print("Written ->", DST)
