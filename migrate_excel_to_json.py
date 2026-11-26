#!/usr/bin/env python3
"""
migrate_excel_to_json.py
One-time migration from the two Excel workbooks into pure Python/CSV/JSON under data/:
  - 2026-27 Men's College Basketball Stats_code.xlsx -> data/team_crosswalk.json + data/raw/*.csv
  - Batch_Simulation_Results.xlsx                    -> data/sims/YYYY-MM-DD.json
Then runs compute_stats.py to build data/stats_master.json.
"""

import json
import math
import os
import re
import openpyxl
import pandas as pd
from compute_stats import compute_master_stats


STATS_XLSX = "2026-27 Men's College Basketball Stats_code.xlsx"
SIM_XLSX = "Batch_Simulation_Results.xlsx"

SHEET_TO_CSV = {
    "Kenpom Ratings": "data/raw/kenpom.csv",
    "Sheet4": "data/raw/height.csv",
    "Warren ELO": "data/raw/warren_elo.csv",
    "ESPN BPI": "data/raw/espn_bpi.csv",
    "EvanMiya Regular Ratings": "data/raw/evanmiya_reg.csv",
    "EvanMiya Injured Ratings": "data/raw/evanmiya_inj.csv",
    "School Stats": "data/raw/cbbref_school.csv",
    "Opponent Stats": "data/raw/cbbref_opp.csv",
}

MODEL_SHEET_ALIASES = {
    "Live Morey Z": ("live morey", (1, 2, 19, 20)),
    "Consensus Morey Z": ("consensus morey", (4, 5, 21, 22)),
    "Live Log5": ("live log5", (7, 8, 23, 24)),
    "Consensus Log5": ("consensus log5", (10, 11, 25, 26)),
    "Adj. Eff.": ("adj eff", (13, 14, 27, 28)),
    "Base": ("regular", (16, 17, 29, 30)),
}


def _clean_str(v):
    return "" if v is None else str(v).strip()


def _opt_int(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)) and math.isfinite(v):
        return int(round(v))
    s = str(v).strip()
    if s.isdigit():
        return int(s)
    return None


def _prob_val(v):
    """Store win probabilities consistently as 0.0-1.0 floats (rounded to 4 decimals)."""
    if isinstance(v, bool) or v is None:
        return 0.0
    try:
        f = float(str(v).strip().replace("%", ""))
        if not math.isfinite(f):
            return 0.0
        if f > 1.0:
            f /= 100.0
        return round(f, 4)
    except ValueError:
        return 0.0


def _norm_acc(raw):
    s = _clean_str(raw).upper()
    if s in ("TEAM+DIFF", "TEAM + DIFF", "TEAM & DIFF"):
        return "TEAM AND DIFF"
    if s == "SUPER ALLS":
        return "SUPER ALL"
    return s


def migrate_stats_workbook():
    if not os.path.isfile(STATS_XLSX):
        print(f"Skipping stats migration: {STATS_XLSX} not found.")
        return

    os.makedirs("data/raw", exist_ok=True)
    wb = openpyxl.load_workbook(STATS_XLSX, data_only=True, read_only=True)
    ws = wb["Worksheet"]
    rows_ws = list(ws.iter_rows(min_row=3, values_only=True))
    wb.close()

    crosswalk = []
    for r in rows_ws:
        if not r or not isinstance(r[0], (int, float)) or not r[1]:
            continue
        kp = str(r[1]).strip()
        crosswalk.append({
            "Rk": int(round(r[0])),
            "Kenpom Name": kp,
            "CBBRef Name": str(r[2]).strip() if r[2] is not None else kp,
            "Warren Name": str(r[3]).strip() if r[3] is not None else kp.lower(),
            "ESPN Name": str(r[4]).strip() if r[4] is not None else kp,
            "EvanMiya Name": str(r[5]).strip() if r[5] is not None else kp,
            "School": str(r[6]).strip() if r[6] is not None else kp,
        })

    with open("data/team_crosswalk.json", "w", encoding="utf-8") as f:
        json.dump(crosswalk, f, indent=2, ensure_ascii=False)
    print(f"[Migrate] Wrote data/team_crosswalk.json ({len(crosswalk)} teams).")

    for sheet, csv_path in SHEET_TO_CSV.items():
        df = pd.read_excel(STATS_XLSX, sheet_name=sheet)
        df.to_csv(csv_path, index=False)
        print(f"[Migrate] {sheet} -> {csv_path} ({len(df)} rows)")


def migrate_sim_workbook():
    if not os.path.isfile(SIM_XLSX):
        print(f"Skipping sim migration: {SIM_XLSX} not found.")
        return

    os.makedirs("data/sims", exist_ok=True)
    wb = openpyxl.load_workbook(SIM_XLSX, data_only=True, read_only=True)
    date_sheets = sorted(s for s in wb.sheetnames if re.match(r"^\d{4}-\d{2}-\d{2}$", s))

    total_games = 0
    for d_sheet in date_sheets:
        ws = wb[d_sheet]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 3:
            continue

        hmap = {}
        for idx, val in enumerate(rows[0]):
            if val is not None:
                key = re.sub(r"\s+", " ", str(val).strip().lower())
                hmap[key] = idx

        school_col = hmap.get("school", 0)
        actual_col = hmap.get("actual score", 18)

        model_cols = {}
        for m_name, (prefix, fb) in MODEL_SHEET_ALIASES.items():
            model_cols[m_name] = (
                hmap.get(f"{prefix} win %", fb[0]),
                hmap.get(f"{prefix} score", fb[1]),
                hmap.get(f"{prefix} accuracy", fb[2]),
                hmap.get(f"{prefix} certainty", fb[3]),
            )

        slate_games = []
        i = 1
        while i + 1 < len(rows):
            r1 = rows[i]
            t1 = _clean_str(r1[school_col]) if r1 and school_col < len(r1) else ""
            if not t1:
                i += 1
                continue
            r2 = rows[i + 1]
            t2 = _clean_str(r2[school_col]) if r2 and school_col < len(r2) else ""
            if not t2:
                i += 1
                continue
            i += 2

            s1 = _opt_int(r1[actual_col]) if actual_col < len(r1) else None
            s2 = _opt_int(r2[actual_col]) if actual_col < len(r2) else None
            if s1 == 0 and s2 == 0:
                s1, s2 = None, None

            models = {}
            for m_name, (win_c, scr_c, acc_c, cert_c) in model_cols.items():
                if max(win_c, scr_c) >= len(r1) or max(win_c, scr_c) >= len(r2):
                    continue
                wp1 = _prob_val(r1[win_c])
                wp2 = _prob_val(r2[win_c])
                ps1 = _opt_int(r1[scr_c])
                ps2 = _opt_int(r2[scr_c])
                if ps1 is None or ps2 is None:
                    continue
                entry = {
                    "wp1": wp1,
                    "wp2": wp2,
                    "ps1": ps1,
                    "ps2": ps2,
                }
                acc = _norm_acc(r1[acc_c]) if acc_c < len(r1) else ""
                cert = _clean_str(r1[cert_c]) if cert_c < len(r1) else ""
                if s1 is not None and s2 is not None and acc:
                    entry["acc"] = acc
                    entry["cert"] = cert
                models[m_name] = entry

            if models:
                slate_games.append({
                    "t1": t1,
                    "t2": t2,
                    "s1": s1,
                    "s2": s2,
                    "models": models,
                })

        out_path = os.path.join("data", "sims", f"{d_sheet}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump({"date": d_sheet, "games": slate_games}, f, indent=2, ensure_ascii=False)
        total_games += len(slate_games)

    wb.close()
    print(f"[Migrate] Exported {len(date_sheets)} date sheets ({total_games} total games) -> data/sims/")


if __name__ == "__main__":
    migrate_stats_workbook()
    migrate_sim_workbook()
    compute_master_stats("data")
