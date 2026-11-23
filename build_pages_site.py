#!/usr/bin/env python3
"""
build_pages_site.py
Generates docs/data.json from the 2026-27 CBB Stats and Batch Simulation workbooks
for the 4-page static site (index.html, models.html, climate.html, bracket.html).

Usage:
    python3 build_pages_site.py
    SKIP_BRACKETOLOGY_SYNC=1 python3 build_pages_site.py
"""

import argparse
from collections import Counter
import datetime
import difflib
import json
import math
import os
import re
import ssl
import subprocess
import sys
import urllib.request
import openpyxl
from bs4 import BeautifulSoup


MODEL_NAMES = [
    "Live Morey Z",
    "Consensus Morey Z",
    "Live Log5",
    "Consensus Log5",
    "Adj. Eff.",
    "Base",
]

CONF_FULL_NAMES = {
    "SEC": "SEC", "ACC": "ACC", "B10": "Big Ten", "B12": "Big 12",
    "BE": "Big East", "P12": "Pac-12", "MWC": "Mountain West", "A10": "Atlantic 10",
    "WCC": "WCC", "AAC": "American", "MVC": "Missouri Valley", "CUSA": "Conference USA",
    "MAC": "MAC", "SB": "Sun Belt", "CAA": "CAA", "BW": "Big West",
    "SC": "SoCon", "Ivy": "Ivy League", "BSky": "Big Sky", "Sum": "Summit League",
    "Horz": "Horizon League", "BSth": "Big South", "Slnd": "Southland",
    "ASUN": "ASUN", "UAC": "UAC", "MAAC": "MAAC", "OVC": "Ohio Valley",
    "Pat": "Patriot League", "AE": "America East", "MEAC": "MEAC",
    "SWAC": "SWAC", "NEC": "NEC",
}


def num(v, d=2, default=0.0):
    if isinstance(v, bool):
        return default
    if isinstance(v, (int, float)) and math.isfinite(v):
        return round(float(v), d)
    if isinstance(v, str):
        s = v.strip().replace("%", "").replace(",", "")
        try:
            f = float(s)
            if math.isfinite(f):
                return round(f, d)
        except ValueError:
            pass
    return default


def opt_int(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)) and math.isfinite(v):
        return int(round(v))
    if isinstance(v, str):
        s = v.strip().lstrip("#")
        if s.isdigit():
            return int(s)
    return None


def pct_val(v, d=1):
    """Convert either 0.549 or 54.9 into 54.9."""
    f = num(v, d + 2, 0.0)
    if 0.0 < abs(f) <= 1.0:
        return round(f * 100.0, d)
    return round(f, d)


def clean_str(v, default=""):
    if v is None:
        return default
    return str(v).strip()


def find_workbook(base_dir, candidates, keyword):
    for name in candidates:
        path = os.path.join(base_dir, name)
        if os.path.isfile(path):
            return path
    for fname in sorted(os.listdir(base_dir)):
        low = fname.lower()
        if low.endswith((".xlsx", ".xlsm")) and not fname.startswith("~$") and keyword in low:
            return os.path.join(base_dir, fname)
    return None


def get_last_git_commit_label(base_dir):
    try:
        out = subprocess.check_output(
            ["git", "log", "-1", "--format=%cd", "--date=format:%b %d, %Y"],
            cwd=base_dir,
            stderr=subprocess.DEVNULL,
            text=True,
        ).strip()
        if out:
            return out
    except Exception:
        pass
    return datetime.datetime.now().strftime("%b %d, %Y")


def build_header_map(ws, header_row=2):
    """Map normalized header text -> 0-based column index."""
    hmap = {}
    for row_idx in (header_row, 1):
        if row_idx > ws.max_row:
            continue
        row = next(ws.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True), ())
        for idx, val in enumerate(row):
            if val is not None:
                key = re.sub(r"\s+", " ", str(val).strip().lower())
                if key and key not in hmap:
                    hmap[key] = idx
    return hmap


def get_col(row, hmap, names, fallback_idx=None):
    for name in names:
        idx = hmap.get(name.lower())
        if idx is not None and idx < len(row):
            return row[idx]
    if fallback_idx is not None and fallback_idx < len(row):
        return row[fallback_idx]
    return None


def extract_teams(stats_path):
    wb = openpyxl.load_workbook(stats_path, data_only=True, read_only=True)
    if "Worksheet" not in wb.sheetnames:
        raise ValueError(f"'Worksheet' tab not found in {stats_path}")

    # 1. EvanMiya Regular Ratings (Rank=0, Team=1, Conf=2)
    em_reg_map = {}
    if "EvanMiya Regular Ratings" in wb.sheetnames:
        for r in wb["EvanMiya Regular Ratings"].iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 3 and r[1]:
                em_reg_map[clean_str(r[1])] = {
                    "em_rk": opt_int(r[0]),
                    "conf_short": clean_str(r[2], "D1"),
                }

    # 2. EvanMiya Injured Ratings (Rank=0, Rank With Injuries=1, Team=2)
    em_inj_map = {}
    for sname in ("EvanMiya Injured Ratings", "EvanMiya Injury Ratings"):
        if sname in wb.sheetnames:
            for r in wb[sname].iter_rows(min_row=2, values_only=True):
                if r and len(r) >= 3 and r[2]:
                    em_inj_map[clean_str(r[2])] = opt_int(r[1])
            break

    # 3. Kenpom Ratings (TeamName=1, RankAdjEM=15)
    kp_rk_map = {}
    if "Kenpom Ratings" in wb.sheetnames:
        for r in wb["Kenpom Ratings"].iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 16 and r[1]:
                kp_rk_map[clean_str(r[1])] = opt_int(r[15])

    # 4. Warren ELO (Team=0, Rank=3)
    elo_rk_map = {}
    if "Warren ELO" in wb.sheetnames:
        for r in wb["Warren ELO"].iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 4 and r[0]:
                elo_rk_map[clean_str(r[0]).lower()] = opt_int(r[3])

    # 5. ESPN BPI (Team=0, CONF=1, BPI RK=4, CONF W-L=9, WIN CONF%=10)
    bpi_map = {}
    if "ESPN BPI" in wb.sheetnames:
        for r in wb["ESPN BPI"].iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 11 and r[0]:
                bpi_map[clean_str(r[0]).lower()] = {
                    "conf": clean_str(r[1]),
                    "bpi_rk": opt_int(r[4]),
                    "proj_conf_wl": clean_str(r[9]),
                    "win_conf_pct": clean_str(r[10]),
                }

    # 6. Sheet4 (TeamName=1, SizeRank=3, Exp=16, Continuity=50)
    s4_map = {}
    if "Sheet4" in wb.sheetnames:
        for r in wb["Sheet4"].iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 51 and r[1]:
                s4_map[clean_str(r[1])] = {
                    "size_rk": opt_int(r[3]),
                    "exp": num(r[16], 2, None),
                    "continuity": num(r[50], 1, None),
                }

    ws = wb["Worksheet"]
    hmap = build_header_map(ws, header_row=2)
    teams = []

    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or not isinstance(r[0], (int, float)) or not r[1]:
            continue

        kp_name = clean_str(r[1])
        warren_name = clean_str(get_col(r, hmap, ["warren name"], 3), kp_name)
        espn_name = clean_str(get_col(r, hmap, ["espn name", "espn"], 4), kp_name)
        em_name = clean_str(get_col(r, hmap, ["evanmiya name", "em name"], 5), kp_name)

        em_info = em_reg_map.get(em_name) or em_reg_map.get(kp_name) or {}
        bpi_info = bpi_map.get(espn_name.lower()) or {}
        s4_info = s4_map.get(kp_name) or {}

        conf_short = em_info.get("conf_short", "D1")
        conf_full = CONF_FULL_NAMES.get(conf_short) or bpi_info.get("conf") or conf_short

        con_tempo = num(get_col(r, hmap, ["consensus tempo"], 172), 1)
        con_ortg = num(get_col(r, hmap, ["consensus ortg"], 173), 2)
        con_drtg = num(get_col(r, hmap, ["consensus drtg"], 174), 2)
        con_net = num(get_col(r, hmap, ["consensus net rtg"], 175), 2)
        if con_net == 0.0 and (con_ortg or con_drtg):
            con_net = round(con_ortg - con_drtg, 2)

        live_tempo = num(get_col(r, hmap, ["live tempo"], 177), 1) or con_tempo
        live_ortg = num(get_col(r, hmap, ["live ortg"], 178), 2)
        live_drtg = num(get_col(r, hmap, ["live drtg"], 179), 2)
        live_net = num(get_col(r, hmap, ["live net rtg"], 180), 2)
        if live_net == 0.0 and (live_ortg or live_drtg):
            live_net = round(live_ortg - live_drtg, 2)

        net_score = num(get_col(r, hmap, ["net score"], 182), 2)
        if net_score == 0.0 and (con_net or live_net):
            net_score = round((con_net + live_net) / 2.0, 2)

        inj_delta = num(get_col(r, hmap, ["injury delta"], 154), 4)
        if 0 < abs(inj_delta) < 0.25:
            inj_delta = round(inj_delta * 100.0, 2)
        else:
            inj_delta = round(inj_delta, 2)

        teams.append({
            "id": int(round(r[0])),
            "team": kp_name,
            "espn_name": espn_name,
            "conf": conf_full,
            "conf_short": conf_short,
            "g": int(num(get_col(r, hmap, ["g"], 7), 0)),
            "w": int(num(get_col(r, hmap, ["w"], 8), 0)),
            "l": int(num(get_col(r, hmap, ["l"], 9), 0)),
            "wl_pct": num(get_col(r, hmap, ["w-l%"], 10), 3),
            "srs": num(get_col(r, hmap, ["srs"], 11), 2),
            "sos": num(get_col(r, hmap, ["sos"], 12), 2),
            "sor": num(get_col(r, hmap, ["sor calc", "sor"], 13), 2),
            "ppg": num(get_col(r, hmap, ["tm."], 23), 1),
            "opp_ppg": num(get_col(r, hmap, ["opp."], 24), 1),
            "pd": num(get_col(r, hmap, ["pd"], 25), 1),
            "fg_pct": pct_val(get_col(r, hmap, ["fg%"], 37), 1),
            "two_pct": pct_val(get_col(r, hmap, ["2p%"], 40), 1),
            "three_pct": pct_val(get_col(r, hmap, ["3p%"], 43), 1),
            "ft_pct": pct_val(get_col(r, hmap, ["ft%"], 46), 1),
            "orb_pg": num(get_col(r, hmap, ["orb"], 47), 1),
            "second_c_pct": pct_val(get_col(r, hmap, ["2c%"], 48), 1),
            "trb": num(get_col(r, hmap, ["trb"], 49), 1),
            "ast": num(get_col(r, hmap, ["ast"], 50), 1),
            "stl": num(get_col(r, hmap, ["stl"], 51), 1),
            "blk": num(get_col(r, hmap, ["blk"], 52), 1),
            "tov": num(get_col(r, hmap, ["tov"], 53), 1),
            "efg_pct": pct_val(get_col(r, hmap, ["efg%"], 87), 1),
            "tov_pct": pct_val(get_col(r, hmap, ["tov%"], 88), 1),
            "orb_pct": pct_val(get_col(r, hmap, ["orb%"], 89), 1),
            "ftr": pct_val(get_col(r, hmap, ["ftr"], 80), 1),
            "opp_efg_pct": pct_val(get_col(r, hmap, ["opp efg%"], 108), 1),
            "opp_tov_pct": pct_val(get_col(r, hmap, ["opp tov%"], 109), 1),
            "opp_orb_pct": pct_val(get_col(r, hmap, ["opp orb%"], 110), 1),
            "opp_ftr": pct_val(get_col(r, hmap, ["opp ftr"], 101), 1),
            "avg_ht": num(get_col(r, hmap, ["avg. ht"], 130), 2),
            "eff_ht": num(get_col(r, hmap, ["eff. ht"], 131), 2),
            "adj_net_ff": num(get_col(r, hmap, ["adj. net ff"], 136), 2),
            "bpi_ortg": num(get_col(r, hmap, ["bpi ortg"], 138), 1),
            "bpi_drtg": num(get_col(r, hmap, ["bpi drtg"], 139), 1),
            "bpi_net": num(get_col(r, hmap, ["bpi net rtg"], 140), 1),
            "bpi_rk": bpi_info.get("bpi_rk"),
            "em_ortg": num(get_col(r, hmap, ["em adj. ortg"], 147), 2),
            "em_drtg": num(get_col(r, hmap, ["em adj. drtg"], 148), 2),
            "em_net": num(get_col(r, hmap, ["em adj. net rtg"], 149), 2),
            "em_rk": em_info.get("em_rk"),
            "em_inj_ortg": num(get_col(r, hmap, ["em adj. inj ortg"], 150), 2),
            "em_inj_drtg": num(get_col(r, hmap, ["em adj. inj drtg"], 151), 2),
            "em_inj_net": num(get_col(r, hmap, ["em adj. net inj rtg"], 152), 2),
            "em_inj_rk": em_inj_map.get(em_name) or em_inj_map.get(kp_name),
            "inj_delta_pct": inj_delta,
            "elo": num(get_col(r, hmap, ["elo"], 158), 1),
            "elo_rk": elo_rk_map.get(warren_name.lower()) or elo_rk_map.get(kp_name.lower()),
            "kp_ortg": num(get_col(r, hmap, ["adj. ortg"], 161), 2),
            "kp_drtg": num(get_col(r, hmap, ["adj. drtg"], 162), 2),
            "kp_net": num(get_col(r, hmap, ["adj. net rtg"], 163), 2),
            "kp_rk": kp_rk_map.get(kp_name),
            "kp_inj_ortg": num(get_col(r, hmap, ["kp inj ortg"], 168), 2),
            "kp_inj_drtg": num(get_col(r, hmap, ["kp inj drtg"], 169), 2),
            "kp_inj_net": num(get_col(r, hmap, ["kp inj net rtg"], 170), 2),
            "con_tempo": con_tempo,
            "con_ortg": con_ortg,
            "con_drtg": con_drtg,
            "con_net": con_net,
            "live_tempo": live_tempo,
            "live_ortg": live_ortg,
            "live_drtg": live_drtg,
            "live_net": live_net,
            "net_score": net_score,
            "size_rk": s4_info.get("size_rk"),
            "exp": s4_info.get("exp"),
            "continuity": s4_info.get("continuity"),
            "win_conf_pct": bpi_info.get("win_conf_pct", ""),
            "proj_conf_wl": bpi_info.get("proj_conf_wl", ""),
        })

    wb.close()

    # Compute deterministic ranks across all teams
    for i, t in enumerate(sorted(teams, key=lambda x: x["con_net"], reverse=True), 1):
        t["con_rk"] = i
    for i, t in enumerate(sorted(teams, key=lambda x: x["con_ortg"], reverse=True), 1):
        t["con_ortg_rk"] = i
    for i, t in enumerate(sorted(teams, key=lambda x: x["con_drtg"] if x["con_drtg"] > 0 else 999), 1):
        t["con_drtg_rk"] = i
    for i, t in enumerate(sorted(teams, key=lambda x: x["live_net"], reverse=True), 1):
        t["live_rk"] = i
    for i, t in enumerate(sorted(teams, key=lambda x: x["live_ortg"], reverse=True), 1):
        t["live_ortg_rk"] = i
    for i, t in enumerate(sorted(teams, key=lambda x: x["live_drtg"] if x["live_drtg"] > 0 else 999), 1):
        t["live_drtg_rk"] = i

    teams.sort(key=lambda x: x["con_rk"])
    return teams


def parse_score_pair(val):
    if val is None:
        return None
    s = str(val).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*$", s)
    if m:
        return int(round(float(m.group(1)))), int(round(float(m.group(2))))
    return None


MODEL_SHEET_ALIASES = {
    "Live Morey Z": ("live morey", (1, 2, 19, 20)),
    "Consensus Morey Z": ("consensus morey", (4, 5, 21, 22)),
    "Live Log5": ("live log5", (7, 8, 23, 24)),
    "Consensus Log5": ("consensus log5", (10, 11, 25, 26)),
    "Adj. Eff.": ("adj eff", (13, 14, 27, 28)),
    "Base": ("regular", (16, 17, 29, 30)),
}


def normalize_accuracy(raw):
    s = clean_str(raw).upper()
    if s in ("TEAM+DIFF", "TEAM + DIFF", "TEAM & DIFF"):
        return "TEAM AND DIFF"
    if s == "SUPER ALLS":
        return "SUPER ALL"
    return s


def extract_simulations(sim_path):
    wb = openpyxl.load_workbook(sim_path, data_only=True, read_only=True)
    date_sheets = sorted(s for s in wb.sheetnames if re.match(r"^\d{4}-\d{2}-\d{2}$", s))
    games = []
    upcoming_games = []
    today_str = datetime.date.today().isoformat()

    for d_sheet in date_sheets:
        ws = wb[d_sheet]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 3:
            continue

        # Map Row 1 headers (normalized) -> column index
        hmap = {}
        for idx, val in enumerate(rows[0]):
            if val is not None:
                key = re.sub(r"\s+", " ", str(val).strip().lower())
                hmap[key] = idx

        school_col = hmap.get("school", 0)
        actual_col = hmap.get("actual score", 18)

        # Resolve (win_col, score_col, acc_col, cert_col) per model
        model_cols = {}
        for m_name, (prefix, fb) in MODEL_SHEET_ALIASES.items():
            win_c = hmap.get(f"{prefix} win %", fb[0])
            scr_c = hmap.get(f"{prefix} score", fb[1])
            acc_c = hmap.get(f"{prefix} accuracy", fb[2])
            cert_c = hmap.get(f"{prefix} certainty", fb[3])
            model_cols[m_name] = (win_c, scr_c, acc_c, cert_c)

        # Walk 2-row game blocks (Row i = Team 1, Row i+1 = Team 2, blank row between games)
        i = 1
        while i + 1 < len(rows):
            r1 = rows[i]
            t1 = clean_str(r1[school_col]) if r1 and school_col < len(r1) else ""
            if not t1:
                i += 1
                continue

            r2 = rows[i + 1]
            t2 = clean_str(r2[school_col]) if r2 and school_col < len(r2) else ""
            if not t2:
                i += 1
                continue

            i += 2  # consume both rows of this game

            s1 = opt_int(r1[actual_col]) if actual_col < len(r1) else None
            s2 = opt_int(r2[actual_col]) if actual_col < len(r2) else None
            has_final_score = (s1 is not None and s2 is not None and not (s1 == 0 and s2 == 0))

            if not has_final_score and d_sheet < today_str:
                continue  # skip old ungraded rows

            m_dict = {}
            for m_name, (win_c, scr_c, acc_c, cert_c) in model_cols.items():
                if max(win_c, scr_c) >= len(r1) or max(win_c, scr_c) >= len(r2):
                    continue
                wp1 = pct_val(r1[win_c], 1)
                wp2 = pct_val(r2[win_c], 1)
                ps1 = opt_int(r1[scr_c])
                ps2 = opt_int(r2[scr_c])
                if ps1 is None or ps2 is None:
                    continue
                acc = normalize_accuracy(r1[acc_c]) if acc_c < len(r1) else ""
                cert = clean_str(r1[cert_c]) if cert_c < len(r1) else ""
                if not cert:
                    max_wp = max(wp1, wp2)
                    cert = "90%+" if max_wp >= 90 else ("75%+" if max_wp >= 75 else ("60%+" if max_wp >= 60 else "50%+"))

                if has_final_score:
                    if acc:
                        m_dict[m_name] = [wp1, wp2, ps1, ps2, acc, cert]
                else:
                    pick = t1 if wp1 >= wp2 else t2
                    m_dict[m_name] = [wp1, wp2, ps1, ps2, pick, cert]

            if m_dict:
                if has_final_score:
                    games.append({
                        "d": d_sheet,
                        "t1": t1,
                        "t2": t2,
                        "s1": s1,
                        "s2": s2,
                        "m": m_dict,
                    })
                else:
                    upcoming_games.append({
                        "d": d_sheet,
                        "t1": t1,
                        "t2": t2,
                        "m": m_dict,
                    })

    # Extract raw summary sheet rows for the "Raw Summary Sheet Cells" mode
    sheet_model_daily = {}
    for m_name in MODEL_NAMES:
        sheet_model_daily[m_name] = []
        if m_name not in wb.sheetnames:
            continue
        ws_m = wb[m_name]
        for r in ws_m.iter_rows(min_row=5, values_only=True):
            if not r or not r[0]:
                continue
            d_val = r[0]
            d_str = d_val.strftime("%Y-%m-%d") if isinstance(d_val, (datetime.datetime, datetime.date)) else clean_str(d_val)
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", d_str):
                continue

            sheet_model_daily[m_name].append({
                "date": d_str,
                "games": int(num(r[1], 0)),
                "90%+": pct_val(r[2], 2) if r[2] is not None else None,
                "75%+": pct_val(r[3], 2) if r[3] is not None else None,
                "60%+": pct_val(r[4], 2) if r[4] is not None else None,
                "50%+": pct_val(r[5], 2) if r[5] is not None else None,
                "SUPER ALLS": int(num(r[6], 0)),
                "ALL": pct_val(r[7], 2),
                "TEAM AND DIFF": pct_val(r[8], 2),
                "TEAM": pct_val(r[9], 2),
                "CORRECT": pct_val(r[10], 2),
                "TOTAL": pct_val(r[11], 2),
                "WRONG": pct_val(r[12], 2),
                "90_r": int(num(r[13], 0)) if len(r) > 13 else 0,
                "90_t": int(num(r[14], 0)) if len(r) > 14 else 0,
                "75_r": int(num(r[15], 0)) if len(r) > 15 else 0,
                "75_t": int(num(r[16], 0)) if len(r) > 16 else 0,
                "60_r": int(num(r[17], 0)) if len(r) > 17 else 0,
                "60_t": int(num(r[18], 0)) if len(r) > 18 else 0,
                "50_r": int(num(r[19], 0)) if len(r) > 19 else 0,
                "50_t": int(num(r[20], 0)) if len(r) > 20 else 0,
                "all_c": int(num(r[21], 0)) if len(r) > 21 else 0,
                "td_c": int(num(r[22], 0)) if len(r) > 22 else 0,
                "tm_c": int(num(r[23], 0)) if len(r) > 23 else 0,
                "corr_c": int(num(r[24], 0)) if len(r) > 24 else 0,
                "tot_c": int(num(r[25], 0)) if len(r) > 25 else 0,
                "wrong_c": int(num(r[26], 0)) if len(r) > 26 else 0,
            })

    wb.close()
    return games, upcoming_games, sheet_model_daily, date_sheets




BRACKET_ALIASES = {
    "uconn": "Connecticut", "conn": "Connecticut",
    "ole miss": "Mississippi", "miss": "Mississippi",
    "nc state": "N.C. State", "ncst": "N.C. State", "n.c. state": "N.C. State",
    "wv": "West Virginia", "wvu": "West Virginia",
    "miami": "Miami FL", "miami (fl)": "Miami FL", "miami fl": "Miami FL", "mia": "Miami FL",
    "miami (oh)": "Miami OH", "miami (ohio)": "Miami OH", "m-oh": "Miami OH",
    "va. tech": "Virginia Tech", "va tech": "Virginia Tech", "vt": "Virginia Tech",
    "n. carolina": "North Carolina", "unc": "North Carolina",
    "okla. st.": "Oklahoma St.", "okst": "Oklahoma St.", "oklahoma state": "Oklahoma St.",
    "st. john's": "St. John's", "saint john's": "St. John's", "sju": "St. John's",
    "saint mary's": "Saint Mary's", "st. mary's": "Saint Mary's", "smc": "Saint Mary's",
    "saint louis": "Saint Louis", "st. louis": "Saint Louis", "slu": "Saint Louis",
    "cincy": "Cincinnati", "cin": "Cincinnati",
    "pitt": "Pittsburgh", "smu": "SMU", "tcu": "TCU", "ucf": "UCF", "vcu": "VCU",
    "lsu": "LSU", "byu": "BYU", "ucla": "UCLA", "usc": "USC", "unlv": "UNLV", "utep": "UTEP",
    "xav": "Xavier", "iowa": "Iowa", "marq": "Marquette", "pur": "Purdue",
    "fsu": "Florida St.", "ind": "Indiana", "nova": "Villanova", "vill": "Villanova",
    "baylor": "Baylor", "bay": "Baylor", "ark": "Arkansas", "aub": "Auburn",
    "mizzou": "Missouri", "mizz": "Missouri", "neb": "Nebraska", "wis": "Wisconsin",
    "msu": "Michigan St.", "mich. st.": "Michigan St.", "mich": "Michigan",
    "osu": "Ohio St.", "ohio state": "Ohio St.", "isu": "Iowa St.", "iowa state": "Iowa St.",
    "ksu": "Kansas St.", "kansas state": "Kansas St.", "ttu": "Texas Tech", "tx tech": "Texas Tech",
    "tamu": "Texas A&M", "texas a&m": "Texas A&M", "vandy": "Vanderbilt", "tenn": "Tennessee",
    "bama": "Alabama", "ala": "Alabama", "uk": "Kentucky", "uga": "Georgia", "uf": "Florida",
    "uva": "Virginia", "lou": "Louisville", "clem": "Clemson", "nd": "Notre Dame", "gt": "Georgia Tech",
    "bc": "Boston College", "cal": "California", "stan": "Stanford", "wake": "Wake Forest",
    "syr": "Syracuse", "dep": "DePaul", "prov": "Providence", "creigh": "Creighton",
    "hall": "Seton Hall", "gtown": "Georgetown", "but": "Butler",
    "asu": "Arizona St.", "arizona state": "Arizona St.", "colo": "Colorado",
    "ore": "Oregon", "wash": "Washington", "rutg": "Rutgers", "minn": "Minnesota",
    "psu": "Penn St.", "nw": "Northwestern", "mary": "Maryland", "md": "Maryland",
    "hou": "Houston", "ku": "Kansas", "ariz": "Arizona", "ill": "Illinois", "duke": "Duke",
    "gonz": "Gonzaga", "utst": "Utah St.", "usu": "Utah St.", "sdsu": "San Diego St.",
    "gcu": "Grand Canyon", "unm": "New Mexico", "boise": "Boise St.", "nev": "Nevada",
    "w. kentucky": "Western Kentucky", "wku": "Western Kentucky",
    "ucsb": "UC Santa Barbara", "uncw": "UNC Wilmington", "uci": "UC Irvine", "ucsd": "UC San Diego",
    "cook": "Bethune Cookman", "bethune-cookman": "Bethune Cookman", "bcu": "Bethune Cookman",
    "tnst": "Tennessee St.", "tsu": "Tennessee St.",
    "hwrd": "Howard", "how": "Howard",
    "liu": "LIU", "long island": "LIU", "long island university": "LIU",
    "harv": "Harvard", "penn": "Penn", "pennsylvania": "Penn",
    "jmu": "James Madison", "queen": "Queens", "queens (nc)": "Queens", "queens university": "Queens",
    "siena": "Siena", "cark": "Central Arkansas", "uca": "Central Arkansas", "c arkansas": "Central Arkansas",
    "stthmn": "St. Thomas", "st. thomas (mn)": "St. Thomas", "st thomas": "St. Thomas",
    "amer": "American", "american university": "American",
    "uvm": "Vermont", "etsu": "East Tennessee St.", "east tennessee state": "East Tennessee St.",
    "mcneese state": "McNeese", "mcneese st.": "McNeese", "mcn": "McNeese",
    "sfa": "Stephen F. Austin", "sf austin": "Stephen F. Austin",
    "illinois-chicago": "Illinois Chicago", "uic": "Illinois Chicago",
    "ut martin": "Tennessee Martin", "utm": "Tennessee Martin",
    "semo": "Southeast Missouri", "se missouri st.": "Southeast Missouri", "southeast missouri state": "Southeast Missouri",
    "a&m-corpus christi": "Texas A&M Corpus Chris", "texas a&m-cc": "Texas A&M Corpus Chris",
    "texas a&m-corpus christi": "Texas A&M Corpus Chris", "amcc": "Texas A&M Corpus Chris",
    "hpu": "High Point", "murr": "Murray St.", "wrst": "Wright St.",
    "sdak": "South Dakota St.", "ndsu": "North Dakota St.",
    "fau": "Florida Atlantic", "usf": "South Florida", "mem": "Memphis", "wich": "Wichita St.",
    "unt": "North Texas", "char": "Charlotte", "tul": "Tulsa", "cofc": "Charleston",
    "col. of charleston": "Charleston", "college of charleston": "Charleston",
    "sam": "Samford", "fur": "Furman", "lib": "Liberty", "kenn": "Kennesaw St.",
    "lip": "Lipscomb", "mont": "Montana", "mtst": "Montana St.", "montana state": "Montana St.",
    "akr": "Akron", "tol": "Toledo", "bgsu": "Bowling Green", "app": "Appalachian St.",
    "app state": "Appalachian St.", "odu": "Old Dominion", "troy": "Troy",
    "gwu": "George Washington", "gw": "George Washington", "gmu": "George Mason",
    "day": "Dayton", "rich": "Richmond", "dav": "Davidson", "uri": "Rhode Island",
    "st. bona": "St. Bonaventure", "st. bonaventure": "St. Bonaventure", "sbu": "St. Bonaventure",
    "loyola-chicago": "Loyola Chicago", "loyola (chi)": "Loyola Chicago", "loyola chi": "Loyola Chicago",
    "st. joe's": "Saint Joseph's", "saint joseph's": "Saint Joseph's",
    "scu": "Santa Clara", "csun": "CSUN", "cal state northridge": "CSUN",
    "utah valley": "Utah Valley", "uvu": "Utah Valley",
    "miami-ohio": "Miami OH",
    "texas-arlington": "UT Arlington",
    "massachusetts-lowell": "UMass Lowell",
    "omaha": "Nebraska Omaha", "seattle u": "Seattle", "sam houston": "Sam Houston St.",
    "grambling": "Grambling St.", "southern u.": "Southern", "army west point": "Army",
    "lamar university": "Lamar", "alcorn": "Alcorn St.", "prairie view": "Prairie View A&M",
    "fdu": "Fairleigh Dickinson", "fgcu": "Florida Gulf Coast", "uiw": "Incarnate Word",
    "ulm": "Louisiana Monroe", "umes": "Maryland Eastern Shore", "utrgv": "UT Rio Grande Valley",
    "ualbany": "Albany", "albany (ny)": "Albany", "umass": "Massachusetts",
    "central conn. st.": "Central Connecticut", "california baptist": "Cal Baptist",
    "csu bakersfield": "Cal St. Bakersfield", "lmu": "Loyola Marymount",
    "lmu (ca)": "Loyola Marymount", "siu edwardsville": "SIUE", "uncg": "UNC Greensboro",
    "unca": "UNC Asheville", "south fla.": "South Florida", "north fla.": "North Florida",
    "southern miss.": "Southern Miss", "gardner-webb": "Gardner Webb",
    "ark.-pine bluff": "Arkansas Pine Bluff", "texas a&m-commerce": "East Texas A&M",
    "tex. a&m-commerce": "East Texas A&M", "mcneese": "McNeese", "nicholls": "Nicholls",
    "southern california": "USC",
    "southern cal": "USC",
    "uni": "Northern Iowa",
    "new orleans": "LSU New Orleans",
    "southeast mo. st.": "Southeast Missouri",
    "boston u.": "Boston University",
    "niu": "Northern Illinois",
    "mississippi val.": "Mississippi Valley St.",
    "fla": "Florida", "tex": "Texas", "ari": "Arizona",
    "mia-fl": "Miami FL", "txtch": "Texas Tech", "stjn": "St. John's",
    "santa barbara": "UC Santa Barbara"
}

NON_TEAM_LABELS = {
    "midwest", "east", "south", "west", "first four", "opening round", "final four",
    "minneapolis", "fort worth", "indianapolis", "greenville", "tampa",
    "oklahoma city", "lexington", "birmingham", "raleigh",
}

DIRECTIONAL_PREFIXES = [
    (r"^c\.?\s+", ["central "]),
    (r"^n\.?\s+", ["north ", "northern "]),
    (r"^s\.?\s+", ["south ", "southern "]),
    (r"^e\.?\s+", ["east ", "eastern "]),
    (r"^w\.?\s+", ["west ", "western "]),
    (r"^se\.?\s+", ["southeast ", "southeastern "]),
    (r"^sw\.?\s+", ["southwest ", "southwestern "]),
    (r"^ne\.?\s+", ["northeast ", "northeastern "]),
    (r"^nw\.?\s+", ["northwest ", "northwestern "]),
    (r"^g\.?\s+", ["george ", "georgia "]),
    (r"^m\.?\s+", ["middle "]),
]

NCAA_WORD_EXPANSIONS = [
    (r"(?<!\w)n\.c\.(?!\w)", "north carolina"),
    (r"(?<!\w)s\.c\.(?!\w)", "south carolina"),
    (r"(?<!\w)fla(?:\.|(?=\s|$))", "florida"),
    (r"(?<!\w)ga(?:\.|(?=\s|$))", "georgia"),
    (r"(?<!\w)ky(?:\.|(?=\s|$))", "kentucky"),
    (r"(?<!\w)ill(?:\.|(?=\s|$))", "illinois"),
    (r"(?<!\w)mich(?:\.|(?=\s|$))", "michigan"),
    (r"(?<!\w)tenn(?:\.|(?=\s|$))", "tennessee"),
    (r"(?<!\w)ala(?:\.|(?=\s|$))", "alabama"),
    (r"(?<!\w)ark(?:\.|(?=\s|$))", "arkansas"),
    (r"(?<!\w)caro(?:\.|(?=\s|$))", "carolina"),
    (r"(?<!\w)wash(?:\.|(?=\s|$))", "washington"),
    (r"(?<!\w)colo(?:\.|(?=\s|$))", "colorado"),
    (r"(?<!\w)ariz(?:\.|(?=\s|$))", "arizona"),
    (r"(?<!\w)conn(?:\.|(?=\s|$))", "connecticut"),
    (r"(?<!\w)la(?:\.|(?=\s|$))", "louisiana"),
    (r"(?<!\w)miss(?:\.|(?=\s|$))", "mississippi"),
    (r"(?<!\w)ind(?:\.|(?=\s|$))", "indiana"),
    (r"(?<!\w)so(?:\.|(?=\s|$))", "southern"),
    (r"(?<!\w)mo(?:\.|(?=\s|$))", "missouri"),
    (r"(?<!\w)okla(?:\.|(?=\s|$))", "oklahoma"),
    (r"(?<!\w)tex(?:\.|(?=\s|$))", "texas"),
]

def _build_ssl_context():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl._create_unverified_context()


SSL_CTX = _build_ssl_context()


def _fetch_html(url, timeout=12):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout, context=SSL_CTX) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def _fetch_html_playwright(url, wait_ms=2500, timeout_ms=25000):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(
                user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
            )
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            page.wait_for_timeout(wait_ms)
            return page.content()
        finally:
            browser.close()


def build_team_matcher(teams):
    lookup = {}

    # Pass 1: Lock in exact canonical KenPom names & safe punctuation/St./State variants
    for t in teams:
        canon = t["team"]
        for v in (
            canon.lower(),
            canon.lower().replace("st.", "state"),
            canon.lower().replace("st.", "st"),
            canon.lower().replace("saint ", "st. "),
            canon.lower().replace("st. ", "saint "),
            canon.lower().replace(".", ""),
            re.sub(r"[^a-z0-9 ]+", "", canon.lower()).strip(),
        ):
            if v:
                lookup[v] = canon
        if t.get("espn_name"):
            lookup.setdefault(t["espn_name"].lower(), canon)

    # Pass 2: Only add 1-word-mascot-stripped ESPN names if unique and not already taken
    espn_minus1 = []
    for t in teams:
        en = (t.get("espn_name") or "").lower()
        parts = en.split()
        if len(parts) >= 2:
            espn_minus1.append((" ".join(parts[:-1]), t["team"]))
    counts = Counter(k for k, _ in espn_minus1)
    for k, canon in espn_minus1:
        if counts[k] == 1 and k not in lookup:
            lookup[k] = canon

    all_keys = list(lookup.keys())

    def _try_direct(low):
        if low in BRACKET_ALIASES:
            return BRACKET_ALIASES[low]
        if low in lookup:
            return lookup[low]
        norm = re.sub(r"\bstate\b", "st.", low).strip()
        if norm in lookup:
            return lookup[norm]
        norm2 = re.sub(r"\bst\b", "st.", low).strip()
        if norm2 in lookup:
            return lookup[norm2]
        stripped = re.sub(r"[^a-z0-9 ]+", "", low).strip()
        if stripped in BRACKET_ALIASES:
            return BRACKET_ALIASES[stripped]
        if stripped in lookup:
            return lookup[stripped]
        return None

    def match_team(raw):
        s = re.sub(r"\([^)]*\)", "", raw)
        s = re.sub(r"\b(AQ|At-Large|automatic qualifier|play-in|first four|opening round)\b", "", s, flags=re.I)
        s = re.sub(r"\s+", " ", s).strip(" -*•†‡#\u00a0")
        if not s or len(s) < 2:
            return None
        low = s.lower()
        if low in NON_TEAM_LABELS:
            return None

        hit = _try_direct(low)
        if hit:
            return hit

        for pat, replacements in DIRECTIONAL_PREFIXES:
            if re.match(pat, low):
                for rep in replacements:
                    hit = _try_direct(re.sub(pat, rep, low))
                    if hit:
                        return hit

        # Expand NCAA.com abbreviated state/direction tokens ("Western Ky.", "Fla. Atlantic", etc.)
        expanded = low.replace("-", " ")
        for pat, rep in NCAA_WORD_EXPANSIONS:
            expanded = re.sub(pat, rep, expanded)
        expanded = re.sub(r"\s+", " ", expanded).strip()
        if expanded != low:
            hit = _try_direct(expanded)
            if hit:
                return hit

        stripped = re.sub(r"[^a-z0-9 ]+", "", low).strip()
        if len(stripped) >= 5:
            close = difflib.get_close_matches(stripped, all_keys, n=1, cutoff=0.90)
            if close:
                return lookup[close[0]]
        return None

    return match_team


def _parse_seed_team_lines(text_lines, match_team):
    parsed = {}
    unmatched = []
    line_re = re.compile(r"^\s*#?(\d{1,2})\.?\s+(.+?)\s*$")
    for line in text_lines:
        m = line_re.match(line)
        if not m:
            continue
        seed = int(m.group(1))
        if not (1 <= seed <= 16):
            continue
        parts = re.split(r"\s*(?:/|\||\bvs\.?\b)\s*", m.group(2))
        for raw_part in parts[:2]:
            cleaned = re.sub(r"\([^)]*\)", "", raw_part).strip(" -*•†‡#\u00a0")
            if not cleaned or cleaned.lower() in NON_TEAM_LABELS or cleaned.isdigit():
                continue
            canon = match_team(cleaned)
            if canon:
                parsed.setdefault(canon, {"seed": seed})
            elif len(cleaned) <= 35:
                unmatched.append(f"{seed}: {cleaned}")
    return parsed, unmatched

SEED_MAX_TEAMS = {s: (8 if s in (10, 11, 12, 15, 16) else 4) for s in range(1, 17)}

def _extract_espn_bracket_from_html(html, match_team):
    soup = BeautifulSoup(html, "html.parser")

    # 1. Remove Bracket Watch <figure> cards, pod-city <p> headers (Omaha, Louisville, etc.), and chrome
    for el in soup.select("figure, svg, script, style, nav, header, footer"):
        el.decompose()
    for p in soup.select("p"):
        cls = " ".join(p.get("class") or [])
        if "text-muted" in cls or "uppercase" in cls:
            p.decompose()

    raw_text = soup.get_text("\n")
    # Cut off the bubble section at the bottom (use rfind so top-of-page labels don't truncate the bracket)
    for cutoff in ("First Four Out", "Next Four Out", "Multi-Bid", "Conference Breakdown", "Women's Bracketology"):
        idx = raw_text.rfind(cutoff)
        if idx > 1500:
            raw_text = raw_text[:idx]
    
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    parsed = {}
    seed_counts = Counter()
    unmatched = []

    def _add(team_name, seed):
        if not team_name or team_name in parsed:
            return False
        if seed_counts[seed] >= SEED_MAX_TEAMS[seed]:
            return False
        parsed[team_name] = {"seed": seed}
        seed_counts[seed] += 1
        return True

    i = 0
    while i < len(lines):
        if re.fullmatch(r"\d{1,2}\.?", lines[i]):
            seed = int(lines[i].rstrip("."))
            if 1 <= seed <= 16 and i + 1 < len(lines):
                nxt = lines[i + 1]
                # Skip trailing seed-movement digits (always followed by another seed number, not a team)
                if re.fullmatch(r"\d{1,2}\.?", nxt) or nxt.lower() in NON_TEAM_LABELS:
                    i += 1
                    continue

                parts = re.split(r"\s*(?:/|\||\bvs\.?\b)\s*", nxt)
                matched_any = False
                for part in parts[:2]:
                    canon = match_team(part)
                    if canon:
                        _add(canon, seed)
                        matched_any = True
                if matched_any:
                    i += 2
                    continue
                elif len(nxt) <= 35:
                    unmatched.append(f"{seed}: {nxt}")
        i += 1

    seen = set()
    dedup_miss = [x for x in unmatched if not (x in seen or seen.add(x))]
    return parsed, dedup_miss


def fetch_latest_external_bracketology(teams, existing_ext=None):
    """Fetch 76-team bracketologies from CBS Sports, ESPN Lunardi (Playwright), and Hoops HD JNG."""
    ext = dict(existing_ext or {})
    if os.environ.get("SKIP_BRACKETOLOGY_SYNC") == "1":
        print("[Bracketology Sync] Skipped (SKIP_BRACKETOLOGY_SYNC=1).")
        return ext

    match_team = build_team_matcher(teams)
    today_label = datetime.datetime.now().strftime("%b %d, %Y")

    # 1. CBS Sports (static fetch first, Playwright fallback if throttled/JS-gated)
    try:
        def _parse_cbs_html(raw_html):
            soup = BeautifulSoup(raw_html, "html.parser")
            raw_text = soup.get_text("\n")
            # Only cut at "On The Bubble" if it appears after the bracket body
            bubble_idx = raw_text.rfind("On The Bubble")
            bracket_chunk = raw_text[:bubble_idx] if bubble_idx > 1500 else raw_text
            lines = [ln.strip() for ln in bracket_chunk.splitlines() if ln.strip()]
            merged = []
            i = 0
            while i < len(lines):
                if re.fullmatch(r"\d{1,2}\.?", lines[i]) and i + 1 < len(lines):
                    if i + 3 < len(lines) and lines[i + 2] == "/":
                        merged.append(f"{lines[i]} {lines[i + 1]} / {lines[i + 3]}")
                        i += 4
                    else:
                        merged.append(f"{lines[i]} {lines[i + 1]}")
                        i += 2
                else:
                    merged.append(lines[i])
                    i += 1
            return _parse_seed_team_lines(merged, match_team)

        cbs_url = "https://www.cbssports.com/college-basketball/bracketology/"
        cbs_parsed, cbs_miss = _parse_cbs_html(_fetch_html(cbs_url))
        if len(cbs_parsed) < 60:
            pw_html = _fetch_html_playwright(cbs_url)
            if pw_html:
                cbs_parsed, cbs_miss = _parse_cbs_html(pw_html)

        if len(cbs_parsed) >= 60:
            ext["cbs"] = cbs_parsed
            top1 = [k for k, v in cbs_parsed.items() if v["seed"] == 1][:4]
            if len(top1) == 4:
                ext.setdefault("top1s", {})["cbs"] = top1
            ext.setdefault("updated", {})["cbs"] = today_label
            print(f"[Bracketology Sync] CBS Sports: {len(cbs_parsed)} teams matched.")
            if cbs_miss:
                print(f"  -> CBS unmatched ({len(cbs_miss)}): {cbs_miss[:15]}")
        else:
            print(f"[Bracketology Sync] CBS Sports only matched {len(cbs_parsed)} teams; keeping previous snapshot.")
    except Exception as exc:
        print(f"[Bracketology Sync] CBS Sports failed ({exc}); keeping previous snapshot.")

    # 2. ESPN Joe Lunardi (2027 URL first via Playwright)
    espn_url = "https://www.espn.com/espn/feature/story/_/page/bracketology/ncaa-bracketology-projecting-2027-march-madness-men-field"
    try:
        html = _fetch_html_playwright(espn_url) or _fetch_html(espn_url)
        espn_parsed, espn_miss = _extract_espn_bracket_from_html(html, match_team)
        if len(espn_parsed) >= 60:
            ext["espn"] = espn_parsed
            top1 = [k for k, v in espn_parsed.items() if v["seed"] == 1][:4]
            if len(top1) == 4:
                ext.setdefault("top1s", {})["espn"] = top1
            ext.setdefault("updated", {})["espn"] = today_label
            print(f"[Bracketology Sync] ESPN Lunardi: {len(espn_parsed)} teams matched.")
            if espn_miss:
                print(f"  -> ESPN unmatched ({len(espn_miss)}): {espn_miss[:25]}")
        else:
            print(f"[Bracketology Sync] ESPN Lunardi only matched {len(espn_parsed)} teams; keeping previous snapshot.")
            if espn_miss:
                print(f"  -> ESPN unmatched ({len(espn_miss)}): {espn_miss[:25]}")
    except Exception as exc:
        print(f"[Bracketology Sync] ESPN Lunardi failed ({exc}); keeping previous snapshot.")

    # 3. Hoops HD JNG
    try:
        html = _fetch_html("https://hoopshd.com/nitty-gritty-rankings/")
        soup = BeautifulSoup(html, "html.parser")
        jng_parsed = {}
        jng_miss = []
        for tr in soup.select("table tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.select("td")]
            if len(cells) >= 2:
                seed_m = re.match(r"^\s*#?(\d{1,2})\b", cells[0])
                if not seed_m:
                    continue
                seed = int(seed_m.group(1))
                if not (1 <= seed <= 16):
                    continue
                for raw_part in re.split(r"\s*(?:/|\bvs\.?\b)\s*", cells[1]):
                    canon = match_team(raw_part)
                    if canon:
                        jng_parsed.setdefault(canon, {"seed": seed})
                    elif raw_part.strip():
                        jng_miss.append(f"{seed}: {raw_part.strip()}")
        if len(jng_parsed) >= 60:
            ext["jngHq"] = jng_parsed
            ext.setdefault("updated", {})["jng"] = today_label
            print(f"[Bracketology Sync] Hoops HD JNG: {len(jng_parsed)} teams matched.")
            if jng_miss:
                print(f"  -> Hoops HD unmatched ({len(jng_miss)}): {jng_miss[:20]}")
        else:
            print(f"[Bracketology Sync] Hoops HD JNG only matched {len(jng_parsed)} teams; keeping previous snapshot.")
    except Exception as exc:
        print(f"[Bracketology Sync] Hoops HD JNG failed ({exc}); keeping previous snapshot.")

    return ext

def _parse_single_ranked_table(table, match_team, max_rank=365):
    """Parse one HTML <table> into (dict[team -> rank], list[unmatched]). Handles T-18 & blank tie rows."""
    ranks = {}
    unmatched = []
    last_rk = None
    for tr in table.select("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.select("td")]
        if len(cells) < 2:
            continue
        rk_cell = cells[0].strip()
        m = re.match(r"^(?:T-?|#)?(\d{1,3})\b", rk_cell, flags=re.I)
        if m:
            rk = int(m.group(1))
            last_rk = rk
        elif rk_cell in ("", "-", "—") and last_rk is not None:
            rk = last_rk  # tied team with blank rank cell
        else:
            continue

        if not (1 <= rk <= max_rank):
            continue

        # Strip trailing "(12)" first-place votes or "(5-0)" records
        raw_team = re.sub(r"\s*\([^)]*\)\s*$", "", cells[1]).strip()
        # On ESPN rankings, team cell may include abbreviation prefix or record; prefer <a> text if present
        a_tags = [a.get_text(" ", strip=True) for a in tr.select("td:nth-of-type(2) a") if a.get_text(strip=True)]
        canon = None
        for cand in ([a_tags[-1]] if a_tags else []) + [raw_team]:
            canon = match_team(re.sub(r"\s*\([^)]*\)\s*$", "", cand).strip())
            if canon:
                break
        if canon:
            ranks.setdefault(canon, rk)
        elif raw_team:
            unmatched.append(f"#{rk}: {raw_team}")
    return ranks, unmatched


def _parse_ranked_table(html, match_team, max_rank=365):
    soup = BeautifulSoup(html, "html.parser")
    best_ranks, best_miss = {}, []
    for tbl in soup.select("table"):
        r, m = _parse_single_ranked_table(tbl, match_team, max_rank=max_rank)
        if len(r) > len(best_ranks):
            best_ranks, best_miss = r, m
    return best_ranks, best_miss


def _extract_orv_blocks(soup):
    """Find 'Others receiving votes' text blocks (handles both inline <p> and heading + sibling layouts)."""
    blocks = []
    seen = set()
    for node in soup.find_all(string=re.compile(r"others\s+receiving\s+votes", re.I)):
        parent = node.parent
        if parent and parent.name in ("b", "strong", "span", "em", "a") and parent.parent:
            parent = parent.parent
        txt = re.sub(r"\s+", " ", parent.get_text(" ", strip=True)) if parent else str(node).strip()
        tail = re.sub(r"^.*?others\s+receiving\s+votes\s*[:.]?\s*", "", txt, flags=re.I).strip()

        # If the matched element was only a heading/label, grab the sibling or wrapper container
        if len(tail) < 10 and parent is not None:
            nxt = parent.find_next_sibling()
            if nxt:
                nxt_txt = re.sub(r"\s+", " ", nxt.get_text(" ", strip=True))
                if re.search(r"\d", nxt_txt):
                    txt = f"Others receiving votes: {nxt_txt}"
                    tail = nxt_txt
            if len(tail) < 10 and parent.parent is not None:
                wrap_txt = re.sub(r"\s+", " ", parent.parent.get_text(" ", strip=True))
                wrap_tail = re.sub(r"^.*?others\s+receiving\s+votes\s*[:.]?\s*", "", wrap_txt, flags=re.I).strip()
                if len(wrap_tail) >= 10:
                    txt = f"Others receiving votes: {wrap_tail}"

        if txt and txt not in seen:
            seen.add(txt)
            blocks.append(txt)
    return blocks


def _parse_orv_text(orv_text, existing_ranks, match_team):
    """
    Parse 'Others receiving votes: Team A 54, Team B 41, Team C 41...'
    Returns (rv_ranks, rv_votes, unmatched) ranked starting at max(26, len(existing_ranks) + 1).
    """
    text = re.split(r"(?:dropped from rankings|also receiving|next\s+games|\n)", orv_text, maxsplit=1, flags=re.I)[0]
    text = re.sub(r"^.*?others\s+receiving\s+votes\s*[:.]?\s*", "", text, flags=re.I).strip()

    entries = []
    unmatched = []
    for chunk in re.split(r"[,;]+", text):
        chunk = chunk.strip().rstrip(".")
        if not chunk:
            continue
        m = re.match(r"^(.*?)\s+(\d+)$", chunk)
        if not m:
            continue
        raw_team, votes = m.group(1).strip(), int(m.group(2))
        canon = match_team(raw_team)
        if canon:
            if canon not in existing_ranks:
                entries.append((canon, votes))
        elif raw_team:
            unmatched.append(f"RV({votes}): {raw_team}")

    # Sort descending by votes (stable sort preserves poll order on ties)
    entries.sort(key=lambda x: x[1], reverse=True)
    start_rk = max(26, len(existing_ranks) + 1)
    rv_ranks, rv_votes = {}, {}
    prev_votes, prev_rk = None, start_rk

    for idx, (canon, votes) in enumerate(entries):
        if canon in rv_ranks:
            continue
        if prev_votes is not None and votes == prev_votes:
            rk = prev_rk
        else:
            rk = start_rk + idx
            prev_rk = rk
            prev_votes = votes
        rv_ranks[canon] = rk
        rv_votes[canon] = votes

    return rv_ranks, rv_votes, unmatched


def fetch_external_rankings(teams, existing_ext=None):
    """Fetch NCAA NET (with RPI fallback), AP Poll (+RV), and USA Today Coaches Poll (+RV)."""
    ext = dict(existing_ext or {})
    if os.environ.get("SKIP_BRACKETOLOGY_SYNC") == "1":
        return ext

    match_team = build_team_matcher(teams)
    today_label = datetime.datetime.now().strftime("%b %d, %Y")

    # 1. NCAA NET (or Warren Nolan NET / RPI fallback)
    net_sources = [
        ("NCAA NET", "https://www.ncaa.com/rankings/basketball-men/d1/ncaa-mens-basketball-net-rankings"),
        ("NCAA NET", "https://www.warrennolan.com/basketball/2027/net"),
        ("RPI (Pre-NET)", "https://www.warrennolan.com/basketball/2027/rpi"),
    ]
    for label, url in net_sources:
        try:
            ranks, miss = _parse_ranked_table(_fetch_html(url), match_team, max_rank=365)
            if len(ranks) >= 100:
                ext["net"] = ranks
                ext["net_label"] = label
                ext.setdefault("updated", {})["net"] = today_label
                print(f"[Rankings Sync] {label}: {len(ranks)} teams matched.")
                if miss:
                    print(f"  -> {label} unmatched ({len(miss)}): {miss[:20]}")
                break
        except Exception:
            continue

    # 2. AP Top 25 + Receiving Votes (NCAA.com first)
    ap_ok = False
    try:
        ap_html = _fetch_html("https://www.ncaa.com/rankings/basketball-men/d1/associated-press")
        ap_soup = BeautifulSoup(ap_html, "html.parser")
        ranks, miss = _parse_ranked_table(ap_html, match_team, max_rank=25)
        if len(ranks) >= 20:
            rv_votes = {}
            orv_blocks = _extract_orv_blocks(ap_soup)
            if orv_blocks:
                rv_r, rv_v, rv_miss = _parse_orv_text(orv_blocks[0], ranks, match_team)
                ranks.update(rv_r)
                rv_votes.update(rv_v)
                miss.extend(rv_miss)
            ext["ap"] = ranks
            ext["ap_votes"] = rv_votes
            ext.setdefault("updated", {})["ap"] = today_label
            print(f"[Rankings Sync] AP Poll: {len(ranks)} teams matched ({len(rv_votes)} RV).")
            if miss:
                print(f"  -> AP Poll unmatched: {miss}")
            ap_ok = True
    except Exception:
        pass

    # 3. USA Today Coaches Poll + Receiving Votes (USA Today official page first, NCAA.com fallback)
    coaches_ok = False
    for c_url in (
        "https://sportsdata.usatoday.com/basketball/ncaab/coaches-poll",
        "https://www.ncaa.com/rankings/basketball-men/d1/usa-today-coaches",
    ):
        try:
            c_html = _fetch_html(c_url)
            c_soup = BeautifulSoup(c_html, "html.parser")

            # On USA Today, team cells contain both a full-name span and a short-code span;
            # strip hidden/mobile abbreviation spans if present, or fall back to full cell text.
            ranks = {}
            miss = []
            for tbl in c_soup.select("table"):
                r_try, m_try = {}, []
                last_rk = None
                for tr in tbl.select("tr"):
                    cells = tr.select("td")
                    if len(cells) < 2:
                        continue
                    rk_txt = cells[0].get_text(" ", strip=True)
                    m = re.match(r"^(?:T-?|#)?(\d{1,2})\b", rk_txt, flags=re.I)
                    if m:
                        rk = int(m.group(1))
                        last_rk = rk
                    elif rk_txt in ("", "-", "—") and last_rk is not None:
                        rk = last_rk
                    else:
                        continue
                    if not (1 <= rk <= 25):
                        continue

                    # Try individual child spans/links first (avoids "Duke Blue DevilsDuke" concatenation)
                    cands = [el.get_text(" ", strip=True) for el in cells[1].select("a, span") if el.get_text(strip=True)]
                    cands.append(cells[1].get_text(" ", strip=True))
                    canon = None
                    for raw_c in cands:
                        cleaned = re.sub(r"\s*\([^)]*\)\s*$", "", raw_c).strip()
                        canon = match_team(cleaned)
                        if canon:
                            break
                    if canon:
                        r_try.setdefault(canon, rk)
                    else:
                        m_try.append(f"#{rk}: {cells[1].get_text(' ', strip=True)}")
                if len(r_try) > len(ranks):
                    ranks, miss = r_try, m_try

            if len(ranks) >= 20:
                rv_votes = {}
                orv_blocks = _extract_orv_blocks(c_soup)
                if orv_blocks:
                    rv_r, rv_v, rv_miss = _parse_orv_text(orv_blocks[0], ranks, match_team)
                    ranks.update(rv_r)
                    rv_votes.update(rv_v)
                    miss.extend(rv_miss)
                ext["coaches"] = ranks
                ext["coaches_votes"] = rv_votes
                ext.setdefault("updated", {})["coaches"] = today_label
                print(f"[Rankings Sync] Coaches Poll: {len(ranks)} teams matched ({len(rv_votes)} RV).")
                if miss:
                    print(f"  -> Coaches Poll unmatched: {miss}")
                coaches_ok = True
                break
        except Exception:
            continue

    # 4. ESPN Rankings fallback (1st table + 1st ORV = AP, 2nd table + 2nd ORV = Coaches)
    if not ap_ok or not coaches_ok:
        try:
            espn_html = _fetch_html("https://www.espn.com/mens-college-basketball/rankings")
            espn_soup = BeautifulSoup(espn_html, "html.parser")
            orv_blocks = _extract_orv_blocks(espn_soup)
            ranked_tables = []
            for tbl in espn_soup.select("table"):
                r, m = _parse_single_ranked_table(tbl, match_team, max_rank=25)
                if len(r) >= 20:
                    ranked_tables.append((r, m))

            if not ap_ok and len(ranked_tables) >= 1:
                ranks, miss = ranked_tables[0]
                rv_votes = {}
                if len(orv_blocks) >= 1:
                    rv_r, rv_v, rv_miss = _parse_orv_text(orv_blocks[0], ranks, match_team)
                    ranks.update(rv_r)
                    rv_votes.update(rv_v)
                    miss.extend(rv_miss)
                ext["ap"] = ranks
                ext["ap_votes"] = rv_votes
                ext.setdefault("updated", {})["ap"] = today_label
                print(f"[Rankings Sync] AP Poll (ESPN): {len(ranks)} teams matched ({len(rv_votes)} RV).")
                if miss:
                    print(f"  -> AP Poll unmatched: {miss}")

            if not coaches_ok and len(ranked_tables) >= 2:
                ranks, miss = ranked_tables[1]
                rv_votes = {}
                if len(orv_blocks) >= 2:
                    rv_r, rv_v, rv_miss = _parse_orv_text(orv_blocks[1], ranks, match_team)
                    ranks.update(rv_r)
                    rv_votes.update(rv_v)
                    miss.extend(rv_miss)
                ext["coaches"] = ranks
                ext["coaches_votes"] = rv_votes
                ext.setdefault("updated", {})["coaches"] = today_label
                print(f"[Rankings Sync] Coaches Poll (ESPN): {len(ranks)} teams matched ({len(rv_votes)} RV).")
                if miss:
                    print(f"  -> Coaches Poll unmatched: {miss}")
        except Exception as exc:
            print(f"[Rankings Sync] ESPN polls fallback failed ({exc}).")

    return ext

def attach_team_logos(teams, prev_teams=None):
    """Attach ESPN CDN logo URLs to each team dict using ESPN's public D1 teams API."""
    prev_logo_map = {
        t["team"]: t.get("logo", "")
        for t in (prev_teams or [])
        if t.get("team") and t.get("logo")
    }
    for t in teams:
        if t["team"] in prev_logo_map:
            t["logo"] = prev_logo_map[t["team"]]

    if os.environ.get("SKIP_BRACKETOLOGY_SYNC") == "1":
        return

    try:
        url = "https://site.api.espn.com/apis/site/v2/sports/basketball/mens-college-basketball/teams?limit=500"
        raw = _fetch_html(url)
        data = json.loads(raw)
        espn_teams = (
            data.get("sports", [{}])[0]
            .get("leagues", [{}])[0]
            .get("teams", [])
        )
        match_team = build_team_matcher(teams)
        by_espn_name = {
            (t.get("espn_name") or "").strip().lower(): t["team"]
            for t in teams if t.get("espn_name")
        }
        logo_map = {}

        for entry in espn_teams:
            tm = entry.get("team", {})
            tid = tm.get("id")
            logos = tm.get("logos") or []
            logo_url = logos[0].get("href") if logos else (
                f"https://a.espncdn.com/i/teamlogos/ncaa/500/{tid}.png" if tid else ""
            )
            if not logo_url:
                continue

            # Try exact espn_name match first, then displayName / location / shortDisplayName
            canon = None
            for cand in (
                tm.get("displayName", ""),
                tm.get("location", ""),
                tm.get("shortDisplayName", ""),
                tm.get("name", ""),
            ):
                if not cand:
                    continue
                low = cand.strip().lower()
                if low in by_espn_name:
                    canon = by_espn_name[low]
                    break
                canon = match_team(cand)
                if canon:
                    break
            if canon and canon not in logo_map:
                logo_map[canon] = logo_url

        matched = 0
        for t in teams:
            if t["team"] in logo_map:
                t["logo"] = logo_map[t["team"]]
                matched += 1
        print(f"[Logos Sync] Attached ESPN logos to {matched}/{len(teams)} teams.")
    except Exception as exc:
        print(f"[Logos Sync] Skipped ({exc}); keeping cached logos.")
        

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    parser = argparse.ArgumentParser(description="Build docs/data.json for the 4-page CBB site.")
    parser.add_argument("--stats", help="Path to 2026-27 CBB Stats Excel file")
    parser.add_argument("--sims", help="Path to Batch_Simulation_Results Excel file")
    parser.add_argument("--out", default=os.path.join(base_dir, "docs", "data.json"), help="Output JSON path")
    args = parser.parse_args()

    stats_path = args.stats or find_workbook(
        base_dir,
        ["2026-27 Men's College Basketball Stats_code.xlsx", "2026-27 Men's College Basketball Stats.xlsx"],
        "stats",
    )
    sim_path = args.sims or find_workbook(
        base_dir,
        ["Batch_Simulation_Results.xlsx", "Batch_Simulation_Results.xlsm"],
        "simulation",
    )

    if not stats_path:
        print("ERROR: Could not locate stats Excel workbook.")
        sys.exit(1)

    # Load previously saved data.json first so cached logos/snapshots are available
    prev_json = {}
    if os.path.isfile(args.out):
        try:
            with open(args.out, "r", encoding="utf-8") as f:
                prev_json = json.load(f)
        except Exception:
            prev_json = {}
    existing_ext = prev_json.get("external_bracketology", {})
    existing_rankings = prev_json.get("external_rankings", {})

    print(f"[Build] Reading stats: {os.path.basename(stats_path)}")
    teams = extract_teams(stats_path)
    attach_team_logos(teams, prev_json.get("teams"))

    games, upcoming_games, sheet_model_daily, date_sheets = [], [], {}, []
    if sim_path:
        print(f"[Build] Reading simulations: {os.path.basename(sim_path)}")
        games, upcoming_games, sheet_model_daily, date_sheets = extract_simulations(sim_path)
    else:
        print("[Build] Warning: No simulation workbook found; exporting teams only.")

    ext_bracket = fetch_latest_external_bracketology(teams, existing_ext)
    ext_rankings = fetch_external_rankings(teams, existing_rankings)

    today_iso = datetime.date.today().isoformat()
    payload = {
        "meta": {
            "updated_label": get_last_git_commit_label(base_dir),
            "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "team_count": len(teams),
            "game_count": len(games),
            "upcoming_count": len(upcoming_games),
            "date_sheets": date_sheets,
        },
        "build_date": today_iso,
        "teams": teams,
        "games": games,
        "upcoming_games": upcoming_games,
        "sheet_model_daily": sheet_model_daily,
        "external_bracketology": ext_bracket,
        "external_rankings": ext_rankings,
    }

    # Atomic, NaN-proof JSON write
    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    tmp_path = args.out + ".tmp"
    payload_json = json.dumps(payload, separators=(",", ":"), ensure_ascii=False, allow_nan=False)

    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(payload_json)
    os.replace(tmp_path, args.out)

    size_kb = os.path.getsize(args.out) / 1024.0
    print(
        f"[Build] Wrote {args.out} ({size_kb:.1f} KB) — "
        f"{len(teams)} teams, {len(games)} graded games, {len(upcoming_games)} upcoming games across {len(date_sheets)} days."
    )

    # Quick check that all 4 site pages exist in docs/
    expected_pages = ["index.html", "models.html", "climate.html", "bracket.html"]
    missing = [p for p in expected_pages if not os.path.isfile(os.path.join(out_dir, p))]
    if missing:
        print(f"[Build] Note: Missing HTML pages in {out_dir}: {', '.join(missing)}")


if __name__ == "__main__":
    main()