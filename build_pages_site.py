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
import datetime
import json
import math
import os
import re
import subprocess
import sys
import urllib.request
import openpyxl


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

    # Optional EvanMiya lookup maps for rank & conference
    em_reg_map = {}
    if "EvanMiya Regular Ratings" in wb.sheetnames:
        ws_em = wb["EvanMiya Regular Ratings"]
        for r in ws_em.iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 3 and r[1]:
                em_reg_map[clean_str(r[1])] = {
                    "em_rk": opt_int(r[0]),
                    "conf_short": clean_str(r[2], "D1"),
                }

    em_inj_map = {}
    for sname in ("EvanMiya Injury Ratings", "EvanMiya Injured Ratings"):
        if sname in wb.sheetnames:
            ws_emi = wb[sname]
            for r in ws_emi.iter_rows(min_row=2, values_only=True):
                if r and len(r) >= 2 and r[1]:
                    em_inj_map[clean_str(r[1])] = opt_int(r[0])
            break

    ws = wb["Worksheet"]
    hmap = build_header_map(ws, header_row=2)
    teams = []

    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or not isinstance(r[0], (int, float)) or not r[1]:
            continue

        kp_name = clean_str(r[1])
        espn_name = clean_str(get_col(r, hmap, ["espn name", "espn"], 4), kp_name)
        em_lookup_name = clean_str(get_col(r, hmap, ["evanmiya name", "em name"], 5), kp_name)
        em_info = em_reg_map.get(em_lookup_name) or em_reg_map.get(kp_name) or {}

        conf_short = clean_str(
            get_col(r, hmap, ["conf", "conference"], None),
            em_info.get("conf_short", "D1"),
        )
        conf_full = CONF_FULL_NAMES.get(conf_short, conf_short)

        con_ortg = num(get_col(r, hmap, ["consensus ortg", "con ortg", "consensus off"], 163), 2)
        con_drtg = num(get_col(r, hmap, ["consensus drtg", "con drtg", "consensus def"], 164), 2)
        con_net = num(get_col(r, hmap, ["consensus net", "con net", "consensus net rtg"], 165), 2)
        if con_net == 0.0 and (con_ortg or con_drtg):
            con_net = round(con_ortg - con_drtg, 2)

        live_ortg = num(get_col(r, hmap, ["live ortg", "live off"], 167), 2)
        live_drtg = num(get_col(r, hmap, ["live drtg", "live def"], 168), 2)
        live_net = num(get_col(r, hmap, ["live net", "live net rtg"], 169), 2)
        if live_net == 0.0 and (live_ortg or live_drtg):
            live_net = round(live_ortg - live_drtg, 2)

        net_score = num(get_col(r, hmap, ["net score", "av score"], 170), 2)
        if net_score == 0.0 and (con_net or live_net):
            net_score = round((con_net + live_net) / 2.0, 2)

        em_ortg = num(get_col(r, hmap, ["em ortg", "evanmiya ortg"], 144), 2)
        em_drtg = num(get_col(r, hmap, ["em drtg", "evanmiya drtg"], 145), 2)
        em_net = num(get_col(r, hmap, ["em net", "evanmiya net", "bpr"], 146), 2)
        em_inj_ortg = num(get_col(r, hmap, ["em inj ortg", "evanmiya inj ortg"], 148), 2)
        em_inj_drtg = num(get_col(r, hmap, ["em inj drtg", "evanmiya inj drtg"], 149), 2)
        em_inj_net = num(get_col(r, hmap, ["em inj net", "evanmiya inj net"], 150), 2)
        inj_delta = num(get_col(r, hmap, ["inj delta", "injury delta", "inj delta %"], 152), 2)
        if abs(inj_delta) > 0 and abs(inj_delta) < 0.25:
            inj_delta = round(inj_delta * 100.0, 2)

        kp_ortg = num(get_col(r, hmap, ["kp ortg", "kenpom ortg", "adjoe"], 156), 2)
        kp_drtg = num(get_col(r, hmap, ["kp drtg", "kenpom drtg", "adjde"], 157), 2)
        kp_net = num(get_col(r, hmap, ["kp net", "kenpom net", "netrtg"], 158), 2)
        kp_inj_ortg = num(get_col(r, hmap, ["kp inj ortg"], 160), 2)
        kp_inj_drtg = num(get_col(r, hmap, ["kp inj drtg"], 161), 2)
        kp_inj_net = num(get_col(r, hmap, ["kp inj net"], 162), 2)

        con_tempo = num(get_col(r, hmap, ["consensus tempo", "con tempo", "adjt"], 162), 1)
        live_tempo = num(get_col(r, hmap, ["live tempo"], 166), 1) or con_tempo

        teams.append({
            "id": int(round(r[0])),
            "team": kp_name,
            "espn_name": espn_name,
            "conf": conf_full,
            "conf_short": conf_short,
            "g": int(num(get_col(r, hmap, ["g", "games"], 7), 0)),
            "w": int(num(get_col(r, hmap, ["w", "wins"], 8), 0)),
            "l": int(num(get_col(r, hmap, ["l", "losses"], 9), 0)),
            "wl_pct": num(get_col(r, hmap, ["w-l%", "wl%"], 10), 3),
            "srs": num(get_col(r, hmap, ["srs"], 11), 2),
            "sos": num(get_col(r, hmap, ["sos"], 12), 2),
            "sor": num(get_col(r, hmap, ["sor"], 13), 2),
            "ppg": num(get_col(r, hmap, ["ppg", "pts/g"], 24), 1),
            "opp_ppg": num(get_col(r, hmap, ["opp ppg", "opp pts/g"], 25), 1),
            "pd": num(get_col(r, hmap, ["pd", "margin"], 26), 1),
            "fg_pct": pct_val(get_col(r, hmap, ["fg%"], 38), 1),
            "two_pct": pct_val(get_col(r, hmap, ["2p%"], 41), 1),
            "three_pct": pct_val(get_col(r, hmap, ["3p%"], 44), 1),
            "ft_pct": pct_val(get_col(r, hmap, ["ft%"], 47), 1),
            "orb_pg": num(get_col(r, hmap, ["orb/g", "orb"], 48), 1),
            "second_c_pct": pct_val(get_col(r, hmap, ["2c%", "second chance %"], 49), 1),
            "trb": num(get_col(r, hmap, ["trb", "reb"], 50), 1),
            "ast": num(get_col(r, hmap, ["ast"], 51), 1),
            "stl": num(get_col(r, hmap, ["stl"], 52), 1),
            "blk": num(get_col(r, hmap, ["blk"], 53), 1),
            "tov": num(get_col(r, hmap, ["tov"], 54), 1),
            "efg_pct": pct_val(get_col(r, hmap, ["efg%"], 88), 1),
            "tov_pct": pct_val(get_col(r, hmap, ["tov%"], 89), 1),
            "orb_pct": pct_val(get_col(r, hmap, ["orb%"], 90), 1),
            "ftr": pct_val(get_col(r, hmap, ["ftr", "ft rate"], 81), 1),
            "opp_efg_pct": pct_val(get_col(r, hmap, ["opp efg%"], 109), 1),
            "opp_tov_pct": pct_val(get_col(r, hmap, ["opp tov%"], 110), 1),
            "opp_orb_pct": pct_val(get_col(r, hmap, ["opp orb%"], 111), 1),
            "opp_ftr": pct_val(get_col(r, hmap, ["opp ftr"], 102), 1),
            "avg_ht": num(get_col(r, hmap, ["avg ht", "avg height"], 131), 2),
            "eff_ht": num(get_col(r, hmap, ["eff ht", "eff height"], 132), 2),
            "adj_net_ff": num(get_col(r, hmap, ["adj net ff"], 137), 2),
            "bpi_ortg": num(get_col(r, hmap, ["bpi ortg", "bpi off"], 140), 1),
            "bpi_drtg": num(get_col(r, hmap, ["bpi drtg", "bpi def"], 141), 1),
            "bpi_net": num(get_col(r, hmap, ["bpi", "bpi net"], 142), 1),
            "bpi_rk": opt_int(get_col(r, hmap, ["bpi rk", "bpi rank"], 143)),
            "em_ortg": em_ortg,
            "em_drtg": em_drtg,
            "em_net": em_net,
            "em_rk": opt_int(get_col(r, hmap, ["em rk"], 147)) or em_info.get("em_rk"),
            "em_inj_ortg": em_inj_ortg,
            "em_inj_drtg": em_inj_drtg,
            "em_inj_net": em_inj_net,
            "em_inj_rk": opt_int(get_col(r, hmap, ["em inj rk"], 151)) or em_inj_map.get(em_lookup_name) or em_inj_map.get(kp_name),
            "inj_delta_pct": inj_delta,
            "elo": num(get_col(r, hmap, ["elo", "warren elo"], 153), 1),
            "elo_rk": opt_int(get_col(r, hmap, ["elo rk", "elo rank"], 154)),
            "kp_ortg": kp_ortg,
            "kp_drtg": kp_drtg,
            "kp_net": kp_net,
            "kp_rk": opt_int(get_col(r, hmap, ["kp rk", "kenpom rk"], 159)),
            "kp_inj_ortg": kp_inj_ortg,
            "kp_inj_drtg": kp_inj_drtg,
            "kp_inj_net": kp_inj_net,
            "con_tempo": con_tempo,
            "con_ortg": con_ortg,
            "con_drtg": con_drtg,
            "con_net": con_net,
            "live_tempo": live_tempo,
            "live_ortg": live_ortg,
            "live_drtg": live_drtg,
            "live_net": live_net,
            "net_score": net_score,
            "size_rk": opt_int(get_col(r, hmap, ["size rk", "height rk"], None)),
            "exp": num(get_col(r, hmap, ["exp", "experience"], None), 2) or None,
            "continuity": num(get_col(r, hmap, ["continuity", "cont%"], None), 1) or None,
            "win_conf_pct": clean_str(get_col(r, hmap, ["win conf%", "win conf"], None)),
            "proj_conf_wl": clean_str(get_col(r, hmap, ["proj conf w-l", "proj conf"], None)),
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


def extract_simulations(sim_path):
    wb = openpyxl.load_workbook(sim_path, data_only=True, read_only=True)
    date_sheets = sorted(s for s in wb.sheetnames if re.match(r"^\d{4}-\d{2}-\d{2}$", s))
    games = []

    # Model column blocks on daily sheets: each model has [wp1, wp2, ps1, ps2, acc, cert]
    # Header-driven detection with standard fallback offsets
    for d_sheet in date_sheets:
        ws = wb[d_sheet]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 3:
            continue

        # Locate model blocks from row 1 or row 2 headers
        model_starts = {}
        for h_row in rows[:2]:
            for col_i, cell in enumerate(h_row):
                if cell is None:
                    continue
                txt = str(cell).strip()
                for m_name in MODEL_NAMES:
                    if txt.lower() == m_name.lower() and m_name not in model_starts:
                        model_starts[m_name] = col_i

        # Fallback standard column layout if headers aren't merged/labeled in row 1
        if not model_starts:
            fallback_cols = {
                "Live Morey Z": 5,
                "Consensus Morey Z": 11,
                "Live Log5": 17,
                "Consensus Log5": 23,
                "Adj. Eff.": 29,
                "Base": 35,
            }
            model_starts = fallback_cols

        for r in rows[2:]:
            if not r or len(r) < 5 or not r[0] or not r[1]:
                continue
            t1 = clean_str(r[0])
            t2 = clean_str(r[1])
            s1 = opt_int(r[2])
            s2 = opt_int(r[3])
            if s1 is None or s2 is None:
                continue

            m_dict = {}
            for m_name, start_col in model_starts.items():
                if start_col + 5 >= len(r):
                    continue
                wp1 = pct_val(r[start_col], 1)
                wp2 = pct_val(r[start_col + 1], 1)
                ps1 = opt_int(r[start_col + 2]) or 0
                ps2 = opt_int(r[start_col + 3]) or 0
                acc = clean_str(r[start_col + 4]).upper()
                cert = clean_str(r[start_col + 5])
                if acc:
                    m_dict[m_name] = [wp1, wp2, ps1, ps2, acc, cert]

            if m_dict:
                games.append({
                    "d": d_sheet,
                    "t1": t1,
                    "t2": t2,
                    "s1": s1,
                    "s2": s2,
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
            if isinstance(d_val, (datetime.datetime, datetime.date)):
                d_str = d_val.strftime("%Y-%m-%d")
            else:
                d_str = clean_str(d_val)
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
    return games, sheet_model_daily, date_sheets


def fetch_latest_external_bracketology(teams, existing_ext=None):
    """Best-effort fetch of CBS Sports bracketology; falls back cleanly if offline."""
    ext = dict(existing_ext or {})
    if os.environ.get("SKIP_BRACKETOLOGY_SYNC") == "1":
        print("[Bracketology Sync] Skipped (SKIP_BRACKETOLOGY_SYNC=1).")
        return ext

    team_lookup = {t["team"].lower(): t["team"] for t in teams}
    for t in teams:
        if t.get("espn_name"):
            team_lookup[t["espn_name"].lower()] = t["team"]

    today_label = datetime.datetime.now().strftime("%b %d, %Y")
    headers = {"User-Agent": "Mozilla/5.0 (compatible; CBBAnalyticsBot/1.0)"}

    try:
        req = urllib.request.Request(
            "https://www.cbssports.com/college-basketball/bracketology/",
            headers=headers,
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            page = resp.read().decode("utf-8", errors="ignore")

        cbs_parsed = {}
        # Match common seed + team patterns in CBS markup
        for seed_str, raw_name in re.findall(r'class="[^"]*seed[^"]*"[^>]*>\s*(\d{1,2})\s*<.*?<a[^>]*>([^<]{2,40})</a>', page, flags=re.DOTALL):
            seed_num = int(seed_str)
            if not (1 <= seed_num <= 16):
                continue
            cand = re.sub(r"\s+", " ", raw_name).strip().lower()
            matched = team_lookup.get(cand)
            if matched and matched not in cbs_parsed:
                cbs_parsed[matched] = {"seed": seed_num}

        if len(cbs_parsed) >= 32:
            ext["cbs"] = cbs_parsed
            top1 = [k for k, v in cbs_parsed.items() if v["seed"] == 1][:4]
            if len(top1) == 4:
                ext.setdefault("top1s", {})["cbs"] = top1
            ext.setdefault("updated", {})["cbs"] = today_label
            print(f"[Bracketology Sync] Updated CBS Sports ({len(cbs_parsed)} teams).")
        else:
            print("[Bracketology Sync] CBS markup changed or incomplete; keeping existing snapshot.")
    except Exception as exc:
        print(f"[Bracketology Sync] Offline or unreachable ({exc}); keeping existing snapshot.")

    return ext


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

    print(f"[Build] Reading stats: {os.path.basename(stats_path)}")
    teams = extract_teams(stats_path)

    games, sheet_model_daily, date_sheets = [], {}, []
    if sim_path:
        print(f"[Build] Reading simulations: {os.path.basename(sim_path)}")
        games, sheet_model_daily, date_sheets = extract_simulations(sim_path)
    else:
        print("[Build] Warning: No simulation workbook found; exporting teams only.")

    # Preserve previously saved external_bracketology if present
    existing_ext = {}
    if os.path.isfile(args.out):
        try:
            with open(args.out, "r", encoding="utf-8") as f:
                existing_ext = json.load(f).get("external_bracketology", {})
        except Exception:
            existing_ext = {}

    ext_bracket = fetch_latest_external_bracketology(teams, existing_ext)

    payload = {
        "meta": {
            "updated_label": get_last_git_commit_label(base_dir),
            "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "team_count": len(teams),
            "game_count": len(games),
            "date_sheets": date_sheets,
        },
        "teams": teams,
        "games": games,
        "sheet_model_daily": sheet_model_daily,
        "external_bracketology": ext_bracket,
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
    print(f"[Build] Wrote {args.out} ({size_kb:.1f} KB) — {len(teams)} teams, {len(games)} graded games across {len(date_sheets)} days.")

    # Quick check that all 4 site pages exist in docs/
    expected_pages = ["index.html", "models.html", "climate.html", "bracket.html"]
    missing = [p for p in expected_pages if not os.path.isfile(os.path.join(out_dir, p))]
    if missing:
        print(f"[Build] Note: Missing HTML pages in {out_dir}: {', '.join(missing)}")


if __name__ == "__main__":
    main()