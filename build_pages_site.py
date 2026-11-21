#!/usr/bin/env python3
"""
Builds an up-to-date static GitHub Pages site (docs/index.html) for theconsensuscbb.com
directly from:
  - 2026-27 Men's College Basketball Stats_code.xlsx
  - Batch_Simulation_Results.xlsx
  - cbb_analytics_dashboard.html (template)

Run locally before pushing, or call `python3 build_pages_site.py` inside
.github/workflows/daily_cbb_sim.yml right before `git add` / `git commit` / `git push`.
"""

import datetime
import json
import os
import re
import sys
import openpyxl

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
STATS_FILE = os.path.join(WORK_DIR, "2026-27 Men's College Basketball Stats_code.xlsx")
SIM_FILE = os.path.join(WORK_DIR, "Batch_Simulation_Results.xlsx")
TEMPLATE_HTML = os.path.join(WORK_DIR, "cbb_analytics_dashboard.html")
DOCS_DIR = os.path.join(WORK_DIR, "docs")
OUT_HTML = os.path.join(DOCS_DIR, "index.html")
CNAME_FILE = os.path.join(DOCS_DIR, "CNAME")
CUSTOM_DOMAIN = "theconsensuscbb.com"

CONF_NAMES = {
    "ACC": "ACC", "B10": "Big Ten", "SEC": "SEC", "B12": "Big 12",
    "BE": "Big East", "P12": "Pac-12", "A10": "Atlantic 10", "MWC": "Mountain West",
    "WCC": "WCC", "AAC": "American", "MVC": "Missouri Valley", "MAC": "MAC",
    "CAA": "CAA", "SB": "Sun Belt", "CUSA": "Conference USA", "SC": "SoCon",
    "Horz": "Horizon League", "BW": "Big West", "BSth": "Big South", "Slnd": "Southland",
    "BSky": "Big Sky", "Ivy": "Ivy League", "MAAC": "MAAC", "Sum": "Summit League",
    "UAC": "UAC", "ASUN": "ASUN", "OVC": "Ohio Valley", "AE": "America East",
    "Pat": "Patriot League", "MEAC": "MEAC", "SWAC": "SWAC", "NEC": "NEC"
}


def num(v, d=2):
    if isinstance(v, (int, float)):
        return round(float(v), d)
    return 0.0


def extract_payload(stats_path, sim_path):
    teams = []
    if os.path.exists(stats_path):
        wb1 = openpyxl.load_workbook(stats_path, data_only=True, read_only=True)

        em_reg_map = {}
        if "EvanMiya Regular Ratings" in wb1.sheetnames:
            for r in list(wb1["EvanMiya Regular Ratings"].iter_rows(values_only=True))[1:]:
                if r and r[1] is not None:
                    tname = str(r[1]).strip()
                    em_reg_map[tname] = {
                        "em_rank": int(r[0]) if isinstance(r[0], (int, float)) else None,
                        "conf": str(r[2]).strip() if r[2] else "D1",
                        "kill_shots_pg": num(r[15], 2),
                        "kill_shots_margin": num(r[17], 2),
                    }

        em_inj_map = {}
        if "EvanMiya Injured Ratings" in wb1.sheetnames:
            for r in list(wb1["EvanMiya Injured Ratings"].iter_rows(values_only=True))[1:]:
                if r and r[2] is not None:
                    em_inj_map[str(r[2]).strip()] = {
                        "inj_rank": int(r[1]) if isinstance(r[1], (int, float)) else None,
                    }

        bpi_map = {}
        if "ESPN BPI" in wb1.sheetnames:
            for r in list(wb1["ESPN BPI"].iter_rows(values_only=True))[1:]:
                if r and r[0] is not None:
                    bpi_map[str(r[0]).strip()] = {
                        "conf_full": str(r[1]).strip() if r[1] else "",
                        "bpi_rk": int(r[4]) if isinstance(r[4], (int, float)) else None,
                        "proj_conf_wl": str(r[9]) if r[9] is not None else "",
                        "win_conf_pct": str(r[10]) if r[10] is not None else "",
                    }

        kp_map = {}
        if "Kenpom Ratings" in wb1.sheetnames:
            for r in list(wb1["Kenpom Ratings"].iter_rows(values_only=True))[1:]:
                if r and r[1] is not None:
                    kp_map[str(r[1]).strip()] = {
                        "kp_rk": int(r[15]) if isinstance(r[15], (int, float)) else None,
                    }

        elo_map = {}
        if "Warren ELO" in wb1.sheetnames:
            for r in list(wb1["Warren ELO"].iter_rows(values_only=True))[1:]:
                if r and r[0] is not None:
                    elo_map[str(r[0]).strip().lower()] = {
                        "elo_rk": int(r[3]) if isinstance(r[3], (int, float)) else None,
                        "elo_delta": int(r[4]) if isinstance(r[4], (int, float)) else 0,
                    }

        hgt_map = {}
        if "Sheet4" in wb1.sheetnames:
            for r in list(wb1["Sheet4"].iter_rows(values_only=True))[1:]:
                if r and r[1] is not None:
                    hgt_map[str(r[1]).strip()] = {
                        "size_rk": int(r[3]) if isinstance(r[3], (int, float)) else None,
                        "exp": num(r[16], 2) if isinstance(r[16], (int, float)) else None,
                        "exp_rk": int(r[17]) if isinstance(r[17], (int, float)) else None,
                        "bench": num(r[18], 2) if isinstance(r[18], (int, float)) else None,
                        "continuity": num(r[50], 1) if isinstance(r[50], (int, float)) else None,
                        "continuity_rk": int(r[51]) if isinstance(r[51], (int, float)) else None,
                    }

        if "Worksheet" in wb1.sheetnames:
            rows1 = list(wb1["Worksheet"].iter_rows(values_only=True))
            for r in rows1[2:]:
                if r and r[0] is not None and r[1] is not None and isinstance(r[0], (int, float)):
                    kp_name = str(r[1]).strip()
                    cbb_name = str(r[2]).strip() if r[2] else kp_name
                    warren_name = str(r[3]).strip().lower() if r[3] else kp_name.lower()
                    espn_name = str(r[4]).strip() if r[4] else kp_name
                    em_name = str(r[5]).strip() if r[5] else kp_name

                    em_info = em_reg_map.get(em_name, {})
                    em_inj_info = em_inj_map.get(em_name, {})
                    bpi_info = bpi_map.get(espn_name, {})
                    kp_info = kp_map.get(kp_name, {})
                    elo_info = elo_map.get(warren_name, {})
                    hgt_info = hgt_map.get(kp_name, {})
                    c_short = em_info.get("conf", "D1")

                    teams.append({
                        "id": int(r[0]),
                        "team": kp_name,
                        "cbb_name": cbb_name,
                        "espn_name": espn_name,
                        "conf_short": c_short,
                        "conf": CONF_NAMES.get(c_short, bpi_info.get("conf_full") or c_short),
                        "g": int(num(r[7], 0)),
                        "w": int(num(r[8], 0)),
                        "l": int(num(r[9], 0)),
                        "wl_pct": num(r[10], 3),
                        "srs": num(r[11], 2),
                        "sos": num(r[12], 2),
                        "sor": num(r[13], 2),
                        "ppg": num(r[24], 1),
                        "opp_ppg": num(r[25], 1),
                        "pd": num(r[26], 1),
                        "fg_pct": num(r[38] * 100 if isinstance(r[38], (int, float)) else 0, 1),
                        "two_pct": num(r[41] * 100 if isinstance(r[41], (int, float)) else 0, 1),
                        "three_pct": num(r[44] * 100 if isinstance(r[44], (int, float)) else 0, 1),
                        "ft_pct": num(r[47] * 100 if isinstance(r[47], (int, float)) else 0, 1),
                        "orb_pg": num(r[48], 1),
                        "second_c_pct": num(r[49] * 100 if isinstance(r[49], (int, float)) else 0, 1),
                        "trb": num(r[50], 1),
                        "ast": num(r[51], 1),
                        "stl": num(r[52], 1),
                        "blk": num(r[53], 1),
                        "tov": num(r[54], 1),
                        "efg_pct": num(r[88] * 100 if isinstance(r[88], (int, float)) else 0, 1),
                        "tov_pct": num(r[89] * 100 if isinstance(r[89], (int, float)) else 0, 1),
                        "orb_pct": num(r[90] * 100 if isinstance(r[90], (int, float)) else 0, 1),
                        "ftr": num(r[81] * 100 if isinstance(r[81], (int, float)) else 0, 1),
                        "opp_efg_pct": num(r[109] * 100 if isinstance(r[109], (int, float)) else 0, 1),
                        "opp_tov_pct": num(r[110] * 100 if isinstance(r[110], (int, float)) else 0, 1),
                        "opp_orb_pct": num(r[111] * 100 if isinstance(r[111], (int, float)) else 0, 1),
                        "opp_ftr": num(r[102] * 100 if isinstance(r[102], (int, float)) else 0, 1),
                        "avg_ht": num(r[131], 2),
                        "eff_ht": num(r[132], 2),
                        "adj_net_ff": num(r[137], 2),
                        "bpi_ortg": num(r[139], 1),
                        "bpi_drtg": num(r[140], 1),
                        "bpi_net": num(r[141], 1),
                        "bpi_rk": bpi_info.get("bpi_rk"),
                        "em_ortg": num(r[148], 2),
                        "em_drtg": num(r[149], 2),
                        "em_net": num(r[150], 2),
                        "em_rk": em_info.get("em_rank"),
                        "em_inj_ortg": num(r[151], 2),
                        "em_inj_drtg": num(r[152], 2),
                        "em_inj_net": num(r[153], 2),
                        "em_inj_rk": em_inj_info.get("inj_rank"),
                        "inj_delta_pct": num(r[155] * 100 if isinstance(r[155], (int, float)) else 0, 2),
                        "elo": num(r[159], 1),
                        "elo_rk": elo_info.get("elo_rk"),
                        "elo_delta": elo_info.get("elo_delta", 0),
                        "kp_ortg": num(r[162], 2),
                        "kp_drtg": num(r[163], 2),
                        "kp_net": num(r[164], 2),
                        "kp_rk": kp_info.get("kp_rk"),
                        "kp_inj_ortg": num(r[169], 2),
                        "kp_inj_drtg": num(r[170], 2),
                        "kp_inj_net": num(r[171], 2),
                        "con_tempo": num(r[173], 1),
                        "con_ortg": num(r[174], 2),
                        "con_drtg": num(r[175], 2),
                        "con_net": num(r[176], 2),
                        "live_tempo": num(r[178], 1),
                        "live_ortg": num(r[179], 2),
                        "live_drtg": num(r[180], 2),
                        "live_net": num(r[181], 2),
                        "net_score": num(r[183], 2),
                        "size_rk": hgt_info.get("size_rk"),
                        "exp": hgt_info.get("exp"),
                        "continuity": hgt_info.get("continuity"),
                        "win_conf_pct": bpi_info.get("win_conf_pct", ""),
                        "proj_conf_wl": bpi_info.get("proj_conf_wl", ""),
                    })

            for i, t in enumerate(sorted(teams, key=lambda x: x["con_net"], reverse=True)):
                t["con_rk"] = i + 1
            for i, t in enumerate(sorted(teams, key=lambda x: x["con_ortg"], reverse=True)):
                t["con_ortg_rk"] = i + 1
            for i, t in enumerate(sorted(teams, key=lambda x: x["con_drtg"])):
                t["con_drtg_rk"] = i + 1
            for i, t in enumerate(sorted(teams, key=lambda x: x["live_net"], reverse=True)):
                t["live_rk"] = i + 1
            for i, t in enumerate(sorted(teams, key=lambda x: x["live_ortg"], reverse=True)):
                t["live_ortg_rk"] = i + 1
            for i, t in enumerate(sorted(teams, key=lambda x: x["live_drtg"])):
                t["live_drtg_rk"] = i + 1
        wb1.close()

    all_games = []
    sheet_model_daily = {}
    date_sheets = []
    if os.path.exists(sim_path):
        wb2 = openpyxl.load_workbook(sim_path, data_only=True, read_only=True)
        date_sheets = sorted([
            s for s in wb2.sheetnames
            if re.match(r"^202\d-\d{2}-\d{2}$", s) and s >= "2026-11-05"
        ])
        models_col_map = [
            ("Live Morey Z", 1, 2, 19, 20),
            ("Consensus Morey Z", 4, 5, 21, 22),
            ("Live Log5", 7, 8, 23, 24),
            ("Consensus Log5", 10, 11, 25, 26),
            ("Adj. Eff.", 13, 14, 27, 28),
            ("Base", 16, 17, 29, 30),
        ]
        for ds in date_sheets:
            rows = list(wb2[ds].iter_rows(values_only=True))
            for i in range(1, len(rows), 3):
                if i + 1 < len(rows) and rows[i][0] is not None and rows[i + 1][0] is not None:
                    r1, r2 = rows[i], rows[i + 1]
                    if isinstance(r1[18], (int, float)) and isinstance(r2[18], (int, float)):
                        g = {
                            "d": ds,
                            "t1": str(r1[0]).strip(),
                            "t2": str(r2[0]).strip(),
                            "s1": int(r1[18]),
                            "s2": int(r2[18]),
                            "m": {},
                        }
                        for m_name, wp_c, sc_c, acc_c, cert_c in models_col_map:
                            g["m"][m_name] = [
                                round(float(r1[wp_c]) * 100, 1) if isinstance(r1[wp_c], (int, float)) else 50.0,
                                round(float(r2[wp_c]) * 100, 1) if isinstance(r2[wp_c], (int, float)) else 50.0,
                                int(round(float(r1[sc_c]))) if isinstance(r1[sc_c], (int, float)) else 0,
                                int(round(float(r2[sc_c]))) if isinstance(r2[sc_c], (int, float)) else 0,
                                str(r1[acc_c]).strip() if r1[acc_c] is not None else "",
                                str(r1[cert_c]).strip() if r1[cert_c] is not None else "",
                            ]
                        all_games.append(g)

        for m_name, _, _, _, _ in models_col_map:
            if m_name not in wb2.sheetnames:
                continue
            rows = list(wb2[m_name].iter_rows(values_only=True))
            d_list = []
            for r in rows[4:]:
                dt = r[1] if len(r) > 1 else None
                if isinstance(dt, datetime.datetime) and isinstance(r[2], (int, float)):
                    d_list.append({
                        "date": dt.strftime("%Y-%m-%d"),
                        "90%+": round(float(r[2]) * 100, 2) if isinstance(r[2], (int, float)) else None,
                        "75%+": round(float(r[3]) * 100, 2) if isinstance(r[3], (int, float)) else None,
                        "60%+": round(float(r[4]) * 100, 2) if isinstance(r[4], (int, float)) else None,
                        "50%+": round(float(r[5]) * 100, 2) if isinstance(r[5], (int, float)) else None,
                        "SUPER ALLS": int(r[6]) if isinstance(r[6], (int, float)) else 0,
                        "ALL": round(float(r[7]) * 100, 2) if isinstance(r[7], (int, float)) else 0.0,
                        "TEAM AND DIFF": round(float(r[8]) * 100, 2) if isinstance(r[8], (int, float)) else 0.0,
                        "TEAM": round(float(r[9]) * 100, 2) if isinstance(r[9], (int, float)) else 0.0,
                        "CORRECT": round(float(r[10]) * 100, 2) if isinstance(r[10], (int, float)) else 0.0,
                        "TOTAL": round(float(r[11]) * 100, 2) if isinstance(r[11], (int, float)) else 0.0,
                        "WRONG": round(float(r[12]) * 100, 2) if isinstance(r[12], (int, float)) else 0.0,
                        "90_r": int(r[13]) if isinstance(r[13], (int, float)) else 0,
                        "90_t": int(r[14]) if isinstance(r[14], (int, float)) else 0,
                        "75_r": int(r[15]) if isinstance(r[15], (int, float)) else 0,
                        "75_t": int(r[16]) if isinstance(r[16], (int, float)) else 0,
                        "60_r": int(r[17]) if isinstance(r[17], (int, float)) else 0,
                        "60_t": int(r[18]) if isinstance(r[18], (int, float)) else 0,
                        "50_r": int(r[19]) if isinstance(r[19], (int, float)) else 0,
                        "50_t": int(r[20]) if isinstance(r[20], (int, float)) else 0,
                        "all_c": int(r[21]) if isinstance(r[21], (int, float)) else 0,
                        "td_c": int(r[22]) if isinstance(r[22], (int, float)) else 0,
                        "tm_c": int(r[23]) if isinstance(r[23], (int, float)) else 0,
                        "corr_c": int(r[24]) if isinstance(r[24], (int, float)) else 0,
                        "tot_c": int(r[25]) if isinstance(r[25], (int, float)) else 0,
                        "wrong_c": int(r[26]) if isinstance(r[26], (int, float)) else 0,
                        "games": int(r[27]) if isinstance(r[27], (int, float)) else 0,
                    })
            sheet_model_daily[m_name] = d_list
        wb2.close()

    return {"teams": teams, "games": all_games, "sheet_model_daily": sheet_model_daily}, date_sheets


def build_site(stats_path=STATS_FILE, sim_path=SIM_FILE, template_path=TEMPLATE_HTML, out_path=OUT_HTML):
    if not os.path.exists(template_path):
        print(f"ERROR: Template file not found: {template_path}")
        sys.exit(1)

    payload, date_sheets = extract_payload(stats_path, sim_path)
    payload_json = json.dumps(payload, separators=(",", ":"))

    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    # 1. Replace embedded RAW_DATA JSON
    pattern = r"const RAW_DATA = \{.*?\};"
    match = re.search(pattern, html, flags=re.DOTALL)
    if not match:
        print("ERROR: Could not locate `const RAW_DATA = {...};` inside HTML template.")
        sys.exit(1)
    html = html[:match.start()] + f"const RAW_DATA = {payload_json};" + html[match.end():]

    # 2. Update header counts & timestamp
    today_label = datetime.datetime.now().strftime("%b %d, %Y")
    html = re.sub(
        r"\d+ D1 Teams • \d+ Graded Sims",
        f"{len(payload['teams'])} D1 Teams • {len(payload['games'])} Graded Sims",
        html,
    )
    html = re.sub(
        r"Updated [A-Z][a-z]{2} \d{1,2}, 202\d",
        f"Updated {today_label}",
        html,
    )

    # 3. Dynamically refresh the Date Window <select id="sim-date-filter"> with all graded dates
    if payload["games"]:
        date_counts = {}
        for g in payload["games"]:
            date_counts[g["d"]] = date_counts.get(g["d"], 0) + 1
        sorted_dates = sorted(date_counts.keys())
        total_g = len(payload["games"])
        w1_g = sum(v for k, v in date_counts.items() if "2026-11-05" <= k <= "2026-11-08")
        w2_g = sum(v for k, v in date_counts.items() if k >= "2026-11-09")

        options = [
            f'<option value="ALL">All {len(sorted_dates)} Days ({sorted_dates[0][5:]} to {sorted_dates[-1][5:]} • {total_g} G)</option>',
            f'<option value="WEEK1">Week 1 (11-05 to 11-08 • {w1_g} G)</option>',
            f'<option value="WEEK2">Since Nov 9 ({w2_g} G)</option>',
        ]
        for d in sorted_dates:
            options.append(f'<option value="{d}">{d} ({date_counts[d]} Games)</option>')

        select_pat = r'(<select id="sim-date-filter"[^>]*>)(.*?)(</select>)'
        html = re.sub(
            select_pat,
            lambda m: m.group(1) + "\n                " + "\n                ".join(options) + "\n              " + m.group(3),
            html,
            flags=re.DOTALL,
        )

        # Also update the trend chart date array dynamically
        dates_js = json.dumps(sorted_dates)
        html = re.sub(
            r"const dates = \['2026-11-05'.*?\];",
            f"const dates = {dates_js};",
            html,
        )

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    # Preserve docs/CNAME for GitHub Pages custom domain
    cname_path = os.path.join(os.path.dirname(out_path), "CNAME")
    if not os.path.exists(cname_path):
        with open(cname_path, "w", encoding="utf-8") as cf:
            cf.write(CUSTOM_DOMAIN + "\n")

    print(f"Built {out_path} ({len(payload['teams'])} teams, {len(payload['games'])} graded games across {len(date_sheets)} dates).")


if __name__ == "__main__":
    build_site()