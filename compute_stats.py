#!/usr/bin/env python3
"""
compute_stats.py
Replaces all 150 Excel formulas on the 'Worksheet' tab in pure Python/pandas.
Reads data/team_crosswalk.json + data/raw/*.csv and writes data/stats_master.json.
"""

import datetime
import json
import math
import os
import re
import numpy as np
import pandas as pd


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


def _to_float(val, default=0.0):
    if val is None or (isinstance(val, float) and not math.isfinite(val)):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("%", "").replace(",", "").replace("+", "")
    try:
        f = float(s)
        return f if math.isfinite(f) else default
    except ValueError:
        return default


def _to_opt_int(val):
    if val is None or (isinstance(val, float) and not math.isfinite(val)):
        return None
    if isinstance(val, (int, float)):
        return int(round(val))
    s = str(val).strip().lstrip("#")
    if s.isdigit():
        return int(s)
    return None


def _clean_key(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip().lower()


def _build_lookup(df, key_col_idx):
    """Build case-insensitive lookup dict: lower(key) -> row tuple (first match wins, like XLOOKUP)."""
    lut = {}
    for row in df.itertuples(index=False, name=None):
        if key_col_idx < len(row):
            k = _clean_key(row[key_col_idx])
            if k and k not in lut:
                lut[k] = row
    return lut


def _get(row, idx, default=0.0):
    if row is None or idx >= len(row):
        return default
    return _to_float(row[idx], default)


def _get_str(row, idx, default=""):
    if row is None or idx >= len(row) or row[idx] is None:
        return default
    if isinstance(row[idx], float) and pd.isna(row[idx]):
        return default
    return str(row[idx]).strip()


def _safe_div(a, b, default=0.0):
    return a / b if b else default


def compute_master_stats(data_dir="data"):
    crosswalk_path = os.path.join(data_dir, "team_crosswalk.json")
    raw_dir = os.path.join(data_dir, "raw")

    with open(crosswalk_path, "r", encoding="utf-8") as f:
        crosswalk = json.load(f)

    # Load all 8 raw CSVs
    df_kp = pd.read_csv(os.path.join(raw_dir, "kenpom.csv"))
    df_ht = pd.read_csv(os.path.join(raw_dir, "height.csv"))
    df_elo = pd.read_csv(os.path.join(raw_dir, "warren_elo.csv"))
    df_bpi = pd.read_csv(os.path.join(raw_dir, "espn_bpi.csv"))
    df_em_reg = pd.read_csv(os.path.join(raw_dir, "evanmiya_reg.csv"))
    df_em_inj = pd.read_csv(os.path.join(raw_dir, "evanmiya_inj.csv"))
    df_sch = pd.read_csv(os.path.join(raw_dir, "cbbref_school.csv"))
    df_opp = pd.read_csv(os.path.join(raw_dir, "cbbref_opp.csv"))

    # Filter out repeated header rows in CBBRef tables if present
    if len(df_sch.columns) > 1:
        df_sch = df_sch[df_sch.iloc[:, 1].astype(str).str.strip().str.lower() != "school"]
    if len(df_opp.columns) > 1:
        df_opp = df_opp[df_opp.iloc[:, 1].astype(str).str.strip().str.lower() != "school"]

    # Build XLOOKUP dictionaries matching exact column indices from Worksheet formulas
    lut_kp = _build_lookup(df_kp, 1)         # TeamName = col B (1)
    lut_ht = _build_lookup(df_ht, 1)         # TeamName = col B (1)
    lut_elo = _build_lookup(df_elo, 0)       # Team = col A (0)
    lut_bpi = _build_lookup(df_bpi, 0)       # Team = col A (0)
    lut_em_reg = _build_lookup(df_em_reg, 1) # Team = col B (1)
    lut_em_inj = _build_lookup(df_em_inj, 2) # Team = col C (2)
    lut_sch = _build_lookup(df_sch, 1)       # School = col B (1)
    lut_opp = _build_lookup(df_opp, 1)       # School = col B (1)

    # Pass 1: Per-team lookups & row-level formulas (before column averages)
    rows = []
    for entry in crosswalk:
        rk = int(entry["Rk"])
        kp_name = entry["Kenpom Name"]
        cbbref_name = entry["CBBRef Name"]
        warren_name = entry["Warren Name"]
        espn_name = entry["ESPN Name"]
        em_name = entry["EvanMiya Name"]
        school = entry.get("School", kp_name)

        r_kp = lut_kp.get(_clean_key(kp_name))
        r_ht = lut_ht.get(_clean_key(kp_name))
        r_elo = lut_elo.get(_clean_key(warren_name)) or lut_elo.get(_clean_key(kp_name))
        r_bpi = lut_bpi.get(_clean_key(espn_name))
        r_em = lut_em_reg.get(_clean_key(em_name)) or lut_em_reg.get(_clean_key(kp_name))
        r_em_inj = lut_em_inj.get(_clean_key(em_name)) or lut_em_inj.get(_clean_key(kp_name))
        r_sch = lut_sch.get(_clean_key(cbbref_name))
        r_opp = lut_opp.get(_clean_key(cbbref_name))

        g = _get(r_sch, 2, 0.0)
        w = _get(r_sch, 3, 0.0)
        l = _get(r_sch, 4, 0.0)
        wl_pct = _get(r_sch, 5, 0.0)
        srs = _get(r_sch, 6, 0.0)
        sos = _get(r_sch, 7, 0.0)

        tm_pts = _safe_div(_get(r_sch, 18), g)
        opp_pts = _safe_div(_get(r_sch, 19), g)
        pd_val = tm_pts - opp_pts
        sor_calc = (1.25 * sos) + ((tm_pts - opp_pts) / 2.0)

        mp = _safe_div(_get(r_sch, 21), g)
        fg = _safe_div(_get(r_sch, 22), g)
        fga = _safe_div(_get(r_sch, 23), g)
        fg_pct = _safe_div(fg, fga)
        p3 = _safe_div(_get(r_sch, 25), g)
        p3a = _safe_div(_get(r_sch, 26), g)
        p3_pct = _safe_div(p3, p3a)
        p2 = fg - p3
        p2a = fga - p3a
        p2_pct = _safe_div(p2, p2a)
        ft = _safe_div(_get(r_sch, 28), g)
        fta = _safe_div(_get(r_sch, 29), g)
        ft_pct = _safe_div(ft, fta)
        orb = _safe_div(_get(r_sch, 31), g)
        second_c_pct = _safe_div(orb, (fga - fg))
        trb = _safe_div(_get(r_sch, 32), g)
        ast = _safe_div(_get(r_sch, 33), g)
        stl = _safe_div(_get(r_sch, 34), g)
        blk = _safe_div(_get(r_sch, 35), g)
        tov = _safe_div(_get(r_sch, 36), g)
        pf = _safe_div(_get(r_sch, 37), g)
        poss = 0.5 * (fga + 0.475 * fta - orb + tov)

        opp_mp = _safe_div(_get(r_opp, 21), g)
        opp_fg = _safe_div(_get(r_opp, 22), g)
        opp_fga = _safe_div(_get(r_opp, 23), g)
        opp_fg_pct = _safe_div(opp_fg, opp_fga)
        opp_3p = _safe_div(_get(r_opp, 25), g)
        opp_3pa = _safe_div(_get(r_opp, 26), g)
        opp_3p_pct = _safe_div(opp_3p, opp_3pa)
        opp_2p = opp_fg - opp_3p
        opp_2pa = opp_fga - opp_3pa
        opp_2p_pct = _safe_div(opp_2p, opp_2pa)
        opp_ft = _safe_div(_get(r_opp, 28), g)
        opp_fta = _safe_div(_get(r_opp, 29), g)
        opp_ft_pct = _safe_div(opp_ft, opp_fta)
        opp_orb = _safe_div(_get(r_opp, 31), g)
        opp_2c_pct = _safe_div(opp_orb, (opp_fga - opp_fg))
        opp_trb = _safe_div(_get(r_opp, 32), g)
        opp_ast = _safe_div(_get(r_opp, 33), g)
        opp_stl = _safe_div(_get(r_opp, 34), g)
        opp_blk = _safe_div(_get(r_opp, 35), g)
        opp_tov = _safe_div(_get(r_opp, 36), g)
        opp_pf = _safe_div(_get(r_opp, 37), g)
        opp_poss = 0.5 * (opp_fga + 0.475 * opp_fta - opp_orb + opp_tov)

        pace = opp_poss + poss
        ortg = _safe_div(100.0, pace) * tm_pts
        opp_ortg = _safe_div(100.0, pace) * opp_pts

        ftr = _safe_div(fta, fga)
        p3ar = _safe_div(p3a, fga)
        ts_pct = _safe_div(tm_pts, 2.0 * (fga + 0.475 * fta))
        trb_pct = _safe_div(trb, trb + opp_trb)
        ast_pct = _safe_div(ast, fg)
        stl_pct = _safe_div(stl, pace)
        blk_pct = _safe_div(blk, opp_fga - opp_3pa)
        efg_pct = _safe_div(fg + 0.5 * p3, fga)
        tov_pct = _safe_div(tov, fga + 0.475 * fta + tov)
        orb_pct = _safe_div(orb, orb + (opp_trb - opp_orb))
        ft_fga = _safe_div(ft, fga)

        opp_ftr = _safe_div(opp_fta, opp_fga)
        opp_3par = _safe_div(opp_3pa, opp_fga)
        opp_ts_pct = _safe_div(opp_pts, 2.0 * (opp_fga + 0.475 * opp_fta))
        opp_trb_pct = _safe_div(opp_trb, trb + opp_trb)
        opp_ast_pct = _safe_div(opp_ast, opp_fg)
        opp_stl_pct = _safe_div(opp_stl, pace)
        opp_blk_pct = _safe_div(opp_blk, fga - p3a)
        opp_efg_pct = _safe_div(opp_fg + 0.5 * opp_3p, opp_fga)
        opp_tov_pct = _safe_div(opp_tov, opp_fga + 0.475 * opp_fta + opp_tov)
        opp_orb_pct = _safe_div(opp_orb, opp_orb + (trb - orb))
        opp_ft_fga = _safe_div(opp_ft, opp_fga)

        # Four Factors Rating (exact Worksheet formulas)
        ff_rating = (0.4 * (efg_pct * 100.0)) - (0.25 * tov_pct) + (0.2 * orb_pct) + (0.15 * (ftr * 100.0))
        opp_ff_rating = (0.4 * (opp_efg_pct * 100.0)) - (0.25 * opp_tov_pct) + (0.2 * opp_orb_pct) + (0.15 * (opp_ftr * 100.0))
        net_ff = ff_rating - opp_ff_rating
        adj_net_ff = net_ff + (0.45 * srs)

        # Height / Experience / Continuity (Sheet4)
        avg_ht = _get(r_ht, 2, 0.0)
        eff_ht = _get(r_ht, 14, 0.0)
        size_rk = _to_opt_int(r_ht[3]) if r_ht is not None and len(r_ht) > 3 else None
        exp_val = _get(r_ht, 16, None) if r_ht is not None and len(r_ht) > 16 else None
        continuity_val = _get(r_ht, 50, None) if r_ht is not None and len(r_ht) > 50 else None

        # KenPom Ratings
        adj_tempo = _get(r_kp, 4, 68.0)
        adj_ortg = _get(r_kp, 8, 100.0)
        adj_drtg = _get(r_kp, 12, 100.0)
        adj_net_rtg = adj_ortg - adj_drtg
        kp_rk = _to_opt_int(r_kp[15]) if r_kp is not None and len(r_kp) > 15 else None

        # ESPN BPI
        bpi_ortg = _get(r_bpi, 6, 0.0)
        bpi_drtg = _get(r_bpi, 7, 0.0)
        bpi_net = bpi_ortg + bpi_drtg
        bpi_rk = _to_opt_int(r_bpi[4]) if r_bpi is not None and len(r_bpi) > 4 else None
        bpi_conf = _get_str(r_bpi, 1, "")
        proj_conf_wl = _get_str(r_bpi, 9, "")
        win_conf_pct = _get_str(r_bpi, 10, "")

        # EvanMiya Regular & Injured (raw O-Rate / D-Rate before league-average offset)
        em_tempo = _get(r_em, 10, adj_tempo)
        em_inj_tempo = _get(r_em_inj, 11, em_tempo)
        em_raw_o = _get(r_em, 3, 0.0)
        em_raw_d = _get(r_em, 4, 0.0)
        em_inj_raw_o = _get(r_em_inj, 4, em_raw_o)
        em_inj_raw_d = _get(r_em_inj, 5, em_raw_d)
        em_rk = _to_opt_int(r_em[0]) if r_em is not None and len(r_em) > 0 else None
        em_inj_rk = _to_opt_int(r_em_inj[1]) if r_em_inj is not None and len(r_em_inj) > 1 else None
        conf_short = _get_str(r_em, 2, "D1") or "D1"
        conf_full = CONF_FULL_NAMES.get(conf_short) or bpi_conf or conf_short

        # Warren ELO
        elo = _get(r_elo, 2, 1500.0)
        elo_rk = _to_opt_int(r_elo[3]) if r_elo is not None and len(r_elo) > 3 else None

        rows.append({
            "Rk": rk,
            "Kenpom Name": kp_name,
            "CBBRef Name": cbbref_name,
            "Warren Name": warren_name,
            "ESPN Name": espn_name,
            "EvanMiya Name": em_name,
            "School": school,
            "conf": conf_full,
            "conf_short": conf_short,
            "G": int(round(g)), "W": int(round(w)), "L": int(round(l)), "W-L%": wl_pct,
            "SRS": srs, "SOS": sos, "SOR Calc": sor_calc,
            "Tm.": tm_pts, "Opp.": opp_pts, "PD": pd_val,
            "Possessions": poss, "Opp Possessions": opp_poss, "Pace": pace, "Tempo": pace,
            "MP": mp, "FG": fg, "FGA": fga, "FG%": fg_pct,
            "2P": p2, "2PA": p2a, "2P%": p2_pct,
            "3P": p3, "3PA": p3a, "3P%": p3_pct,
            "FT": ft, "FTA": fta, "FT%": ft_pct,
            "ORB": orb, "2C%": second_c_pct, "TRB": trb,
            "AST": ast, "STL": stl, "BLK": blk, "TOV": tov, "PF": pf,
            "Opp MP": opp_mp, "Opp FG": opp_fg, "Opp FGA": opp_fga, "Opp FG%": opp_fg_pct,
            "Opp 2P": opp_2p, "Opp 2PA": opp_2pa, "Opp 2P%": opp_2p_pct,
            "Opp 3P": opp_3p, "Opp 3PA": opp_3pa, "Opp 3P%": opp_3p_pct,
            "Opp FT": opp_ft, "Opp FTA": opp_fta, "Opp FT%": opp_ft_pct,
            "Opp ORB": opp_orb, "Opp 2C%": opp_2c_pct, "Opp TRB": opp_trb,
            "Opp AST": opp_ast, "Opp STL": opp_stl, "Opp BLK": opp_blk, "Opp TOV": opp_tov, "Opp PF": opp_pf,
            "ORtg": ortg, "FTr": ftr, "3PAr": p3ar, "TS%": ts_pct, "TRB%": trb_pct,
            "AST%": ast_pct, "STL%": stl_pct, "BLK%": blk_pct, "eFG%": efg_pct,
            "TOV%": tov_pct, "ORB%": orb_pct, "FT/FGA": ft_fga,
            "Opp ORtg": opp_ortg, "Opp FTr": opp_ftr, "Opp 3PAr": opp_3par, "Opp TS%": opp_ts_pct,
            "Opp TRB%": opp_trb_pct, "Opp AST%": opp_ast_pct, "Opp STL%": opp_stl_pct,
            "Opp BLK%": opp_blk_pct, "Opp eFG%": opp_efg_pct, "Opp TOV%": opp_tov_pct,
            "Opp ORB%": opp_orb_pct, "Opp FT/FGA": opp_ft_fga,
            "Avg. Ht": avg_ht, "Eff. Ht": eff_ht,
            "FF Rating": ff_rating, "Opp FF Rating": opp_ff_rating,
            "Net FF Rating": net_ff, "Adj. Net FF": adj_net_ff,
            "BPI Ortg": bpi_ortg, "BPI DRtg": bpi_drtg, "BPI Net Rtg": bpi_net, "bpi_rk": bpi_rk,
            "EM Adj. Tempo": em_tempo, "EM Adj. Inj Tempo": em_inj_tempo,
            "_em_raw_o": em_raw_o, "_em_raw_d": em_raw_d,
            "_em_inj_raw_o": em_inj_raw_o, "_em_inj_raw_d": em_inj_raw_d,
            "em_rk": em_rk, "em_inj_rk": em_inj_rk,
            "ELO": elo, "elo_rk": elo_rk,
            "Adj. Tempo": adj_tempo, "Adj. ORtg": adj_ortg, "Adj. DRtg": adj_drtg,
            "Adj. Net Rtg": adj_net_rtg, "kp_rk": kp_rk,
            "size_rk": size_rk, "exp": exp_val, "continuity": continuity_val,
            "win_conf_pct": win_conf_pct, "proj_conf_wl": proj_conf_wl,
        })

    # Pass 2: Column-average dependent formulas (exact match to Excel AVERAGE(Table1[...]))
    avg_ortg = float(np.mean([r["ORtg"] for r in rows]))
    avg_opp_ortg = float(np.mean([r["Opp ORtg"] for r in rows]))
    avg_elo = float(np.mean([r["ELO"] for r in rows]))
    avg_adj_ortg = float(np.mean([r["Adj. ORtg"] for r in rows]))
    avg_adj_drtg = float(np.mean([r["Adj. DRtg"] for r in rows]))

    for r in rows:
        em_adj_ortg = r["_em_raw_o"] + avg_ortg
        em_adj_drtg = avg_opp_ortg - r["_em_raw_d"]
        em_adj_net = em_adj_ortg - em_adj_drtg

        em_adj_inj_ortg = r["_em_inj_raw_o"] + avg_ortg
        em_adj_inj_drtg = avg_opp_ortg - r["_em_inj_raw_d"]
        em_adj_inj_net = em_adj_inj_ortg - em_adj_inj_drtg

        inj_delta = (
            em_adj_inj_ortg if em_adj_ortg == 0
            else (em_adj_inj_ortg - em_adj_ortg) / em_adj_ortg
        )

        elo_vs_avg = r["ELO"] - avg_elo
        adj_ortg_vs_avg = r["Adj. ORtg"] - avg_adj_ortg
        adj_drtg_vs_avg = r["Adj. DRtg"] - avg_adj_drtg

        kp_inj_ortg = adj_ortg_vs_avg + (adj_ortg_vs_avg * inj_delta) + avg_ortg
        kp_inj_drtg = (adj_drtg_vs_avg - (adj_drtg_vs_avg * inj_delta)) + avg_opp_ortg
        kp_inj_net = kp_inj_ortg - kp_inj_drtg

        con_tempo = (r["Adj. Tempo"] + r["EM Adj. Tempo"]) / 2.0
        con_ortg = (
            ((r["BPI Ortg"] * 1.25) + (em_adj_ortg - avg_ortg) + adj_ortg_vs_avg + (elo_vs_avg / 25.0)) / 4.0
        ) + 100.0
        con_drtg = (
            (-(r["BPI DRtg"] * 1.25) - (avg_opp_ortg - em_adj_drtg) + adj_drtg_vs_avg - (elo_vs_avg / 25.0)) / 4.0
        ) + 100.0
        con_net = con_ortg - con_drtg

        live_tempo = (r["Adj. Tempo"] + r["EM Adj. Inj Tempo"]) / 2.0
        live_ortg = (kp_inj_ortg + em_adj_inj_ortg) / 2.0
        live_drtg = (kp_inj_drtg + em_adj_inj_drtg) / 2.0
        live_net = live_ortg - live_drtg

        net_score = (con_net + live_net + (1.25 * r["SOR Calc"]) + r["Adj. Net FF"]) / 4.0

        r.update({
            "EM Adj. ORtg": em_adj_ortg,
            "EM Adj. DRtg": em_adj_drtg,
            "EM Adj. Net Rtg": em_adj_net,
            "EM Adj. Inj Ortg": em_adj_inj_ortg,
            "EM Adj. Inj Drtg": em_adj_inj_drtg,
            "EM Adj. Net inj rtg": em_adj_inj_net,
            "Injury Delta": inj_delta,
            "ELO vs Avg": elo_vs_avg,
            "Adj. Ortg vs Avg": adj_ortg_vs_avg,
            "Adj. Drtg vs Avg": adj_drtg_vs_avg,
            "KP Inj Ortg": kp_inj_ortg,
            "KP Inj Drtg": kp_inj_drtg,
            "KP Inj Net rtg": kp_inj_net,
            "Consensus Tempo": con_tempo,
            "Consensus ORtg": con_ortg,
            "Consensus DRtg": con_drtg,
            "Consensus Net Rtg": con_net,
            "Live Tempo": live_tempo,
            "Live ORtg": live_ortg,
            "Live DRtg": live_drtg,
            "Live Net Rtg": live_net,
            "Net Score": net_score,
        })
        for tmp_k in ("_em_raw_o", "_em_raw_d", "_em_inj_raw_o", "_em_inj_raw_d"):
            del r[tmp_k]

    out_path = os.path.join(data_dir, "stats_master.json")
    payload = {
        "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "team_count": len(rows),
        "teams": rows,
    }
    tmp = out_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, separators=(",", ":"), ensure_ascii=False, allow_nan=False)
    os.replace(tmp, out_path)
    print(f"[Compute Stats] Wrote {out_path} ({len(rows)} teams).")
    return rows


if __name__ == "__main__":
    compute_master_stats()
