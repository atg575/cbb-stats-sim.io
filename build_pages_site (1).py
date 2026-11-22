#!/usr/bin/env python3
"""
Builds an up-to-date static GitHub Pages site (docs/index.html) for theconsensuscbb.com
directly from:
  - 2026-27 Men's College Basketball Stats_code.xlsx
  - Batch_Simulation_Results.xlsx
  - cbb_analytics_dashboard.html (template)

Features:
  - NaN/Infinity-safe numeric extraction so embedded JSON always parses cleanly.
  - Accurate last-push / last-commit timestamp via Git with local clock fallback.
  - Optional, fault-tolerant external bracketology refresh (CBS Sports / ESPN) that
    only overwrites embedded data when a complete field is cleanly parsed.
  - Dynamic date-filter & trend-chart synchronization from Batch_Simulation_Results.xlsx.
  - Pre-write HTML integrity checks and automatic docs/CNAME & favicon preservation.

Run locally before pushing, or call `python3 build_pages_site.py` inside
.github/workflows/daily_cbb_sim.yml right before `git add` / `git commit` / `git push`.
"""

import datetime
import html as html_lib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
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

# Common display-name aliases used by CBS Sports / ESPN mapped to KenPom team names
TEAM_ALIASES = {
    "uconn": "Connecticut",
    "uconn huskies": "Connecticut",
    "nc state": "N.C. State",
    "north carolina state": "N.C. State",
    "ole miss": "Mississippi",
    "miami": "Miami FL",
    "miami (fl)": "Miami FL",
    "miami (oh)": "Miami OH",
    "miami ohio": "Miami OH",
    "st. john's": "St. John's",
    "saint john's": "St. John's",
    "saint mary's": "Saint Mary's",
    "st. mary's": "Saint Mary's",
    "smu": "SMU",
    "tcu": "TCU",
    "ucf": "UCF",
    "ucla": "UCLA",
    "usc": "USC",
    "unlv": "UNLV",
    "vcu": "VCU",
    "byu": "BYU",
    "lsu": "LSU",
    "liu": "LIU",
    "long island": "LIU",
    "long island university": "LIU",
    "unc wilmington": "UNC Wilmington",
    "uncw": "UNC Wilmington",
    "pitt": "Pittsburgh",
    "fau": "Florida Atlantic",
    "uab": "UAB",
    "utep": "UTEP",
    "utsa": "UTSA",
    "fiu": "FIU",
    "umass": "Massachusetts",
    "uic": "Illinois Chicago",
    "siue": "SIUE",
    "siu edwardsville": "SIUE",
    "umbc": "UMBC",
    "njit": "NJIT",
    "vmi": "VMI",
    "csun": "CSUN",
    "cal state northridge": "CSUN",
    "bethune-cookman": "Bethune Cookman",
    "gardner-webb": "Gardner Webb",
    "mcneese state": "McNeese",
    "michigan state": "Michigan St.",
    "iowa state": "Iowa St.",
    "ohio state": "Ohio St.",
    "penn state": "Penn St.",
    "florida state": "Florida St.",
    "utah state": "Utah St.",
    "san diego state": "San Diego St.",
    "boise state": "Boise St.",
    "arizona state": "Arizona St.",
    "oklahoma state": "Oklahoma St.",
    "kansas state": "Kansas St.",
    "mississippi state": "Mississippi St.",
    "colorado state": "Colorado St.",
    "oregon state": "Oregon St.",
    "washington state": "Washington St.",
    "fresno state": "Fresno St.",
    "san jose state": "San Jose St.",
    "murray state": "Murray St.",
    "wright state": "Wright St.",
    "wichita state": "Wichita St.",
    "montana state": "Montana St.",
    "tennessee state": "Tennessee St.",
    "south dakota state": "South Dakota St.",
    "north dakota state": "North Dakota St.",
    "kent state": "Kent St.",
    "ball state": "Ball St.",
    "illinois state": "Illinois St.",
    "indiana state": "Indiana St.",
    "missouri state": "Missouri St.",
    "texas state": "Texas St.",
    "georgia state": "Georgia St.",
    "appalachian state": "Appalachian St.",
    "app state": "Appalachian St.",
    "arkansas state": "Arkansas St.",
    "east tennessee state": "East Tennessee St.",
    "etsu": "East Tennessee St.",
    "sam houston": "Sam Houston St.",
    "sam houston state": "Sam Houston St.",
    "kennesaw state": "Kennesaw St.",
    "jacksonville state": "Jacksonville St.",
    "new mexico state": "New Mexico St.",
    "youngstown state": "Youngstown St.",
    "morehead state": "Morehead St.",
    "cleveland state": "Cleveland St.",
    "norfolk state": "Norfolk St.",
    "morgan state": "Morgan St.",
    "coppin state": "Coppin St.",
    "delaware state": "Delaware St.",
    "south carolina state": "South Carolina St.",
    "jackson state": "Jackson St.",
    "alcorn state": "Alcorn St.",
    "alabama state": "Alabama St.",
    "grambling": "Grambling St.",
    "grambling state": "Grambling St.",
    "northwestern state": "Northwestern St.",
    "portland state": "Portland St.",
    "sacramento state": "Sacramento St.",
    "idaho state": "Idaho St.",
    "weber state": "Weber St.",
    "tarleton": "Tarleton St.",
    "tarleton state": "Tarleton St.",
    "long beach state": "Long Beach St.",
    "cal state fullerton": "Cal St. Fullerton",
    "cal state bakersfield": "Cal St. Bakersfield",
    "california baptist": "Cal Baptist",
    "texas a&m-corpus christi": "Texas A&M Corpus Chris",
    "seattle u": "Seattle",
    "loyola maryland": "Loyola MD",
    "queens university": "Queens",
    "st. thomas (mn)": "St. Thomas",
    "saint thomas": "St. Thomas",
    "pennsylvania": "Penn",
    "hawai'i": "Hawaii",
    "ualbany": "Albany",
    "ul monroe": "Louisiana Monroe",
    "omaha": "Nebraska Omaha",
    "ut martin": "Tennessee Martin",
    "utrgv": "UT Rio Grande Valley",
}


def num(v, d=2):
    """Safely round finite numeric values, converting NaN/Inf/None to 0.0."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v):
        return round(float(v), d)
    return 0.0


def safe_int(v, default=0):
    """Safely convert finite numeric values to int."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v):
        return int(round(float(v)))
    return default


def safe_pct(v, d=2):
    """Safely convert a 0..1 Excel ratio into a 0..100 rounded percentage."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v):
        return round(float(v) * 100.0, d)
    return 0.0


def get_last_git_commit_label():
    """Returns a human-readable timestamp for the latest Git commit, or today's date."""
    try:
        proc = subprocess.run(
            ["git", "log", "-1", "--format=%cd", "--date=format:%b %d, %Y"],
            cwd=WORK_DIR,
            capture_output=True,
            text=True,
            timeout=5,
        )
        label = proc.stdout.strip()
        if proc.returncode == 0 and label:
            return label
    except Exception:
        pass
    return datetime.datetime.now().strftime("%b %d, %Y")


def _normalize_team_key(name):
    """Normalizes a team name for fuzzy matching across external bracketology sites."""
    s = html_lib.unescape(str(name or "")).strip().lower()
    s = re.sub(r"[^a-z0-9&'.\-\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _build_team_lookup(teams):
    """Builds a lookup map from normalized names/aliases to canonical KenPom team names."""
    lookup = {}
    for t in teams:
        canonical = t.get("team", "").strip()
        if not canonical:
            continue
        for candidate in (canonical, t.get("espn_name", ""), t.get("cbb_name", "")):
            norm = _normalize_team_key(candidate)
            if norm:
                lookup[norm] = canonical
                # Also index without periods (e.g., "michigan st" -> "Michigan St.")
                lookup[norm.replace(".", "")] = canonical

    for alias_key, canonical in TEAM_ALIASES.items():
        if canonical in {t.get("team") for t in teams}:
            lookup[_normalize_team_key(alias_key)] = canonical
            lookup[_normalize_team_key(alias_key).replace(".", "")] = canonical

    return lookup


def _resolve_team_name(raw_name, lookup):
    """Resolves a scraped team string to a canonical team name in our dataset."""
    norm = _normalize_team_key(raw_name)
    if not norm:
        return None
    if norm in lookup:
        return lookup[norm]
    no_dot = norm.replace(".", "")
    if no_dot in lookup:
        return lookup[no_dot]
    # Try converting trailing 'st' <-> 'state'
    if no_dot.endswith(" st"):
        alt = no_dot[:-3] + " state"
        if alt in lookup:
            return lookup[alt]
    elif no_dot.endswith(" state"):
        alt = no_dot[:-6] + " st"
        if alt in lookup:
            return lookup[alt]
    return None


def _http_get_text(url, timeout=10):
    """Fetches webpage text with a browser-like User-Agent and strict timeout."""
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        return resp.read().decode(charset, errors="replace")


def fetch_latest_external_bracketology(teams):
    """
    Optionally refreshes CBS Sports / ESPN bracketology snapshots if reachable.
    Set environment variable `SKIP_BRACKETOLOGY_SYNC=1` to skip network calls.
    Returns a dict with optional keys `cbs`, `espn`, and `updated`, or `{}` on failure.
    """
    if os.environ.get("SKIP_BRACKETOLOGY_SYNC", "").strip().lower() in ("1", "true", "yes"):
        print("[Bracketology Sync] Skipped via SKIP_BRACKETOLOGY_SYNC.")
        return {}

    if not teams:
        return {}

    lookup = _build_team_lookup(teams)
    today_str = datetime.datetime.now().strftime("%b %d, %Y")
    updates = {"updated": {}}

    # --- 1. CBS Sports Bracketology (best-effort parse) ---
    try:
        cbs_html = _http_get_text("https://www.cbssports.com/college-basketball/bracketology/", timeout=10)
        cbs_parsed = {}
        current_region = "East"
        region_headings = ["East", "South", "Midwest", "West"]

        # Scan region blocks if standard headings are present
        region_spans = []
        for reg in region_headings:
            for m in re.finditer(rf"\b{reg}\s+Region\b", cbs_html, flags=re.IGNORECASE):
                region_spans.append((m.start(), reg))
        region_spans.sort(key=lambda x: x[0])

        def region_for_pos(pos):
            reg_name = current_region
            for r_pos, r_lbl in region_spans:
                if r_pos <= pos:
                    reg_name = r_lbl
                else:
                    break
            return reg_name

        # Match typical CBS seed + team markup patterns (e.g. seed 1..16 followed by team name)
        seed_team_pat = re.compile(
            r'class="[^"]*seed[^"]*"[^>]*>\s*(\d{1,2})\s*<.*?class="[^"]*team[^"]*"[^>]*>\s*([^<]{2,45})\s*<',
            flags=re.IGNORECASE | re.DOTALL,
        )
        for m in seed_team_pat.finditer(cbs_html):
            seed_val = int(m.group(1))
            if not (1 <= seed_val <= 16):
                continue
            raw_team = m.group(2).strip()
            # Handle play-in pairings written as "Team A/Team B"
            parts = [p.strip() for p in raw_team.split("/") if p.strip()]
            is_play_in = len(parts) > 1
            reg_name = region_for_pos(m.start())
            for part in parts:
                canonical = _resolve_team_name(part, lookup)
                if canonical and canonical not in cbs_parsed:
                    cbs_parsed[canonical] = {
                        "seed": seed_val,
                        "region": reg_name,
                        "playIn": is_play_in,
                    }

        # Only apply CBS update if we cleanly parsed a near-complete tournament field
        if len(cbs_parsed) >= 64:
            updates["cbs"] = cbs_parsed
            updates["updated"]["cbs"] = today_str
        else:
            print(f"[Bracketology Sync] CBS HTML layout did not yield a full field ({len(cbs_parsed)} matched); preserving existing snapshot.")
    except (urllib.error.URLError, TimeoutError, Exception) as exc:
        print(f"[Bracketology Sync] CBS fetch skipped ({exc}); preserving existing snapshot.")

    # --- 2. ESPN Joe Lunardi Bracketology (best-effort parse) ---
    try:
        espn_url = "https://www.espn.com/espn/feature/story/_/page/bracketology/ncaa-bracketology-2027-march-madness-men-field-predictions"
        espn_html = _http_get_text(espn_url, timeout=10)
        espn_parsed = {}
        espn_pat = re.compile(
            r'data-seed="(\d{1,2})"[^>]*>.*?>\s*([A-Za-z0-9&\'\.\-\s]{2,40})\s*<',
            flags=re.IGNORECASE | re.DOTALL,
        )
        for m in espn_pat.finditer(espn_html):
            seed_val = int(m.group(1))
            if not (1 <= seed_val <= 16):
                continue
            canonical = _resolve_team_name(m.group(2), lookup)
            if canonical and canonical not in espn_parsed:
                espn_parsed[canonical] = f"#{seed_val} Seed ({today_str})"

        if len(espn_parsed) >= 16:
            updates["espn"] = espn_parsed
            updates["updated"]["espn"] = today_str
        else:
            print(f"[Bracketology Sync] ESPN HTML layout did not yield enough seeds ({len(espn_parsed)} matched); preserving existing snapshot.")
    except (urllib.error.URLError, TimeoutError, Exception) as exc:
        print(f"[Bracketology Sync] ESPN fetch skipped ({exc}); preserving existing snapshot.")

    if not updates.get("cbs") and not updates.get("espn"):
        return {}
    return updates


def extract_payload(stats_path, sim_path):
    """Reads both Excel workbooks and builds the JSON-serializable dashboard payload."""
    teams = []
    if os.path.exists(stats_path):
        wb1 = openpyxl.load_workbook(stats_path, data_only=True, read_only=True)

        em_reg_map = {}
        if "EvanMiya Regular Ratings" in wb1.sheetnames:
            for r in list(wb1["EvanMiya Regular Ratings"].iter_rows(values_only=True))[1:]:
                if r and len(r) > 17 and r[1] is not None:
                    tname = str(r[1]).strip()
                    em_reg_map[tname] = {
                        "em_rank": safe_int(r[0], None) if isinstance(r[0], (int, float)) else None,
                        "conf": str(r[2]).strip() if r[2] else "D1",
                        "kill_shots_pg": num(r[15], 2),
                        "kill_shots_margin": num(r[17], 2),
                    }

        em_inj_map = {}
        if "EvanMiya Injured Ratings" in wb1.sheetnames:
            for r in list(wb1["EvanMiya Injured Ratings"].iter_rows(values_only=True))[1:]:
                if r and len(r) > 2 and r[2] is not None:
                    em_inj_map[str(r[2]).strip()] = {
                        "inj_rank": safe_int(r[1], None) if isinstance(r[1], (int, float)) else None,
                    }

        bpi_map = {}
        if "ESPN BPI" in wb1.sheetnames:
            for r in list(wb1["ESPN BPI"].iter_rows(values_only=True))[1:]:
                if r and len(r) > 10 and r[0] is not None:
                    bpi_map[str(r[0]).strip()] = {
                        "conf_full": str(r[1]).strip() if r[1] else "",
                        "bpi_rk": safe_int(r[4], None) if isinstance(r[4], (int, float)) else None,
                        "proj_conf_wl": str(r[9]) if r[9] is not None else "",
                        "win_conf_pct": str(r[10]) if r[10] is not None else "",
                    }

        kp_map = {}
        if "Kenpom Ratings" in wb1.sheetnames:
            for r in list(wb1["Kenpom Ratings"].iter_rows(values_only=True))[1:]:
                if r and len(r) > 15 and r[1] is not None:
                    kp_map[str(r[1]).strip()] = {
                        "kp_rk": safe_int(r[15], None) if isinstance(r[15], (int, float)) else None,
                    }

        elo_map = {}
        if "Warren ELO" in wb1.sheetnames:
            for r in list(wb1["Warren ELO"].iter_rows(values_only=True))[1:]:
                if r and len(r) > 4 and r[0] is not None:
                    elo_map[str(r[0]).strip().lower()] = {
                        "elo_rk": safe_int(r[3], None) if isinstance(r[3], (int, float)) else None,
                        "elo_delta": safe_int(r[4], 0),
                    }

        hgt_map = {}
        if "Sheet4" in wb1.sheetnames:
            for r in list(wb1["Sheet4"].iter_rows(values_only=True))[1:]:
                if r and len(r) > 51 and r[1] is not None:
                    hgt_map[str(r[1]).strip()] = {
                        "size_rk": safe_int(r[3], None) if isinstance(r[3], (int, float)) else None,
                        "exp": num(r[16], 2) if isinstance(r[16], (int, float)) and math.isfinite(r[16]) else None,
                        "exp_rk": safe_int(r[17], None) if isinstance(r[17], (int, float)) else None,
                        "bench": num(r[18], 2) if isinstance(r[18], (int, float)) and math.isfinite(r[18]) else None,
                        "continuity": num(r[50], 1) if isinstance(r[50], (int, float)) and math.isfinite(r[50]) else None,
                        "continuity_rk": safe_int(r[51], None) if isinstance(r[51], (int, float)) else None,
                    }

        if "Worksheet" in wb1.sheetnames:
            rows1 = list(wb1["Worksheet"].iter_rows(values_only=True))
            for r in rows1[2:]:
                if not r or len(r) < 182 or r[0] is None or r[1] is None:
                    continue
                if not isinstance(r[0], (int, float)) or not math.isfinite(r[0]):
                    continue

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

                con_net = num(r[176], 2)
                live_net = num(r[181], 2)
                raw_net_score = r[183] if len(r) > 183 else None
                net_score = (
                    num(raw_net_score, 2)
                    if isinstance(raw_net_score, (int, float)) and math.isfinite(raw_net_score)
                    else round((con_net + live_net) / 2.0, 2)
                )

                teams.append({
                    "id": safe_int(r[0]),
                    "team": kp_name,
                    "cbb_name": cbb_name,
                    "espn_name": espn_name,
                    "conf_short": c_short,
                    "conf": CONF_NAMES.get(c_short, bpi_info.get("conf_full") or c_short),
                    "g": safe_int(r[7]),
                    "w": safe_int(r[8]),
                    "l": safe_int(r[9]),
                    "wl_pct": num(r[10], 3),
                    "srs": num(r[11], 2),
                    "sos": num(r[12], 2),
                    "sor": num(r[13], 2),
                    "ppg": num(r[24], 1),
                    "opp_ppg": num(r[25], 1),
                    "pd": num(r[26], 1),
                    "fg_pct": safe_pct(r[38], 1),
                    "two_pct": safe_pct(r[41], 1),
                    "three_pct": safe_pct(r[44], 1),
                    "ft_pct": safe_pct(r[47], 1),
                    "orb_pg": num(r[48], 1),
                    "second_c_pct": safe_pct(r[49], 1),
                    "trb": num(r[50], 1),
                    "ast": num(r[51], 1),
                    "stl": num(r[52], 1),
                    "blk": num(r[53], 1),
                    "tov": num(r[54], 1),
                    "efg_pct": safe_pct(r[88], 1),
                    "tov_pct": safe_pct(r[89], 1),
                    "orb_pct": safe_pct(r[90], 1),
                    "ftr": safe_pct(r[81], 1),
                    "opp_efg_pct": safe_pct(r[109], 1),
                    "opp_tov_pct": safe_pct(r[110], 1),
                    "opp_orb_pct": safe_pct(r[111], 1),
                    "opp_ftr": safe_pct(r[102], 1),
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
                    "inj_delta_pct": safe_pct(r[155], 2),
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
                    "con_net": con_net,
                    "live_tempo": num(r[178], 1),
                    "live_ortg": num(r[179], 2),
                    "live_drtg": num(r[180], 2),
                    "live_net": live_net,
                    "net_score": net_score,
                    "size_rk": hgt_info.get("size_rk"),
                    "exp": hgt_info.get("exp"),
                    "exp_rk": hgt_info.get("exp_rk"),
                    "bench": hgt_info.get("bench"),
                    "continuity": hgt_info.get("continuity"),
                    "continuity_rk": hgt_info.get("continuity_rk"),
                    "kill_shots_pg": em_info.get("kill_shots_pg", 0.0),
                    "kill_shots_margin": em_info.get("kill_shots_margin", 0.0),
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
            for i, t in enumerate(sorted(teams, key=lambda x: x["net_score"], reverse=True)):
                t["av_rk"] = i + 1
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
                    if (
                        len(r1) > 30
                        and len(r2) > 18
                        and isinstance(r1[18], (int, float))
                        and isinstance(r2[18], (int, float))
                        and math.isfinite(r1[18])
                        and math.isfinite(r2[18])
                    ):
                        g = {
                            "d": ds,
                            "t1": str(r1[0]).strip(),
                            "t2": str(r2[0]).strip(),
                            "s1": safe_int(r1[18]),
                            "s2": safe_int(r2[18]),
                            "m": {},
                        }
                        for m_name, wp_c, sc_c, acc_c, cert_c in models_col_map:
                            wp1 = safe_pct(r1[wp_c], 1) if isinstance(r1[wp_c], (int, float)) and math.isfinite(r1[wp_c]) else 50.0
                            wp2 = safe_pct(r2[wp_c], 1) if isinstance(r2[wp_c], (int, float)) and math.isfinite(r2[wp_c]) else 50.0
                            g["m"][m_name] = [
                                wp1,
                                wp2,
                                safe_int(r1[sc_c]),
                                safe_int(r2[sc_c]),
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
                if isinstance(dt, datetime.datetime) and len(r) > 27 and isinstance(r[2], (int, float)) and math.isfinite(r[2]):
                    d_list.append({
                        "date": dt.strftime("%Y-%m-%d"),
                        "90%+": safe_pct(r[2], 2) if isinstance(r[2], (int, float)) and math.isfinite(r[2]) else None,
                        "75%+": safe_pct(r[3], 2) if isinstance(r[3], (int, float)) and math.isfinite(r[3]) else None,
                        "60%+": safe_pct(r[4], 2) if isinstance(r[4], (int, float)) and math.isfinite(r[4]) else None,
                        "50%+": safe_pct(r[5], 2) if isinstance(r[5], (int, float)) and math.isfinite(r[5]) else None,
                        "SUPER ALLS": safe_int(r[6]),
                        "ALL": safe_pct(r[7], 2),
                        "TEAM AND DIFF": safe_pct(r[8], 2),
                        "TEAM": safe_pct(r[9], 2),
                        "CORRECT": safe_pct(r[10], 2),
                        "TOTAL": safe_pct(r[11], 2),
                        "WRONG": safe_pct(r[12], 2),
                        "90_r": safe_int(r[13]),
                        "90_t": safe_int(r[14]),
                        "75_r": safe_int(r[15]),
                        "75_t": safe_int(r[16]),
                        "60_r": safe_int(r[17]),
                        "60_t": safe_int(r[18]),
                        "50_r": safe_int(r[19]),
                        "50_t": safe_int(r[20]),
                        "all_c": safe_int(r[21]),
                        "td_c": safe_int(r[22]),
                        "tm_c": safe_int(r[23]),
                        "corr_c": safe_int(r[24]),
                        "tot_c": safe_int(r[25]),
                        "wrong_c": safe_int(r[26]),
                        "games": safe_int(r[27]),
                    })
            sheet_model_daily[m_name] = d_list
        wb2.close()

    return {"teams": teams, "games": all_games, "sheet_model_daily": sheet_model_daily}, date_sheets


def build_site(stats_path=STATS_FILE, sim_path=SIM_FILE, template_path=TEMPLATE_HTML, out_path=OUT_HTML):
    """Compiles the static docs/index.html file from Excel data and the dashboard template."""
    if not os.path.exists(template_path):
        print(f"ERROR: Template file not found: {template_path}")
        sys.exit(1)

    payload, date_sheets = extract_payload(stats_path, sim_path)
    if not payload["teams"]:
        print(f"ERROR: Extracted 0 teams from {stats_path}; aborting to preserve existing {out_path}.")
        sys.exit(1)

    # allow_nan=False enforces valid JSON so the browser never hits a SyntaxError
    payload_json = json.dumps(payload, separators=(",", ":"), ensure_ascii=False, allow_nan=False)

    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    # 1. Replace embedded RAW_DATA JSON
    pattern = r"const RAW_DATA = \{.*?\};"
    match = re.search(pattern, html, flags=re.DOTALL)
    if not match:
        print("ERROR: Could not locate `const RAW_DATA = {...};` inside HTML template.")
        sys.exit(1)
    html = html[:match.start()] + f"const RAW_DATA = {payload_json};" + html[match.end():]

    # 2. Update header counts & accurate last-push timestamp (using lambdas to avoid backslash issues)
    commit_label = get_last_git_commit_label()
    html = re.sub(
        r"\d+ D1 Teams • \d+ Graded Sims",
        lambda _: f"{len(payload['teams'])} D1 Teams • {len(payload['games'])} Graded Sims",
        html,
    )
    if 'id="site-last-updated"' in html:
        html = re.sub(
            r'(<p id="site-last-updated"[^>]*>)(.*?)(</p>)',
            lambda m: f"{m.group(1)}theconsensuscbb.com • Updated {commit_label}{m.group(3)}",
            html,
            flags=re.DOTALL,
        )
    else:
        html = re.sub(
            r"Updated [A-Z][a-z]{2} \d{1,2}, 202\d",
            lambda _: f"Updated {commit_label}",
            html,
        )

    # 2b. Optionally refresh external bracketology sources (CBS / ESPN) if reachable
    ext_updates = fetch_latest_external_bracketology(payload["teams"])
    if ext_updates:
        if ext_updates.get("cbs"):
            cbs_json = json.dumps(ext_updates["cbs"], separators=(",", ":"), ensure_ascii=False, allow_nan=False)
            html = re.sub(
                r"cbs:\s*\{.*?\},\s*\n\s*cbsFirstFourOut:",
                lambda _: f"cbs: {cbs_json},\n      cbsFirstFourOut:",
                html,
                flags=re.DOTALL,
            )
            top1_cbs = [k for k, v in ext_updates["cbs"].items() if v.get("seed") == 1][:4]
            if len(top1_cbs) == 4:
                top1_json = json.dumps(top1_cbs, ensure_ascii=False)
                html = re.sub(r"(top1s:\s*\{\s*cbs:\s*)\[.*?\]", lambda m: f"{m.group(1)}{top1_json}", html, flags=re.DOTALL)
            print(f"[Bracketology Sync] Refreshed CBS Sports ({len(ext_updates['cbs'])} teams).")

        if ext_updates.get("espn"):
            espn_json = json.dumps(ext_updates["espn"], separators=(",", ":"), ensure_ascii=False, allow_nan=False)
            html = re.sub(
                r"espn:\s*\{.*?\},\s*\n\s*jngHq:",
                lambda _: f"espn: {espn_json},\n      jngHq:",
                html,
                flags=re.DOTALL,
            )
            print(f"[Bracketology Sync] Refreshed ESPN Lunardi ({len(ext_updates['espn'])} teams).")

        for skey, sdate in ext_updates.get("updated", {}).items():
            html = re.sub(
                r"(updated:\s*\{[^}]*?\b" + re.escape(skey) + r':\s*)"[^"]*"',
                lambda m, sd=sdate: f'{m.group(1)}"{sd}"',
                html,
                flags=re.DOTALL,
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
            lambda _: f"const dates = {dates_js};",
            html,
        )

    # 4. Pre-write sanity check so broken output never overwrites docs/index.html
    if "const RAW_DATA = " not in html or html.count("<script") != html.count("</script>"):
        print("ERROR: Generated HTML failed integrity check; aborting write to preserve existing docs/index.html.")
        sys.exit(1)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    # 5. Preserve docs/CNAME and copy favicon.png if available
    cname_path = os.path.join(os.path.dirname(out_path), "CNAME")
    if not os.path.exists(cname_path):
        with open(cname_path, "w", encoding="utf-8") as cf:
            cf.write(CUSTOM_DOMAIN + "\n")

    src_favicon = os.path.join(WORK_DIR, "favicon.png")
    dst_favicon = os.path.join(os.path.dirname(out_path), "favicon.png")
    if os.path.exists(src_favicon) and not os.path.exists(dst_favicon):
        try:
            shutil.copy2(src_favicon, dst_favicon)
        except Exception as exc:
            print(f"[Favicon] Warning: could not copy favicon.png ({exc}).")

    print(
        f"Built {out_path} ({len(payload['teams'])} teams, {len(payload['games'])} graded games "
        f"across {len(date_sheets)} dates, timestamp '{commit_label}')."
    )


if __name__ == "__main__":
    build_site()