#!/usr/bin/env python3
"""
Dual-Mode Server (Local Mac + Render Cloud) & Auto-Reload Watcher
for the 2026-27 CBB Analytics & Simulation Hub.

Local Mac usage:
  cd /Users/alexvazquez/Coding/cbb_stats && python3 cbb_local_server.py

Render Cloud usage (via render.yaml):
  Start Command: python3 cbb_local_server.py
  Environment Variables:
    - CBB_ADMIN_TOKEN (required to protect POST /api/run-script publicly)
    - PORT (automatically injected by Render, defaults to 10000 on Render / 8765 locally)
"""

from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import os
import subprocess
import sys
import webbrowser
import openpyxl

IS_RENDER = bool(os.environ.get("RENDER"))
HOST = os.environ.get("HOST", "0.0.0.0" if IS_RENDER else "127.0.0.1")
PORT = int(os.environ.get("PORT", os.environ.get("CBB_PORT", "10000" if IS_RENDER else "8765")))
ADMIN_TOKEN = os.environ.get("CBB_ADMIN_TOKEN", "").strip()
WORK_DIR = os.path.dirname(os.path.abspath(__file__))

STATS_FILE = "2026-27 Men's College Basketball Stats_code.xlsx"
SIM_FILE = "Batch_Simulation_Results.xlsx"

ALLOWED_SCRIPTS = {
    "batch_pull.py",
    "batch_sim.py",
    "batch_result.py",
    "daily_table_extender.py",
    "weekly_table_extender.py",
    "kenpom_puller.py",
    "height_puller.py",
    "evan-miya-ratings.py",
    "espn_bpi_puller.py",
    "elo_ratings.py",
    "old_stat_puller.py",
}

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


def extract_live_payload():
    """Reads both Excel workbooks from disk and returns the fresh JSON payload."""
    f1 = os.path.join(WORK_DIR, STATS_FILE)
    f2 = os.path.join(WORK_DIR, SIM_FILE)

    teams = []
    if os.path.exists(f1):
        wb1 = openpyxl.load_workbook(f1, data_only=True, read_only=True)
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
    if os.path.exists(f2):
        wb2 = openpyxl.load_workbook(f2, data_only=True, read_only=True)
        date_sheets = sorted([s for s in wb2.sheetnames if s.startswith("202") and s >= "2026-11-05"])
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
        wb2.close()

    return {"teams": teams, "games": all_games, "sheet_model_daily": sheet_model_daily}


class CBBRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WORK_DIR, **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Admin-Token")
        super().end_headers()

    def _send_json(self, code, payload_obj):
        body = json.dumps(payload_obj, separators=(",", ":")).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.path = "/cbb_analytics_dashboard.html"
            return super().do_GET()

        if self.path.startswith("/healthz"):
            return self._send_json(200, {"status": "ok", "env": "render" if IS_RENDER else "local"})

        if self.path.startswith("/api/status"):
            f1 = os.path.join(WORK_DIR, STATS_FILE)
            f2 = os.path.join(WORK_DIR, SIM_FILE)
            return self._send_json(200, {
                "ok": True,
                "env": "render" if IS_RENDER else "local",
                "auth_required": bool(ADMIN_TOKEN),
                "work_dir": WORK_DIR,
                "stats_mtime": os.path.getmtime(f1) if os.path.exists(f1) else 0,
                "sim_mtime": os.path.getmtime(f2) if os.path.exists(f2) else 0,
            })

        if self.path.startswith("/api/payload"):
            return self._send_json(200, extract_live_payload())

        return super().do_GET()

    def do_POST(self):
        if self.path.startswith("/api/run-script"):
            if ADMIN_TOKEN:
                supplied = self.headers.get("X-Admin-Token", "").strip()
                if supplied != ADMIN_TOKEN:
                    return self._send_json(401, {
                        "returncode": 401,
                        "stdout": "",
                        "stderr": "Unauthorized: Enter your CBB_ADMIN_TOKEN in the Script Runner tab to execute scripts on Render."
                    })

            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(raw)
            script = req.get("script", "")
            args = req.get("args", [])

            if script not in ALLOWED_SCRIPTS:
                return self._send_json(400, {
                    "returncode": 1,
                    "stdout": "",
                    "stderr": f"Script '{script}' is not in ALLOWED_SCRIPTS."
                })

            script_path = os.path.join(WORK_DIR, script)
            if not os.path.exists(script_path):
                return self._send_json(404, {
                    "returncode": 1,
                    "stdout": "",
                    "stderr": f"File not found: {script_path}"
                })

            cmd = [sys.executable, script_path] + [str(a) for a in args if a]
            proc = subprocess.run(cmd, cwd=WORK_DIR, capture_output=True, text=True)
            return self._send_json(200, {
                "returncode": proc.returncode,
                "stdout": proc.stdout,
                "stderr": proc.stderr,
            })

        self.send_response(404)
        self.end_headers()


if __name__ == "__main__":
    mode_label = "Render Cloud" if IS_RENDER else "Local Mac"
    url = f"http://{HOST}:{PORT}/cbb_analytics_dashboard.html"
    print("=" * 68)
    print(f"🏀 CBB 2026-27 Server ({mode_label}) listening on {HOST}:{PORT}")
    print(f"📂 Working directory: {WORK_DIR}")
    print(f"🔐 Admin Token Protection: {'ENABLED' if ADMIN_TOKEN else 'DISABLED (local mode)'}")
    print("=" * 68)
    if not IS_RENDER:
        try:
            webbrowser.open(f"http://127.0.0.1:{PORT}/cbb_analytics_dashboard.html")
        except Exception:
            pass
    server = ThreadingHTTPServer((HOST, PORT), CBBRequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down CBB Server.")
        server.server_close()