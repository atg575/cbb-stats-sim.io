# ... existing code ...
import datetime
import json
import math
import os
import re
import subprocess
import sys
import openpyxl
# ... existing code ...
def num(v, d=2):
    if isinstance(v, (int, float)) and math.isfinite(v):
        return round(float(v), d)
    return 0.0
# ... existing code ...
    payload, date_sheets = extract_payload(stats_path, sim_path)
    payload_json = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)

    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    # 1. Replace embedded RAW_DATA JSON
    pattern = r"const RAW_DATA = \{.*?\};"
    match = re.search(pattern, html, flags=re.DOTALL)
    if not match:
        print("ERROR: Could not locate `const RAW_DATA = {...};` inside HTML template.")
        sys.exit(1)
    html = html[:match.start()] + f"const RAW_DATA = {payload_json};" + html[match.end():]

    # 2. Update header counts & accurate last-push timestamp
    commit_label = get_last_git_commit_label()
    html = re.sub(
        r"\d+ D1 Teams • \d+ Graded Sims",
        lambda _: f"{len(payload['teams'])} D1 Teams • {len(payload['games'])} Graded Sims",
        html,
    )
    html = re.sub(
        r'(<p id="site-last-updated"[^>]*>)(.*?)(</p>)',
        lambda m: f"{m.group(1)}Updated {commit_label}{m.group(3)}",
        html,
        flags=re.DOTALL,
    )

    # 2b. Optionally refresh external bracketology sources (CBS / ESPN) if reachable
    ext_updates = fetch_latest_external_bracketology(payload["teams"])
    if ext_updates:
        if ext_updates.get("cbs"):
            cbs_json = json.dumps(ext_updates["cbs"], separators=(",", ":"), ensure_ascii=False)
            html = re.sub(
                r"cbs:\s*\{.*?\},\s*\n\s*cbsFirstFourOut:",
                lambda _: f"cbs: {cbs_json},\n      cbsFirstFourOut:",
                html,
                flags=re.DOTALL,
            )
            top1_cbs = [k for k, v in ext_updates["cbs"].items() if v.get("seed") == 1][:4]
            if len(top1_cbs) == 4:
                top1_json = json.dumps(top1_cbs, ensure_ascii=False)
                html = re.sub(r"cbs:\s*\[.*?\]", lambda _: f"cbs: {top1_json}", html)
            print(f"[Bracketology Sync] Refreshed CBS Sports ({len(ext_updates['cbs'])} teams).")
        if ext_updates.get("espn"):
            espn_json = json.dumps(ext_updates["espn"], separators=(",", ":"), ensure_ascii=False)
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
# ... existing code ...
        # Also update the trend chart date array dynamically
        dates_js = json.dumps(sorted_dates)
        html = re.sub(
            r"const dates = \['2026-11-05'.*?\];",
            lambda _: f"const dates = {dates_js};",
            html,
        )

    # Sanity-check generated HTML before writing
    if "const RAW_DATA = " not in html or html.count("<script") != html.count("</script>"):
        print("ERROR: Generated HTML failed integrity check; aborting write to preserve existing docs/index.html.")
        sys.exit(1)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
# ... existing code ...