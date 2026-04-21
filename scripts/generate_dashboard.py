"""Generate docs/dashboard.html from memory/status/*.json files."""
import json
import os
from datetime import datetime

SCRIPT_LABELS = {
    "run_morning_part1": "Morning Part 1 — ThaiBMA scrape + Morningstar",
    "run_morning_part2": "Morning Part 2 — ZRR / Corp Bond / Composite",
    "run_evening":       "Evening — VBA macros + GPO workflow",
    "send_gpo_email":    "GPO Email — standalone Bloomberg report",
    "run_sql_queries":   "SQL Queries — NAV / Holdings / VaR / Bloomberg",
}

STATUS_COLOR = {
    "success": ("#10b981", "✅", "SUCCESS"),
    "failed":  ("#ef4444", "❌", "FAILED"),
    "skipped": ("#f59e0b", "⏭", "SKIPPED"),
    "unknown": ("#64748b", "❓", "NO DATA"),
}


def load_statuses(status_dir):
    results = {}
    for key in SCRIPT_LABELS:
        path = os.path.join(status_dir, f"{key}.json")
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                results[key] = json.load(f)
        else:
            results[key] = {
                "script": key,
                "status": "unknown",
                "detail": "Never run",
                "timestamp": "—",
            }
    return results


def build_card(key, data):
    label = SCRIPT_LABELS.get(key, key)
    status = data.get("status", "unknown")
    color, icon, badge = STATUS_COLOR.get(status, STATUS_COLOR["unknown"])
    ts = data.get("timestamp", "—")
    detail = data.get("detail", "")
    return f"""
        <div class="card">
            <div class="card-header" style="border-left: 4px solid {color};">
                <span class="badge" style="background:{color};">{icon} {badge}</span>
                <span class="script-name">{label}</span>
            </div>
            <div class="card-body">
                <div class="meta">🕐 {ts}</div>
                <div class="detail">{detail}</div>
            </div>
        </div>"""


def build_html(statuses):
    cards = "\n".join(build_card(k, v) for k, v in statuses.items())
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    all_statuses = [v.get("status") for v in statuses.values()]
    if "failed" in all_statuses:
        overall_color, overall_text = "#ef4444", "⚠️ ONE OR MORE SCRIPTS FAILED"
    elif all(s in ("success", "skipped") for s in all_statuses):
        overall_color, overall_text = "#10b981", "✅ ALL SCRIPTS COMPLETED"
    else:
        overall_color, overall_text = "#64748b", "❓ AWAITING FIRST RUN"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="300">
<title>LHFund Risk — Run Dashboard</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: #0f0f23;
    color: #eeeeee;
    min-height: 100vh;
  }}
  header {{
    background: #1a1a2e;
    border-bottom: 2px solid #e94560;
    padding: 20px 32px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }}
  header h1 {{ font-size: 1.4rem; color: #e94560; letter-spacing: 1px; }}
  header .sub {{ font-size: 0.8rem; color: #64748b; margin-top: 4px; }}
  .overall {{
    margin: 24px 32px 8px;
    padding: 14px 20px;
    border-radius: 8px;
    background: #1e1e3a;
    border: 1px solid {overall_color};
    color: {overall_color};
    font-weight: 600;
    font-size: 1rem;
    letter-spacing: 0.5px;
  }}
  .grid {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(420px, 1fr));
    gap: 20px;
    padding: 16px 32px 40px;
  }}
  .card {{
    background: #1e1e3a;
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #2a2a4a;
    transition: transform 0.15s;
  }}
  .card:hover {{ transform: translateY(-2px); }}
  .card-header {{
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 14px 18px;
    background: #16213e;
  }}
  .badge {{
    font-size: 0.7rem;
    font-weight: 700;
    padding: 3px 10px;
    border-radius: 20px;
    color: #fff;
    white-space: nowrap;
    letter-spacing: 0.5px;
  }}
  .script-name {{
    font-size: 0.9rem;
    font-weight: 600;
    color: #eeeeee;
  }}
  .card-body {{ padding: 14px 18px; }}
  .meta {{
    font-size: 0.78rem;
    color: #94a3b8;
    margin-bottom: 6px;
  }}
  .detail {{
    font-size: 0.82rem;
    color: #cbd5e1;
    word-break: break-word;
  }}
  footer {{
    text-align: center;
    padding: 16px;
    font-size: 0.75rem;
    color: #64748b;
    border-top: 1px solid #1e1e3a;
  }}
  @media print {{
    body {{ background: #fff; color: #000; }}
    header {{ background: #f1f5f9; border-color: #e94560; }}
    .card {{ background: #f8fafc; border-color: #e2e8f0; }}
    .card-header {{ background: #f1f5f9; }}
  }}
</style>
</head>
<body>
<header>
  <div>
    <h1>LHFund Risk — Run Dashboard</h1>
    <div class="sub">Auto-refreshes every 5 minutes &nbsp;|&nbsp; Generated: {generated}</div>
  </div>
</header>
<div class="overall">{overall_text}</div>
<div class="grid">
{cards}
</div>
<footer>risk-python monitoring dashboard &nbsp;·&nbsp; {generated}</footer>
</body>
</html>
"""


def main():
    proj_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    status_dir = os.path.join(proj_root, "memory", "status")
    docs_dir = os.path.join(proj_root, "docs")
    os.makedirs(docs_dir, exist_ok=True)

    statuses = load_statuses(status_dir)
    html = build_html(statuses)

    out_path = os.path.join(docs_dir, "dashboard.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Dashboard written to {out_path}")


if __name__ == "__main__":
    main()
