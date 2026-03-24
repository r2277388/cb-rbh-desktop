from __future__ import annotations

import html
import socket
import subprocess
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

from flask import Flask, abort, redirect, render_template_string, request, url_for


ROOT_DIR = Path(__file__).resolve().parent
VENV_PYTHON = ROOT_DIR / "venv" / "Scripts" / "python.exe"
LOG_DIR = ROOT_DIR / "launcher_logs"
HOST = "127.0.0.1"


@dataclass(frozen=True)
class Action:
    action_id: str
    label: str
    command: list[str]
    description: str
    section: str


ACTIONS: dict[str, Action] = {
    "amazon_po_archive": Action(
        "amazon_po_archive",
        "Amazon PO Archive Manager",
        [str(VENV_PYTHON), "-m", "tools.po_archive_manager"],
        "Launch the PO archive helper.",
        "Main Processes",
    ),
    "amazon_po_report": Action(
        "amazon_po_report",
        "Amazon PO Report",
        [str(VENV_PYTHON), "amazon_po/main.py"],
        "Run the Amazon PO report.",
        "Main Processes",
    ),
    "amazon_preorders": Action(
        "amazon_preorders",
        "Amazon PreOrders",
        [str(VENV_PYTHON), "amazon_preorders/main.py"],
        "Run the Amazon PreOrders process.",
        "Main Processes",
    ),
    "amazon_customer_orders": Action(
        "amazon_customer_orders",
        "Amazon Customer Orders",
        [str(VENV_PYTHON), "amazon_customer_orders/main.py"],
        "Run the Amazon Customer Orders process.",
        "Main Processes",
    ),
    "amazon_sql_upload": Action(
        "amazon_sql_upload",
        "Amazon Sellthru SQL Upload",
        [str(VENV_PYTHON), "amazon_sql_upload/main.py"],
        "Run the Amazon sellthru upload workflow.",
        "Main Processes",
    ),
    "amazon_rolling_reports": Action(
        "amazon_rolling_reports",
        "Amazon Rolling Reports",
        [str(VENV_PYTHON), "amazon_rolling_reports/main.py"],
        "Run the normal Amazon rolling reports process.",
        "Main Processes",
    ),
    "amazon_rolling_check": Action(
        "amazon_rolling_check",
        "Amazon Rolling Check",
        [str(VENV_PYTHON), "amazon_rolling_reports/check_last_10_weeks.py"],
        "Check the Amazon upload table for the last 10 weeks.",
        "Amazon Rolling Reports",
    ),
    "amazon_ams_manager": Action(
        "amazon_ams_manager",
        "Amazon AMS Manager",
        [str(VENV_PYTHON), "amazon_ams/manage_ams.py"],
        "Run the Amazon AMS manager.",
        "Main Processes",
    ),
    "bn_rolling_reports": Action(
        "bn_rolling_reports",
        "Barnes & Noble Rolling Reports",
        [str(VENV_PYTHON), "bn_rolling_reports/main.py"],
        "Run the Barnes & Noble rolling reports process.",
        "Main Processes",
    ),
    "frontlist_supercharged": Action(
        "frontlist_supercharged",
        "Frontlist Supercharged Data",
        [str(VENV_PYTHON), "FLTracking_Supercharged/main.py"],
        "Run the frontlist supercharged process.",
        "Main Processes",
    ),
    "uk_rolling_file_combining": Action(
        "uk_rolling_file_combining",
        "UK Rolling File Combining",
        [str(VENV_PYTHON), "UK_Rolling_File_Combining/main.py"],
        "Run the UK rolling file combine process.",
        "Main Processes",
    ),
    "hachette_orders": Action(
        "hachette_orders",
        "Hachette Orders - Shipping Estimates",
        [str(VENV_PYTHON), "hachette_orders/main.py"],
        "Run the Hachette shipping estimates process.",
        "Main Processes",
    ),
    "invobs": Action(
        "invobs",
        "Consolidate Inventory for the INVOBS",
        [str(VENV_PYTHON), "invobs_consolidated_inventory/main.py"],
        "Run the INVOBS inventory consolidation process.",
        "Main Processes",
    ),
    "xgboost_model": Action(
        "xgboost_model",
        "XGBoost Model",
        [str(VENV_PYTHON), "xgboost_model/main.py"],
        "Run the XGBoost launcher.",
        "Main Processes",
    ),
    "desk_procedures": Action(
        "desk_procedures",
        "Desk Procedures",
        [str(VENV_PYTHON), "desk_procedures/main.py"],
        "Open the desk procedures menu.",
        "Main Processes",
    ),
    "ssr_preparation": Action(
        "ssr_preparation",
        "SSR Daily Reporting",
        [str(VENV_PYTHON), "ssr_daily_summary/ssr_preparation.py"],
        "Run the daily SSR reporting preparation.",
        "SSR Daily Summary",
    ),
    "ssr_summary": Action(
        "ssr_summary",
        "SSR Aggregate Totals",
        [str(VENV_PYTHON), "ssr_daily_summary/ssr_summary.py"],
        "Run the SSR aggregate totals process.",
        "SSR Daily Summary",
    ),
    "ssr_visualization": Action(
        "ssr_visualization",
        "SSR Visualization",
        [str(VENV_PYTHON), "ssr_daily_summary/ssr_visualizations.py"],
        "Build the SSR visualization HTML output.",
        "SSR Daily Summary",
    ),
    "table_all_updates": Action(
        "table_all_updates",
        "All Updates",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "1"],
        "Run all table freshness checks.",
        "Check Table Updates",
    ),
    "table_ssr_summary": Action(
        "table_ssr_summary",
        "Tables for SSR Summary",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "2"],
        "Run the SSR Summary table check.",
        "Check Table Updates",
    ),
    "table_ebs_sales": Action(
        "table_ebs_sales",
        "Ebs.Sales Prior 5 Days",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "3"],
        "Run the Ebs.Sales prior 5 days check.",
        "Check Table Updates",
    ),
    "table_amazon": Action(
        "table_amazon",
        "Amazon",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "4"],
        "Run the Amazon table check.",
        "Check Table Updates",
    ),
    "table_bookscan": Action(
        "table_bookscan",
        "Bookscan",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "5"],
        "Run the Bookscan table check.",
        "Check Table Updates",
    ),
    "table_bn": Action(
        "table_bn",
        "Barnes & Noble",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "6"],
        "Run the Barnes & Noble table check.",
        "Check Table Updates",
    ),
    "table_freight_costs": Action(
        "table_freight_costs",
        "Freight Costs",
        [str(VENV_PYTHON), "table_check/check_table_updates.py", "7"],
        "Run the freight costs table check.",
        "Check Table Updates",
    ),
}


SECTION_ORDER = [
    "Main Processes",
    "SSR Daily Summary",
    "Check Table Updates",
    "Amazon Rolling Reports",
]


PAGE_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Chronicle Launcher</title>
  <style>
    :root {
      --bg: #f5efe4;
      --bg-alt: #e7dccb;
      --card: rgba(255, 251, 245, 0.92);
      --ink: #1e2f2f;
      --muted: #566867;
      --line: rgba(30, 47, 47, 0.14);
      --accent: #0f766e;
      --accent-2: #c2410c;
      --shadow: 0 18px 40px rgba(39, 32, 19, 0.12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Georgia, "Palatino Linotype", serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(194, 65, 12, 0.12), transparent 28%),
        radial-gradient(circle at top right, rgba(15, 118, 110, 0.18), transparent 25%),
        linear-gradient(180deg, var(--bg) 0%, var(--bg-alt) 100%);
      min-height: 100vh;
    }
    .shell {
      width: min(1220px, calc(100% - 32px));
      margin: 24px auto 48px;
    }
    .hero {
      padding: 28px;
      border: 1px solid var(--line);
      border-radius: 24px;
      background: linear-gradient(135deg, rgba(255,255,255,0.75), rgba(255,255,255,0.5));
      box-shadow: var(--shadow);
      backdrop-filter: blur(8px);
    }
    .eyebrow {
      text-transform: uppercase;
      letter-spacing: 0.12em;
      font-size: 12px;
      color: var(--accent);
      margin: 0 0 8px;
    }
    h1 {
      margin: 0;
      font-size: clamp(32px, 5vw, 54px);
      line-height: 0.95;
      font-weight: 700;
    }
    .hero p {
      max-width: 760px;
      margin: 14px 0 0;
      color: var(--muted);
      font-size: 17px;
      line-height: 1.5;
    }
    .banner {
      margin-top: 18px;
      padding: 14px 16px;
      border-radius: 16px;
      background: rgba(255,255,255,0.78);
      border: 1px solid var(--line);
      font-size: 14px;
    }
    .grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 18px;
      margin-top: 22px;
    }
    .section {
      padding: 20px;
      border-radius: 22px;
      background: var(--card);
      border: 1px solid var(--line);
      box-shadow: var(--shadow);
    }
    .section h2 {
      margin: 0 0 16px;
      font-size: 22px;
      border-bottom: 1px solid var(--line);
      padding-bottom: 10px;
    }
    .actions {
      display: grid;
      gap: 12px;
    }
    .action {
      padding: 14px;
      border-radius: 16px;
      border: 1px solid var(--line);
      background: rgba(255,255,255,0.82);
    }
    .action form { margin: 0; }
    .action button {
      appearance: none;
      border: 0;
      width: 100%;
      text-align: left;
      font: inherit;
      color: white;
      background: linear-gradient(135deg, var(--accent), #155e75);
      border-radius: 12px;
      padding: 12px 14px;
      cursor: pointer;
      transition: transform 140ms ease, filter 140ms ease;
    }
    .action button:hover {
      transform: translateY(-1px);
      filter: saturate(1.08);
    }
    .action h3 {
      margin: 0 0 8px;
      font-size: 18px;
    }
    .action p {
      margin: 10px 0 0;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.45;
    }
    .meta {
      margin-top: 10px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      font-size: 13px;
      color: var(--muted);
    }
    .pill {
      padding: 5px 9px;
      border-radius: 999px;
      background: rgba(15,118,110,0.09);
      border: 1px solid rgba(15,118,110,0.18);
    }
    .jobs {
      margin-top: 24px;
      padding: 22px;
      border-radius: 22px;
      background: rgba(28, 35, 35, 0.92);
      color: #f8f1e6;
      box-shadow: var(--shadow);
    }
    .jobs h2 {
      margin: 0 0 14px;
      font-size: 24px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }
    th, td {
      text-align: left;
      padding: 10px 8px;
      border-bottom: 1px solid rgba(255,255,255,0.11);
      vertical-align: top;
    }
    th {
      color: #d8cbc0;
      font-weight: 600;
    }
    a {
      color: #7dd3fc;
      text-decoration: none;
    }
    a:hover { text-decoration: underline; }
    .status-running { color: #fde68a; }
    .status-finished { color: #86efac; }
    .status-failed { color: #fca5a5; }
    .status-starting { color: #c4b5fd; }
    @media (max-width: 720px) {
      .shell { width: min(100% - 20px, 1220px); }
      .hero, .section, .jobs { padding: 16px; border-radius: 18px; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <p class="eyebrow">Local Process Launcher</p>
      <h1>Chronicle Browser Menu</h1>
      <p>Click a process to run it. Jobs start in the background and stream to a log file so the page stays responsive.</p>
      <div class="banner">
        Web mode launches scripts directly. Terminal-only file preview prompts are skipped here, so use the terminal menu when you want the pre-run file confirmation screens.
      </div>
    </section>

    <div class="grid">
      {% for section in sections %}
      <section class="section">
        <h2>{{ section.name }}</h2>
        <div class="actions">
          {% for action in section.actions %}
          <article class="action">
            <h3>{{ action.label }}</h3>
            <form method="post" action="{{ url_for('run_action', action_id=action.action_id) }}">
              <button type="submit">Run {{ action.label }}</button>
            </form>
            <p>{{ action.description }}</p>
          </article>
          {% endfor %}
        </div>
      </section>
      {% endfor %}
    </div>

    <section class="jobs">
      <h2>Recent Jobs</h2>
      {% if jobs %}
      <table>
        <thead>
          <tr>
            <th>Started</th>
            <th>Action</th>
            <th>Status</th>
            <th>Exit</th>
            <th>Log</th>
          </tr>
        </thead>
        <tbody>
          {% for job in jobs %}
          <tr>
            <td>{{ job.started }}</td>
            <td>{{ job.label }}</td>
            <td class="status-{{ job.status }}">{{ job.status }}</td>
            <td>{{ job.returncode if job.returncode is not none else "-" }}</td>
            <td><a href="{{ url_for('view_log', job_id=job.job_id) }}" target="_blank" rel="noreferrer">Open log</a></td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
      {% else %}
      <p>No jobs have been started yet.</p>
      {% endif %}
    </section>
  </div>
</body>
</html>
"""


app = Flask(__name__)
_jobs_lock = threading.Lock()
_jobs: Dict[str, dict] = {}


def _find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((HOST, 0))
        return sock.getsockname()[1]


def _sections() -> list[dict]:
    sections = []
    for section_name in SECTION_ORDER:
        section_actions = [action for action in ACTIONS.values() if action.section == section_name]
        if section_actions:
            sections.append({"name": section_name, "actions": section_actions})
    return sections


def _recent_jobs() -> list[dict]:
    with _jobs_lock:
        jobs = list(_jobs.values())
    jobs.sort(key=lambda job: job["started_ts"], reverse=True)
    return jobs[:20]


def _run_job(job_id: str, action: Action) -> None:
    log_path = LOG_DIR / f"{job_id}.log"
    with _jobs_lock:
        _jobs[job_id]["status"] = "running"
    with log_path.open("a", encoding="utf-8", errors="replace") as log_file:
        log_file.write(f"Command: {' '.join(action.command)}\n")
        log_file.write(f"Started: {_jobs[job_id]['started']}\n\n")
        log_file.flush()
        process = subprocess.Popen(
            action.command,
            cwd=ROOT_DIR,
            stdout=log_file,
            stderr=subprocess.STDOUT,
            text=True,
        )
        returncode = process.wait()
    with _jobs_lock:
        _jobs[job_id]["returncode"] = returncode
        _jobs[job_id]["status"] = "finished" if returncode == 0 else "failed"


@app.get("/")
def index():
    return render_template_string(
        PAGE_TEMPLATE,
        sections=_sections(),
        jobs=_recent_jobs(),
    )


@app.post("/run/<action_id>")
def run_action(action_id: str):
    action = ACTIONS.get(action_id)
    if action is None:
        abort(404)

    LOG_DIR.mkdir(exist_ok=True)
    job_id = f"{int(time.time() * 1000)}-{action_id}"
    started = time.strftime("%Y-%m-%d %I:%M:%S %p")
    log_path = LOG_DIR / f"{job_id}.log"

    with _jobs_lock:
        _jobs[job_id] = {
            "job_id": job_id,
            "label": action.label,
            "status": "starting",
            "returncode": None,
            "started": started,
            "started_ts": time.time(),
            "log_path": str(log_path),
        }

    worker = threading.Thread(target=_run_job, args=(job_id, action), daemon=True)
    worker.start()
    return redirect(url_for("index"))


@app.get("/logs/<job_id>")
def view_log(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if job is None:
        abort(404)

    log_path = Path(job["log_path"])
    if log_path.exists():
        content = log_path.read_text(encoding="utf-8", errors="replace")
    else:
        content = "Log file not found yet."
    return (
        "<pre style='white-space:pre-wrap;font-family:Consolas,monospace;padding:16px;'>"
        + html.escape(content)
        + "</pre>"
    )


def launch_browser_menu() -> None:
    port = _find_free_port()
    url = f"http://{HOST}:{port}/"
    timer = threading.Timer(0.7, lambda: webbrowser.open(url))
    timer.daemon = True
    timer.start()
    print(f"Opening graphical launcher at {url}")
    app.run(host=HOST, port=port, debug=False, use_reloader=False)
