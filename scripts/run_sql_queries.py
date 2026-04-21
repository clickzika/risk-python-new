"""Daily SQL query runner.

Connects to SQL Server, runs each query in the catalog, and saves results
as Excel files to SQL_OUTPUT_DIR. Skips public holidays.
"""
import os
import sys
from datetime import date, timedelta

import pandas as pd
import pyodbc

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import (
    DB_MAIN_SERVER,
    SQL_OUTPUT_DIR,
    VAR_LOOKBACK_DAYS,
    VAR_CONFIDENCE_PCT,
)
from risk_logger import get_logger, send_failure_alert, is_holiday, write_status

log = get_logger("run_sql_queries")

_sql_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sql")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sql_path(*parts):
    return os.path.join(_sql_dir, *parts)


def _read_sql(path):
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read()


def _connect(database):
    return pyodbc.connect(
        f"DRIVER={{SQL Server}};"
        f"SERVER={DB_MAIN_SERVER};"
        f"DATABASE={database};"
        "Trusted_Connection=yes;",
        timeout=30,
    )


def _prev_business_day(ref=None):
    """Return the most recent weekday on or before ref (default: today)."""
    d = ref or date.today()
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def _date_range_30d():
    end = _prev_business_day()
    start = end - timedelta(days=30)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def run_query(conn, sql, params=None):
    """Execute sql (with optional .execute params) and return a DataFrame."""
    return pd.read_sql(sql, conn, params=params)


def save_result(df, filename):
    os.makedirs(SQL_OUTPUT_DIR, exist_ok=True)
    today_str = date.today().strftime("%Y%m%d")
    out_path = os.path.join(SQL_OUTPUT_DIR, f"{today_str}_{filename}.xlsx")
    df.to_excel(out_path, index=False)
    log.info(f"Saved {len(df)} rows → {out_path}")
    return out_path


# ---------------------------------------------------------------------------
# Individual query runners
# ---------------------------------------------------------------------------

def run_holdings():
    log.info("Running: daily holdings")
    sql = _read_sql(_sql_path("holdings", "holding_daily.sql"))
    with _connect("INV_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "holdings_daily")


def run_mf_nav():
    log.info("Running: mutual fund NAV")
    start, end = _date_range_30d()
    sql = _read_sql(_sql_path("nav", "mf_nav.sql"))
    sql = sql.replace("###AAA###", start).replace("###BBB###", end)
    with _connect("FIN_REG_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "mf_nav")


def run_pf_nav():
    log.info("Running: private fund NAV")
    start, end = _date_range_30d()
    sql = _read_sql(_sql_path("nav", "pf_nav.sql"))
    sql = sql.replace("###AAA###", start).replace("###BBB###", end)
    with _connect("INV_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "pf_nav")


def run_pvd_nav():
    log.info("Running: provident fund NAV")
    start, end = _date_range_30d()
    sql = _read_sql(_sql_path("nav", "pvd_nav.sql"))
    sql = sql.replace("###AAA###", start).replace("###BBB###", end)
    with _connect("INV_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "pvd_nav")


def run_var():
    log.info("Running: VaR calculation")
    as_of = _prev_business_day().strftime("%Y-%m-%d")
    sql = _read_sql(_sql_path("var", "final_var.sql"))
    sql = (sql
           .replace("###XXX###", as_of)
           .replace("##YY#", str(VAR_LOOKBACK_DAYS))
           .replace("#X#Y", str(VAR_CONFIDENCE_PCT)))
    with _connect("LHF_PERFORMANCE") as conn:
        df = run_query(conn, sql)
    return save_result(df, "var")


def run_bloomberg():
    log.info("Running: Bloomberg MTM FX/EQ")
    sql = _read_sql(_sql_path("bloomberg", "bloomberg.sql"))
    with _connect("LHF_SYSTEM") as conn:
        df = run_query(conn, sql)
    return save_result(df, "bloomberg_mtm")


# ---------------------------------------------------------------------------
# Catalog — ordered list of (name, fn) pairs
# ---------------------------------------------------------------------------

QUERY_CATALOG = [
    ("holdings",  run_holdings),
    ("mf_nav",    run_mf_nav),
    ("pf_nav",    run_pf_nav),
    ("pvd_nav",   run_pvd_nav),
    ("var",       run_var),
    ("bloomberg", run_bloomberg),
]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if is_holiday():
        log.info("Today is a public holiday — skipping SQL queries.")
        write_status("run_sql_queries", "skipped", "Public holiday")
        return

    log.info("=== SQL query runner started ===")
    log.info(f"Output dir: {SQL_OUTPUT_DIR}")

    failed = []
    saved_files = []

    for name, fn in QUERY_CATALOG:
        try:
            out = fn()
            saved_files.append(os.path.basename(out))
        except Exception as e:
            log.error(f"Query '{name}' failed: {e}", exc_info=True)
            failed.append(f"{name}: {e}")

    if failed:
        summary = "; ".join(failed)
        log.warning(f"Completed with {len(failed)} failure(s): {summary}")
        write_status("run_sql_queries", "failed", summary)
        send_failure_alert("run_sql_queries", summary)
    else:
        detail = f"Saved {len(saved_files)} files: {', '.join(saved_files)}"
        log.info(f"=== SQL query runner completed === {detail}")
        write_status("run_sql_queries", "success", detail)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.critical(f"Script failed: {e}", exc_info=True)
        write_status("run_sql_queries", "failed", str(e))
        send_failure_alert("run_sql_queries", str(e))
        sys.exit(1)
