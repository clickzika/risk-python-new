import logging
import os
import sys
from datetime import date, datetime


def send_failure_alert(script_name: str, error_msg: str) -> None:
    """Send an Outlook email alert when a script fails."""
    try:
        import win32com.client
        ol = win32com.client.Dispatch('Outlook.Application')
        mail = ol.CreateItem(0)
        mail.Subject = f"[ALERT] {script_name} FAILED — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        mail.To = 'risk@lhfund.co.th'
        mail.Body = (
            f"Script: {script_name}\n"
            f"Error: {error_msg}\n"
            f"Check logs/{script_name}_*.log for full traceback."
        )
        mail.Send()
    except Exception as alert_err:
        logging.getLogger(script_name).error(f"Failed to send failure alert: {alert_err}")


def is_holiday(check_date: date = None) -> bool:
    """Return True if check_date (default: today) is a Thai public holiday.

    Queries FIN_REG_LHF.dbo.holiday via SQL Server. Falls back to False
    on any connection error so scripts still run if DB is unreachable.
    """
    if check_date is None:
        check_date = date.today()
    try:
        import pyodbc
        conn = pyodbc.connect(
            "DRIVER={SQL Server};"
            "SERVER=192.168.102.7\\DB2008;"
            "DATABASE=FIN_REG_LHF;"
            "Trusted_Connection=yes;",
            timeout=5,
        )
        cursor = conn.cursor()
        cursor.execute(
            "SELECT COUNT(*) FROM dbo.holiday WHERE holiday_date = ?",
            check_date,
        )
        count = cursor.fetchone()[0]
        conn.close()
        return count > 0
    except Exception as e:
        logging.getLogger("risk_logger").warning(
            f"Holiday check failed ({e}) — assuming not a holiday"
        )
        return False


def get_logger(script_name: str) -> logging.Logger:
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)

    date_str = datetime.now().strftime("%Y%m%d")
    log_file = os.path.join(log_dir, f"{script_name}_{date_str}.log")

    logger = logging.getLogger(script_name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger
