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

# ---------------------------------------------------------------------------
# Inlined SQL queries
# ---------------------------------------------------------------------------

_SQL_HOLDINGS = """\
SELECT  DATEADD(day, -1, CAST(GETDATE() AS DATE)) as Date , *

  FROM [INV_LHF].[INVEST].[HOLDING]
 -- where PORTFOLIOCODE = 'LHVN'
"""

_SQL_MF_NAV = """\
SET NOCOUNT ON

DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);

SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(FundCode)
    FROM  [FIN_REG_LHF].[dbo].[View_NAVReturnExcel]
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
, 1, 1, '');

SET @sql = '
SELECT NAVDate, ' + @columns + '
FROM (
    SELECT NAVDate, FundCode, NAVPerUnit
    FROM  [FIN_REG_LHF].[dbo].[View_NAVReturnExcel]
    WHERE NAVDate BETWEEN ''###AAA###'' AND ''###BBB###''
) AS SourceTable
PIVOT (
    SUM(NAVPerUnit)
    FOR FundCode IN (' + @columns + ')
) AS PivotTable
ORDER BY NAVDate;'

EXEC sp_executesql @sql;
"""

_SQL_PF_NAV = """\
DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);


WITH NAVTemp AS (
select 	t3.VALUEDATE,TRIM(t2.PORTFOLIOCODE) as PORTFOLIOCODE,t3.TOTALNAVAMOUNT AS TOTALNAVAMOUNT
		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit

		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,12) AS NAVPerUnit_ProvidentFund
		, t1.AVAILUNIT
		,t3.TOTALAVAMOUNT as TOTALAVAMOUNT

		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit_PropertyFund
		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,4) AS NAVPerUnit_PrivateFund
	FROM        INV_LHF.INVEST.HOLDINGUNIT AS t1
	inner JOIN	INV_LHF.INVEST.PORTFOLIO AS t2 ON t1.PORTFOLIOID = t2.PORTFOLIOID and t2.ACTIVEFLAG = 'A'
	LEFT OUTER JOIN	INV_LHF.INVEST.TOTALNAV AS t3 ON t1.VALUEDATE = t3.VALUEDATE AND t1.PORTFOLIOID = t3.PORTFOLIOID
	WHERE     (t1.AVAILUNIT > 0)
	and (t2.PORTFOLIOCODE like 'PF%')
	AND (t1.VALUEDATE between  '###AAA###' and '###BBB###')
	GROUP BY t3.VALUEDATE,t2.PORTFOLIOCODE,t3.TOTALNAVAMOUNT, t3.TOTALAVAMOUNT , t1.AVAILUNIT
	)

SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(PORTFOLIOCODE)
    FROM  NAVTemp
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
, 1, 1, '');


SET @sql = '

WITH NAVTemp AS (
select 	t3.VALUEDATE,TRIM(t2.PORTFOLIOCODE) as PORTFOLIOCODE,t3.TOTALNAVAMOUNT AS TOTALNAVAMOUNT
		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit

		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,12) AS NAVPerUnit_ProvidentFund
		, t1.AVAILUNIT
		,t3.TOTALAVAMOUNT as TOTALAVAMOUNT

		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit_PropertyFund
		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,4) AS NAVPerUnit_PrivateFund
	FROM        INV_LHF.INVEST.HOLDINGUNIT AS t1
	inner JOIN	INV_LHF.INVEST.PORTFOLIO AS t2 ON t1.PORTFOLIOID = t2.PORTFOLIOID and t2.ACTIVEFLAG = ''A''
	LEFT OUTER JOIN	INV_LHF.INVEST.TOTALNAV AS t3 ON t1.VALUEDATE = t3.VALUEDATE AND t1.PORTFOLIOID = t3.PORTFOLIOID
	WHERE     (t1.AVAILUNIT > 0)
	--and (t2.PORTFOLIOCODE in ())
	AND (t1.VALUEDATE between  ''###AAA###'' and ''###BBB###'')
	GROUP BY t3.VALUEDATE,t2.PORTFOLIOCODE,t3.TOTALNAVAMOUNT, t3.TOTALAVAMOUNT , t1.AVAILUNIT
	)


	select VALUEDATE , ' + @columns + ' from
	(
		SELECT VALUEDATE , NAVPerUnit_PrivateFund , PORTFOLIOCODE
		FROM  NAVTemp
		WHERE VALUEDATE BETWEEN ''###AAA###'' AND ''###BBB###''
	) AS SourceTable
	PIVOT (
		SUM(NAVPerUnit_PrivateFund)
		FOR PORTFOLIOCODE IN (' + @columns + ')
	) AS PivotTable
	ORDER BY VALUEDATE;';


EXEC sp_executesql @sql;
"""

_SQL_PVD_NAV = """\
DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);


WITH NAVTemp AS (
select 	t3.VALUEDATE,TRIM(t2.PORTFOLIOCODE) as PORTFOLIOCODE,t3.TOTALNAVAMOUNT AS TOTALNAVAMOUNT
		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit

		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,12) AS NAVPerUnit_ProvidentFund
		, t1.AVAILUNIT
		,t3.TOTALAVAMOUNT as TOTALAVAMOUNT

		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit_PropertyFund
		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,4) AS NAVPerUnit_PrivateFund
	FROM        INV_LHF.INVEST.HOLDINGUNIT AS t1
	inner JOIN	INV_LHF.INVEST.PORTFOLIO AS t2 ON t1.PORTFOLIOID = t2.PORTFOLIOID and t2.ACTIVEFLAG = 'A'
	LEFT OUTER JOIN	INV_LHF.INVEST.TOTALNAV AS t3 ON t1.VALUEDATE = t3.VALUEDATE AND t1.PORTFOLIOID = t3.PORTFOLIOID
	WHERE     (t1.AVAILUNIT > 0)
	and (t2.PORTFOLIOCODE like 'PVD%')
	AND (t1.VALUEDATE between  '###AAA###' and '###BBB###')
	GROUP BY t3.VALUEDATE,t2.PORTFOLIOCODE,t3.TOTALNAVAMOUNT, t3.TOTALAVAMOUNT , t1.AVAILUNIT
	)

SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(PORTFOLIOCODE)
    FROM  NAVTemp
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
, 1, 1, '');


SET @sql = '

WITH NAVTemp AS (
select 	t3.VALUEDATE,TRIM(t2.PORTFOLIOCODE) as PORTFOLIOCODE,t3.TOTALNAVAMOUNT AS TOTALNAVAMOUNT
		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit

		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,12) AS NAVPerUnit_ProvidentFund
		, t1.AVAILUNIT
		,t3.TOTALAVAMOUNT as TOTALAVAMOUNT

		,ROUND(ROUND(t3.TOTALNAVAMOUNT / t1.AVAILUNIT,5), 4 ,1) AS NAVPerUnit_PropertyFund
		,round(
			CAST(
				cast(t3.TOTALNAVAMOUNT as float) / cast( t1.AVAILUNIT as float)
				AS decimal(32,13))
			,4) AS NAVPerUnit_PrivateFund
	FROM        INV_LHF.INVEST.HOLDINGUNIT AS t1
	inner JOIN	INV_LHF.INVEST.PORTFOLIO AS t2 ON t1.PORTFOLIOID = t2.PORTFOLIOID and t2.ACTIVEFLAG = ''A''
	LEFT OUTER JOIN	INV_LHF.INVEST.TOTALNAV AS t3 ON t1.VALUEDATE = t3.VALUEDATE AND t1.PORTFOLIOID = t3.PORTFOLIOID
	WHERE     (t1.AVAILUNIT > 0)
	--and (t2.PORTFOLIOCODE in ())
	AND (t1.VALUEDATE between  ''###AAA###'' and ''###BBB###'')
	GROUP BY t3.VALUEDATE,t2.PORTFOLIOCODE,t3.TOTALNAVAMOUNT, t3.TOTALAVAMOUNT , t1.AVAILUNIT
	)


	select VALUEDATE , ' + @columns + ' from
	(
		SELECT VALUEDATE , NAVPerUnit_ProvidentFund , PORTFOLIOCODE
		FROM  NAVTemp
		WHERE VALUEDATE BETWEEN ''###AAA###'' AND ''###BBB###''
	) AS SourceTable
	PIVOT (
		SUM(NAVPerUnit_ProvidentFund)
		FOR PORTFOLIOCODE IN (' + @columns + ')
	) AS PivotTable
	ORDER BY VALUEDATE;';


EXEC sp_executesql @sql;
"""

_SQL_VAR = """\
Set NOCOUNT ON

Declare @Date1 as char(10) = '###XXX###'
Declare @Numdate as decimal = '##YY#'
Declare @Confidence as decimal = '#X#Y'




;WITH RankedNAV AS (
    SELECT *,
		   LAG(NAVPerUnitDiv, 1) OVER (PARTITION BY FundCode ORDER BY NAVDate ASC) AS Nav_previous,
		   LAG(NAVDate, 1) OVER (PARTITION BY FundCode ORDER BY NAVDate ASC) AS DateNav_previous,
		   ((NavPerUnitDiv / (LAG(NAVPerUnitDiv, 1) OVER (PARTITION BY FundCode ORDER BY NAVDate ASC))) -1)*100 AS DailyReturn
    FROM [LHF_PERFORMANCE].[dbo].[ViewFundNavAll]

    WHERE NAVDate <= @Date1  AND DATENAME(dw, NAVDate) NOT IN ('Saturday', 'Sunday')
	AND NAVPerUnitDiv is not null  and NAVPerUnitDiv != '0'  and NAVDate not in (SELECT HolidayDate
 FROM [192.168.102.7\\DB2008].[FIN_REG_LHF].[dbo].holiday where [CalendarID] = 1)
),
CTEa AS (
	SELECT *,ROW_NUMBER() OVER (PARTITION BY FundCode ORDER BY NAVDate DESC) AS RowNum
	FROM RankedNAV
	where Nav_previous is not null --and FundCode = 'LHMM-A'
		AND FundType = 'Mutual_Fund' AND DailyReturn != '0'
),
CTEb AS (
	SELECT *
	FROM CTEa
	WHERE RowNum <= @Numdate
),
CTEc AS (

    SELECT *,
    PERCENTILE_CONT((100-@Confidence)/100) WITHIN GROUP (ORDER BY DailyReturn)
    OVER (PARTITION BY FundCode) AS VaR
    FROM CTEb
)
SELECT NAVDate , FundCode , VaR
FROM CTEc
where NAVDate = @Date1


--ORDER BY FundCode, NAVDate DESC
"""

_SQL_BLOOMBERG = """\
SET NOCOUNT ON

DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);


SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(PRICECODE)
    FROM  LHF_SYSTEM.DBO.LHF_BBG_DL_MTM_FX_EQ_MORNING
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
, 1, 1, '');

SET @sql = '
SELECT MTMDATE, ' + @columns + '
FROM (
    SELECT MTMDATE, PRICECODE, PX_LAST
    FROM  LHF_SYSTEM.DBO.LHF_BBG_DL_MTM_FX_EQ_MORNING
) AS SourceTable
PIVOT (
    SUM(PX_LAST)
    FOR PRICECODE IN (' + @columns + ')
) AS PivotTable
ORDER BY MTMDATE;';

EXEC sp_executesql @sql;
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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
    with _connect("INV_LHF") as conn:
        df = run_query(conn, _SQL_HOLDINGS)
    return save_result(df, "holdings_daily")


def run_mf_nav():
    log.info("Running: mutual fund NAV")
    start, end = _date_range_30d()
    sql = _SQL_MF_NAV.replace("###AAA###", start).replace("###BBB###", end)
    with _connect("FIN_REG_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "mf_nav")


def run_pf_nav():
    log.info("Running: private fund NAV")
    start, end = _date_range_30d()
    sql = _SQL_PF_NAV.replace("###AAA###", start).replace("###BBB###", end)
    with _connect("INV_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "pf_nav")


def run_pvd_nav():
    log.info("Running: provident fund NAV")
    start, end = _date_range_30d()
    sql = _SQL_PVD_NAV.replace("###AAA###", start).replace("###BBB###", end)
    with _connect("INV_LHF") as conn:
        df = run_query(conn, sql)
    return save_result(df, "pvd_nav")


def run_var():
    log.info("Running: VaR calculation")
    as_of = _prev_business_day().strftime("%Y-%m-%d")
    sql = (_SQL_VAR
           .replace("###XXX###", as_of)
           .replace("##YY#", str(VAR_LOOKBACK_DAYS))
           .replace("#X#Y", str(VAR_CONFIDENCE_PCT)))
    with _connect("LHF_PERFORMANCE") as conn:
        df = run_query(conn, sql)
    return save_result(df, "var")


def run_bloomberg():
    log.info("Running: Bloomberg MTM FX/EQ")
    with _connect("LHF_SYSTEM") as conn:
        df = run_query(conn, _SQL_BLOOMBERG)
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
