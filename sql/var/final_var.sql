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
 FROM [192.168.102.7\DB2008].[FIN_REG_LHF].[dbo].holiday where [CalendarID] = 1)
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