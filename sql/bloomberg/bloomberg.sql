SET NOCOUNT ON

DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);


SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(PRICECODE)
    FROM  LHF_SYSTEM.DBO.LHF_BBG_DL_MTM_FX_EQ_MORNING
    --WHERE FundType = 'Mutual_Fund'
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

-- รันคำสั่ง
EXEC sp_executesql @sql;
