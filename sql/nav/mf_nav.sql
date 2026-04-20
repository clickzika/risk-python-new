SET NOCOUNT ON

DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);

-- สร้าง column string สำหรับ PIVOT
SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(FundCode)
    FROM  [FIN_REG_LHF].[dbo].[View_NAVReturnExcel]
    --WHERE FundType = 'Mutual_Fund'
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
, 1, 1, '');

-- สร้างคำสั่ง SQL แบบ Dynamic
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

-- รันคำสั่ง
EXEC sp_executesql @sql;
