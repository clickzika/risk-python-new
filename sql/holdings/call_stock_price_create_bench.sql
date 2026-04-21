SET NOCOUNT ON

DECLARE @columns NVARCHAR(MAX),
        @sql NVARCHAR(MAX);


SELECT @columns = STUFF((
    SELECT DISTINCT ',' + QUOTENAME(BenchmarkCode)
    FROM  [dbo].[ViewBenchmark]
    --WHERE FundType = 'Mutual_Fund'
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
, 1, 1, '');

SET @sql = '
SELECT ValueDate, ' + @columns + '
FROM (
    	SELECT [dbo].[PERFORMANCE_BENCHMARK_PRICE].ValueDate , [dbo].[ViewBenchmark].BenchmarkCode , [dbo].[PERFORMANCE_BENCHMARK_PRICE].Price
    FROM [dbo].[PERFORMANCE_BENCHMARK_PRICE]
    join [dbo].[ViewBenchmark] ON [dbo].[PERFORMANCE_BENCHMARK_PRICE].BenchmarkID = [dbo].[ViewBenchmark].BenchmarkID
) AS SourceTable
PIVOT (
    SUM(Price)
    FOR BenchmarkCode IN (' + @columns + ')
) AS PivotTable
ORDER BY ValueDate;';

-- รันคำสั่ง
EXEC sp_executesql @sql;
