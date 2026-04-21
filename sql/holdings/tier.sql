Set NOCOUNT ON

Declare @Date as char(10) = '###XXX###'
Declare @Date2 as char(10) = '###YYY###'

select ValueDate,Trim(PortfolioCode) as Port ,PerTier1/100 as PerTier1,COALESCE(PerTier2, 0)/100 as PerTier2,PerSumTier1Tier2/100 as  PerSumTier1Tier2
from [LHF_PERFORMANCE].[dbo].[LHF_Tier]
where ValueDate between @Date and @Date2
order by ValueDate
