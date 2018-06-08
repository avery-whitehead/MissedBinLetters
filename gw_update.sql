--GARDEN WASTE
DECLARE @gw_cutoff datetime = '2018-06-07 12:00:00.00';
UPDATE HDC_AF_GW_Missed_Collections.dbo.Missed_Collections
SET GWSacksLetterSent = 1
WHERE HDC_AF_GW_Missed_Collections.dbo.Missed_Collections.id NOT IN (12, 13, 17)
AND GWSacksRequested = 'yes' AND GWSacksLetterSent = 0
and AddedDateTime < @gw_cutoff