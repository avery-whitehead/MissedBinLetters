--RECYCLING
DECLARE @rec_cutoff datetime = '2018-06-07 12:00:00.00';
UPDATE HDC_AF_GW_Missed_Collections.dbo.Missed_Collections
SET RECYSacksLetterSent = 1 
WHERE m.id NOT IN (12, 13, 17)
AND RecSacksRequested = 'yes' and m.RECYSacksLetterSent = 0
AND AddedDateTime < @rec_cutoff