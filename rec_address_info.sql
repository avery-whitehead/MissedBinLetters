--RECYCLING
DECLARE @rec_cutoff datetime = '2018-06-07 12:00:00.00';
SELECT 
	'The Occupier' AS occupier,
	REPLACE(REPLACE(l.ADDRESS_BLOCK, 'North Yorkshire' + CHAR(13) + CHAR(10), ''), CHAR(13) + CHAR(10), '<br>') AS address,
	case_ref,
	RecSacksRequested AS rec_sacks_requested
FROM HDC_AF_GW_Missed_Collections.dbo.Missed_Collections m
JOIN wasscollections.dbo.MV_HDC_LLPG_ADDRESSES_CURRENT l
	ON m.UPRN = l.UPRN
  	WHERE m.id NOT IN (12, 13, 17)
  	AND RecSacksRequested = 'yes' and m.RECYSacksLetterSent = 0
  	AND AddedDateTime < @rec_cutoff