--GARDEN WASTE
DECLARE @gw_cutoff datetime = '2018-06-07 12:00:00.00';
SELECT 
	'The Occupier' AS occupier,
	REPLACE(REPLACE(l.ADDRESS_BLOCK, 'North Yorkshire' + CHAR(13) + CHAR(10), ''), CHAR(13) + CHAR(10), '<br>') AS address,
	case_ref,
	ADDRESS_STR_ORG_POSTAL as addr_str,
	GWSacksRequested AS gw_sacks_requested,
	GWLics.num_subs
FROM HDC_AF_GW_Missed_Collections.dbo.Missed_Collections m
JOIN wasscollections.dbo.MV_HDC_LLPG_ADDRESSES_CURRENT l
	ON m.UPRN = l.UPRN
	JOIN (
		SELECT
			uprn,
			SUM(NumRequestedSubs) as num_subs
		FROM [wasscollections].[dbo].[v_Subs-OrderStages_Latest-Paid]
		WHERE subyear = 'Y2'
		GROUP BY uprn
	) GWLics
	ON m.UPRN = GWLics.UPRN
	WHERE m.id NOT IN (12, 13, 17)
  	AND GWSacksRequested = 'yes'
  	AND GWSacksLetterSent = 0
  	AND AddedDateTime < @gw_cutoff