SELECT
	'The Occupier' AS occup,
	ADDRESS_STR AS addr,
	UPRN AS uprn
FROM (
	SELECT
		l.UPRN,
		l.ADDRESS_STR,
		CASE
			WHEN oldREF = 0 THEN NULL
			ELSE oldREF
		END AS oldREF,
		CASE
			WHEN oldRECY = 0 THEN NULL
			ELSE oldRECY
		END AS oldRECY,
		CASE
			WHEN oldMIX = 0 THEN NULL
			ELSE oldMIX
		END AS oldMIX,
		CASE
			WHEN oldGLASS = 0 THEN NULL
			ELSE oldGLASS
		END AS oldGLASS,
		CASE
			WHEN oldGW = 0 THEN NULL
			ELSE oldGW
		END AS oldGW,
		CASE
			WHEN newREF = 0 THEN NULL
			ELSE newREF
		END AS newREF,
		CASE
			WHEN newRECY = 0 THEN NULL
			ELSE newRECY
		END AS newRECY,
		CASE
			WHEN newMIX = 0 THEN NULL
			ELSE newMIX
		END AS newMIX,
		CASE
			WHEN newGLASS = 0 THEN NULL
			ELSE newGLASS
		END AS newGLASS,
		CASE WHEN newGW = 0 THEN NULL
			ELSE newGW
		END AS newGW,
		CASE
			WHEN oldREF <> newREF THEN 'X'
			ELSE NULL
		END AS REF,
		CASE
			WHEN oldRECY > newRECY THEN 'X'
			ELSE NULL
		END AS RECY,
		CASE
			WHEN oldMIX <> newMIX THEN 'X'
			ELSE NULL
		END AS MIX,
		CASE
			WHEN oldGLASS <> newGLASS THEN 'X'
			ELSE NULL
		END AS GLASS,
		CASE
			WHEN oldGW <> newGW THEN 'X'
			ELSE NULL
		END AS GW
	FROM dbo.MV_HDC_LLPG_ADDRESSES_CURRENT l
	LEFT JOIN (
		SELECT
			UPRN,
			ISNULL([REF], 0) AS oldREF,
			ISNULL([RECY], 0) AS oldRECY,
			ISNULL([MIX], 0) AS oldMIX,
			ISNULL([GLASS], 0) AS oldGLASS,
			ISNULL([GW], 0) AS oldGW
		FROM (
			SELECT
				UPRN,
				psr.serviceid,
				MIN(r.ScheduleDayID) AS ScheduleDayID
			FROM dbo.PropertyServiceRounds_180531_061215 psr
			LEFT JOIN (
				SELECT *
				FROM dbo.Rounds
				WHERE roundera = 2
			) r
			ON psr.roundid = r.RoundID
			AND psr.ServiceID = r.ServiceID
			WHERE psr.RoundEra = 2
			GROUP BY uprn, psr.serviceid
		) a
		PIVOT (
			MIN(ScheduleDayID)
			FOR serviceid
			IN ([REF], [RECY], [MIX], [GLASS], [GW])
		) pvt 
	) o
	ON l.uprn = o.uprn
	LEFT JOIN (
		SELECT
			UPRN,
			ISNULL([REF], 0) AS newREF,
			ISNULL([RECY], 0) AS newRECY,
			isnull([MIX], 0) AS newMIX,
			isnull([GLASS], 0) AS newGLASS,
			isnull([GW], 0) AS newGW
		FROM (
			SELECT
				UPRN,
				psr.serviceid,
				MIN(r.ScheduleDayID) AS ScheduleDayID
			FROM dbo.PropertyServiceRounds psr
			LEFT JOIN (
				SELECT *
				FROM dbo.Rounds
				WHERE roundera = 2
			) r
			ON psr.roundid = r.RoundID
			AND psr.ServiceID = r.ServiceID
			WHERE psr.RoundEra = 2
			GROUP BY uprn, psr.serviceid
		) a
		PIVOT (
			MIN(ScheduleDayID)
			FOR serviceid
			IN ([REF], [RECY], [MIX], [GLASS], [GW])
		) pvt 
	) n
	ON l.uprn = n.uprn
	WHERE oldREF <> newREF
		OR oldRECY > newRECY
		OR oldMIX <> newMIX
		OR oldGLASS <> newGLASS
		OR oldGW <> newGW
) a
WHERE REF IS NOT NULL
	OR RECY IS NOT NULL
	OR MIX IS NOT NULL
	OR GLASS IS NOT NULL