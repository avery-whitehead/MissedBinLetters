DECLARE @html nvarchar(MAX)
DECLARE @uprn varchar(12) = '200003299567'
DECLARE @backdatedBy varchar(4) = '-1'
DECLARE @text varchar(4) = 'no'
DECLARE @numWeeks int = 1
DECLARE @tableHeight varchar(6) = '250px'
DECLARE @pccCSS varchar(max) = ''
DECLARE @datelimit date = '2018-11-30'

IF (DATEADD(WEEK, @numWeeks, GETDATE()) > @datelimit)
	SET @numWeeks = DATEDIFF(ww, GETDATE(), @datelimit) + @backdatedBy

IF OBJECT_ID('tempdb..#temp') IS NOT NULL
	DROP TABLE #temp

SELECT *
INTO #temp
FROM dbo.[tvfPropertyCollectionsCalendar_U] (
	@uprn,
	@backdatedBy,
	@numWeeks,
	@text
)

DECLARE @tabCols nvarchar(MAX) = ''
DECLARE @colCount int = 0
DECLARE @gwCol int = 0
IF (
	SELECT COUNT(Refuse)
	FROM #temp
	WHERE LEN(Refuse) > 0
) > 0
	BEGIN
		SET @tabCols = @tabCols + ',Refuse'
		SET @colCount = @colCount + 1
	END

IF (
	SELECT COUNT(Recycling)
	FROM #temp
	WHERE LEN(Recycling) > 0
) > 0
	BEGIN
		SET @tabCols = @tabCols + ',Recycling'
		SET @colCount = @colCount + 1
	END

IF (
	SELECT COUNT(Mix)
	FROM #temp
	WHERE LEN(Mix) > 0
) > 0
    BEGIN
       SET @tabCols = @tabCols + ',Mix as NEWRecycling'
       SET @colCount = @colCount + 1
	END

IF (
	SELECT COUNT(Glass)
	FROM #temp
	WHERE LEN(Glass) > 0
) > 0
    BEGIN
		SET @tabCols = @tabCols + ',Glass'
		SET @colCount = @colCount + 1
    END

IF (
	SELECT COUNT(Garden)
	FROM #temp
	WHERE len(Garden) > 0
) > 0
	BEGIN
		SET @tabCols = @tabCols + ',Garden'
		SET @colCount = @colCount + 1
		SET @gwCol = 1
    END

DECLARE @tableRowWidth varchar(10) = (100 - 2) / (@colCount + 1)
DECLARE @queryStr nvarchar(MAX) = ''
SET @queryStr = N'select WeekCommencing ' + @tabCols + ' from #temp'

EXEC dbo.p_pccQueryToHtmlTable
	@html = @html OUTPUT,
	@query = @queryStr,
	@orderBy = NULL
SELECT
	REPLACE(
		REPLACE(
			REPLACE(
				REPLACE(
					REPLACE(
						REPLACE(
							REPLACE(
								REPLACE(
									REPLACE(
										REPLACE(
											REPLACE(@html, '|^', '<'),
										'|$', '>'),
									'|/', '/'),
								'<td xsi:nil="true"/>', '<td>'),
							'<td>', '<td id="licences" class="hdc-td-greenborder" style="width:' + @tableRowWidth + '%;">'),
						'<tr xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">', '<tr>'),
					'<th>', '<th class="hdc-td-greenborder">'),
				'WeekCommencing', 'Week Starting'),
			'500px', @tableHeight),
		'<div class="table"', '<div id="collections-table" class="table ' + @pccCSS + '"'),
	'NEWRecycling', 'Recycling')
AS pccTableHTML
