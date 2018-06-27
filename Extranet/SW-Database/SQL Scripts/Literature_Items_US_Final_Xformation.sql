
CREATE PROCEDURE [dbo].[Literature_Items_US_Final_Xformation] AS

DECLARE
	@id int,
	@rev_code varchar(20),
	@rev varchar(10)

/*Set the Active_Flag field in the Oracle Nightly View to False 
(unused field from our perspective, but we can use it to flag active item numbers)*/

UPDATE dbo.Literature_Items_US SET ACTIVE_FLAG = 0

/*Reset Active_Flag field in Oracle Nightly View to True for all LIVE Item Numbers */

UPDATE dbo.Literature_Items_US SET Active_Flag=-1 
WHERE (Deliverable_ID IN (SELECT Deliverable_ID 
FROM   dbo.Literature_Items_US 
WHERE ((STATUS = 'Active') 
AND (STATUS_NAME = 'Final Loaded' OR STATUS_NAME = 'Reprint' OR STATUS_NAME = 'Web Only') 
AND ([ACTION] = 'Complete' OR [ACTION] = 'N/A'))))

/*Reset all invalid Item_Numbers and Revision to NULL if they are not numeric*/

UPDATE dbo.Calendar 
SET Item_Number = NULL, Revision_Code = NULL 
WHERE ([ID] IN (SELECT [ID]  
FROM dbo.calendar  
WHERE (Item_Number LIKE '%[a-z]%')))
 
/*Blanket Reset of PP Status to Archive except those with Status Override*/

UPDATE dbo.Calendar  
SET Status=2, 
	Status_Comment='This Item Number and Revision is not Active in the Oracle Marketing On-Line Deliverables.'  
WHERE ([ID] IN (SELECT [ID]  
FROM   dbo.Calendar  
WHERE (Status <> 0 and cast(Item_Number as INT)>=1000000 and cast(Item_Number as int) <=7999999 and Status_Override <> -1 )))

/*Update Status Override Commet*/

UPDATE dbo.Calendar  
SET Status_Comment='AMS LIVE Override Enabled. This Item Number and Revision is not Active in the Oracle Marketing On-Line Deliverables.'  
WHERE ([ID] IN (SELECT [ID]  
FROM   dbo.Calendar  
WHERE (Status 1 and cast(Item_Number as INT)>=1000000 and cast(Item_Number as int) <=7999999 and Status_Override = -1 )))

/*Auto Update Item Number Revision_Code to Oracle Revision for Non PDF/POD/PRINT items */

CREATE TABLE #Lit_temp([ID] int, Revision_Code varchar(20), Revision varchar(10))

INSERT INTO #Lit_temp 
SELECT dbo.Calendar.[ID] AS [ID], dbo.Calendar.Revision_Code, dbo.Literature_Items_US.REVISION  
FROM   dbo.Literature_Items_US RIGHT OUTER JOIN  
dbo.Calendar ON dbo.Literature_Items_US.REVISION <> dbo.Calendar.Revision_Code AND   
dbo.Literature_Items_US.ITEM = dbo.Calendar.Item_Number  
WHERE (dbo.Literature_Items_US.ACTIVE_FLAG = - 1) AND (dbo.Literature_Items_US.PDF = 0) AND (dbo.Literature_Items_US.POD = 0) AND   
(dbo.Literature_Items_US.[PRINT] = 0)  
ORDER BY [ID]

WHILE (SELECT count(*) from #lit_temp) > 0
BEGIN
	Set ROWCOUNT 1
	SELECT @id=[ID], @rev_code=Revision_Code, @rev=Revision FROM #Lit_temp
	DELETE FROM #Lit_temp
	Set ROWCOUNT 0
	UPDATE dbo.Calendar SET Revision_Code=@rev WHERE [ID]=@id	
END

DROP TABLE #Lit_temp

/* Reset Status for Active Oracle Item_Numbers */

UPDATE dbo.Calendar SET Status=1, Status_Comment=NULL, Status_Override=0  
WHERE ([ID] IN (SELECT [ID]  
FROM dbo.Calendar LEFT OUTER JOIN   
dbo.Literature_Items_US ON dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM AND dbo.Calendar.Revision_Code = dbo.Literature_Items_US.REVISION  
WHERE   (dbo.Literature_Items_US.ACTIVE_FLAG = -1) and (dbo.Calendar.Status <> 0)))

/*Archive Non-Exsisting Oracle Items from Calendar DB */

UPDATE dbo.Calendar  
SET Status=2, Status_Comment='This Item Number does not exist in Oracle Marketing On-Line Deliverables.', Status_Override=0  
WHERE ([ID] IN (SELECT [ID]  
FROM  dbo.Calendar LEFT OUTER JOIN  
dbo.Literature_Items_US_Oracle ON dbo.Calendar.Item_Number = dbo.Literature_Items_US_Oracle.ITEM_NUMBER  
WHERE (dbo.Calendar.Status = 1 OR  
dbo.Calendar.Status = 2) AND (dbo.Calendar.Item_Number IS NOT NULL OR  
dbo.Calendar.Item_Number <> '') AND (dbo.Literature_Items_US_Oracle.ITEM_NUMBER IS NULL OR  
dbo.Literature_Items_US_Oracle.ITEM_NUMBER = '' OR  
dbo.Literature_Items_US_Oracle.ITEM_NUMBER = 'No Value') 
AND (cast(dbo.Calendar.Item_Number as INT)>=1000000 and cast(dbo.Calendar.Item_Number as int) <=7999999 )))

/*Set status to archive because there was a hard expiration on the asset.*/

UPDATE dbo.Calendar  
SET    Status = 2, Status_Comment = 'An asset container archive date has been reached for this Item Number.', Status_Override=0  
WHERE (Status = 1 OR Status = 2) AND (BDate = EDate) AND (XDays > 0) AND (XDate < GetDate())  
OR (Status = 1 OR Status = 2) AND (BDate <> EDate) AND (XDate < GetDate())

GO
