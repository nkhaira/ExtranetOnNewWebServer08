<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

    Call Connect_SiteWide
    
    ' --------------------------------------------------------------------------------------
    ' The following code does some cleanup and transformations for items that have expired or items retired from the
    ' Oracle Literature DB and moves them to Archive Status.
    ' --------------------------------------------------------------------------------------  

    ' Set the Active_Flag field in the Oracle Nightly View to False (no-meaning field from the PP perspective, so we can use it to flag active item numbers)

    SQL = "UPDATE dbo.Literature_Items_US SET ACTIVE_FLAG = 0"
    conn.execute SQL
    
    ' Reset Active_Flag field in Oracle Nightly View to True for all LIVE Item Numbers
    
    SQL = "UPDATE dbo.Literature_Items_US SET Active_Flag=-1, PDF=-1 " &_
          " WHERE (Deliverable_ID IN (SELECT Deliverable_ID " &_
          "   FROM   dbo.Literature_Items_US " &_
          "     WHERE ((STATUS = 'Active') " &_
          "   		AND (STATUS_NAME = 'Final Loaded' OR STATUS_NAME = 'Reprint') " &_
          "   		AND ([ACTION] = 'Complete' OR [ACTION] = 'N/A' OR [ACTION] = 'Web Only'))))"
    conn.execute SQL

    ' Reset all invalid Item_Numbers and Revision to NULL if they are not numeric
    
    SQL = "UPDATE dbo.Calendar " &_
          "   SET Item_Number = NULL, Revision_Code = NULL " &_
          "     WHERE (ID IN (SELECT ID " &_
          "     FROM dbo.calendar " &_
          "     WHERE (Item_Number LIKE '%[a-z]%')))"
    conn.execute SQL

    ' Blanket Reset of PP Status to Archive
    
    SQL = "UPDATE dbo.Calendar " &_
          "   SET Status=2, Status_Comment='The Item Number and Revision of the Asset Container in the Partner Portal is not Active in the Oracle Marketing On-Line Deliverables.' " &_
          "     WHERE (ID IN (SELECT ID " &_
          "     FROM   dbo.Calendar " &_
          "     WHERE (cast(Item_Number as INT)>=1000000 and cast(Item_Number as int) <=7999999 )))"
    conn.execute SQL
    
    ' Auto Update Item Number Revision_Code to Oracle Revision for Non PDF/POD/PRINT items
    SQL = "SELECT dbo.Calendar.ID AS ID, dbo.Calendar.Revision_Code, dbo.Literature_Items_US.REVISION " &_
          "FROM   dbo.Literature_Items_US RIGHT OUTER JOIN " &_
          "       dbo.Calendar ON dbo.Literature_Items_US.REVISION <> dbo.Calendar.Revision_Code AND  " &_
          "       dbo.Literature_Items_US.ITEM = dbo.Calendar.Item_Number " &_
          "WHERE (dbo.Literature_Items_US.ACTIVE_FLAG = - 1) AND (dbo.Literature_Items_US.PDF = 0) AND (dbo.Literature_Items_US.POD = 0) AND  " &_
          "      (dbo.Literature_Items_US.[PRINT] = 0) " &_
          "ORDER BY ID"
    Set rsItem = Server.CreateObject("ADODB.Recordset")
    rsItem.Open SQL, conn, 3, 3
    
    do while not rsItem.EOF
      SQL = "UPDATE dbo.Calendar SET Revision_Code='" & rsItem("REVISION") & "' WHERE ID=" & rsItem("ID")
      conn.execute SQL
      rsItem.MoveNext
    loop
    
    rsItem.close
    set rsItem = nothing

    ' Reset Status for Active Oracle Item_Numbers
    
    SQL = "UPDATE dbo.Calendar SET Status=1, Status_Comment=NULL " &_
          "   WHERE (ID IN (SELECT ID " &_
          "     FROM dbo.Calendar LEFT OUTER JOIN  " &_
          "          dbo.Literature_Items_US ON dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM AND dbo.Calendar.Revision_Code = dbo.Literature_Items_US.REVISION " &_
          "     WHERE   (dbo.Literature_Items_US.ACTIVE_FLAG = -1)))"
    conn.execute SQL

    ' Archive Non-Exsisting Oracle Items from Calendar DB
    
    SQL = "UPDATE dbo.Calendar " &_
          "   SET Status=2, Status_Comment='This Item Number does not exist in Oracle Marketing On-Line Deliverables.' " &_
          "   WHERE (ID IN (SELECT ID " &_
          "   FROM  dbo.Calendar LEFT OUTER JOIN " &_
          "         dbo.Literature_Items_US_Oracle ON dbo.Calendar.Item_Number = dbo.Literature_Items_US_Oracle.ITEM_NUMBER " &_
          "           WHERE (dbo.Calendar.Status = 1 OR " &_
          "              dbo.Calendar.Status = 2) AND (dbo.Calendar.Item_Number IS NOT NULL OR " &_
          "               dbo.Calendar.Item_Number <> '') AND (dbo.Literature_Items_US_Oracle.ITEM_NUMBER IS NULL OR " &_
          "               dbo.Literature_Items_US_Oracle.ITEM_NUMBER = '' OR " &_
          "               dbo.Literature_Items_US_Oracle.ITEM_NUMBER = 'No Value') AND (cast(dbo.Calendar.Item_Number as INT)>=1000000 and cast(dbo.Calendar.Item_Number as int) <=7999999 )))"
    conn.execute SQL
    
    ' Set status to archive because there was a hard expiration on the asset.
    
    SQL = "UPDATE dbo.Calendar " &_
          "SET    Status = 2, Status_Comment = 'An asset container archive date has been reached for this Item Number.' " &_
          "WHERE (Status = 1 OR Status = 2) AND (BDate = EDate) AND (XDays > 0) AND (XDate < '" & Date() & "') " &_
          "   OR (Status = 1 OR Status = 2) AND (BDate <> EDate) AND (XDate < '" & Date() & "')"
    conn.execute SQL
    
    Call Disconnect_SiteWide
%>
