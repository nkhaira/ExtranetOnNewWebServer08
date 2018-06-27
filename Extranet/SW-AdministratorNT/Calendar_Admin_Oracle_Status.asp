<%
' Set Status Initial

if isnumeric(Record_ID) then

  SQLStatus = "UPDATE dbo.Calendar " &_
              "SET Status_Comment='This Item Number and Revision is not Active in the Oracle Marketing On-Line Deliverables.' " &_
              "WHERE (cast(Item_Number as INT)>=1000000 AND cast(Item_Number as int) <=7999999) " &_
              "AND ID=" & Record_ID
  
  conn.Execute(SQLStatus)
  
  ' Reset Status for Active Oracle Item_Numbers
  
  SQLStatus = "UPDATE dbo.Calendar " &_
              "SET Status_Comment=NULL " &_
              "WHERE ([ID] IN (SELECT [ID] " &_
              "FROM dbo.Calendar LEFT OUTER JOIN " &_
              "dbo.Literature_Items_US ON dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM AND dbo.Calendar.Revision_Code = dbo.Literature_Items_US.REVISION " &_
              "WHERE dbo.Literature_Items_US.ACTIVE_FLAG = -1 " &_
              "AND dbo.Calendar.ID=" & Record_ID & "))"
  
  conn.Execute(SQLStatus)
  
  ' Flag Non-Exsisting Oracle Items in Calendar DB
  
  SQLStatus = "UPDATE dbo.Calendar " &_
              "SET Status_Comment='This Item Number does not exist in Oracle Marketing On-Line Deliverables.' " &_
              "WHERE ([ID] IN (SELECT [ID] " &_
              "FROM  dbo.Calendar LEFT OUTER JOIN " &_
              "dbo.Literature_Items_US_Oracle ON dbo.Calendar.Item_Number = dbo.Literature_Items_US_Oracle.ITEM_NUMBER " &_
              "WHERE (dbo.Calendar.Item_Number IS NOT NULL OR dbo.Calendar.Item_Number <> '') " &_
              "AND (dbo.Literature_Items_US_Oracle.ITEM_NUMBER IS NULL OR dbo.Literature_Items_US_Oracle.ITEM_NUMBER = '' OR dbo.Literature_Items_US_Oracle.ITEM_NUMBER = 'No Value') " &_
              "AND (cast(dbo.Calendar.Item_Number as INT)>=1000000 AND cast(dbo.Calendar.Item_Number as int) <=7999999 ) " &_
              "AND dbo.Calendar.ID=" & Record_ID & "))"
              
  conn.Execute(SQLStatus)
  
  set SQLStatus = nothing
  
end if
%>