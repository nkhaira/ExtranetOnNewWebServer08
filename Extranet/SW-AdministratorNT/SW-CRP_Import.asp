<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_eStore.asp" -->
<%

  Call Connect_eStoreDatabase
  response.write "Begin Import<BR>"
  SQL = "SELECT * " &_
        "FROM crp order by model"
        
  Set rsCRP = Server.CreateObject("ADODB.Recordset")
  rsCRP.open SQL,eConn,3,1,1
  
  do while not rsCRP.EOF
  
    SQL = "INSERT INTO vcturbo_replaceable_parts_xref (Model,Part_Description,Part_Exception,Serial_Range,Item_Number,Family,Update_Date,Update_By) "
    SQL = SQL & "VALUES("
    
    if not isblank(rsCRP("Model")) then
      SQL = SQL & "'" & replace(rsCRP("Model"),"'","''") & "', "
    else
      SQL = SQL & "NULL,"
    end if
    if not isblank(rsCRP("Part Description")) then
      SQL = SQL & "'" & replace(rsCRP("Part Description"),"'","''") & "', "
    else
      SQL = SQL & "NULL,"
    end if

    if not isblank(rsCRP("Part Exception")) then
      SQL = SQL & "'" & replace(rsCRP("Part Exception"),"'","''") & "', "
    else
      SQL = SQL & "NULL,"
    end if

    if not isblank(rsCRP("Serial Number Range")) then
      SQL = SQL & "'" & replace(rsCRP("Serial Number Range"),"'","''") & "', "
    else
      SQL = SQL & "NULL,"
    end if

    if not isblank(rsCRP("Part Number")) then
      SQL = SQL & "'" & replace(rsCRP("Part Number"),"'","''") & "', "
    else
      SQL = SQL & "NULL,"
    end if
    
    if not isblank(rsCRP("Family")) then
      SQL = SQL & "'" & replace(rsCRP("Family"),"'","''") & "', "
    else
      SQL = SQL & "NULL,"
    end if
    
    SQL = SQL & "'" & Date() & "', "
    
    SQL = SQL & "'110')"
    
    'response.write SQL & "<P>"
    
    eConn.execute SQL
    
    rsCRP.MoveNext
    
  loop
  
  rsCRP.close
  set rsCRP = Nothing
  response.write "End Import<BR>"
  Call Disconnect_eStoreDatabase
%>