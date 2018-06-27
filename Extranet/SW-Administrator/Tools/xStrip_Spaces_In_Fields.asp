<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' Author: Kelly Whitlock

Session.timeout      = 240 ' Set to 4 Hours
Server.ScriptTimeout = 10 * 60

Call Connect_SiteWide

Dim Field2Strip(1)
Field2Strip(0) = "Revision_Code"
Field2Strip(1) = "Language"

for x = 0 to 1

  SQL = "SELECT ID, " & Field2Strip(x) & " " &_
        "FROM dbo.Calendar " &_
        "WHERE " & Field2Strip(x) & " LIKE '% %' "
  
  Set rsID = Server.CreateObject("ADODB.Recordset")
  rsID.Open SQL, conn, 3, 3
  
  do while not rsID.EOF
  
    StripIt = Trim(rsID(Field2Strip(x)))
    
    SQLU = "UPDATE Calendar SET " & Field2Strip(x) & "='" & StripIt & "' WHERE ID=" & rsID("ID")
    response.write SQLU & "<P>"
    response.flush
    conn.execute SQLU
    
    rsID.MoveNext
  
  loop
  
  response.write "<P>Done"
  
  rsID.close
  set rsID = nothing
  
next

Call Disconnect_SiteWide
%>