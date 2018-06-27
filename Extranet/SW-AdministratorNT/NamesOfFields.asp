<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

if LCase(Session("LOGON_USER")) <> "whitlock" then
  response.end
  
else

  Call Connect_SiteWide
  
  Table_Name = "UserData"
  i = 0
  
  SQL = "SELECT TOP 1 * FROM " & Table_Name
  Set rsFields = Server.CreateObject("ADODB.Recordset")
  rsFields.Open SQL, conn, 3, 3
  
  do while i < rsFields.Fields.Count
    Set Field_Name = rsFields.Fields(i)
    response.write Field_Name.Name & " (" & Field_Name.Type & ") = " & Field_Name.Value & "<BR>"
    
    SQL = "INSERT INTO Field_Names (Table_Name,Field_Name,Field_Type)"
    SQL = SQL & " VALUES ('" & Table_Name & "','" & Field_Name.Name & "'," & CInt(Field_Name.Type) & ")"
    
  '  response.write SQL & "<BR>"
    conn.execute(SQL)
    
    i = i + 1
  loop
  
  rsFields.Close
  set rsFields = nothing
  
  Call Disconnect_SiteWide
end if  
%>
