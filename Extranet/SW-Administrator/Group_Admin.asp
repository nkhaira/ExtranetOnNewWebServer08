<%@ Language="VBScript" CODEPAGE="65001" %>

<%
'
' Author: D. Whitlock
'
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<% 

' Update Record

Call Connect_SiteWide

SQL = "UPDATE SubGroups SET "

SQL = SQL & "SubGroups.X_Description="& "'" & replacequote(request("X_Description")) & "'"

if request("Enabled") = "on" then           
  SQL = SQL & ",SubGroups.Enabled=" & CInt(True)
else
  SQL = SQL & ",SubGroups.Enabled=" & CInt(False)
end if    
if request("Default_Select") = "on" then           
  SQL = SQL & ",SubGroups.Default_Select=" & CInt(True)
else
  SQL = SQL & ",SubGroups.Default_Select=" & CInt(False)
end if    
  
SQL = SQL & " WHERE SubGroups.ID=" & CInt(request("ID"))

conn.execute (SQL)

Call Disconnect_SiteWide
      
' Success, Go back and re-display record with updated data

BackURL = "default.asp?ID=edit_group&Site_ID=" & CInt(request("Site_ID")) & "&Group_ID=" & CInt(request("ID"))
  
response.redirect BackURL
                   
%>
