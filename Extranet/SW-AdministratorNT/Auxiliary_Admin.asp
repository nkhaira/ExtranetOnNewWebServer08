<%@ Language="VBScript" CODEPAGE="65001" %>
<%
'
' Author: D. Whitlock
'

' -------------------------------------------------------------------------------------- 
' Functions
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
  
' --------------------------------------------------------------------------------------

' Update Record

SQL = "UPDATE Auxiliary SET "

SQL = SQL & "Auxiliary.Description="& "'" & replacequote(request("Description")) & "'"

SQL = SQL & ",Auxiliary.Input_Method=" & CInt(request("Input_Method"))

if isblank(request("Radio_Text")) or CInt(request("Input_Method")) = 0 then
  SQL = SQL & ",Auxiliary.Radio_Text='Yes, No'"
else
  SQL = SQL & ",Auxiliary.Radio_Text='" & replacequote(request("Radio_Text")) & "'"
end if  

if request("Enabled") = "on" then           
  SQL = SQL & ",Auxiliary.Enabled=" & CInt(True)
else
  SQL = SQL & ",Auxiliary.Enabled=" & CInt(False)
end if    
  
if request("Required") = "on" then           
  SQL = SQL & ",Auxiliary.Required=" & CInt(True)
else
  SQL = SQL & ",Auxiliary.Required=" & CInt(False)
end if    

if request("User_Edit") = "on" then           
  SQL = SQL & ",Auxiliary.User_Edit=" & CInt(True)
else
  SQL = SQL & ",Auxiliary.User_Edit=" & CInt(False)
end if    

if request("Registration") = "on" then           
  SQL = SQL & ",Auxiliary.Registration=" & CInt(True)
else
  SQL = SQL & ",Auxiliary.Registration=" & CInt(False)
end if    


SQL = SQL & " WHERE (((Auxiliary.ID)=" & CInt(request("ID")) & "))"

Call Connect_SiteWide
conn.execute (SQL)
Call Disconnect_SiteWide
      
' Success, Go back and re-display record with updated data

BackURL = "default.asp?ID=edit_aux&Site_ID=" & CInt(request("Site_ID")) & "&Aux_ID=" & CInt(request("ID"))
  
response.redirect BackURL
                   
%>
