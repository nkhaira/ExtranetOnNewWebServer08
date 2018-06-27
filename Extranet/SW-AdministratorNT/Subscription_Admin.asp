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

SQL = "UPDATE Site SET Subscription_"

Subscription_Text   = request.form("Subscription_Text")
Subscription_Option = request.form("Subscription_Option_ID")

if not isblank(Subscription_Text) then
  select case Subscription_Option
    case 1
      SQL = SQL & "Subject='" & Subscription_Text & "' "
    case 2
      SQL = SQL & "Header='" & Subscription_Text & "' "
    case 3
      SQL = SQL & "Footer='" & Subscription_Text & "' " 
  end select
else
  select case Subscription_Option
    case 1
      SQL = SQL & "Subject=NULL" & " "
    case 2
      SQL = SQL & "Header=NULL" & " "
    case 3
      SQL = SQL & "Footer=NULL" & " "
  end select
end if  
  
SQL = SQL & "WHERE Site.ID=" & CInt(request.form("Site_ID"))

conn.execute (SQL)

Call Disconnect_SiteWide
      
' Success, Go back and re-display record with updated data

BackURL = "default.asp?ID=edit_subscription&Site_ID=" & request.form("Site_ID") & "&Subscription_Option_ID=" & request.form("Subscription_Option_ID")
  
response.redirect BackURL
%>
