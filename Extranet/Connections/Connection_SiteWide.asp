<!-- #include file="servers.asp" -->
<%
Dim conn
Dim strConnectionString_Sitewide
Dim strDatabaseName_Sitewide
Dim strServerName_Sitewide
Dim strLogin_Sitewide
Dim MookieSQL, rsMookie

Server.ScriptTimeOut = 30
strServerName_Sitewide = UCase(Request.ServerVariables("SERVER_NAME"))
strDatabaseName_Sitewide = "FLUKE_SITEWIDE"
strLogin_Sitewide = "WEBUSER"

strConnectionString_Sitewide = GetConnectionString(strDatabaseName_Sitewide, strServerName_Sitewide, strLogin_Sitewide)

function GetConnSiteWide()

	on error resume next
	
	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 30
	localConnect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strConnectionString_SiteWide & ";"

	set GetConnSiteWide = localConnect
	
	if CInt(err.number) > 0 then
'		response.redirect("/default-unavailable.asp")
	end if
  
  on error goto 0
  
end function

'====================================================================================================================

sub Connect_SiteWide()

	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 30
	localConnect.Open strConnectionString_SiteWide

  	set conn = localConnect
  	conn.CommandTimeout = 240
  ' Test Connection, if error send email advisory
  
  MookieSQL = "Select Top 1 ID FROM Site"
  set rsMookie = Server.CreateObject("ADODB.Recordset")
  
  on error resume next
  rsMookie.open MookieSQL, conn, 3, 3
  if err.Number <> 0 then
    response.redirect "/Site_Error_Advisory.asp"
  end if
  
  on error goto 0
  
  rsMookie.close
  set rsMookie  = nothing
  set MookieSQL = nothing

end sub

'====================================================================================================================

sub Disconnect_SiteWide()

  conn.close
	set conn = nothing

end sub

%>