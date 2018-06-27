<%

'Dim dbConnSiteWide
Dim conn
Dim strConnectionString_SiteWide
Server.ScriptTimeOut = 1500

'====================================================================================================================

Dim strServerNameNavigation

strServerNameNavigation = UCase(Request("SERVER_NAME"))

if InStr(1, strServerNameNavigation, "MILHOUSE") or strServerNameNavigation = "216.244.65.42" then
	strConnectionString_SiteWide = "DRIVER={SQL Server}; SERVER=216.244.65.42,1433;UID=sa;DATABASE=fluke_sitewide;pwd="
else
	strConnectionString_SiteWide = "DRIVER={SQL Server}; SERVER=216.244.76.70; UID=sitewide_usr;DATABASE=fluke_sitewide;pwd=tuggy_boy"
end if

Function GetConnSiteWide()

	on error resume next
	
	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 30
	localConnect.Open strConnectionString_SiteWide
	set GetConnSiteWide = localConnect
	
'response.write("err: " & err.description & "<BR>")
'response.end
	if cInt(err.number) > 0 then
'		response.redirect("/default-unavailable.asp")
	end if
End Function

Sub Connect_SiteWide()
  set conn = GetConnSiteWide()
'	set dbConnSiteWide = GetConnSiteWide()
End Sub

'====================================================================================================================

Sub Disconnect_SiteWide()
  conn.Close
'	dbConnSiteWide.Close
'	set dbConnSiteWide = nothing
	set Conn = nothing  
End Sub

%>