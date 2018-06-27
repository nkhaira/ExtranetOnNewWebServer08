
<!-- #include virtual="/connections/servers.asp" -->
<%
dim conn
dim strConnectionString_Sitewide
Dim strDatabaseName_Sitewide
Dim strServerName_Sitewide
Dim strLogin_Sitewide

Server.ScriptTimeOut = 30
strServerName_Sitewide = UCase(Request.ServerVariables("SERVER_NAME"))
strDatabaseName_Sitewide = "FLUKE_SITEWIDE"
strLogin_Sitewide = "WEBUSER"

strConnectionString_Sitewide = GetConnectionString(strDatabaseName_Sitewide, strServerName_Sitewide, strLogin_Sitewide)

Function GetConnSiteWide()

	on error resume next
	
	Dim localConnect

	set localConnect = Server.CreateObject("ADODB.Connection")
	localConnect.ConnectionTimeout = 30 * 10
	localConnect.Open strConnectionString_SiteWide
	set GetConnSiteWide = localConnect
	
	if CInt(err.number) > 0 then
'		response.redirect("/default-unavailable.asp")
	end if
End Function

'====================================================================================================================

Sub Connect_SiteWide()

  set conn = GetConnSiteWide()
  
  ' Test Connection, if error send email advisory
  
  MookieSQL = "Select Top 1 ID FROM Site"
  Set rsMookie = Server.CreateObject("ADODB.Recordset")
  on error resume next
  rsMookie.Open MookieSQL, conn, 3, 3
  if err.Number <> 0 then
    response.redirect "/Site_Error_Advisory.asp"
  end if
  rsMookie.close
  set rsMookie  = nothing
  set MookieSQL = nothing

End Sub

'====================================================================================================================

Sub Disconnect_SiteWide()

  conn.Close
	set Conn = nothing

End Sub

%>
<%

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 2 * 60

Connect_SiteWide

SQL = "SELECT DISTINCT Account_ID " &_
      "FROM         dbo.Activity " &_
      "WHERE     (Region = 0) AND (Account_ID > 1) " &_
      "ORDER BY Account_ID"
      
Set rsA = Server.CreateObject("ADODB.Recordset")
rsA.Open SQL, conn, 3, 3
  
do while not rsA.EOF
  
  SQL = "Select Region from UserData where id=" & rsA("Account_ID")
  Set rsB = Server.CreateObject("ADODB.Recordset")
  rsB.Open SQL, conn, 3, 3
  
  if not rsB.EOF then
  
    SQL = "Update Activity set Region=" & rsB("Region") & " WHERE Account_ID=" & rsA("Account_ID")
response.write rsA("Account_ID") & "<BR>"
response.flush
    conn.execute SQL
    
  end if
  
  rsB.close
  set rsB = Nothing
  
  rsA.MoveNext
  
loop

rsA.close
set rsA = Nothing

response.write "Done"


Disconnect_SiteWide


%>