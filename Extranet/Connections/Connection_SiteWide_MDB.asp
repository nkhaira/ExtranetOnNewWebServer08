<%

' Access 97 Database ADODB Connection
' Used by:  /SW-Administration
'           /Calibration-Sales
'           /FNet-Sales
'           /FInd-Sales
'           /PCal-Sales
'           /Met-Support-Gold
'           /Sandbox
'           /Register

Dim conn

'If IsObject(Session("SiteWide_conn")) Then
'    Set conn = Session("SiteWide_conn")
'Else
'		Set conn = Server.CreateObject("ADODB.Connection")
'    conn.open "SiteWide","",""
'    Set Session("SiteWide_conn") = conn
'End If

Sub Connect_SiteWide()
  if IsObject(Session("SiteWide_conn")) Then
      Set conn = Session("SiteWide_conn")
  else
  		Set conn = Server.CreateObject("ADODB.Connection")
      conn.ConnectionTimeOut = 120
	    conn.CommandTimeout = 120      
      conn.open "SiteWide","",""
      Set Session("SiteWide_conn") = conn
  end If
End Sub

Sub Disconnect_SiteWide()
	if not IsObject(conn) then
		conn.Close
		set conn = nothing
	end if
End Sub

%>
