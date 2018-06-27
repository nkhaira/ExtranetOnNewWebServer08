<%
' --------------------------------------------------------------------------------------
' Globals
' --------------------------------------------------------------------------------------
Dim g_strCore_ID
Dim g_strSitewideUser
Dim g_iSitewide_ID
Dim g_bCoreExists

' --------------------------------------------------------------------------------------
' Initialize Globals
' --------------------------------------------------------------------------------------
Site_ID = 11	' Implicitly declared and used in /SW-Common/SW-Site_Information.asp
g_strCore_ID = Request("strCore_ID")
g_strSitewideUser = Session("Logon_User")


' --------------------------------------------------------------------------------------
' Validate security. If neither a sitewide account nor a formdata account exists, send back to www
' --------------------------------------------------------------------------------------
Function Validate_Security(g_strSitewideUser, g_strCore_ID, Site_ID, g_iSitewide_ID, g_bCoreExists, g_strEmail, g_strPassword)
	Dim strDevTest

	g_iSitewide_ID = GetLogonID(g_strSitewideUser, Site_ID, g_strEmail, g_strPassword)
	g_bCoreExists = IsDuplicate(g_strCore_ID)

	if g_iSitewide_ID = 0 and g_bCoreExists = false then
  
    SeverName = UCase(request("SERVER_NAME"))
		if InStr(ServerName, "DEV") > 0  or InStr(ServerName, "DTMEVTVSDV15") > 0 or InStr(ServerName, "DTMEVTVSDV18") > 0 then
			strDevTest = "dev."
		elseif InStr(ServerName, "TST") > 0  or InStr(ServerName, "TEST") > 0 then
			strDevTest = "test."
		end if

		response.redirect "http://www." & strDevTest & "fluke.com/calibrators/registration/register.asp?AGID=" & Request("AGID") & "&SID=" & Request("SID") & "&redir=http://www." & strDevTest & "fluke.com/calibrators/software/csoftware.asp" & Server.URLEncode("?action=metcaldownload")

	end if
End Function

' --------------------------------------------------------------------------------------

Function GetLogonID(strLogon, iSite_ID, g_strEmail, g_strPassword)
  Dim iSiteWide_ID
  Dim cmd, prm, rsLogin

  iSiteWide_ID = 0

  if trim(strLogon) <> "" then
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Admin_GetUserID_ByLogin"

	Set prm = cmd.CreateParameter("@strLogin", adVarchar, adParamInput, 50, strLogon & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iSite_ID", adInteger, adParamInput, , iSite_ID & "")
	cmd.Parameters.Append prm

	Set rsLogin = Server.CreateObject("ADODB.Recordset")
	rsLogin.CursorLocation = adUseClient
	rsLogin.CursorType = adOpenDynamic
	rsLogin.open cmd

	if not rsLogin.eof then
		iSiteWide_ID = rsLogin("ID")
		g_strEmail = rsLogin("email")
		g_strPassword = rsLogin("Password")
	end if

	set prm = nothing
	set cmd = nothing

  end if

  GetLogonID = iSiteWide_ID

End Function

'--------------------------------------------------------------------------
Function IsDuplicate(strCore_ID)
' Also found in www\forms
'Purpose: checks for a duplicate ID in the core table

'sID: the ID to be checked
'-------------------------------------------------------------------------

'	On Error Resume Next

	Dim dbCmd
	Dim rs_CoreTest

	IsDuplicate = false

	if strCore_ID <> "" then
		Connect_FormDatabase
	
		set dbCmd = Server.CreateObject("ADODB.Command")
		dbCmd.ActiveConnection = dbConnFormData
		dbCmd.CommandType = adCmdStoredProc
		dbCmd.CommandText = "FormFill_GetUser_Core"
		set tmpParameter = dbCmd.CreateParameter("@CoreID", advarchar, adParamInput, 25, strCore_ID)
		dbCmd.Parameters.Append tmpParameter
		set rs_CoreTest = dbCmd.execute 

		If not rs_CoreTest.EOF Then
			IsDuplicate = True
		End if

		set dbCmd = nothing
		Set rs_CoreTest = Nothing
		Disconnect_FormDatabase
	end if
End Function

%>