<!-- #include file="servers.asp" -->
<%
Dim dbConnSurvey

Sub Connect_SurveyDatabase()
	Dim strConnectionString
	Dim strDatabaseName
	Dim strServerName
	Dim strLogin
	
	strServerName = UCase(Request("SERVER_NAME"))
	strDatabaseName = "FLUKE_SURVEY"
	strLogin = "WEBUSER"
	
	strConnectionString = GetConnectionString(strDatabaseName, strServerName, strLogin)
	
	set dbConnSurvey = Server.CreateObject("ADODB.Connection")
	
	dbConnSurvey.ConnectionTimeOut = 120
	dbConnSurvey.CommandTimeout = 120
	
	dbConnSurvey.Open strConnectionString
End Sub

Sub Disconnect_SurveyDatabase()
	if IsObject(dbConnSurvey) then
		dbConnSurvey.Close
		set dbConnSurvey = nothing
	end if
End Sub
%>