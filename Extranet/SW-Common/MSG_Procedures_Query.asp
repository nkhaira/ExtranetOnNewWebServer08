<!--#include virtual="/connections/adovbs.inc"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

Dim strQuery
Dim strActions

strActions = "CORE,WARRANTED,USER"
strQuery = uCase(request("action") & "")

if strQuery <> "" and InStr(strActions, strQuery) then
	Dim cmd, prm
	Dim rsProcedures


	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc

	cmd.CommandText = "Metcal_GetProcedures_WebQuery"
	Set prm = cmd.CreateParameter("@strAction", adVarChar, adParamInput, 9, strQuery & "")
	cmd.Parameters.Append prm

	Set rsProcedures = Server.CreateObject("ADODB.Recordset")
	rsProcedures.CursorLocation = adUseClient
	rsProcedures.CursorType = adOpenStatic
	rsProcedures.open cmd
	set prm = nothing
	set cmd = nothing

	if not rsProcedures.EOF then
	  response.write("<table border=0><tr>")
	  for each foo in rsProcedures.Fields
		if uCase(foo.name) <> "PROCEDURE_ID" then
			response.write("<td><b>" & foo.name & "</b></td>")
		end if
	  next
	  response.write("</tr>")

	  do while not rsProcedures.EOF
		response.write("<tr>")
		for each foo in rsProcedures.Fields
			if uCase(foo.name) <> "PROCEDURE_ID" then
				response.write("<td>" & rsProcedures(foo.name) & "</td>")
			end if
		next
		response.write("</tr>")
		rsProcedures.movenext
	  loop
	  response.write("</tr></table>")
	end if
end if
%>