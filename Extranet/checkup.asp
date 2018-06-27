<!-- #include virtual="/adminTools/Connections/adovbs.inc" -->
<!-- #include virtual="/AdminTools/Connections/Connection_sql.asp" -->


<%
response.write "<HTML><HEAD></head><BODY><TABLE>"&vbcrlf

on error resume next
CheckDBconnections
CheckEmail
Checktranslate
response.write "<TR><TD>I'm here</td></tr>"
response.write "</TABLE></body></html>"&vbcrlf


'** START SUBROUTINES ******************************************

sub CheckEmail
	on error resume next
	response.write "<tr><td colspan=3><HR></td><td>"
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	
	if isObject(Mailer) then
		response.write "<tr><td>ASPQueMAIL</td><td>Version " & Mailer.Version &"</td>"
		response.write "<td><font color=""blue"">GOOD</font></td></TR>"&vbcrlf
	else
		err.clear
		response.write "<tr><td>ASPQueMAIL</td><td>&nbsp;</td>"
		response.write "<td><font color=""Red"">BAD</font></td></TR>"&vbcrlf
	end if
	Set Mailer = nothing
end Sub


'********************************************************
sub Checktranslate
	on error resume next
	response.write "<tr><td colspan=3><HR></td><td>"
	response.write "<tr><td>TsiAToW Convert</td>"
	Set Translate = Server.CreateObject("TsiAToW.Convert")
	
	if isObject(Translate) then
		response.write "<td></td><td><font color=""blue"">GOOD</font></td></TR>" & vbcrlf
	else
		response.write "<td></td><td><font color=""Red"">" & Err.Description
		response.write "</font></td></TR>" & vbcrlf
		err.clear
	end if
	Set Translate = nothing
end sub


'********************************************************
sub CheckDBconnections
	Dim strListDB
	Dim arrListDB
	Dim arrThisDB
	Dim strDB
	Dim strServer
	Dim strUser
	Dim strUserDisplay
	
	strListDB = "WEBUSER|FLUKE_SURVEY:" &_
				"WEBUSER|FLUKE_PRODUCTS:" &_
				"ADMIN|FLUKE_PRODUCTS:" &_
				"ADMIN|FLUKE_CALNEWSLETTER:" &_
				"WEBUSER|FLUKE_CALNEWSLETTER:" &_
				"WEBUSER|METERMAN:" &_
				"WEBUSER|FLUKE_VIRTUALDIRECTORIES:" &_
				"WEBUSER|FLUKE_SITEWIDE:" &_
				"WEBUSER|FLUKE_WHERETOBUY:" &_
				"WEBUSER|WTB:" &_
				"WEBUSER|FLUKE_TMSTORE:" &_
				"WEBUSER|FLUKE_PROMO:" &_
				"WEBUSER|FLUKE_BUYING:" &_
				"WEBUSER|FLUKE_FORMDATA"
	
	strServer = UCase(Request("SERVER_NAME"))
	arrListDB = split(strListDB, ":") 
	on error resume next
	
	for n=0 to Ubound(arrListDB)
		arrThisDB = split(arrListDB(n), "|") 
		strDB = arrThisDB(1) 
		strUser = arrThisDB(0) 
		if strUser = "WEBUSER" then
			strUserDisplay = "read only"
		else
			strUserDisplay = "read / write"
		end if
		response.write "<TR><TD>" & strDB & "</td><td>" & strUserDisplay & "</td>"
		Set objConn = DB_connect(strDB,strUser,strServer)
		
		if err.number <> 0 then
			response.write "<td><font color=""Red"">" & err.description &_
				 " </font></td></TR>" & vbcrlf
			err.clear
		elseif objConn.state = adStateOpen then 
			response.write "<td><font color=""blue"">GOOD</font></td></TR>" & vbcrlf
			Db_Disconnect(objConn)
		else
			response.write "<td><font color=""Red"">BAD</font></td></TR>" & vbcrlf
		end if
		
	next
end sub
	

  %>

