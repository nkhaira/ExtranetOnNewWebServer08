<%

dim qstring
dim filestring
dim filename
dim fstring
dim st
dim currentst
dim strUser
dim path
dim datetime
dim iStartPos
dim firstname
dim lastname
dim country
dim company
dim ID

Function loginDB()
	dim dbConn
	
	Set dbConn = Server.CreateObject("ADODB.Connection")	
	dbConn.Open "ServiceUsers"
	set LoginDB = dbConn
End Function
fstring=request("file")
filestring=Cstr(fstring)

st = 1

if DateDiff("s",session("last"),now()) < 10 then
	response.redirect filestring
else
	set db=loginDB()
	strUser = Request.ServerVariables("LOGON_USER")

	iStartPos = InStrRev(strUser, "\") + 1
	if iStartPos <> 0 then
		'Machine name is in login.
		strUser = Mid(strUser, iStartPos)
	end if

	qstring="SELECT FirstName, LastName, Company, Country, Groups FROM UserData WHERE NTLogin='" & strUser & "'"

'response.write("NTLogin: " & strUser & "<BR>")
'response.write("Request login: " & request.servervariables("LOGON_USER") & "<BR>")
'response.end

	set qobject = db.execute(qstring)
	groupstring = qobject("Groups")
	if instr(groupstring,"Admin") > 0 then
		firstname=qobject("FirstName")
		lastname=qobject("LastName")
		country=qobject("Country")
		company=qobject("Company")
		qobject.close
	
		'Parse login here, remove machine name
		iStartPos = InStr(strUser, "/") + 1
		if iStartPos <> 0 then
			'Machine name is in login.
			strUser = Mid(strUser, iStartPos)
		end if
		while (st<>0)
			currentst=st
			st = InStr(st,filestring,"/")
			if st <> 0 then st = st + 1
		Wend
		filename=mid(filestring,currentst,len(filestring))
		path=mid(filestring,1,currentst-1)
		datetime=now()
		qstring="SELECT TrackNumber, UserID, FileName FROM DownloadData WHERE UserID='" & strUser & "' AND FileName='" & filename & "' AND Filepath='" & path & "'"
		set qobject=db.execute(qstring)
		if qobject.EOF then
			qstring = "INSERT INTO DownloadData (UserID, FileName, FirstName, LastName, DownloadDateTime, FilePath, Company, Country, Transactioncode) VALUES('" & strUser & "','" & filename & "','" & firstname & "','" & lastname & "','" & datetime & "','" & path & "','" & Company & "','" & Country & "','D')"
		else
			ID=qobject("TrackNumber")
			'qstring = "UPDATE DownloadData SET Transactioncode = 'D', FilePath= '" & path & "', DownloadDateTime = '" & datetime  & "' WHERE TrackNumber=" & Cint(ID)
			qstring = "UPDATE DownloadData SET Transactioncode = 'D', FilePath= '" & path & "', DownloadDateTime = '" & datetime  & "' WHERE TrackNumber=" & ID
		end if
		qobject.close
		db.execute(qstring)
	end if
	set db=nothing
	session("last")=now()
	response.redirect filestring
end if



%>