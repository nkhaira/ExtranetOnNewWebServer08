	
<%
	dim yr
	dim month
	dim addto
	dim present
	dim cutdate
	dim qstring
	dim x
	dim limit
	dim modelstr
	dim qedit
	qedit = 0
	x = 1
	
	if request("Returns") <> "0" then
		limit=Cint(request("Returns"))
	else
		limit=10
	end if

	Function changecard(str)
		dim i
		dim tempstr
		for i = 1 to len(str)
			if Mid(str,i,1) = "*" then
				tempstr=tempstr & "%"
				if i <> 1 then
					exit for	
				end if
				if i = len(str)-1 then
					exit for
				end if
			else if Mid(str,i,1) = "'" then
				tempstr=tempstr
			else
				tempstr=tempstr & mid(str,i,1)
			end if
			end if
		next
		changecard = tempstr
	end function

	Function loginDB()
		dim dbConn
		Set dbConn = Server.CreateObject("ADODB.Connection")	
		dbConn.Open "servicedocs"
		set LoginDB = dbConn
	End Function
	
	set db=loginDB()
	Set Session("RS") = Server.CreateObject("ADODB.Recordset")
	qstring = "SELECT * FROM ""Models"""
	
	if request("Model") <> "" then
		modelstr = CStr(request("Model"))
		if InStr(modelstr,"*") then
			modelstr = changecard(modelstr)
			qstring = qstring & " WHERE ""Model"" LIKE '" & modelstr & "'"
		else
			modelstr = changecard(modelstr)
			qstring = qstring & " WHERE ""Model""='" & modelstr & "'"
		end if
		qedit = qedit + 1
	end if
	
	

	Session("RS").open qstring,db,3,1,1
	
	if limit > Session("RS").recordcount then
		numpages=1
	else
		numpages=Session("RS").recordcount \ limit
		numpages=cint(numpages)
		extrapage = Session("RS").recordcount MOD limit
		if extrapage > 0 then
			numpages=numpages + 1
		end if
	end if
	
'	if numpages > 25 then
'		numpages = 25
'	end if
	
	Session("Perpage")=limit
	Session("partpages")=numpages

	response.redirect ("/SW-Common/SvcIndex_Model_Results.asp?whatpage=1&region=" & request("region") & "&view=" & request("view"))
  
	'response.write Cstr(extrapage) & " extra<br>"
	'response.write Cstr(numpages) & " numpages<br>"
	'response.write Cstr(Session("RS").recordcount) & " records<br>"
	'response.write Cstr(limit) & " limit<br>"
	'response.write Cstr(Session("RS").recordcount \ limit) & " calculation<br>"
	'response.write Cstr(Cint(Session("RS").recordcount \ limit)) & " integer<br>"
	%>