<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>500 custom error page</title>
</head>

<body>
<%
Response.Write String(513, "_")

Set objASPError = Server.GetLastError()

intNumber = objASPError.Number
strDesc=objASPError.Description
strSource = Server.HTMLEncode (objASPError.Source)
strFile = objASPError.File
intLine = objASPError.Line
intColumn = objASPError.Column
strCategory = objASPError.Category
strASPCode = objASPError.ASPCode
strASPDesc = objASPError.ASPDescription


strMsg =  "<H2>500 Custom Error: " & strCategory & " error " & intNumber & "</H2>"
strMsg = strMsg & strDesc & "<BR> <B>"
strMsg = strMsg & strFile & "</B>"

If intLine > 0 Then
	strMsg = strMsg & ", line " & intLine
End If
	
'if it's not a syntax error, the column will be -1
	If intColumn > 0 Then
		strMsg = strMsg & ", column " & intColumn
	End if
	
strMsg = strMsg & "<BR>"

' The source is only returned for compilation errors
If strSource <> "" Then
	strMsg = strMsg & "<BR><Font Face ='courier'>" & strSource & "<BR>"
		For intCount = 1 To intColumn
			strMsg = strMsg & "-"
		Next
	strMsg = strMsg & "^</font><BR>"
	
End if

strMsg = strMsg & "<BR>"

'ASP-specific errors provide additional information in the 
'ASPCode and ASPDescription properties
If strASPCode <> "" Then
strMsg = strMsg & "ASP Error: " & strASPCode & ". " & strASPDesc

End if

Response.Write strMsg
Set objASPError = Nothing

%>

</body>
</html>
