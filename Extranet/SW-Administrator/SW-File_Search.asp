
    <%@LANGUAGE="VBSCRIPT"%>
    <%
    ' --------------------------------------------------------------------------------------
    ' Name: ASP File Search
    '
    ' Inputs:SearchText = The file containing the text you are searching for.
    '     Directory = The Directory (and it's subdirectories) to search (default is c:\)
    '
    ' Returns: A list of files matching the string you are searching for.
    ' --------------------------------------------------------------------------------------

    Response.AddHeader "Pragma", "No-Cache"	'try Not To cache page
    Response.CacheControl = "Private"		'try Not To cache page
  	server.scripttimeout = 300		'script will time out after 5 minutes
    %>

    <HTML>
    <HEAD><TITLE>Search for files containing...</TITLE>
    <style>
    body {font:10pt Arial;background-color:papayawhip;color:antiquewhite;font-weight:bold;margin-top:0px;margin-left:0px;margin-right:0px}
    A:link {color:black;text-decoration:none}
    A:hover {color:red;text-decoration:underline}
    A:visited {color:black;text-decoration:none}
    td {color:black;border-bottom:1pt solid black;font:9pt Arial}
    th {color:black;border-bottom:1pt solid black;font:9pt Arial;font-weight:bold}
    </style>
    </HEAD>
    <BODY>
    <DIV style="background-color:tan">
    <CENTER>
    Search for files containing key words<BR>
    <%
    Dim filecounter, searchtext, directory	'dimention variables
    Dim fcount, fsize
    filecounter = 0				'initialize filecounter To zero
    searchtext = Trim(request("SearchText"))	'get the querystring SearchText
    directory = Trim(request("Directory"))		'get the querystring Directory
    if directory = "" Then directory = "c:\"	'if no directory the Set to c:\
    						'Write the Search Form To the page
    response.write "<FORM action='SW-File_Search.asp' method=get>Search For:" & _
    	" <INPUT type=text name=SearchText size=20 value=" & Chr(34) & searchtext & _
    	Chr(34) & "> Change Directory: <INPUT type=text name=Directory size=20 value=" & _
    	Chr(34) & directory & Chr(34) & "> <INPUT style='background-color:blanchedalmond;" & _
    	"color:chocolate' type=submit value='Start Search'></FORM><BR></DIV>"
    if searchtext <> "" Then	'if there is a file To search For then search
    response.write "<TABLE border=0 width='100%'>"
    response.write "<TR><TH width='60%'>File Name</TH><TH width='10%'>File Size</TH><TH width='30%'>Date Modified</TH></TR>"
    		'create the recordset object To store
    		'the filepath, filename, filesize and last modified Date
    Set rs = createobject("adodb.recordset")
    rs.fields.append "FilePath",200,255
    	 rs.fields.append "FileName",200,255
    rs.fields.append "FileSize",200,255
    	 rs.fields.append "FileDate",7,255
    rs.open
    	Recurse directory	'call the subroutine to traverse the directories
    	Sub Recurse(Path)
    			'create the file system object
    		Dim fso, Root, Files, Folders, File, i, FoldersArray(1000)
    		Set fso = Server.CreateObject("Scripting.FileSystemObject")
    		Set Root = fso.getfolder(Path)
    		Set Files = Root.Files
    		Set Folders = Root.SubFolders
    		fcount = 0			'zero out the file count variable
    			'traverse through the subdirectories In the current directory
    		For Each Folder In Folders
    			FoldersArray(i) = Folder.Path
    			i = i + 1
    		Next
    			'traverse through the files In the current folder or subfolder
    		For Each File In Files
    				'check if the search String is found
    			num = InStr(UCase(File.Name), UCase(searchtext))
    				'if it is Then update the recordset and sort it
    			if num <> 0 Then
    			filecounter = filecounter + 1
    			rs.addnew
    		rs.fields("FilePath") = File.Path
    			rs.fields("FileName") = File.Name
    			rs.fields("FileSize") = File.Size
    			rs.fields("FileDate") = File.DateLastModified
    			rs.update
    		rs.Sort = "FileName ASC"
    			End if
    		Next
    			'recurse through the current directory until 
    			'all subfolders have been traversed
    		For i = 0 To UBound(FoldersArray)
    			if FoldersArray(i) <> "" Then 
    				Recurse FoldersArray(i)				
    			Else
    				Exit For
    			End if
    		Next
    	End Sub
    		'if files were found Then write them To the document
    	if filecounter <> 0 Then
    			filecounter = 0
    		Do While Not rs.eof
    			filecounter = filecounter + 1
    			response.write "<TR><TD width='50%' valign=top><A href=""" & rs.fields("FilePath") & """>" & rs.fields("FileName") & "</TD><TD width='10%' align=right valign=top>"
    					'get the file size so we can
    					'assign the proper Bytes, KB or MB value
    				fsize = CLng(rs.fields("FileSize"))
    				'if less than 1 kilobyte Then it's Bytes
    			if fsize >= 0 And fsize <= 999 Then
    				fnumber = FormatNumber(fsize,0) & " Bytes"
    			End if
    				'if 1 KB but less Then 1 MB then assign KB
    			if fsize >= 1000 And fsize <= 999999 Then
    				fnumber = FormatNumber((fsize / 1000),2) & " KB"
    			End if
    				'if 1 MB or more Then assign MB
    			if fsize >= 1000000 Then
    				fnumber = FormatNumber((fsize / 1000000),2) & " MB"
    			End if
    				'write Each file and corresponding info To the document
    			response.write fnumber & "</TD><TD width='30%' align='center'>" & rs.fields("FileDate") & "</TD></TR>"
    			rs.movenext
    		Loop
    		response.write "</TABLE>"	'end the table
    	Else
    			'no files were found
    	End if
    End if
    %>
    </BODY>
    </HTML>
