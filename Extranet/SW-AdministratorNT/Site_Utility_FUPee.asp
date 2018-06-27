<%@Language="VBScript" Codepage=65001%>
<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->
<%

' --------------------------------------------------------------------------------------

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

oFileUpEE.ProgressIndicator(saClient)    = false
oFileUpEE.ProgressIndicator(saWebServer) = false

on error resume next
oFileUpEE.ProcessRequest Request, False, False
if err.Number <> 0 then
	response.write "<B>WebServer ProcessRequest Error</B><BR>" & Err.Description & " (" & Err.Source & ")"
	response.end
end if
on error goto 0

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
' Find Base Directory Root
' --------------------------------------------------------------------------------------

Dim unlock

if Session("SWFUP_Path") = "" then
  Session("SWFUP_Path") = Server.mappath(request.Servervariables("SCRIPT_NAME"))
  FUP_Path = split(Session("SWFUP_Path"),"\")
  FUP_Node = UBound(FUP_Path)
  for x = 0 to FUP_Node
    if LCase(FUP_Path(x)) <> "extranet" then
      FUP_Root = FUP_Root & FUP_Path(x) & "\"
    elseif LCase(FUP_Path(x)) = "extranet" then
      FUP_Root = FUP_Root & FUP_Path(x)
      exit for
    end if
  next    
  Session("FUP_Root") = FUP_Root
end if  

' --------------------------------------------------------------------------------------
' File Upload
' --------------------------------------------------------------------------------------

if request.querystring("upload") = "@" then

  Dim Uploader, File
  Set Uploader = New FileUploader

  ' This starts the upload process

  with response
    .write "<HTML>" & vbCrLf
    .write "<TITLE>SiteWide File Utility Program</TITLE>" & vbCrLf & vbCrLf
    .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">"
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
    .write "    <TD><FONT FACE=Arial SIZE=-1>File Upload Information:</FONT></TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    
    .write "  <TR>" & vbCrLf
    .write "    <TD BGCOLOR=BLACK>" & vbCrLf
    
            oFileUpEE.Files("File_Name").SaveAs Path_Destination
        oFileUpEE.Files.Remove("File_Name")
       
    ' Check if any files were uploaded

    if oFileUpEE.Files.Count = 0 then
    	.write "No file(s) were uploaded."
    else
      .write "<TABLE>" & vbCrLf
  
    	' Loop through the uploaded files
      	
      for each oFormItem in oFileUpEE.File
        oFileUpEE.Files(oFormItem.Name).ClientFileName.Save request.querystring("txtpath")      
     		.write "<TR><TD>&nbsp;</TD></TR>"
        .write "<TR><TD><FONT color=gray>File Uploaded: </FONT></TD><TD>" & oFileUpEE.Files(oFormItem.Name).ClientFileName & "</TD></TR>"
    		.write "<TR><TD><FONT color=gray>Size: </FONT></TD><TD>" & Int(oFileUpEE.Files(oFormItem.Name).Size / 1024) + 1 & " kb</TD></TR>"
     		.write "<TR><TD><FONT color=gray>Type: </FONT></TD><TD>" & oFileUpEE.Files(oFormItem.Name).ContentType & "</TD></TR>"
      next
      .write "<TR><TD>&nbsp;</TD></TR>"
      .write "</TABLE>"
    end if
  
    .write "    </TD>"
    .write "  </TR>"
    .write "</TABLE>"
    .write "  <BR>"

    .write "<A HREF=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & request.querystring("txtpath") & """><FONT FACE=""Webdings"" TITLE="" Back "" SIZE=+2 >7</FONT></A>" & vbCrLf
    .write "  </DIV>"
    .write "</BODY>"
    .write "</HTML>"
    
  end with  

  response.end
  
end if

' --------------------------------------------------------------------------------------
' Logon
' --------------------------------------------------------------------------------------

on error resume next

response.Buffer = true

password1 = "password"
password2 = "!SiteWide"

' Production root:  D:\IIS\websites\Extranet\

if request.querystring("logoff") = "@" then
	session("SWFUPassword")        = "" ' Logged off
	session("dbcon")               = ""	' Database Connection
	session("txtpath")             = ""	' any pathinfo
end if

if (session("SWFUPassword") <> password1 and session("SWFUPassword") <> password2) and request.form("code") = "" then

  with response

    .write "<HTML>" & vbCrLf
    .write "<TITLE>SiteWide File Utility Program Login</TITLE>" & vbCrLf
    .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">"
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
    .write "<FONT FACE=Arial SIZE=-2 COLOR=#FF8300>Administrator's Logon</FONT>" & vbCrLf
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
    .write "            <INPUT NAME=SUBMIT TYPE=SUBMIT VALUE=""Logon"">" & vbCrLf
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
    
    if oFileUpEE.form("logoff") = "@" then
      .write "<FONT COLOR=GRAY SIZE=-2 FACE=ARIAL>"
      .write "<SPAN STYLE='CURSOR: HAND;' ONCLICK=WINDOW.CLOSE(THIS);>Close Window"
      .write "</FONT>"
    end if
  
    .write "</DIV>"

		.end
    
  end with

end if

' --------------------------------------------------------------------------------------
' Logon Verify
' --------------------------------------------------------------------------------------

if request.form("code") = password1 or session("SWFUPassword") = password1 then
  session("SWFUPassword") = password1
  unlock = false
elseif request.form("code") = password2 or session("SWFUPassword") = password2 then
  session("SWFUPassword") = password2
  unlock = true
else
  response.write "<BR><B><P align=center><FONT color=white ><b>Access Denied</B></FONT></p>"
	response.end
end if

server.scriptTimeout = 180
set fso    = Server.CreateObject("Scripting.FileSystemObject")
mapPath    = Server.mappath(request.Servervariables("SCRIPT_NAME"))
mapPathLen = len(mapPath)

if session(myScriptName) = "" then
	for x = mapPathLen to 0 step -1
	  myScriptName = mid(mapPath, x)
  	if instr(1, myScriptName,"\") > 0  then
	  	myScriptName = mid(mapPath, x + 1)
		  x = 0
  		session(myScriptName) = myScriptName
	  end if
	next
else
	myScriptName = session(myScriptName)
end if

wwwRoot = left(mapPath, mapPathLen - len(myScriptName))
' Temp Directory to which files will be DUMPED To and From
Target = wwwRoot & "..\SW-FileUp_Temp\"

if len(request.querystring("txtpath")) = 3 then
  pathname = left(request.querystring("txtpath"),2) & "\" & request.form("Fname")
else
  pathname = request.querystring("txtpath") & "\" & request.form("Fname")
end if

if oFileUpEE.Form("txtpath") = "" then
	MyPath = request.querystring("txtpath")
else
	MyPath = oFileUpEE.Form("txtpath")
end if

' --------------------------------------------------------------------------------------
' Path Correction Routine
' --------------------------------------------------------------------------------------

if len(MyPath) = 1  then MyPath = Session("FUP_Root") & "\"
if len(MyPath) = 2  then MyPath = Session("FUP_Root") & "\"
if MyPath      = "" then MyPath = Session("FUP_Root") & "\"

if not fso.FolderExists(MyPath) then
	response.write "<FONT FACE=Arial size=+2 color=""Black"">Non-existing path specified.<BR>Please use browser back button to continue !"
	response.end
end if

set folder = fso.GetFolder(MyPath)
if fso.GetFolder(Target) = false then
	response.write "<FONT FACE=Arial size=-2 color=Black>Temporary Target Directory does not exist. </FONT><FONT FACE=Arial size=-1 color=White>" & Target & "<BR></FONT>"
else
	set fileCopy = fso.GetFolder(Target)
end if
if Not(folder.IsRootFolder) then
  if len(folder.ParentFolder) > 3 then
    if err.number = 76 then
        showPath="\"
        Rootfolder=true
        err.Clear
    else
        showPath = folder.ParentFolder & "\" & folder.name
    end if
  else
	showPath = folder.ParentFolder & folder.name
  end if
else
        Rootfolder=true
  showPath = left(MyPath,2)
end if
'response.Write Session("FUP_Root")
if showPath = "\" then showPath = Session("FUP_Root")
MyPath   = showPath
showPath = MyPath & "\"

' --------------------------------------------------------------------------------------
' File Download
' --------------------------------------------------------------------------------------

set drv = fso.GetDrive(left(MyPath, 2))

if request.Form("cmd") = "Download" then

  if request.Form("Fname") <> "" then
  
    response.Buffer = True
  	response.Clear
    
	  strFileName = request.querystring("txtpath") & "\" & request.Form("Fname")
    
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
  	err.description = "No files have been selected for download."
  end if
  
end if

' --------------------------------------------------------------------------------------
' Directory / File Commands
' --------------------------------------------------------------------------------------

with response

  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf
  .write "<TITLE>" & MyPath & "</TITLE>" & vbCrLf
  .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">" & vbCrLf
  .write "<STYLE>" & vbCrLf
  .write "<!--" & vbCrLf
  .write "A:link {font-style: text-decoration: none; color: #c8c8c8}" & vbCrLf
  .write "A:visited {font-style: text-decoration: none; color: #777777}" & vbCrLf
  .write "A:active {font-style: text-decoration: none; color: #ff8300}" & vbCrLf
  .write "A:hover {font-style: text-decoration: cursor: hand; color: #ff8300}" & vbCrLf
  .write "*		{scrollbar-base-color:#777777;"
  .write "scrollbar-track-color:#777777;scrollbar-darkshadow-color:#777777;scrollbar-face-color:#505050;"
  .write "scrollbar-arrow-color:#ff8300;scrollbar-shadow-color:#303030;scrollbar-highlight-color:#303030;}" & vbCrLf
  .write "input,select,table {font-family:verdana,arial;font-size:11px;text-decoration:none;border:1px solid #000000;}" & vbCrLf
  .write "//-->" & vbCrLf
  .write "</STYLE>" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY BGCOLOR=BLACK TEXT=WHITE TOPAPRGIN=""0"">" & vbCrLf

  .flush

end with

' --------------------------------------------------------------------------------------
' Command Tree and Execute
' --------------------------------------------------------------------------------------
  
select case request.form("cmd")
	case ""                 ' No Command
  
		if request.form("dirStuff") <> "" then
			response.write "<FONT FACE=Arial size=-2>You need to click [Create] or [Delete] for folder operations to be</FONT>"
		else
			response.write "<FONT face=webdings size=+3 color=#ff8300>Â</FONT>"
		end if
    
	case "   Copy   "     	' Copy FROM Folder
  
		if request.Form("Fname") = "" then
  		response.write "<FONT FACE=Arial size=-2 color=#ff8300>Copying: " & request.querystring("txtpath") & "\???</FONT><BR>"
			err.number = 424
		else
			response.write "<FONT FACE=Arial size=-2 color=#ff8300>Copying: " & request.querystring("txtpath") & "\" & request.Form("Fname") & "</FONT><BR>"
			fso.CopyFile request.querystring("txtpath") & "\" & request.Form("Fname"),Target & request.Form("Fname")
			response.Flush
		end if

	case "  Copy "          ' Copy TO Folder
  
		if request.Form("ToCopy") <> "" and request.Form("ToCopy") <> "------------------------------" then
			response.write "<FONT FACE=Arial size=-2 color=#ff8300>Copying: " & oFileUpEE.Form("txtpath") & "\" & request.Form("ToCopy") & "</FONT><BR>"
			response.Flush
			fso.CopyFile Target & request.Form("ToCopy"), oFileUpEE.Form("txtpath") & "\" & request.Form("ToCopy")
		else
		  response.write "<FONT FACE=Arial size = -2 color=#ff8300>Copying: " & oFileUpEE.Form("txtpath") & "\???</FONT><BR>"
			err.number = 424
		end if
    
	case "Delete"
  
		if request.form("dirStuff") <> "" then  ' Delete Folder
      CommandToDo = "Deleting folder..."
	    fso.DeleteFolder MyPath & "\" & request.form("DirName")
 		elseif request.Form("Fname") <> "" then
      CommandToDo = "Deleting: " & oFileUpEE.Form("txtpath") & "\" & request.Form("Fname")
			response.flush
 			fso.DeleteFile oFileUpEE.Form("txtpath") & "\" & request.Form("Fname")
    else
 			CommandToDo = "Deleting: " & oFileUpEE.Form("txtpath") & "\???"
 			err.number = 424
 		end if
    
  case "Edit/Create"
  
    with response

      .write "<CENTER>" & vbCrLf
      .write "<BR>" & vbCrLf
      .write "<TABLE BGCOLOR=""#505050"" CELLPADDING=""8"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD BGCOLOR=""#000000"" VALIGN=""bottom"">" & vbCrLf
    	.write "      <FONT FACE=Arial SIZE=-2 COLOR=#FF8300>NOTE: The following edit box may not display special characters from files. Therefore the contents displayed may not be considered correct or accurate.</FONT>" & vbCrLf
    	.write "    </TD>" & vbCrLf
      .write "  </TR>" & vbCrLf
      .write "<TR>" & vbCrLf
      .write "  <TD><TT>Path=> " & pathname & "<BR><BR>" & vbCrLf
      
    end with
      
  	' Fetch file information 

    Set f = fso.GetFile(pathname) 

    with response

      .write "File Type: " & f.Type & "<BR>"
      .write "File Size: " & FormatNumber(f.size,0) & " bytes<BR>"
      .write "File Created: " & FormatDateTime(f.datecreated,1) & "&nbsp;" & FormatDateTime(f.datecreated,3) & "<BR>"
      .write "Last Modified: " & FormatDateTime(f.datelastmodified,1) & "&nbsp;" & FormatDateTime(f.datelastmodified,3) & "<BR>"
      .write "Last Accessed: " & FormatDateTime(f.datelastaccessed,1) & "&nbsp;" & FormatDateTime(f.datelastaccessed,3) & "<BR>"
      .write "File Attributes: " & f.attributes & "<BR>"

    	Set f = Nothing
  	  .write "<center><FORM action=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & MyPath & """ METHOD=""POST"">"

      ' Read the file
    
	  	Set f = fso.OpenTextFile(pathname)
  		if not f.AtEndOfStream then fstr = f.readall
  		f.Close
  		Set f   = Nothing
  		Set fso = Nothing
    
		  .write "<TABLE>"
      .write "  <TR>"
      .write "    <TD>" & VBCRLF
		  .write "      <FONT TITLE=""Use this text area to view or change the contents of this document. Click [Save As] to store the updated contents to the web server."" FACE=arial SIZE=1 ><B>Document Comments</B></FONT><BR>" & VBCRLF
		  .write "      <TEXTAREA NAME=FILEDATA ROWS=16 COLS=85 WRAP=OFF>" & Server.HTMLEncode(fstr) & "</TEXTAREA>" & VBCRLF
		  .write "    </TD>"
      .write "  </TR>"
      .write "</TABLE>" & VBCRLF

      .write "      <BR>"
      .write "      <CENTER>"
      .write "      <TT>LOCATION <INPUT TYPE=""TEXT"" SIZE=48 MAXLENGTH=255 NAME=""PATHNAME"" VALUE=""" & pathname & """>"
      .write "      <INPUT TYPE=""SUBMIT"" NAME=CMD VALUE=""Save As"" TITLE=""This write to the file specifed and overwrite it without warning."">"
      .write "      <INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""Cancel"" TITLE=""If you recieve an error while saving, then most likely you do not have write access OR the file attributes are set to ReadOnly."">"
      .write "      </FORM>"
      .write "    </TD>"
      .write "  </TR>"
      .write "</TABLE>"
      .write "<BR>"

      .end
      
    end with

	case "Create"
  
		CommandToDo = "Creating folder..."
		fso.CreateFolder MyPath & "\" & request.form("DirName")

	case "Save As"

		CommandToDo = "Saving file..."
		Set f = fso.CreateTextFile(request.Form("pathname"))
		f.write request.Form("FILEDATA")
		f.close

end select

' --------------------------------------------------------------------------------------
' Drive Information
  ' --------------------------------------------------------------------------------------

if oFileUpEE.form("getDRVs")="@" then

  with response

    .write "<P>" & vbCrLf
    .write "<CENTER>" & vbCrLf & vbCrLf
    .write "<TABLE BGCOLOR=""#505050"" CELLPADDING=4>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD>" & vbCrLf
    .write "      <FONT FACE=Arial SIZE=-1>Available Drive Information:</FONT>" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD BGCOLOR=BLACK>" & vbCrLf & vbCrLf
    .write "      <TABLE>" & vbCrLf
    .write "        <TR>" & vbCrLf
    .write "          <TD><TT> Drive</TD>" & vbCrLf
    .write "          <TD><TT> Type</TD>" & vbCrLf
    .write "          <TD><TT> Path</TD>" & vbCrLf
    .write "          <TD><TT> ShareName</TD>" & vbCrLf
    .write "          <TD><TT> Size[MB]</TD>" & vbCrLf
    .write "          <TD><TT> ReadyToUse</TD>" & vbCrLf
    .write "          <TD><TT> VolumeLabel</TD><TD>" & vbCrLf
    .write "        </TR>" & vbCrLf
    
    for each thingy in fso.Drives
    
      .write "      <TR>" & vbCrLf
      .write "        <TD><TT> " & thingy.DriveLetter & "</TD>" & vbCrLf
      .write "        <TD><TT> " & thingy.DriveType   & "</TD>" & vbCrLf
      .write "        <TD><TT> " & thingy.Path        & "</TD>" & vbCrLf
      .write "        <TD><TT> " & thingy.ShareName   & "</TD>" & vbCrLf
      .write "        <TD><TT> " & ((thingy.TotalSize)/1024000) & "</TD>" & vbCrLf
      .write "        <TD><TT> " & thingy.IsReady     & " </TD>" & vbCrLf
      .write "        <TD><TT> " & thingy.VolumeName

    next

    .write "          </TD>" & vbCrLf
    .write "        </TR>" & vbCrLf
    .write "      </TABLE>" & vbCrLf & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "<BR>"
    .write "<A HREF=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & MyPath & """><FONT FACE=""webdings"" TITLE="" BACK "" SIZE=+2>7</FONT></A>"
    .write "</CENTER>"
	
    .end
    
  end with  

end if

with response

  .write vbCrLf
  .write "<HEAD>" & vbCrLf
  .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">" & vbCrLf
  .write "<SCRIPT LANGUAGE=""VBScript"">" & vbCrLf
  .write "sub getit(thestuff)" & vbCrLf
  .write "  if right(""" & showPath & """,1) <> ""\"" then" & vbCrLf
  .write "    document.myform.txtpath.value = """ & showPath & """ & ""\"" & thestuff" & vbCrLf
  .write "  else" & vbCrLf
  .write "    document.myform.txtpath.value = """ & showPath & """ & thestuff" & vbCrLf
  .write "  end if" & vbCrLf
  .write "  document.myform.submit()" & vbCrLf
  .write "end sub" & vbCrLf
  .write "</SCRIPT>" & vbCrLf
  .write "</HEAD>" & vbCrLf & vbCrLf

end with

with response

  .write "<B><FONT COLOR=GRAY >SiteWide File Utility Program</FONT></B><FONT COLOR=#FF8300 FACE=WEBDINGS SIZE=6 >!</FONT><B><FONT COLOR=GRAY >Spider</FONT></B>" & vbCrLf

  ' --------------------------------------------------------------------------------------
  ' Report errors
  ' --------------------------------------------------------------------------------------

  if CommandToDo <> "" then
    .write "<P><FONT FACE=Arial size=-1 color=White>Task: " & CommandToDo & "</FONT>" & vbCrLf
    CommandToDo = ""
  end if
  
  select case err.number
	  case "0"
  	  .write "<BR><FONT FACE=Arial size=-1 color=white>Task Successful</FONT>" & vbCrLf

  	case "58"
    	.write "<BR><FONT FACE=Arial size=-1 color=#FFFF99>Error: Folder already exists OR no folder name specified.</FONT>" & vbCrLf

  	case "70"
    	.write "<BR><FONT FACE=Arial size=-1 color=#FFFF99>Error: Permission Denied, folder/file is ReadOnly or contains files.</FONT>" & vbCrLf

  	case "76"
    	.write "<BR><FONT FACE=Arial size=-1 color=#FFFF99>Error: Path not found.</FONT>" & vbCrLf

  	case "424"
    	.write "<BR><FONT FACE=Arial size=-1 color=#FFFF99>Error: Missing, Insufficient data OR file is ReadOnly.</FONT>" & vbCrLf
	
  	case else
    	.write "<BR><FONT FACE=Arial size=-1 color=#FFFF99>Error: " & err.description & "</FONT>" & vbCrLf
  end select

  ' --------------------------------------------------------------------------------------
  ' Directory Information and Logoff
  ' --------------------------------------------------------------------------------------

  .write "<FONT FACE=COURIER>" & vbCrLf

  .write "<FORM METHOD=""POST"" ACTION=""" & request.Servervariables("SCRIPT_NAME") & """ NAME=""myform"">" & vbCrLf

  .write "  <TABLE BGCOLOR=#505050 CELLPADDING=2 WIDTH=""90%"">" & vbCrLf
  .write "    <TR>" & vbCrLf
  .write "      <TD WIDTH=""99%"" BGCOLOR=#303030><FONT FACE=WINGDINGS COLOR=Gray VALIGN=Top SIZE=+1>;</FONT>&nbsp;"   & vbCrLf
  .write "      <INPUT TYPE=BUTTON NAME=DriveInfo VALUE=""View Drive Information"" ONCLICK=""window.location.href='" & request.Servervariables("SCRIPT_NAME") & "?getDRVs=@&txtpath=" & MyPath & "'""></TD>" & vbCrLf
  .write "      <TD BGCOLOR=#505050 WIDTH=""1%"" ><INPUT TYPE=BUTTON NAME=Logoff VALUE=""Logoff"" ONCLICK=""window.location.href='" & request.Servervariables("SCRIPT_NAME") & "?logoff=@'""></TD>" & vbCrLf
  .write "    </TR>" & vbCrLf
  .write "  </TABLE>" & vbCrLf
  .write "  <TABLE BGCOLOR=#505050 BORDER=0 WIDTH=""90%"">" & vbCrLf
  .write "    <TR>" & vbCrLf
  .write "      <TD BGCOLOR=#505050><FONT FACE=Arial SIZE=-2 COLOR=#99FFFF>Current Physical Path</FONT></TD>" & vbCrLf
  .write "      <TD ALIGN=RIGHT><FONT FACE=Arial SIZE=-2 COLOR=#99FFFF>Volume Label:&nbsp;&nbsp;</FONT>" & drv.VolumeName & "&nbsp;</TD>" & vbCrLf
  .write "    </TR>" & vbCrLf
  .write "    <TR>" & vbCrLf
  .write "      <TD COLSPAN=2 CELLPADDING=2 BGCOLOR=#303030>" & vbCrLf
  .write "        <FONT FACE=WINGDINGS COLOR=GRAY>1</FONT>" & vbCrLf
  .write "        <FONT FACE=Arial SIZE=+1>" & vbCrLf
  .write "        <INPUT TYPE=TEXT WIDTH=40 SIZE=60 NAME=TXTPATH VALUE=""" & showPath & """>&nbsp;&nbsp;" & vbCrLf
  .write "        <INPUT TYPE=SUBMIT NAME=CMD VALUE=""  View  "">" & vbCrLf
  .write "        </FONT>" & vbCrLf
  .write "      </TD>" & vbCrLf
  .write "    </TR>" & vbCrLf
  .write "  </TABLE>" & vbCrLf

  .write "</FORM>" & vbCrLf
  
  ' --------------------------------------------------------------------------------------
  ' Add or Delete Directory at Current Directory
  ' --------------------------------------------------------------------------------------
  
  if unlock = true then
  
    .write "<FORM METHOD=POST ACTION=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & MyPath & """>" & vbCrLf
    .write "  <TABLE BGCOLOR=#505050 BORDER=0 WIDTH=""90%"">" & vbCrLf
    .write "    <TR>" & vbCrLf
    .write "      <TD BGCOLOR=#505050><FONT FACE=Arial SIZE=-2 COLOR=#99FFFF>Directory</FONT></TD>" & vbCrLf
    .write "    </TR>" & vbCrLf
    .write "    <TR>" & vbCrLf
    .write "      <TD CELLPADDING=2 BGCOLOR=#303030>" & vbCrLf
    .write "        <FONT FACE=WINGDINGS COLOR=GRAY>0</FONT>" & vbCrLf
    .write "        <FONT FACE=Arial SIZE=+1>" & vbCrLf
    .write "      <INPUT TYPE=TEXT SIZE=60 NAME=DIRNAME>&nbsp;&nbsp;"
    .write "      <INPUT TYPE=SUBMIT NAME=CMD VALUE=Create>&nbsp;&nbsp;"
    .write "      <INPUT TYPE=SUBMIT NAME=CMD VALUE=Delete>&nbsp;&nbsp;"
    .write "      <INPUT TYPE=HIDDEN NAME=DIRSTUFF VALUE=@>" & vbCrLf
    .write "        </FONT>" & vbCrLf
    .write "      </TD>" & vbCrLf
    .write "    </TR>" & vbCrLf
    .write "  </TABLE>" & vbCrLf
    .write "</FORM>" & vbCrLf
  
  end if
  
end with

' --------------------------------------------------------------------------------------
' Directory View Tree
' --------------------------------------------------------------------------------------

with response

  .write "<TABLE CELLPADDING=2 WIDTH=""90%"" BGCOLOR=#505050>" & vbCrLf
  .write "  <TR>" & vbCrLf
  .write "    <TD VALIGN=TOP WIDTH=""50%"" BGCOLOR=#303030>Folders:<P>" & vbCrLf
  
  fo = 0
  if Rootfolder=false then
	.write "<FONT face=wingdings color=Gray >0</FONT> <FONT COLOR=#c8c8c8><span style='cursor: hand;' OnClick=""getit('..')"">..</span></FONT><BR>" & vbCrLf
  end if	
  for each fold in folder.SubFolders
    fo = fo + 1
  	.write "<FONT face=wingdings color=Gray >0</FONT> <FONT COLOR=#eeeeee><span style='cursor: hand;' OnClick=""getit('" & fold.name & "')"">" & fold.name & "</span></FONT><BR>" & vbCrLf
  next  
  
  .write "    </TD>" & vbCrLf
  
  .write "    <TD VALIGN=TOP BGCOLOR=#303030 NOWRAP>Files:<P>" & vbCrLf

  .flush

  .write "      <FORM METHOD=POST NAME=FRMCOPYSELECTED ACTION=""" & request.Servervariables("SCRIPT_NAME") & "?txtpath=" & MyPath & """>" & vbCrLf

  fi = 0
  MLenFileName = 0
  MLenFileSize = 0
  for each file in folder.Files
    if len(file.name) > MLenFileName then MLenFileName = len(file.name)
    LenFileSize = len(trim(Cstr(Int(file.size / 1024)+1)))
    if LenFileSize > MLenFileSize then MLenFileSize = LenFileSize
    fi = fi+1
  next
  
  if fi > 0 then
    Padding = ".........................................................................."
    
    .write "      <SELECT NAME=Fname SIZE=" & fi+2 & " STYLE=""font-family:Courier; font-size:10px; background-color: rgb(48,48,48); color: rgb(210,210,210)"">" & vbCrLf
    .write "<OPTION> </OPTION>" & vbCrLf
    for each file in folder.Files
      TempFileName = file.name
      TempFileNameLen = len(TempFileName)
      TempFileSizeLen = len(trim(Cstr(Int(file.size / 1024)+1)))
      .write "<OPTION VALUE=""" & TempFileName & """>"
      .write TempFileName
      .write Mid(Padding,1,MLenFileName - (TempFileNameLen -1))
      .write Mid(Padding,1,MLenFileSize - (TempFileSizeLen))    
      .write Int(file.size / 1024)+1 & " kb "
      
      ' Decode File Attributes
      if file.Attributes = 0 then
      
        .write "...."
  
      else
        if file.Attributes AND 1 then
          .write "R"
        else
          .write "."
        end If
    
        if file.Attributes AND 32 then
          .write "A"
        else
          .write "."
        end if
  
        if file.Attributes AND 2 then
          .write "H"
        else
          .write "."
        end if
    
        if file.Attributes AND 4 then
          .write "S"
        else
          .write "."
        end if
    
        if 1=2 then  ' Keep Code for future this detail not needed now
          if file.Attributes AND 8 then
            .write "V"
          else
            .write "."
          end if
      
          if file.Attributes AND 16 then
            .write "D"
          else
            .write "."
          end if
    
          if file.Attributes AND 64 then
            .write "L"
          else
            .write "."
          end if
      
          if file.Attributes AND 128 then
            .write "C"
          else
            .write "."
          end if
        end if
        
        .write " "
    
      end if
      
      .write FormatDate(1, FormatDateTime(file.datelastmodified,2)) & " "
      if len(FormatDateTime(file.datelastmodified,3)) = 10 then
        .write "0"
      end if
      .write FormatDateTime(file.datelastmodified,3)
  
      .write "</OPTION>" & vbCrLf
  
    next
  	
    .write "      </SELECT>" & vbCrLf
    
  else
    .write "Folder Empty"
  end if
	
  .write "      <P>" & vbCrLf
  
  .write "      <INPUT TYPE=hidden NAME=txtpath VALUE=""" & MyPath & """>" & vbCrLf
  
  if unlock = true and fi > 0 then
  
    .write "      <INPUT TYPE=Submit NAME=cmd VALUE=""Delete"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    .write "      <INPUT TYPE=SUBMIT NAME=cmd VALUE=""   Copy   "">&nbsp;&nbsp;"
    .write "      <INPUT TYPE=SUBMIT NAME=cmd VALUE=""Edit/Create"">&nbsp;&nbsp;"
    .write "      <INPUT TYPE=SUBMIT NAME=cmd VALUE=""Download"">" & vbCrLf
  
  end if  

	.write "      </FORM>" & vbCrLf
  
  if unlock = true then
  
    .write "      <P>"
    
    .write "      <HR COLOR=GRAY SIZE=1>"
    .write "      <FORM METHOD=POST NAME=frmCopyFile ACTION=""" & request.Servervariables("SCRIPT_NAME") & """>" & vbCrLf
  
    .write "      <FONT FACE=Arial SIZE=-2 COLOR=#99FFFF>Copy temp directory file to current directory:</FONT></BR>"  
  
    .write "      <SELECT SIZE=1 NAME=ToCopy>" & vbCrLf
    .write "        <OPTION>Select from list</OPTION>" & vbCrLf
  
    for each file in fileCopy.Files
    	.write "<OPTION>" & file.name & "</OPTION>" & vbCrLf
    next
    
    .write "      </SELECT>" & vbCrLf
  
    .write "      <INPUT TYPE=HIDDEN NAME=txtpath VALUE=""" & MyPath & """>" & vbCrLf
    .write "      <INPUT TYPE=Submit NAME=cmd VALUE=""  Copy "" >" & vbCrLf
  
    .write "      </FORM>" & vbCrLf
    
  end if

  .write "    </TD>" & vbCrLf
  .write "  </TR>" & vbCrLf
  .write "  <TR>" & vbCrLf
  .write "    <TD ALIGN=CENTER><B>Total Folders: " & fo & "</B></TD>" & vbCrLf
  .write "    <TD ALIGN=CENTER><B>Total Files: " & fi & "</B></TD>" & vbCrLf
  .write "  </TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf

end with

' --------------------------------------------------------------------------------------
' File Upload Dialog
' --------------------------------------------------------------------------------------

with response
 
  if unlock = true then
  
    .write "<FORM METHOD=""POST"" ENCTYPE=""multipart/form-data"" ACTION=""" & request.Servervariables("SCRIPT_NAME") & "?upload=@&txtpath=" & MyPath & """>" & vbCrLf
    .write "<TABLE BGCOLOR=""#505050"" CELLPADDING=""8"">" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD BGCOLOR=#303030 VALIGN=""bottom"">" & vbCrLf
    .write "      <FONT SIZE=+1 FACE=WINGDINGS COLOR=GRAY>2</FONT><FONT FACE=""Arial"" SIZE=-2 COLOR=""#99FFFF""> Select Files to Upload:<BR>" & vbCrLf
    .write "      <INPUT TYPE=""FILE"" SIZE=""53"" NAME=""FILE1""><BR>" & vbCrLf
    .write "      <INPUT TYPE=""FILE"" SIZE=""53"" NAME=""FILE2""><BR>" & vbCrLf
  	.write "      <INPUT TYPE=""FILE"" SIZE=""53"" NAME=""FILE3"">&nbsp;&nbsp;"
    .write "      <INPUT TYPE=""submit"" VALUE=""Upload"" NAME=""Upload"">" & vbCrLf
  	.write "      </FONT>" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf
    .write "</FORM>" & vbCrLf

  end if
  
end with

' --------------------------------------------------------------------------------------
' End of Main Page
' --------------------------------------------------------------------------------------

with response

  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf

end with
  
' --------------------------------------------------------------------------------------
' Functions, Subroutines and Class
' --------------------------------------------------------------------------------------

Function FormatDate(iDateFormat, TempDate)

  Dim strDay
  Dim strMonth
  
  if Len(DatePart("d", TempDate)) = 1 then
  	strDay = "0" & DatePart("d", TempDate)
  else
  	strDay = DatePart("d", TempDate)
  end if
  
  if Len(DatePart("m", TempDate)) = 1 then
  	strMonth = "0" & DatePart("m", TempDate)
  else
  	strMonth = DatePart("m", TempDate)
  end if
  
  if iDateFormat = 1 then
  	FormatDate = strMonth & "/" & strDay & "/" & DatePart("yyyy", TempDate)
  elseif iDateFormat = 2 then
  	FormatDate = strDay & "/" & strMonth & "/" & DatePart("yyyy", TempDate)
  else
  	FormatDate = DatePart("yyyy", TempDate) & "/" & strMonth & "/" & strDay
  end if

end function

' --------------------------------------------------------------------------------------

Function BufferContent(data)
	Dim strContent(64)
	Dim i
	ClearString strContent
	For i = 1 To LenB(data)
		AddString strContent,Chr(AscB(MidB(data,i,1)))
	Next
	BufferContent = fnReadString(strContent)
End Function

' --------------------------------------------------------------------------------------

Sub ClearString(part)
	Dim index
	For index = 0 to 64
		part(index)=""
	Next
End Sub

' --------------------------------------------------------------------------------------

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

' --------------------------------------------------------------------------------------

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

' --------------------------------------------------------------------------------------

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

' --------------------------------------------------------------------------------------

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
		'Set oFile = oFS.CreateTextFile(sPath & FileName, True)
		' output mechanism modified for buffering
        on error resume next
		'oFile.Write BufferContent(FileData)
		SaveStreamData (sPath & FileName),FileData
        if err.number <> 0 then
          response.write err.number & ": " & err.description & "<BR>"
        end if
    on error goto 0
		'oFile.Close
	End Sub

	Public sub SaveToDatabase(ByRef oField)
		If LenB(FileData) = 0 then Exit Sub
		If IsObject(oField) Then
			oField.AppendChunk FileData
		end if
	End Sub
	'Added by zensar on 25-08-2006
	Function SaveStreamData(FileName, ByteArray)
          Const adTypeText = 2
          Const adSaveCreateOverWrite = 2
          
          'Create Stream object
          Dim FileStream
          Set FileStream = CreateObject("ADODB.Stream")
          
          'Specify stream type.
          FileStream.Type=adTypeText
          
          'Open the stream And write text data To the object
          FileStream.Open
          FileStream.WriteText FileData
          
          'Save text data To disk
          FileStream.SaveToFile FileName, adSaveCreateOverWrite
    End Function
    ''''''''''''''''''''''''''''''''''
End Class
' --------------------------------------------------------------------------------------
%>