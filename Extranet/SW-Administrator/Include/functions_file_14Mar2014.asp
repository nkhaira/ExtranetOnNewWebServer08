<%
'
'  File Functions
'  Author:  Kelly Whitlock
'  Date:    7/19/2005

' --------------------------------------------------------------------------------------

Function FileExists(strFileName)

  Dim ckfile
  Dim strFilePath
    
  Set ckfile = CreateObject("Scripting.FileSystemObject")  
  strFilePath = request.ServerVariables("APPL_PHYSICAL_PATH")
  strFilePath = replace(LCase(strFilePath),"sw-administrator\","")
  strFilePath = mid(strFilePath, 1, instrrev(strFilePath, "\")) & strFileName %><%
  
'  response.write strFilePath
'  response.end
  
  if ckfile.FileExists(strFilePath) then
    set ckfile = nothing
    FileExists = True
  else
    set ckfile = nothing  
    FileExists = False
  end if

end function

' --------------------------------------------------------------------------------------

function FileStream(strFileName, bShowDialog)

  adTypeBinary = 1
  
  Set objStream = Server.CreateObject("ADODB.Stream")
  objStream.Open
  objStream.Type = adTypeBinary
  objStream.LoadFromFile strFileName
  
  intFileSize = objStream.Size
  
  strFile = Mid(strFileName, InstrRev(strFileName, "/") + 1) %><%
  
  ' Output File Name
  
  strFile   = Replace(strFile," ","_")  ' Convert spaces to underscores
  FileOrig  = strFileName
  FileRoot  = LCase(Mid(strFileName, 1, InstrRev(strFileName, ".") - 1))
  FileExtn  = LCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))
  
  ' Get Content Type for this file
  
  Call Connect_SiteWide
  
  SQL = "SELECT ContentType FROM Asset_Type WHERE File_Extension='" & FileExtn & "'"
  Set rsType = Server.CreateObject("ADODB.Recordset")
  rsType.Open SQL, conn, 3, 3
  
  if not rsType.EOF then
    ContentType = rsType("ContentType")
  else
    ContentType = "application/octet-stream"
  end if
  
  rsType.close
  set rsType = nothing
  
  Call Disconnect_SiteWide
  
  ' Stream the File
  
  with response

    if CInt(bShowDialog) = CInt(True) then
'      .AddHeader "Content-Disposition", "attachment; filename=" & strFile
      .AddHeader "Content-Disposition", "attachment; filename=" & FileRoot & "." & FileExtn
    end if

    .AddHeader     "Content-Length", intFileSize  
    .Charset     = "UTF-8"
    .ContentType = ContentType
    .BinaryWrite   objStream.Read
    .Flush

  end with
  
  objStream.Close
  Set objStream = Nothing

end function
%>