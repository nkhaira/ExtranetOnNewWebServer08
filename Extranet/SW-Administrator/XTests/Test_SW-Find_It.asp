<%@Language="VBScript" Codepage=65001%>
<%

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

if isblank(request("File_Name")) then

  Call Connect_SiteWide
  
  with response
  
  .write "<FORM NAME=""Stream_ID"" METHOD=""GET"">"
  .write "<SELECT NAME=""File_Name"">"
  
  SQL = "SELECT Title, File_Name from Calendar WHERE Status=1 AND File_Name IS NOT NULL AND SITE_ID=82"
  
  Set rsAsset = Server.CreateObject("ADODB.Recordset")
  rsAsset.Open SQL, conn, 3, 3
  
  do while not rsAsset.EOF
    .write "<OPTION VALUE=""" & rsAsset("File_Name") & """>" & rsAsset("Title") & "</OPTION>" & vbCrLf
    rsAsset.MoveNext
  loop
  
  .write "</SELECT><P>"

  rsAsset.close
  set rsAsset = nothing
  
  .write "<INPUT TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""Stream The File"">"
  .write "</FORM><P>"
  
  .write "After you have streamed the file, press the [Back] button to return to this file select form."
  
  end with
  
  Call Disconnect_SiteWide
  
else 

  strFileName = "/portweb/" & request("File_Name")
  response.write strFileName
  strFileName = Server.Mappath(strFileName)


' File Stream Code Here

  adTypeBinary = 1
  Set objStream = Server.CreateObject("ADODB.Stream")
  chunk = 2048
  
  objStream.Open
  objStream.Type = adTypeBinary
  objStream.LoadFromFile strFileName
  
  intFileSize = objStream.Size
  
  strFile = Mid(strFileName, InstrRev(strFileName, "\") + 1) %><%
  
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
  
    .clear
  
    if CInt(bShowDialog) = CInt(True) then
      .AddHeader "Content-Disposition", "attachment; filename=" & strFile
    else
      .AddHeader "Content-Disposition", "inline; filename=" & strFile
    end if

    .AddHeader     "Content-Length", intFileSize  
    .Charset     = "UTF-8"
    .ContentType = ContentType
    
    if bDownloadOnly then
      .CacheControl = "no-store"
    end if

    for i = 1 to intFileSize \ chunk 
       if not .IsClientConnected then exit for 
        .BinaryWrite objStream.Read(chunk)
        .flush
    next 
 
    if intFileSize mod chunk > 0 Then 
        if .IsClientConnected then 
          .BinaryWrite objStream.Read(intFileSize Mod chunk)
          .flush
        end if 
    end if 

    ' .BinaryWrite   objStream.Read
    .flush
  
    objStream.Close
    Set objStream = Nothing
    
  end with
  
  for x = 1 to 1000
  next
  
end if
%>
