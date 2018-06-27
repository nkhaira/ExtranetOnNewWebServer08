<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
' --------------------------------------------------------------------------------------
' The purpose of this script is to delete Asset, Archive and Thumbnail files for each sub-site
' that do not not exist in File_Name, Archive_Name, and Thumbnail in SiteWide.Calendar
'
' This script can be run anytime syncing is required.
'
' Note: Set Site_ID value to ID in Calendar.Site of the sub-site you wish to work on.
'
' Author: Kelly Whitlock
' --------------------------------------------------------------------------------------

Dim Site_ID, Site_Code, DoWhat, DoWork
Site_ID = 3

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 60 * 10        ' Seconds * Minutes

' --------------------------------------------------------------------------------------

SourceFolderName   = Server.MapPath("/POD/")
IncludeSubfolders  = false

if isblank(request("DoWhat")) then
  response.write "What do you want to do?  List or Delete"
  response.end
elseif UCase(request("DoWhat")) = "LIST" then
  DoWork = false
elseif UCase(request("DoWhat")) = "DELETE" then
  DoWork = true
end if

Call Connect_SiteWide

Call ListFilesInFolder(SourceFolderName, IncludeSubfolders)
  
' --------------------------------------------------------------------------------------

sub ListFilesInFolder(SourceFolderName, IncludeSubfolders)

  Dim objFSO
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

  ' Get the folder object associated with the directory
  Dim objFolder
  Set objFolder = objFSO.GetFolder(SourceFolderName)

  response.Write "<P><FONT COLOR=RED><B>\" & objFolder.Name & "</B></FONT><BR>"

  ' Loop through the Files collection
  Dim objFile
  
  for each objFile in objFolder.Files
    
    response.Write objFile.Name & " "
    
    SQL = "SELECT DISTINCT File_Name_POD FROM Calendar WHERE File_Name_POD='POD/" & objFile.Name & "'"
        
    Set rsFile = Server.CreateObject("ADODB.Recordset")
    rsFile.Open SQL, conn, 3, 3
    
    if rsFile.EOF then
      if DoWork then
        objFile.Delete
      end if  
      response.write "<FONT COLOR=RED><B>Deleted</B></FONT>"
    end if
    response.write "<BR>"
    response.flush
    
    rsFile.close
    set rsFile = nothing
  
  next

  ' Clean up
  Set objFolder = Nothing
  Set objFile = Nothing
  Set objFSO = Nothing

 end sub

Call Disconnect_SiteWide

response.write "<P>Done"
%>