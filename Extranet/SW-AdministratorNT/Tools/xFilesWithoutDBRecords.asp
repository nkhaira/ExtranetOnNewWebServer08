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

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 60 * 10        ' Seconds * Minutes

if isblank(request("DoWhat")) then
  response.write "What do you want to do?  List or Delete"
  response.end
elseif UCase(request("DoWhat")) = "LIST" then
  DoWork = false
elseif UCase(request("DoWhat")) = "Delete" then
  DoWork = true
end if

if isblank(request("Site")) then
  response.write "What Site ID?"
  response.end
elseif isnumeric(request("Site")) then
  Site_ID = request("Site")
else
  response.write "Invalid Site ID"
  response.end
end if

Call Connect_SiteWide

SQL = "SELECT Site_Code FROM dbo.Site WHERE ID=" & Site_ID

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

if not rsID.EOF then
  Site_Code = rsID("Site_Code")
else
  response.write "Invalid Site ID. Site ID not found"
  response.end
end if  

' --------------------------------------------------------------------------------------

Dim DirToCheck(2)

DirToCheck(0) = "Asset"
DirToCheck(1) = "Archive"
DirToCheck(2) = "Thumbnail"

for AssetDir = 0 to 2

  SourceFolderName   = Server.MapPath("/find-sales/download/" & DirToCheck(AssetDir))
  IncludeSubfolders  = false
  
  Call ListFilesInFolder(SourceFolderName, DirToCheck(AssetDir), IncludeSubfolders, Site_ID)
  
next
  
' --------------------------------------------------------------------------------------

sub ListFilesInFolder(SourceFolderName, DirToCheck, IncludeSubfolders, Site_ID)

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
    
    select case DirToCheck
      case "Asset"
        SQL = "SELECT DISTINCT File_Name FROM Calendar WHERE Site_ID=" & Site_ID & " AND File_Name='Download/" & DirToCheck & "/" & objFile.Name & "'"
      case "Archive"
        SQL = "SELECT DISTINCT Archive_Name FROM Calendar WHERE Site_ID=" & Site_ID & " AND Archive_Name='Download/" & DirToCheck & "/" & objFile.Name & "'"
      case "Thumbnail"
        SQL = "SELECT DISTINCT Thumbnail FROM Calendar WHERE Site_ID=" & Site_ID & " AND Thumbnail='Download/" & DirToCheck & "/" & objFile.Name & "'"     
    end select
        
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
  
  Next

  ' Clean up
  Set objFolder = Nothing
  Set objFile = Nothing
  Set objFSO = Nothing

 end sub

Call Disconnect_SiteWide

response.write "<P>Done"
%>