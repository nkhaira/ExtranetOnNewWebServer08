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

if isblank(request("Site")) then
  response.write "What Site ID?"
  response.end
elseif isnumeric(request("DoWhat")) then
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

Dim TypeToCheck(2)

TypeToCheck(0) = "File_Name"
TypeToCheck(1) = "File_Name_POD"
TypeToCheck(2) = "Thumbnail"

Dim PathToCheck(2)

PathToCheck(0) = "/" & Site_Code & "/"
PathToCheck(1) = "/"
PathToCheck(2) = "/" & Site_Code & "/"

Dim Counter(2)
Counter(0) = 0
Counter(1) = 0
Counter(2) = 0

set FileObj = Server.CreateObject("Scripting.FileSystemObject")
    
for AssetDir = 0 to 2

  SQL = "SELECT " & TypeToCheck(AssetDir) & " AS FileToCheck FROM Calendar WHERE " & TypeToCheck(AssetDir) & " IS NOT NULL AND Site_ID=" & Site_ID & " ORDER BY Item_Number"
  Set rsID = Server.CreateObject("ADODB.Recordset")
  rsID.Open SQL, conn, 3, 3

  response.write "<P>" & TypeToCheck(AssetDir) & "</B><P>"

  do while not rsID.EOF

    select case AssetDir
      case 0, 2
        FileToCheck = "/" & Site_Code & "/" & rsID("FileToCheck")
      case 1
        FileToCheck = "/" & rsID("FileToCheck")      
    end select  
    
    MyFilePath = Server.MapPath(FileToCheck)

    if not FileObj.FileExists(MyFilePath) then
      response.write rsID("FileToCheck") & "<BR>"
'      response.write MyFilePath  & "<P>"      
      response.flush
      Counter(AssetDir) = Counter(AssetDir) + 1
    end if

    rsID.MoveNext
    
  loop
  
  rsID.Close
  set rsID = nothing
 
next

set FileObj = Nothing  

response.write "<P>"  

for AssetDir = 0 to 2
  response.write TypeToCheck(AssetDir) & " Count: " & Counter(AssetDir) & "<BR>"
next
  
' --------------------------------------------------------------------------------------

Call Disconnect_SiteWide

response.write "<P>Done"
%>