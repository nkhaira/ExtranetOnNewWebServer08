<%@Language="VBScript" Codepage=65001%>

<%
' --------------------------------------------------------------------------------------
'
' Author: Kelly Whitlock
' Date:   10/10/2005
' --------------------------------------------------------------------------------------

server.scripttimeout = 600  ' 10 Minutes

%>
<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/include/functions_File.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Dim Site_Code
if not isblank(request.form("Site_Code")) then
  Site_Code = request.form("Site_Code")
else
  Site_Code = ""
end if

Dim SubDirectory(2)
SubDirectory(0) = "Asset"
SubDirectory(1) = "Archive"
SubDirectory(2) = "Thumbnail"

Call Connect_SiteWide

SQL = "SELECT ID, Site_Code, Site_Description from Site WHERE Enabled=-1 and Site_Alias is NULL and NoShow=0"
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

with response

  .write "<FONT FACE=Arial Size=2>"

  .write "<FORM NAME=""Site_Picker"" METHOD=""POST"" ACTION=""" & request.ServerVariables("SCRIPT_NAME") & """>" & vbCrLf
  .write "<SELECT NAME=""Site_Code"">" & vbCrLf

  do while not rsSite.EOF
  
    .write "<OPTION VALUE=""" & rsSite("Site_Code") & """"
    if Site_Code = rsSite("Site_Code") then
      .write " SELECTED"
      Site_ID = rsSite("ID")
    end if  
    .write ">" & rsSite("Site_Description") & "</OPTION>" & vbCrLf

    rsSite.MoveNext
    
  loop
  
  .write "</SELECT>" & vbCrLf
  .write "<INPUT TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""GO"">" & vbCrLf
  .write "</FORM>" & vbCrLf
  
end with

rsSite.close
set rsSite = nothing

if not isblank(Site_Code) then

  UP = "HTTP://" & request.ServerVariables("HTTP_HOST") & "/" & Site_Code
  PP = Server.MapPath("/" & Site_Code)

  response.write "<B>" & UP & "</B><P>"
  
  response.write "The following physical files located in the directories below can be deleted since they do not have a corresponding record in the SiteWide.Calendar table.<P>"
  
  for subD = 0 to 2
  
    response.write "<P><B>/" & SubDirectory(subD) & "</B><P>"
  
    URL_Path = UP & "/Download/" & SubDirectory(subD) & "/"
    PHY_Path = PP & "\Download\" & SubDirectory(subD) & "\"
  
    Call CheckDirectory(URL_Path,PHY_Path,Site_ID,SubDirectory(subD),conn)
    
  next
  
  response.write "<P>"  

end if
  
response.write "</FONT>"

Call Disconnect_SiteWide

response.end

' Functions and Subroutines
         
function CheckDirectory(UP,PP,Site_ID,SubDPath,conn)

	sPP = PP 'Physical Path
	sUP = UP 'URL Path

	set fso = CreateObject("Scripting.FileSystemObject")
	set f = fso.GetFolder(sPP)  
	set fc = f.Files 
	set ff = f.SubFolders
   
	For Each fl in fc

    FileToCheck = "download/" & SubDPath & "/" & fl.name
    
    select case SubDPath
      case "Asset"
        SQLFile = "SELECT File_Name FROM Calendar WHERE File_Name='" & FileToCheck & "' AND Site_ID=" & Site_ID
      case "Archive"
        SQLFile = "SELECT File_Name FROM Calendar WHERE Archive_Name='" & FileToCheck & "' AND Site_ID=" & Site_ID
      case "Thumbnail"
        SQLFile = "SELECT File_Name FROM Calendar WHERE Thumbnail='" & FileToCheck & "' AND Site_ID=" & Site_ID      
    end select  

    Set rsFile = Server.CreateObject("ADODB.Recordset")
    rsFile.Open SQLFile, conn, 3, 3
    
    if not rsFile.EOF then
      response.write "<A HREF=" & sUP & fl.name & ">" & fl.name & "</A> " & "<BR>" & vbCrLf
      response.flush
    end if
    
    rsFile.close
    set rsFile  = nothing
    set SQLFile = nothing  
    
	Next  

	Set ff  = nothing
	Set fso = nothing
	Set f   = nothing
	Set fc  = nothing
  Set fl  = nothing
  
end function
%>
