<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------

'response.buffer = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/Pop-Up.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<%

Dim BackURL
Dim LimitView
Dim ErrorString

BackURL = Session("BackURL")    

Dim strPath                 ' Path of directory to show
Dim objFSO                  ' FileSystemObject variable
Dim objFolder               ' Folder variable
Dim objItem                 ' Variable used to loop through the contents of the folder

Dim Sub_Directories
Sub_Directories = False     ' Set to True to show sub directories

strPath = Mid(Request.ServerVariables("PATH_INFO"),1, InstrRev(Request.ServerVariables("PATH_INFO"), "/"))

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Directory Listing",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Service Documents",Login_Language,conn) & "</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>" & Translate("Directory Listing",Login_Language,conn) & "</FONT><BR>"
response.write "<FONT CLASS=Heading4>" & strPath & "</FONT>"
response.write "<FONT CLASS=Medium>"
response.write "<BR><BR>"

' --------------------------------------------------------------------------------------
' Directory / File List
' --------------------------------------------------------------------------------------

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))

response.write "<TABLE WIDTH=""100%"" BORDER=1 CELLPADDING=0 CELLSPACING=0 BORDERCOLOR=""Black"" BGCOLOR=""Black"">"
response.write "<TR>"
response.write "<TD>"
    
response.write "<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH=""100%"">"
response.write "<TR>"
response.write "<TD BGCOLOR=""Black""><FONT CLASS=SmallBoldGold>" & Translate("Name",Login_Language,conn) & "</FONT></TD>"
response.write "<TD BGCOLOR=""Black"" ALIGN=CENTER WIDTH=""10%""><FONT CLASS=SmallBoldGold>" & Translate("Size",Login_Language,conn) & "</FONT></TD>"
response.write "<TD BGCOLOR=""Black"" ALIGN=CENTER WIDTH=""10%""><FONT CLASS=SmallBoldGold>" & Translate("Date",Login_Language,conn) & "</FONT></TD>"
response.write "<TD BGCOLOR=""Black""><FONT CLASS=SmallBoldGold>" & Translate("Type",Login_Language,conn) & "</FONT></TD>"
response.write "</TR>"

' --------------------------------------------------------------------------------------
' Directories
' --------------------------------------------------------------------------------------

if show_directories = true then
  For Each objItem In objFolder.SubFolders
  	if InStr(1, objItem, "_vti", 1) = 0 then
      response.write "<TR BGCOLOR=""White"">"
        response.write "<TD ALIGN=Left><FONT Class=Small><A HREF=""" & strPath & objItem.Name & """>" & UCase(objItem.Name) & "</A></FONT></TD>"
      	response.write "<TD ALIGN=Right><FONT Class=Small>"
          if CDbl(cdbl(objItem.Size / 1024)) = 0 then
            response.write "1"
          else
            response.write FormatNumber(CDbl(cdbl(objItem.Size / 1024)),0)
          end if
          response.write " KB</FONT></TD>"
    		response.write "<TD ALIGN=Right><FONT Class=Small>" & FormatDate(1, objItem.DateLastModified) & "</FONT></TD>"
    		response.write "<TD ALIGN=Left><FONT Class=Small>" & Translate(objItem.Type,Login_Language,conn) & "</FONT></TD>"
      response.write "</TR>"
    end if
  Next
end if

' --------------------------------------------------------------------------------------
' Files
' --------------------------------------------------------------------------------------

For Each objItem In objFolder.Files
  select case LCase(mid(objItem.Name,instr(1,objItem.Name,".")))
    case ".asp", ".htm", ".html", ".css", ".tag", ".shtml", ".cfm", ".cfml", ".asa", ".js", ".inc", ".apf", ".scc"
    case else
      response.write "<TR BGCOLOR=""White"">"
      response.write "<TD ALIGN=Left>"
        response.write "<FONT Class=Small><A HREF=""" & strPath & objItem.Name & """>"
        Call Write_Icon
        response.write "</A>&nbsp;&nbsp;"
        response.write UCase(objItem.Name) & "</FONT></TD>"
    	response.write "<TD ALIGN=Right><FONT Class=Small>"
        if CDbl(Cdbl(objItem.Size / 1024)) = 0 then
          response.write "1"
        else
          response.write FormatNumber(CDbl(Cdbl(objItem.Size / 1024)),0)
        end if
        response.write " KB</FONT></TD>"
      
      'select case DateDiff("d", Date, objItem.DateLastModified))
      '  case 31..365
      '  case 22..30
      '  case 15..21
      '  case 8..14
      '  case else
      'end select
      
        
  		response.write "<TD ALIGN=Right><FONT Class=Small>" & FormatDate(1, objItem.DateLastModified) & "</FONT></TD>"
  		response.write "<TD ALIGN=Left><FONT Class=Small>" & Translate(objItem.Type,Login_Language,conn) & "</FONT></TD>"
    response.write "</TR>"
  end select
Next

response.write "</TABLE>"
response.write "</TR>"
response.write "</TD>"
response.write "</TABLE>"

Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

' --------------------------------------------------------------------------------------

response.write "<BR><BR>"

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

'if isobject(rsIcon) then
'  rsIcon.close
'  set rsIcon  = nothing
'  set SQLIcon = nothing
'end if

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------

sub Write_Icon

  Icon_Extension = UCase(mid(objItem.Name,Instr(1,objItem.Name,".")+1))
  Icon_Image     = "<IMG SRC=""/images/Button-TXT.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"

'  if not isobject(rsIcon) then
'    SQLIcon = "SELECT * FROM Asset_Type" ' WHERE File_Extension='" & Icon_Extension & "'"
'    Set rsIcon = Server.CreateObject("ADODB.Recordset")
'    rsIcon.Open SQLIcon, conn, 3, 3
'  else

'    rsIcon.MoveFirst
    
'    do while not rsIcon.EOF then
'      if UCase(rsIcon("File_Extension") = Icon_Extension then
'        Icon_Image = "<IMG SRC=""" & rsIcon("Icon_File") & """ BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
'        exit do
'      end if
'      rsIcon.MoveNext
'   end if

'   response.write Icon_Image

'    SQLIcon = "SELECT * FROM Asset_Type"  WHERE File_Extension='" & Icon_Extension & "'"
'    Set rsIcon = Server.CreateObject("ADODB.Recordset")
'    rsIcon.Open SQLIcon, conn, 3, 3

'    if not rsIcon.EOF then
'      response.write "<IMG SRC=""" & rsIcon("Icon_File") & """ BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
'    else
      response.write Icon_Image
'    end if

'    rsIcon.close
'    set rsIcon  = nothing
'    set SQLIcon = nothing

end sub

%>