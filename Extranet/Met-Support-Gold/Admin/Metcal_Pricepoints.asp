<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Setup Connection
' --------------------------------------------------------------------------------------
response.buffer = true

Call Connect_SiteWide

%>

<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Administrator/CK_Admin_Credentials.asp"-->
<!-- #include virtual="/met-support-gold/admin/CK_Credentials.asp"-->
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->

<%
' --------------------------------------------------------------------------------------
' Build Nav and Header information
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_Width
Dim BackURL

Top_Navigation  = False 
Side_Navigation = True
Screen_Title    = Site_Description & " - " & "Calibration Procedure Price Point Check"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "Calibration Procedure Price Point Check" & "</FONT>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

' --------------------------------------------------------------------------------------
' Start building form to display procedure details
' --------------------------------------------------------------------------------------

response.write "<SPAN CLASS=Heading3>Calibration Procedure Price Point Check</SPAN>"
response.write "<BR><BR>"

response.write "<SPAN CLASS=Small>"

Call Nav_Border_Begin
response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" ONCLICK=""return document.location='metcal_admin.asp?KeyWord=" & strSearchkeyword & "&Calibrator=" & strSearchCalibrator & "'"" value=""MetCal Administration Menu"">"
Call Nav_Border_End

response.write "<P>"
response.flush

SQL = "SELECT DISTINCT Price_Point_ID, ZipFileName " &_
      "FROM         dbo.Metcal_Procedures " &_
      "GROUP BY Price_Point_ID, ZipFileName " &_
      "ORDER BY ZipFileName"

Set rsPP = Server.CreateObject("ADODB.Recordset")
rsPP.Open SQL, conn, 3, 3

Dim FileName, FileNameCount, curFileName, oldFileName, PPBad
FileName      = ""
FileNameCount = 0
PPBad         = false
curFileName   = ""
oldFileName   = "blank"

do while not rsPP.EOF

  curFileName = rsPP("ZipFileName")
  
  if UCase(curFileName) <> "NA" and not isblank(curFileName) then
  
    if oldFileName = curFileName and FileNameCount = 0 then
      if PPBad = false then
        FileName = FileName & curFileName
      else
        FileName = FileName & "," & curFileName
      end if
      
      FileNameCount = FileNameCount + 1
      PPBad = true
    elseif oldFileName = curFileName then
      FileNameCount = FileNameCount + 1
    else
      oldFileName = curFileName
      FileNameCount = 0
    end if
    
  end if
    
  rsPP.MoveNext

loop

rsPP.close
set rsPP = nothing

if PPBad = true then
  response.write "<SPAN CLASS=SmallBold>The following procedures using the same Zip File Names have mismatched price points.  Click on the Zip File Name below to list all procedures using this Zip File Name.</SPAN><P>"
  LinkFileName = split(FileName,",")
  LinkFileNameCount = Ubound(LinkFileName)
  
  for x = 0 to LinkFileNameCount
    response.write "<A HREF=""/met-support-gold/admin/metcal_procedures.asp?FileName=" & LinkFileName(x) & """>" & LinkFileName(x) & "<A><BR>"
  next
  'response.write replace(FileName,",","<BR>")
else
  response.write "<SPAN CLASS=SmallBold>There are <SPAN CLASS=SMALLBOLDRED>no</SPAN> procedures using the same Zip File Names have mismatched price points.</SPAN><P>"
end if

response.write "</SPAN>"
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide
response.flush
%>
