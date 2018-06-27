<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Setup Connection
' --------------------------------------------------------------------------------------
response.buffer = true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------
<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
<!-- #include virtual="/SW-Administrator/CK_Admin_Credentials.asp"-->
<!-- #include virtual="/met-support-gold/admin/CK_Credentials.asp"-->

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
Screen_Title    = Site_Description & " - " & "Calibration Procedure Download Administration"
Bar_Title       = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>" & "Calibration Procedure Download Administration" & "</SPAN>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->

<%
' --------------------------------------------------------------------------------------
' Get Posted values, build various variables needed later
' --------------------------------------------------------------------------------------
Dim strSubCategory, strHREF
Dim iCategory_ID, strDescription
Dim rsCategories

' If no category was passed in then go back to the admin page
strSubCategory = Trim(Request("SUBCATEGORY"))

if strSubCategory= "" then
	response.redirect "metcal_admin.asp"
end if

' --------------------------------------------------------------------------------------
' Start building table to display records for desired subcategory
' --------------------------------------------------------------------------------------
response.write "<SPAN CLASS=Heading3>Calibration Procedure Download Admin - " & strSubcategory & "</SPAN>"
response.write "<BR><BR>"
Call Nav_Border_Begin
response.write "<input type=""button"" class=""NavLeftHighlight1"" onclick=""return document.location='metcal_admin.asp'"" value=""Back to Metcal Admin"">&nbsp;&nbsp;&nbsp;&nbsp;"
response.write "<input type=""button"" class=""NavLeftHighlight1"" onclick=""return document.location='metcal_category.asp?action=new&subcategory=" & strSubCategory & "'"" value=""Create New"">"
Call Nav_Border_End
response.write "<P>"
response.write "<SPAN CLASS=Medium>"

' Get Recordset
set rsCategories = GetCategories(strSubCategory)

Call Table_Begin
response.write "<TABLE BORDER=1 BORDERCOLOR=""Gray"" CELLSPACING=0 CELLPADDING=2 WIDTH=""100%"" BGCOLOR=""#EEEEEE"">"
response.write "<TR>"
response.write "<TD ALIGN=CENTER CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"" WIDTH=""5%"">ID</TD>"
response.write "<TD CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"">Description</TD>"
response.write "<TD ALIGN=CENTER CLASS=MEDIUMBOLDGOLD BGCOLOR=""#000000"" WIDTH=""5%"">Action</TD>"
response.write "</TR>"
response.write "<TBODY>"

    ToggleColor = "#FFFFFF"
    ToggleStr = ""

  do while not rsCategories.eof
    iCategory_ID = rsCategories("Category_ID")
    strDescription = rsCategories("Description")
   	strHREF = "<a href=""metcal_category.asp?category_id=" & iCategory_ID & "&action=edit&subcategory=" & strSubCategory & """>"
    response.write "<TR VALIGN=TOP>"

      response.write "<TD ALIGN=CENTER CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """>" & strHREF & iCategory_ID & "</TD>"
      response.write "<TD CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """>" & strHREF & strDescription & "</TD>"
      response.write "<TD ALIGN=CENTER CLASS=MEDIUM BGCOLOR=""" & ToggleColor & """><a href=""javascript:CheckDelete(" & iCategory_ID & ", '" & strSubCategory & "')"">Delete</a></TD>"
    response.write "</TR>"

    rsCategories.moveNext
  Loop

  Response.write "</TBODY>"
  response.write "</TABLE>"
Call Table_End
	WriteScripts
%>



<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush
%>

<%
' --------------------------------------------------------------------------------------
' Get all records for a given subcategory. One stored proc returns them all.
' --------------------------------------------------------------------------------------

Function GetCategories(strSubCategory)
  Dim cmd, prm, rsCategories

'response.write("strSubCategory: " & strSubCategory & "<BR>")

  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "Admin_MetCal_Categories_GetList"

  Set prm = cmd.CreateParameter("@strSubCategory", adVarChar, adParamInput, 50, strSubCategory & "")
  cmd.Parameters.Append prm

  Set rsCategories = Server.CreateObject("ADODB.Recordset")
  rsCategories.CursorLocation = adUseClient
  rsCategories.CursorType = adOpenDynamic
  rsCategories.open cmd

  set prm = nothing
  set cmd = nothing

  set GetCategories = rsCategories
End Function

Function WriteScripts
%>
		<SCRIPT LANGUAGE="JavaScript">
		var OldNumber

		function CheckDelete (ID, strSubCategory)
		{
		var msg = "\n Are you sure you want to delete this entry type?";
		if (confirm(msg))
			location.replace("metcal_category.asp?action=delete&category_ID=" + ID + "&subcategory=" + strSubCategory);
		}
		</script>
<%
End Function
%>