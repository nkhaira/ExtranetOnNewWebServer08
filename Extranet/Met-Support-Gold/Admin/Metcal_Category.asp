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
Screen_Title    = Site_Description & " - " & "Calibration Procedure Download"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & "Calibration Procedure Download" & "</FONT>"
Content_Width   = 95  ' Percent
BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->

<%
' --------------------------------------------------------------------------------------
' Get Posted values, build various variables needed later. If this is a post, save the record.
' --------------------------------------------------------------------------------------
Dim strSubCategory, iCategory_ID, strAction, strPostFlag, strDescription, iSubCategory_ID
Dim strOracleID, strListPrice
Dim rsCategory

strPostFlag = Request("PostFlag")
strAction = Trim(Request("ACTION"))
strSubCategory = Trim(Request("SUBCATEGORY"))
iCategory_ID = Request("Category_ID")
iSubCategory_ID = Request("SUBCATEGORY_ID")
strDescription = Trim(Request("Description"))
strOracleID = Trim(Request("OracleID"))
strListPrice = Trim(Request("ListPrice"))

'response.write("strPostFlag: " & strPostFlag & "<BR>")
'response.write("strAction: " & strAction & "<BR>")
'response.write("strSubCategory: " & strSubCategory & "<BR>")
'response.write("iCategory_ID: " & iCategory_ID & "<BR>")
'response.write("iSubCategory_ID: " & iSubCategory_ID & "<BR>")
'response.write("strDescription: " & strDescription & "<BR>")
'response.write("strOracleID: " & strOracleID & "<BR>")
'response.write("strListPrice: " & strListPrice & "<BR>")

' If no category was passed in then go back to the admin page
if strSubCategory = "" and uCase(strAction) <> "DELETE" then
	response.redirect "metcal_admin.asp"
end if

' If form was posted (a save) then process
if strPostFlag = 1 then
  if SaveRecord(strSubCategory, strAction, iCategory_ID, iSubCategory_ID, strDescription, strOracleID, strListPrice) then
	response.redirect "metcal_Categories.asp?subcategory=" & strSubCategory
  end if
end if

' If user wants to delete then process
if uCase(strAction) = "DELETE" then
	if DeleteRecord(iCategory_ID) then
		response.redirect "metcal_categories.asp?subcategory=" & strSubCategory
	end if
else

' --------------------------------------------------------------------------------------
' Start building table to display records for desired subcategory
' --------------------------------------------------------------------------------------
response.write "<FONT CLASS=Heading3>Calibration Procedure Download Admin - " & strSubcategory & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

' Get Recordset
if uCase(strAction) = "EDIT" THEN
	set rsCategory = GetCategory(iCategory_ID)
	iSubCategory_ID = rsCategory("SubCategory_ID")
	strDescription = rsCategory("Description")
	strOracleID = rsCategory("Oracle_ID")
	strListPrice = rsCategory("ListPrice")
else
	strDescription = ""
	iCategory_ID = ""
	iSubCategory_ID = ""
end if

%>

<FORM ACTION="metcal_Category.asp" METHOD="POST">
<input type="hidden" name="category_id" value="<%=iCategory_ID%>">
<input type="hidden" name="action" value="<%=strAction%>">
<input type="hidden" name="subcategory" value="<%=strSubCategory%>">
<input type="hidden" name="subcategory_id" value="<%=iSubCategory_ID%>">
<input type="hidden" name="PostFlag" value="1">
<input type="hidden" name="ListPrice" value="<%=strListPrice%>">
<CENTER>
<% Call Table_Begin %>
<TABLE BORDER=0 WIDTH="100%" BORDERCOLOR="#666666" BGCOLOR="#FFCC00" CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="#FFCC00" WIDTH="100%">
	<TR>
	  <TD CLASS=MediumBold>ID</TD>
    
	  <TD CLASS=MediumBold><%=iCategory_ID%></TD>
        </TR>                
        <TR>
	<TR>
	  <TD CLASS=MediumBold>Description</TD>
	  <TD CLASS=MediumBold><INPUT TYPE=TEXTBOX VALUE="<%=strDescription%>" NAME=Description SIZE=40 MAXLENGTH=50></TD>
	</TR>                
<% if uCase(strSubCategory) = "PRICE POINTS" then %>
	<TR>
	  <TD CLASS=MediumBold>Oracle ID</TD>
	  <TD CLASS=MediumBold><INPUT TYPE=TEXTBOX VALUE="<%=strOracleID%>" NAME=OracleID SIZE=40 MAXLENGTH=50></TD>
    </TR>                
	<TR>
	  <TD CLASS=MediumBold>Price</TD>
	  <TD CLASS=MediumBold><%if IsNumeric(strListPrice) then response.write(FormatCurrency(strListPrice\100))%></TD>
    </TR>                
<% end if %>
	<TR>
          <TD COLSPAN=2 BGCOLOR="BLACK">
            <TABLE WIDTH="100%">
              <TR>
                <TD>
		  <INPUT TYPE="submit" VALUE=" Save " CLASS=NavLeftHighlight1>
		  <INPUT TYPE="button" VALUE=" Cancel " onclick="return document.location='metcal_Categories.asp?subcategory=<%=strSubCategory%>'" CLASS=NavLeftHighlight1>
		</TD>
              </TR>
            </TABLE>
          </TD>   
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
<% Call Table_End %>
</FORM>

</CENTER>
<%
end if
%>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.flush
%>

<%
Function GetCategory(iCategory_ID)
  Dim cmd, prm, rsCategory

  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "Admin_MetCal_Category_Get"

  Set prm = cmd.CreateParameter("@iCategory_ID", adInteger,adParamInput ,, cInt(iCategory_ID) & "")
  cmd.Parameters.Append prm

  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.CursorLocation = adUseClient
  rsCategory.CursorType = adOpenDynamic
  rsCategory.open cmd

  set prm = nothing
  set cmd = nothing

  set GetCategory = rsCategory
End Function

Function SaveRecord(strSubCategory, strAction, iCategory_ID, iSubCategory_ID, strDescription, strOracleID, strListPrice)
'  on error resume next

  Dim cmd, prm, rsCategory

'response.write("strSubCategory: " & strSubCategory & "<BR>")
'response.write("strAction: " & strAction & "<BR>")
'response.write("iID: " & iID & "<BR>")
'response.write("iSubCategory_ID: " & iSubCategory_iID & "<BR>")
'response.write("strDescription: " & strDescription & "<BR>")
'response.write("strOracleID: " & strOracleID & "<BR>")
'response.write("strListPrice: " & strListPrice & "<BR>")

	if iCategory_ID = "" then
		iCategory_ID = 0
	end if
	if iSubCategory_ID = "" then
		iSubCategory_ID = 0
	end if

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Admin_MetCal_Category_Edit"

	Set prm = cmd.CreateParameter("@strSubCategory", adVarchar,adParamInput ,50, uCase(strSubCategory) & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strAction", adVarchar,adParamInput ,50, uCase(strAction) & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iCategory_ID", adInteger,adParamInput ,, cInt(iCategory_ID) & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@iSubCategory_ID", adInteger,adParamInput ,, cInt(iSubCategory_ID) & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strDescription", adVarchar,adParamInput ,50, strDescription & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strOracleID", adVarchar,adParamInput ,50, strOracleID & "")
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("@strListPrice", adVarchar,adParamInput ,50, strListPrice & "")
	cmd.Parameters.Append prm

	Set rsCategory = Server.CreateObject("ADODB.Recordset")
	rsCategory.CursorLocation = adUseClient
	rsCategory.CursorType = adOpenDynamic
	rsCategory.open cmd
	
	set prm = nothing
	set cmd = nothing

	if err.description = "" then
		SaveRecord = true
	else
		SaveRecord = false
	end if
End Function

Function DeleteRecord(iCategory_ID)
  on error resume next
  
  Dim cmd, prm

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "Admin_MetCal_Category_Delete"

	Set prm = cmd.CreateParameter("@iCategory_ID", adInteger, adParamInput, , iCategory_ID & "")
	cmd.Parameters.Append prm

	Set rsSubCategory = Server.CreateObject("ADODB.Recordset")
	rsSubCategory.CursorLocation = adUseClient
	rsSubCategory.CursorType = adOpenDynamic
	rsSubCategory.open cmd
	
	set prm = nothing
	set cmd = nothing

	if err.description = "" then
		DeleteRecord = true
	else
		response.write("<font class=""mediumbold"">There was an error deleting the record.</font><font class=""medium""><BR>Error Number: " & err.number & "<BR>Error Description: " & err.description & "</font><BR><BR>")
		if err.number = "-2147217900" then
			response.write("<font class=""mediumbold""><BR>The record you are trying to delete is currently associated with an active procedure.</font>")
		end if
	end if
End Function
%>