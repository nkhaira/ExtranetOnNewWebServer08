<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             SW-Manuals.asp
'
' Manual Search Engine developed by Jeff Patrick  (www.fluke.com/products/manuals.asp)
' --------------------------------------------------------------------------------------

response.buffer = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/Connections/Connection_Products.asp"-->

<%

Call Connect_SiteWide
Call ConnectProducts

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

if request("Site_ID") <> "" and isnumeric(request("Site_ID")) then
  Site_ID = request("Site_ID")
elseif session("Site_ID") <> "" and isnumeric(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
  Session.Abandon
  Session("ErrorString") = "<LI>" & Translate("Your session has expired. For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please type in the site's URL in your browser&acute;s address line, then re-enter your User Name and Password, or",Login_Language,conn) & "</LI><LI>" & Translate("Use the Site Search feature below.",Login_Language,conn) & "</LI>"
  response.redirect "/register/default.asp"
end if

Dim BackURL
BackURL = Session("BackURL")    

Set Session("rs") = nothing

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code = rsSite("Site_Code")
Site_Description = rsSite("Site_Description")
Logo = rsSite("Logo")
Footer_Disabled = rsSite("Footer_Disabled")
  
rsSite.close
set rsSite=nothing

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Site_Description & " - " & Translate("Fluke Product Manuals - Search",Alt_Language,conn)
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Fluke Product Manuals - Search",Login_Language,conn) & "</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->

<%
response.write "<FONT CLASS=Heading3>" & Translate("Fluke Product Manuals - Search",Login_Language,conn)
response.write "</FONT><BR><BR>"
%>

<%
' Start creating the page

strPID = ""
strFID = ""
strSection_ID = "ALL"

ManualPageStart



' If we know what product is being requested
	if strPID <> "" AND strPID <> 0 then
		bAddRC = WriteManualTable("PROD", strPID)
' If we know what family is being requested
	elseif strFID <> "" AND strFID <> 0 then
		bAddRC = WriteManualTable("FAM", strFID)
' If user chose to see All the manuals
	elseif strSection_ID <> "" and uCase(strSection_ID) = "ALL" then
		bAddRC = WriteManualTable("ALL", 0)
' If user chose to see all the accessories manuals
	elseif strSection_ID <> "" and uCase(strSection_ID) = "ACCESSORIES" then
		bAddRC = WriteManualTable("ACCESSORIES", 0)
' If user chose to see all discontinued products
	elseif strSection_ID <> "" and uCase(strSection_ID) = "DISCONTINUED" then
		bAddRC = WriteManualTable("DISCONTINUED", 0)
' If user is searching for a manual
	elseif Request("SearchString") <> "" then
		bAddRC = WriteManualTable("SEARCHSTRING", Request("SearchString"))
' We'll need to drill down through the Applicaton groups and Sections
	else
		' If we know what Section is being requested
		if strSection_ID <> "" then
			' Get a list of families and products with associated manuals.
			set dbRS = dbconnProducts.execute("select * from prd_section where section_id=" & strSection_ID)
			Response.Write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>Product Category: " & dbRS("description")&"</b></font><br><br>"
			dbRS.close
			set dbRS = nothing
			if DisplayProductFamilyDropdown(strAppGrp_ID, strSection_ID, strSection_ID) = False then
				Response.Write "There are no manuals associated with any products or families in this section."
			end if
		' If we know what application group is being requested.
		else
			' Get a list of Sections that have familes or products with associated manuals.
			aListOfSections = GetSectionsManuals(strAGID)
			' Check to see if there were any sections returned.
			if aListOfSections(0, 0) = "" then
				' If the passed in AGID doesn't find anything, try once more with 0 - the default.
				aListOfSections = GetSectionsManuals(0)
				if aListOfSections(0, 0) = "" then
					Response.Write "There are no manuals associated with any products, families or sections in this group."
				else
					OutputSections aListOfSections, 0, strAGID, strSID
				end if
			else
				OutputSections aListOfSections, strAGID, strAGID, strSID
			end if
		end if
	end if

' Finish off the bottom navigation

Call DisconnectProducts

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

'****************************************************************************************************
Sub ManualPageStart

	Dim rs
	Dim strCBColor
  
  strAGID=0
		
	set rs = dbConnProducts.Execute("SELECT HeadlineColor FROM prd_ApplicationGroup WHERE ApplicationGroup_ID=" & strAGID)
	strCBColor = rs("HeadlineColor")
	rs.close
	set rs = nothing

	strHeadline = "Manuals test "
	strHeadlineGraphic = "<img src=/images/Prdtmanu.gif>"
'	StartNav strHeadline, "", "", AGID, SID, 0

	ConnectProducts
	GetIDs
	WriteHTMLHead Title, MetaDescription, MetaKeywords

%>
<SCRIPT LANGUAGE="JavaScript">
function DownloadManual(oForm)
{
	var strNewLocation
	var strLanguage
		
	var iManual_Ver_ID = oForm[2].value;
	var strTitle = oForm[3].value;
	var strManualTypeDesc = oForm[4].value;
	var strPrefix = oForm[5].value;
	strPrefix = FilloutValues(strPrefix, 7);
	var strManualType = oForm[6].value;
	strManualType = FilloutValues(strManualType, 2);
	var strRevision = oForm[7].value;
	strRevision = FilloutValues(strRevision, 2);
	var strSupplement = oForm[8].value;
	strSupplement = FilloutValues(strSupplement, 2);
	var strFileExtension = oForm[9].value;
	strFileExtension = "." + FilloutValues(strFileExtension, 3);
	
	for (var iCounter=0; iCounter < oForm[0].length; iCounter++){
		if (oForm[0][iCounter].selected == true){
			strLanguage = oForm[0][iCounter].value;
			strLanguageDescription = oForm[0][iCounter].text;
		}
	}
	strLanguage = FilloutValues(strLanguage, 3);
	strNewLocation = strPrefix + strManualType + strLanguage + strRevision + strSupplement + strFileExtension;
	strValues = "&Manual_Ver_ID=" + iManual_Ver_ID + "&language=" + strLanguageDescription;
	if (strSupplement != "00"){
		strValues = strValues + "&Supplement=true";
	}else{
		strValues = strValues + "&Supplement=false";
	}
	iTop = (window.screen.height - 200) / 2;
	iLeft = (window.screen.width - 400) / 2;

	window.open("manualsdownload.asp?location=" + strNewLocation + strValues, null, "height=200,width=400,status=no,toolbar=no,menubar=no,location=no,top=" + iTop + ",left=" + iLeft);


	return false;
}

function FilloutValues(strSource, iReqLen)
{
	var strRepeatVal = "";
	var strOutput
	
	strOutput = strSource;

	if (strSource.length < iReqLen){
		for(var iCounter = 1; iCounter <= iReqLen - strSource.length; iCounter++){
			strRepeatVal = strRepeatVal + "_";
		}
		strOutput = strOutput + strRepeatVal;
	}
	return strOutput
}
</script>

<%
	if bShowNav then
		Response.Write("<BODY BGCOLOR=#FFFFFF leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 background=""/images/bg.gif"">")
		WriteTopNav 
		WriteLeftNav
	else
		response.write("<BODY BGCOLOR='White' leftmargin=8 topmargin=8 marginwidth=0 marginheight=0>")
	end if

	WriteContent

	if strPID <> "" then
		RegisterProduct strPID
		strHeadlineText = Session("pHeadlineText")
		strHeadlineGraphic = Session("pHeadline")
		strHeadline = Session("FName")
		response.write "<H2><FONT COLOR=#" & Session("CBColor") & ">" & TranslateCopy(strHeadline) & "<BR>Manuals</FONT></H2>"
		Response.Write(Session("pTopNav"))
	elseif strFID <> "" then
		RegisterFamily strFID, false
		strHeadlineText = Session("fHeadlineText")
		strHeadlineGraphic = Session("fHeadline")
		strHeadline = Session("FName")
		response.write "<H2><FONT COLOR=#" & Session("CBColor") & ">" & TranslateCopy(strHeadline) & "<BR>Manuals</FONT></H2>"
		Response.Write(Session("fTopNav"))
	end if
	Response.Write("<BR>")
	
End Sub

'****************************************************************************************************
Function GetSectionsManuals(iAGID)
	Dim aLOS()
	Dim bUsed
		
	bUsed = False
	
	'set rsLOS = dbConnProducts.Execute("SELECT Section_ID FROM prd_ApplicationSection WHERE ApplicationGroup_ID = " & iAGID)

	set dbCmd = Server.CreateObject("ADODB.Command")
	dbCmd.ActiveConnection = dbConnProducts
	dbCmd.CommandType = adCmdStoredProc
	dbCmd.CommandText = "Manual_SearchApplication"
	set tmpParameter = dbCmd.CreateParameter("@VSID", adInteger, adParamInput, , Session("VS_ID"))
	dbCmd.Parameters.Append tmpParameter
	set tmpParameter = dbCmd.CreateParameter("@AGID", adInteger, adParamInput, , iAGID)
	dbCmd.Parameters.Append tmpParameter
	set tmpParameter = dbCmd.CreateParameter("@PEDDAte", adVarChar, adParamInput, 50, now())
	dbCmd.Parameters.Append tmpParameter
	set  rsLOS = dbCmd.execute 
	if not rsLOS.EOF then
		GetSectionsManuals = rsLOS.GetRows
	else
		ReDim tmpArray(2,2)
		GetSectionsManuals = tmpArray
	end if
End Function

'****************************************************************************************************
Function DisplayProductFamilyDropdown(strAppGrp_ID, strSection_ID, iSID)
	Dim bIsFirstTime
	Dim dbCmd
	Dim tmpParameter
	Dim dbRS
	Dim strProdFam_ID
		
	bIsFirstTime = true
	
	set dbCmd = Server.CreateObject("ADODB.Command")
	dbCmd.ActiveConnection = dbConnProducts
	dbCmd.Commandtype = adCmdStoredProc
	dbCmd.CommandText = "Manual_GetProdFam"
	set tmpParameter = dbCmd.CreateParameter("@VS_ID", adInteger, adParamInput, , session("VS_ID"))
	dbCmd.Parameters.Append tmpParameter
	set tmpParameter = dbCmd.CreateParameter("@SID", adInteger, adParamInput, , iSID)
	dbCmd.Parameters.Append tmpParameter
	set tmpParameter = dbCmd.CreateParameter("@PEDDAte", adVarChar, adParamInput, 50, now())
	dbCmd.Parameters.Append tmpParameter
	set dbRS = dbCmd.execute 
	
	if dbRS.EOF then
		DisplayProductFamilyDropdown = false
	else
		DisplayProductFamilyDropdown = true
		Response.Write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>Products</b></font><br>"
		Response.Write "<form action=""/SW-Common/SW-Manuals.asp?SID=" & Request("SID") & "&AGID=" & Request("AGID") & """ method=""post"">"
		Response.Write "<SELECT name=PID_FID>"

		do while not dbRS.EOF
			if uCase(dbRS("Type")) = "ISPRODUCT" then
				strProdFam_ID = "p" & dbRS("family_id")
			else
				strProdFam_ID = "f" & dbRS("family_ID")
			end if
			Response.Write "<OPTION value=""" & strProdFam_ID & """>" & TranslateCopy(dbRS("ProdFam_Name"))
			dbRS.MoveNext
		loop

		Response.Write "</SELECT>"
		Response.Write "<input type=hidden name=strAppGrp_ID value=""" & strAppGrp_ID & """>"
		Response.Write "<input type=hidden name=strSection_ID value=""" & strSection_ID & """>"
		Response.Write"&nbsp;&nbsp;<input type=""submit"" name=""Go"" value=""Go"">"
		Response.Write "</form>"
		Response.Write "<P>"
	end if
End Function

'***************************************************************************************************
Function WriteManualTable(sType, strValue)
	on error resume next
	
	Dim strModel_Code, strTitle, strManualType, strManualTypeDesc, strRevision
	Dim iManualID
	Dim iSupplementID
	Dim strDownload, strFileExtention, strSupplement
	Dim rsManuals
	Dim aFoundLanguages
	Dim bFoundManual
	Dim iCounter
	Dim iLastManualID
	Dim iLastSupplementID					
	Dim rsFamily
	
	' Get a recordset of manuals for this product or family
	set rsManuals = GetManuals(sType, strValue)
	if rsManuals.EOF and sType = "PROD" then
		set rsFamily = GetProductFamily(strValue)
		if not rsFamily.EOF then
			sType = "FAM"
			strValue = rsFamily("Family_ID")
			set rsManuals = GetManuals("FAM", strValue)
		end if
		rsFamily.close
		set rsFamily = nothing
	end if
	
	' Print product or famlily descriptions as sub headline
'	if sType = "PROD" then
'		Response.write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>Product: " & TranslateCopy(rsManuals("ProdFamDescription")) & "</b></font><P>"
'	elseif sType = "FAM" then
'		Response.write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>Family: " & TranslateCopy(rsManuals("ProdFamDescription")) & "</b></font><P>"
	if sType = "ALL" then
		Response.write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>All Manuals</b></font><P>"
	elseif sType = "ACCESSORIES" then
		Response.write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>All Accessories</b></font><P>"
	elseif sType = "DISCONTINUED" OR sType = "SEARCHSTRING" then
		Response.write "<FONT FACE=""Verdana, Arial, Helvetica"" SIZE=3><b>All Manuals matching your search</b></font><P>"
	end if
	
%>
		<font face="arial, helvetica" size=2>
		<table width = "650px" bordercolor=#000000 bgcolor=#EEEEEE border=1 cellpadding=4 cellspacing=0>
			<tr>
				<td valign=top align=left BGCOLOR=#000000 width = "275px"><font FACE="Verdana, Arial, Helvetica" Font color=#FFCC00 SIZE="2">Title</td>
				<td valign=top align=left BGCOLOR=#000000 width = "225px"><font FACE="Verdana, Arial, Helvetica" Font color=#FFCC00 SIZE="2">Type</td>
				<td valign=top align=left BGCOLOR=#000000 width = "100px"><font FACE="Verdana, Arial, Helvetica" Font color=#FFCC00 SIZE="2">Language</td>
				<td valign=top align=left BGCOLOR=#000000 width = "50px"><font FACE="Verdana, Arial, Helvetica" Font color=#FFCC00 SIZE="2">Download</td>
			</tr>
<%
			' For each manual print out info formatted
			iCounter = 0
			iLastManualID = 0
			iLastSupplementID = 0
			
			if IsObject(rsManuals) then
				do while NOT rsManuals.EOF
					iManualID = rsManuals("Manual_ID")
					iManual_Ver_ID = rsManuals("Manual_Ver_ID")
					iSupplementID = rsManuals("Supplement_N")
					strModel_Code = FilloutValues(TranslateCopy(rsManuals("Model_Code")), 7)
					strTitle = TranslateCopy(rsManuals("Title"))
					strManualType = TranslateCopy(rsManuals("ManualType"))
					strManualTypeDesc = TranslateCopy(rsManuals("ManualTypeDesc"))
					strRevision = TranslateCopy(rsManuals("Revision_N"))
					strSupplement = TranslateCopy(rsManuals("Supplement_N"))
					strFileExtention = TranslateCopy(rsManuals("FileExtention"))
					' Get the various languages this manual is printed available in
					aLanguages = GetLanguages(rsManuals("Manual_ID"), rsManuals("Manual_Ver_ID"))
					strDownload = "<input name=""Download"" type=submit value=""Download"">"

					strManualName = strModel_Code & strManualType & "LANGUAGE" & strRevision & strSupplement & "." & strFileExtention
					' Now, make sure that these physical manuals actually exist on the server
					aFoundLanguages = FindManualFiles(aLanguages, strManualName)

					' If we found manuals then we can print them
					if IsArray(aFoundLanguages) then
						bFoundManual = true
						strSupplementNote = ""
%>
						<form action="/SW-Common/SW-Manuals.asp" name="ManualDownload<%=iCounter%>" onsubmit="return DownloadManual(this);">
							<%=strSupplementNote%>
							<tr>
								<td><FONT FACE=""Verdana, Arial, Helvetica"" SIZE=2>
									<% 
									if iManualID = iLastManualID then
										Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
									end if
									Response.Write strTitle
									%>
									</font>
								</td>
								<td><FONT FACE=""Verdana, Arial, Helvetica"" SIZE=2>
									<%
									'if iManualID = iLastManualID then
									if cInt(iSupplementID) > 0 then
										strManualTypeDesc = strManualTypeDesc & " Supplement"
									end if
									Response.Write strManualTypeDesc
									%></font>
								</td>
								<td><FONT FACE=""Verdana, Arial, Helvetica"" SIZE=2><%=TranslateLanguages(aFoundLanguages)%></font></td>
								<td><FONT FACE=""Verdana, Arial, Helvetica"" SIZE=2><%=strDownload%></font>
									<input type="hidden" name="Manual_Ver_ID" value="<%=iManual_Ver_ID%>">
									<input type="hidden" name="Title" value="<%=strTitle%>">
									<input type="hidden" name="ManualTypeDesc" value="<%=strManualTypeDesc%>">
									<input type="hidden" name="Prefix" value="<%=strModel_Code%>">
									<input type="hidden" name="ManualType" value="<%=strManualType%>">
									<input type="hidden" name="Revision" value="<%=strRevision%>">
									<input type="hidden" name="Supplement" value="<%=strSupplement%>">
									<input type="hidden" name="FileExtension" value="<%=strFileExtention%>">
									<input type="hidden" name="DownloadFile" value="Download">
								</td>
							</tr>
						</form>
<%
					end if
					iCounter = iCounter + 1
					iLastManualID = iManualID
					iLastSupplementID = iSupplementID
					rsManuals.MoveNext
				loop
			end if
			rsManuals.close
			set rsManuals = Nothing
%>
	</table>
<%			if not bFoundManual then %>
				<br>Sorry, no manuals were located for this product
<%			end if %>

</font>
<%
End Function
'***************************************************************************************************

Function GetManuals(sType, strValue)
	on error resume next
	
	Dim strStoredProc
	Dim strInputParameter
	Dim dbCmd
	Dim rsManual
	
	sType = uCase(sType)

	if sType = "PROD" then
		strStoredProc = "Manual_GetProductManuals"
		strInputParameter = "@Product_ID"
	elseif sType = "FAM" then
		strStoredProc = "Manual_GetFamilyManuals"
		strInputParameter = "@Family_ID"
	' Assume we're getting all manuals
	elseif sType = "ALL" then
		strStoredProc = "Manual_GetAllManuals"
	elseif sType = "ACCESSORIES" then
		strStoredProc = "Manual_GetAllAccessories"
	elseif sType = "DISCONTINUED" then
		strStoredProc = "Manual_GetDiscontinued"
	else
		strStoredProc = "Manuals_SearchManuals"
		strInputParameter = "@searchstring"
	end if

	set dbCmd = Server.CreateObject("ADODB.Command")
	dbCmd.ActiveConnection = dbConnProducts
	dbCmd.CommandType = adCmdStoredProc
	dbCmd.CommandText = strStoredProc
	set tmpParameter = dbCmd.CreateParameter("@VS_ID", adInteger, adParamInput, , Session("VS_ID"))
	dbCmd.Parameters.Append tmpParameter
	if sType = "PROD" or sType = "FAM" then
		set tmpParameter = dbCmd.CreateParameter(strInputParameter, adInteger, adParamInput, , strValue)
		dbCmd.Parameters.Append tmpParameter
	elseif sType = "SEARCHSTRING" then
		set tmpParameter = dbCmd.CreateParameter(strInputParameter, adVarChar, adParamInput, 50, strValue)
		dbCmd.Parameters.Append tmpParameter
	end if
	
	set tmpParameter = dbCmd.CreateParameter("@PEDDate", adVarChar, adParamInput, 50, now())
	dbCmd.Parameters.Append tmpParameter
	set rsManual = dbCmd.execute
	set dbCmd = nothing
	set GetManuals = rsManual
End Function

Function GetProductFamily(iID)
on error resume next
	Dim rsFamily
	Dim dbCmd
	Dim tmpParameter
	
	set dbCmd = Server.CreateObject("ADODB.Command")
	dbCmd.ActiveConnection = dbConnProducts
	dbCmd.CommandType = adCmdStoredProc
	dbCmd.CommandText = "Manual_GetProductFamily"
	set tmpParameter = dbCmd.CreateParameter("@VS_ID", adInteger, adParamInput, , Session("VS_ID"))
	dbCmd.Parameters.Append tmpParameter
	set tmpParameter = dbCmd.CreateParameter("@Product_ID", adInteger, adParamInput, , iID)
	dbCmd.Parameters.Append tmpParameter
	set rsFamily = dbCmd.execute
	set dbCmd = nothing
	set GetProductFamily = rsFamily
End Function

Function GetLanguages(strManual_ID, strManual_Ver_ID)
	Dim bdCmd
	Dim rsLanguages
	Dim strLanguages
	
	set dbCmd = Server.CreateObject("ADODB.Command")
	dbCmd.ActiveConnection = dbConnProducts
	dbCmd.CommandType = adCmdStoredProc
	dbCmd.CommandText = "Manual_GetLanguages"
	set tmpParameter = dbCmd.CreateParameter("@Manual_ID", adInteger, adParamInput, , strManual_ID)
	dbCmd.Parameters.Append tmpParameter
	set tmpParameter = dbCmd.CreateParameter("@Manual_Ver", adInteger, adParamInput, , strManual_Ver_ID)
	dbCmd.Parameters.Append tmpParameter
	set rsLanguages = dbCmd.execute
	set dbCmd = nothing
	
	GetLanguages = rsLanguages.GetRows

	set dbCmd = nothing
	set rsLanguages = nothing
	
End Function

Function TranslateLanguages(aLanguages)
	Dim strLanguages
	Dim iCounter
	
	strLanguages = "<select name=""ManualLanguage_Code"">"

	for iCounter = 0 to uBound(aLanguages,2)
		strLanguages = strLanguages & "<option value=""" & Trim(aLanguages(0, iCounter)) & """>" & aLanguages(1, iCounter)
	next
	strLanguages = strLanguages & "</select>"
	
	TranslateLanguages= strLanguages
End Function

Function FindManualFiles(aLanguages, strManualName)
	on error resume next
	
	Dim aFoundLanguages
	Dim iCounter
	Dim oFileSys
	Dim strPath
	Dim BACKSLASH
	Dim iMaxFileSize
	Dim strManualName_Original
	
	strManualName_Original = strManualName	
	BACKSLASH = chr(92)	
	ReDim aFoundLanguages(1,0)
				
	for iCounter = 0 to uBound(aLanguages, 2)
		' Get the language code and make it part of the manual name
		strLanguage = aLanguages(2, iCounter)
		strManualName = replace(strManualName_Original, "LANGUAGE", strLanguage)
		set oFileSys = Server.CreateObject("Scripting.FileSystemObject")
		strPath = Request.ServerVariables("PATH_TRANSLATED")
		strPath = left(strPath, InstrRev(strPath, "products") - 1)
		strPath = strPath & "download\manuals" & BACKSLASH

		set oFile = oFileSys.GetFile(strPath & strManualName)

		if uCase(Trim(err.description)) <> "FILE NOT FOUND" then
			if not bFoundManual then
				ReDim aFoundLanguages(1, 0)
			else
				iMax = uBound(aFoundLanguages, 2) + 1
				ReDim Preserve aFoundLanguages(1, iMax)
			end if
			bFoundManual = true
			
			if uCase(strLanguage) = "ENG" then
				g_strFileSize = oFile.size
			end if
			aFoundLanguages(0, uBound(aFoundLanguages, 2)) = strLanguage
			' Get the actual name of the language
			aFoundLanguages(1, uBound(aFoundLanguages,2)) = aLanguages(0, iCounter)
		end if
		err.clear
	next

	g_strFileSize = cInt(g_strFileSize / 1000) & "KB"

	set oFileSys = nothing
	set oFile = nothing
	
	if bFoundManual= true then
		FindManualFiles = aFoundLanguages
	else
		FindManualFiles = ""
	end if
End Function

Function FilloutValues(strSource, iReqLen)
	Dim iCounter
	Dim strRepeatVal
	
	if len(strSource) < iReqLen then
		for iCounter = 1 to iReqLen - len(strSource)
			strRepeatVal = strRepeatVal & "_"
		next
		strSource = strSource & strRepeatVal
	end if
	FilloutValues = strSource
End Function

Function SendUserToDownload
	Dim strNewURL
	Dim strPrefix
	Dim strManualType
	Dim strRevision
	Dim strSupplement
	Dim strFileExtension
	Dim strDownloadFile
	Dim strLanguage

	strPrefix = FilloutValues(Request("Prefix"), 8)
	strManualType = FilloutValues(Request("ManualType"), 2)
	strRevision = FilloutValues(Request("Revision"), 2)
	strSupplement = FilloutValues(Request("Supplement"), 2)
	strFileExtension = FilloutValues(Request("FileExtension"), 3)
	strLanguage = FilloutValues(Request("ManualLanguage_Code"), 3)
	strNewURL = "ftp://" & Request.ServerVariables("SERVER_NAME") & "/download/manuals/"
	strNewURL = strNewURL & strPrefix & strManualType & strLanguage & strRevision & strSupplement & "." & strFileExtension
	response.redirect strNewURL
End Function

%>
