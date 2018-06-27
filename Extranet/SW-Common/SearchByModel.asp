<%
' Search by Model - Product Databas
%>

<!--#include virtual="/connections/Connection_Products.asp"-->

<%

  Call ConnectProducts

	sBaseURL = "http://www.fluke.com/lookup.asp?footnote=y&logo=y&"
  dim oRS
  dim sSearchString
  dim sSQL
  dim sBaseURL
  dim sURL
  dim aResults

  sSearchString = Trim(Request("searchstring"))
  
  if sSearchString <> "" then
  
    sSQL = "EXEC SearchProducts '" & Replace(sSearchString,"'","''") & "'"
    set oRS = dbConnProducts.Execute(sSQL)

    if oRS.EOF then
      response.Write "<FONT CLASS=MediumBoldRed>" & Translate("No Products were found matching your Search Criteria.",Login_Language,conn) & "</B></FONT>" & vbCrLf
    else			
      response.write "<FONT CLASS=MediumBoldRed>" & Translate("Products Matching",Login_Language,conn) & "</FONT>: " & sSearchString
    end if
    
    response.write "<BR><BR><FONT CLASS=Medium>" & Translate("To respecify your Search Criteria, click on the <FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;Search&nbsp;&nbsp;</FONT> button in the site&acute;s navigation bar.",Login_Language,conn) & "</FONT><BR><BR>" & vbCrLf              

    if not oRS.EOF then
      
      'Fields returned in query
      '
      'ID, Name, Description, MetaDescription, IsFamily, IsProduct, IsOption
      
      oRS.MoveFirst
      aResults = oRS.GetRows
				
      'Matching titles first					

      response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
        
      for i = 0 to Field_Max
        Field_Flag(i) = False
        Field_Data(i) = ""
      next
      
      for j = LBound(aResults,2) to UBound(aResults,2)
      	if InStr(LCase(aResults(1,j)),LCase(sSearchString)) >= 1 then
      		call WriteModelResult(j)
      	end if
      next
      
      'Other results

      for j = LBound(aResults,2) to UBound(aResults,2)
      	if InStr(LCase(aResults(1,j)),LCase(sSearchString)) < 1 then
      		call WriteModelResult(j)
      	end if
      next
    end if
    
    response.write "</TABLE>"
    
    set oRS = Nothing
    
    else
      Response.Write "<FONT CLASS=MediumBoldRed>" & Translate("Please enter a nonempty search string in the Search Criteria area.",Login_Language,conn) & "</FONT>"
    end if
    
    response.write "<BR><BR><BR><BR>"
 
  Call DisconnectProducts
  
' --------------------------------------------------------------------------------------
  
Function TranslateChars(strCopy)

	strOut = strCopy

  strOut = replace(strOut, ";(tm)","<SUP><FONT Size=1>TM</FONT></SUP>")
  strOut = replace(strOut, ";(TM)","<SUP><FONT Size=1>TM</FONT></SUP>")
  strOut = replace(strOut,";(r)","&reg;")
  strOut = replace(strOut,";(R)","&reg;")
  strOut = replace(strOut,";(c)","&copy;")  
  strOut = replace(strOut,";(C)","&copy;")

	TranslateChars = strOut

End Function

' --------------------------------------------------------------------------------------

Sub WriteModelResult(index)

	if aResults(4,index) = 1 Then
		sURL = sBaseURL & "FID=" & aResults(0,index)
	else
		sURL = sBaseURL & "PID=" & aResults(0,index)
	end if
	
	sName = aResults(1,index) & " " & aResults(2,index)
	sDescription = aResults(3,index)

	if Len(sDescription) > 0 then
		if Len(sDescription) >= 255 then
			sDescription = sDescription & "..."
		end if
	end if
  
  

  Field_Data(xID) = "0"
  Field_Flag(xID) = False
    
  Field_Data(xStatus) = "1"
  Field_Flag(xStatus) = True
  
  Field_Data(xTitle) = TranslateChars(sName)
  Field_Flag(xTitle) = True
  
  Field_Data(xDescription) = TranslateChars(sDescription)
  Field_Flag(xDescription) = True

  Field_Data(xLink) = sURL
  Field_Flag(xLink) = True  
  
  Field_Data(xLink_PopUp_Disabled) = "0"
  Field_Flag(xLink_PopUp_Disabled) = True

  Field_Data(xLanguage) = "eng"  
  Field_Flag(xLanguage) = True

  Call Display_Category_Item

End Sub

' --------------------------------------------------------------------------------------
%>

