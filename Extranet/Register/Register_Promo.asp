<%@ Language="VBScript" CODEPAGE="65001" %>

<%

' --------------------------------------------------------------------------------------
' Author: K. Whitlock
' Date:   11/12/2000
' Title:  Register Promotion Auto-submit/link
' --------------------------------------------------------------------------------------

Dim Reg_Fields
Dim Reg_Fields_Max
Dim Session_Reg_Fields
Dim Reg_Values
Dim Reg_Values_Max
Dim Session_Reg_Values
Dim Site_Reg_Fields

Site_Reg_Fields     = request("Site_Reg_Fields")
on error resume next
Session_Reg_Fields  = Session("Reg_Fields")       ' Set by Register_Admin.asp
on error resume next
Session_Reg_Values  = Session("Reg_Values")       ' Set by Register_Admin.asp
on error goto 0

if not isnull(Session_Reg_Fields) and Session_Reg_Fields <> "" then
  
  Reg_Fields     = Split(Session_Reg_Fields,",")
  Reg_Fields_Max = UBound(Reg_Fields)
  Reg_Values     = Split(Session_Reg_Values,",")
  Reg_Values_Max = Ubound(Reg_Values)
  
  with response
    .write "<HTML>"
    .write "<HEAD>"
    .write "<TITLE></TITLE>"
    .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">"
    .write "</HEAD>"
    .write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>"
    .write "<FORM ACTION=""" & request("Promotion_URL") & """ METHOD=""POST"">"
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & request("Site_ID") & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_Code"" VALUE=""" & request("Site_Code") & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_Description"" VALUE=""" & request("Site_Description") & """>" & VbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Promotion_Complete_URL"" VALUE=""" & request("Promotion_Complete_URL") & """>" & VbCrLf
  end with

  for i = 0 to Reg_Fields_Max
    if instr(1,LCase(Site_Reg_Fields),LCase(Reg_Fields(i))) > 0 then
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & Reg_Fields(i) & """ VALUE=""" & Reg_Values(i) & """>" & vbCrLf
    end if  
  next
  
  with response
    .write "</FORM>"
    .write "</BODY>"
    .write "</HTML>  "
  end with
  
else

  response.redirect "/register/Default.asp"
  
end if

set Session("Reg_Fields") = nothing
set Session("Reg_Values") = nothing


%>