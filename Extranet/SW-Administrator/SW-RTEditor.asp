<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   July 1, 2003
' Rich Text Editor
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->  
<!--#Include virtual="/include/RTEditor/class.devedit.asp"-->
<%
' --------------------------------------------------------------------------------------

Dim Border_Toggle, Length, Rows, Cols
Dim Action, Max_Col, Max_Row, Max_Tot
Dim Site_ID, Site_Code, Style_Sheet

Length = 50

Border_Toggle = 0

if 1=2 then
for each item in request.querystring
  response.write item & ": " & request.querystring(item) & "<BR>"
next
response.write "<P>---------<P>"
for each item in request.form
  response.write item & ": " & request.form(item) & "<BR>"
next  
end if  

'response.end

Dim RTE
Set RTE = new DevEdit
RTE.SetName "RTE"

Dim Form, Element, Content, Fetch
if not isblank(request.form("Fetch")) then
  Fetch = request.form("Fetch")
else
  Fetch = 0
end if  

if not isblank(request.form("Form")) and not isblank(request.form("Element")) and not isblank(RTE.GetValue(false)) then
  Content = RTE.GetValue(false)
  Content = Replace(Content,"'","\'")
  Content = Replace(Content,vbCrLf,"")
  Content = Replace(Content,"     "," ")
  Content = Replace(Content,"   "," ")
  Content = Replace(Content,"  "," ")
  Content = Replace(Content,"<FONT size=+0>","<FONT color=#FFFFFF>")
  %>
  <SCRIPT LANGUAGE="JAVASCRIPT">
    window.opener.<%=request.form("Form")%>.<%=request.form("Element")%>.value='<%=Content%>';
    self.close();
    window.opener.focus();
  </SCRIPT>
  <%  
elseif not isblank(request("Form")) and not isblank(request("Element")) and isblank(request.form("Content")) and Fetch = 0 then

  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<TITLE>SiteWide - Rich Text Editor Refresh</TITLE>" & vbCrLf
  response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">" & vbCrLf
  response.write "</HEAD>" & vbCrLf
  response.write "<BODY BGCOLOR=""White"" onLoad='document." & request("Form") & ".submit()'>" & vbCrLf
  response.write "<FORM NAME=""" & request("Form") & """ ACTION=""/SW-Administrator/SW-RTEditor.asp"" METHOD=""POST"">" & vbCrLf  
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Form"" VALUE=""" & request("Form") & """>" & vbCrLf
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Element"" VALUE=""" & request("Element") & """>" & vbCrLf
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Content"" VALUE="""">" & vbCrLf
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Fetch"" VALUE=""" & Fetch + 1 & """>" & vbCrLf                      
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & request("Site_ID") & """>" & vbCrLf                  
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_Code"" VALUE=""" & request("Site_Code") & """>" & vbCrLf          
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Length"" VALUE=""" & request("Length") & """>" & vbCrLf                  
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Cols"" VALUE=""" & request("Cols") & """>" & vbCrLf                    
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Rows"" VALUE=""" & request("Rows") & """>" & vbCrLf                  
  response.write "</FORM>" & vbCrLf
  %>
  <SCRIPT Language = "JavaScript">
      var Content = opener.document.<%=request("Form")%>.<%=request("Element")%>.value;
      document.<%=request("Form")%>.Content.value = Content;
      document.<%=request("Form")%>.submit();
  </SCRIPT>  
  <%  
  response.write "</BODY>" & vbCrLf
  response.write "</HTML>" & vbCrLf
    
else    

  Site_ID   = request.form("Site_ID")
  Site_Code = request.form("Site_Code")
  
  Style_Sheet = "/SW-Common/SW-Style_RTE_Editor.css"
  if not isblank(Site_Code) then
    if FileExists(Site_Code & "\" & "SW-Style.css") then %><%
      Style_Sheet = "/" & Site_Code & "/SW-Style.css"
    end if
  end if    

  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<TITLE>SiteWide - Rich Text Editor</TITLE>" & vbCrLf
  response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">" & vbCrLf
  
  response.write "<LINK REL=STYLESHEET HREF=""" & Replace(Style_Sheet,"SW-Style_RTE_Editor.css","SW-Style.css") & """>" & vbCrLf
  
  response.write "</HEAD>" & vbCrLf
  response.write "<BODY BGCOLOR=""White"">" & vbCrLf
  
  if not isblank(request("Action")) then Action  = LCase(request("Action"))else Action = "general"
  if not isblank(request("Cols"))   then Max_Col = request("Cols") else Max_Col = 42
  if not isblank(request("Rows"))   then Max_Row = request("Rows") else Max_Row = 6
  if not isblank(request("Length")) then Length  = request("Length") else Length = 4000

  Form    = request.form("Form")
  Element = request.form("Element")
  Content = Replace(request.form("Content"),vbCrLf,"")

  Max_Tot = Max_Col * Max_Row

  ' --------------------------------------------------------------------------------------
  ' Main
  ' --------------------------------------------------------------------------------------

  with response

  Call Configure_RTE_Control(RTE)
  RTE.SetTextAreaDimensions Max_Col, Max_Row

  .write "<FORM NAME=""RTE"" METHOD=""POST"" ONSUBMIT=""return(CheckLength());"">"
  .write "<INPUT TYPE=""HIDDEN"" NAME=""Form"" VALUE=""" & Form & """>" & vbCrLf
  .write "<INPUT TYPE=""HIDDEN"" NAME=""Element"" VALUE=""" & Element & """>" & vbCrLf
  .write "<INPUT TYPE=""HIDDEN"" NAME=""Length"" VALUE=""" & Length & """>" & vbCrLf                  
  .write "<DIV ALIGN=CENTER>"

  if not isblank(Site_Code) then
    RTE.SetSnippetStyleSheet(Style_Sheet)    
    RTE.SetValue Content
    RTE.ShowControl "760", "400", "/" & Site_Code & "/download/thumbnail"
  else  
    RTE.SetValue Content
    RTE.ShowControl "760", "400", ""
  end if

  .write "<TABLE BORDER=0 ALIGN=CENTER CELLPADDING=10>"
  .write "<TR><TD>"
  Call Nav_Border_Begin    
  .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""Update Source"">"
  Call Nav_Border_End
  .write "</TD><TD>"
  Call Nav_Border_Begin
  .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Close"" VALUE=""Close HTML Editor"" ONCLICK='self.close(); window.opener.focus();'>"
  Call Nav_Border_End
  .write "</TD></TR>"
  .write "</TABLE>"            
  .write "<P>"
 
  .write "</DIV>"
  .write "</FORM>" & vbCrLf        

  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf

  end with
  
end if  
  
Function Configure_RTE_Control(myControl)
  
  set RTE_Control = myControl
  
  ' Configure Control Parameters here
    
  RTE_Control.SetLanguage(de_AMERICAN)
  RTE_Control.SetFontList("Arial,Verdana")
  RTE_Control.SetFontSizeList("1,2,3,4,5,6")
  'RTE_Control.SetImageDisplayType de_IMAGE_TYPE_THUMBNAIL
  RTE_Control.SetDocumentType de_DOC_TYPE_SNIPPET
 'RTE_Control.SetPathType de_PATH_ABSOLUTE
  
  RTE_Control.DisableXHTMLFormatting
 'RTE_Control.DisableSingleLineReturn

 'RTE_Control.DisableSourceMode
 'RTE_Control.DisablePreviewMode

 'RTE_Control.HideFullScreenButton
 'RTE_Control.HideSpellingButton
 'RTE_Control.HideRemoveTextFormattingButton
 'RTE_Control.HideBoldButton
 'RTE_Control.HideUnderlineButton
 'RTE_Control.HideItalicButton
  RTE_Control.HideStrikethroughButton
  
 'RTE_Control.HideNumberListButton
 'RTE_Control.HideBulletListButton
  
 'RTE_Control.HideDecreaseIndentButton
 'RTE_Control.HideIncreaseIndentButton
  
 'RTE_Control.HideLeftAlignButton
 'RTE_Control.HideCenterAlignButton
 'RTE_Control.HideRightAlignButton
 'RTE_Control.HideJustifyButton
  
 'RTE_Control.HideSuperScriptButton
 'RTE_Control.HideSubScriptButton
  
  RTE_Control.HideTextBoxButton
  
 'RTE_Control.HideHorizontalRuleButton
  
 'RTE_Control.HideLinkButton
 'RTE_Control.HideMailLinkButton
  RTE_Control.HideAnchorButton
  
 'RTE_Control.HideHelpButton
  
 'RTE_Control.HideFontList
 'RTE_Control.HideSizeList
  RTE_Control.HideFormatList
 'RTE_Control.HideStyleList
  
 'RTE_Control.HideForeColorButton
 'RTE_Control.HideBackColorButton

 'RTE_Control.HideTableButton
  RTE_Control.HideFormButton

 'RTE_Control.HideImageButton
  RTE_Control.DisableImageDeleting
 'RTE_Control.DisableImageUploading

 'RTE_Control.HideSymbolButton
 'RTE_Control.HidePropertiesButton
 'RTE_Control.HideCleanHTMLButton
 'RTE_Control.HideAbsolutePositionButton  
 'RTE_Control.HideGuidelinesButton
  RTE_Control.EnableGuidelines
 
 'RTE_Control.HideCopyButton
 'RTE_Control.HideCutButton
 'RTE_Control.HidePasteButton
 'RTE_Control.HideUndoButton
 'RTE_Control.HideRedoButton
 RTE_Control.HideFindButton   
    
end Function

'--------------------------------------------------------------------------------------

%>
<SCRIPT LANGUAGE="JavaScript">
function CheckLength() {
  var me = RTE_frame.foo.document.body.innerText;
  if (me.length  >= <%=Length%>) {
    var me_error = "Attention!\r\rThe length of your document containing text and 'HTML tag' characters is: " + me.length + ".\r\rThis has exceeded the maximum length limit of: " + <%=Length%> + " of text or 'HTML tag' characters.\r\rPlease edit your document to reduce the character count before [Updating Source]."
    alert(me_error);
    return false;
  }
  return true;
}
</SCRIPT>

