<%@ Language="VBScript" CODEPAGE="65001" %>

<% 

' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Translation Table Editor
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Show_Date             ' System Date PST
Show_Date = True

Dim Site_ID
Site_ID = 0
Dim Site_Code
Dim Site_Description
Dim Login_Language
Dim Translation_ID

if isblank(request("Language")) then
  Login_Language = "eng"
else
  Login_Language = request("Language")
end if
    
if isblank(request("Language_ID")) then
  Language_ID = "0"
else
  Language_ID = request("Language_ID")
end if

if isblank(request("Translation_ID")) then
  Translation_ID = 0
elseif isnumeric(request("Translation_ID")) then
  Translation_ID = CInt(request("Translation_ID"))
else
  Translation_ID = request("Translation_ID")
end if

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
Call Connect_SiteWide

Site_Description = "Site Wide Translation Table"
Screen_Title    = Site_Description & " - Editor"
Bar_Title       = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>Editor</SPAN>" 
Side_Navigation = False
Content_Width   = 95  ' Percent

%>  
<!--#include virtual="/SW-Common/SW-Header.asp"-->  

<!-- BEGIN CONTENT -->

<%
    response.write "<BR><BR><BR>"
    response.write "<DIV ALIGN=CENTER>"
    response.write "<FORM NAME=""Translations"" METHOD=""POST"">"
    response.write "<TABLE WIDTH=""95%"" BORDER=1>"

    response.write "<TR><TD CLASS=Small WIDTH=""15%"">Select Language: </TD>"
    response.write "<TD CLASS=Small>"
    
    SQL = "SELECT * FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
    Set rsLanguage = Server.CreateObject("ADODB.Recordset")
    rsLanguage.Open SQL, conn, 3, 3
    
    response.write "<SELECT NAME=""Language_ID"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='/SW-Common/Translation_Editor.asp" & "?Translation_ID=" & Translation_ID & "&Group_ID=" & Group_ID & "&Language_ID='+this.options[this.selectedIndex].value"">" & vbCrLf
       
    Do while not rsLanguage.EOF
   	  response.write "<OPTION VALUE=""" & rsLanguage("ID") & """"
      if Language_ID = rsLanguage("ID") then
        response.write " SELECTED"
        Meta_CharSet = rsLanguage("Name_CharSet")
      end if  
      response.write ">" & rsLanguage("Description") & "</OPTION>"              
      rsLanguage.MoveNext 
    loop
    
    rsLanguage.close
    set rsLanguage=nothing
    
    response.write "</TD></TR>"
    response.write "<TD CLASS=Small>Meta CharSet Tag: </TD>"
    
    
    response.write "<TD CLASS=Small>" & Meta_CharSet & "</TD>"
    response.write "</TR>"
    
    
    response.write "</TABLE></DIV><BR>"

    SQL = "SELECT * " &_
          "FROM Translations " &_
          "WHERE Language_ID=0 OR Language_ID=" & Language_ID & " " &_
          "ORDER BY Grouping, Language_ID"

'    response.write SQL & "<BR><BR>"
    
    Set rsTranslate = Server.CreateObject("ADODB.Recordset")
    rsTranslate.Open SQL, conn, 3, 3

    %>
    <DIV ALIGN=CENTER>

    <TABLE WIDTH="95%" BORDER=1 BGCOLOR="Silver">
    <TR>
    <TD WIDTH="100%">
    <TABLE WIDTH="100%" Border=0 CELLPADDING=4>
    
    <%
    counter = 0
    do while not rsTranslate.eof or counter < 100
    
      counter = counter + 1
    
      Eng_String = rsTranslate("Content")
      Eng_Grouping = rsTranslate("Grouping")
      Eng_ID       = rsTranslate("Translation_ID") 
      
      rsTranslate.MoveNext
      if not rsTranslate.EOF then
        if CInt(Eng_Grouping) = CInt(rsTranslate("Grouping")) then
          Edt_String = rsTranslate("Content")
          Edt_ID     = rsTranslate("Translation_ID") 
          
          with response
            .write "<TR>"
            .write "<TD CLASS=Small BGCOLOR=""#FEFEFE"">" & Eng_String & "</TD>"
            .write "</TR>"
            .write "<TD Small BGCOLOR=""#FEFEFE""><BLOCKQUOTE>" & Eng_String & "</BLOCKQUOTE></TD>"
            .write "</TR>"
            .write "<TR>"
            .write "<TD CLASS=Small BGCOLOR=""#FEFEFE"">" & Edt_String & "</TD>"
            .write "</TR>"
          end with
          rsTranslate.MoveNext

        else
          Edt_String = ""
          Edt_ID     = -1
        
          with response
            .write "<TR>"
            .write "<TD CLASS=Small BGCOLOR=""#FEFEFE"">" & Eng_String & "</TD>"
            .write "</TR>"
            .write "<TD BGCOLOR=""#FEFEFE""><BLOCKQUOTE>" & Eng_String & "</BLOCKQUOTE></TD>"
            .write "</TR>"
            .write "<TR>"
            .write "<TD CLASS=Medium BGCOLOR=""#FEFEFE"">" & Edt_String & "</TD>"
            .write "</TR>"
          end with
        end if
      end if  

'    response.write "<TR><TD COLSPAN=2 HEIGHT=2 CLASS=Small>&nbsp;</TD></TR>"
    
    loop
    
    rsTranslate.close
    set rsTranslate = nothing
    
    response.write "</TABLE>"
    response.write "</TR>"
    response.write "</TD>"
    response.write "</DIV>"
    response.write "</FORM>"
    
    Call Disconnect_SiteWide
%>
      
<!-- END CONTENT -->


<%'<!--#include virtual="/SW-Common/SW-Footer.asp"-->%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->




