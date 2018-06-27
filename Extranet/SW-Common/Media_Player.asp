<%
  ' Author:   K. David Whitlock
  ' Date:     02/01/2000
  ' Title:    TV Cabinet and EMBEDDING for MS Media Player
  
  %>
  <!--#include virtual="/include/functions_string.asp"-->
  <%
  
  Dim Media_File
  Dim Media_Title
  Dim Media_Description
  Dim Media_Type
  
  if not isblank(request("Media_File")) then
    Media_File        = request("Media_File")
    Media_Title       = request("Media_Title")
    Media_Description = request("Media_Description")
    Media_Type        = UCase(Mid(Media_File, InstrRev(Media_File, ".") + 1))
  end if

  response.write "<HTML>"
  response.write "<HEAD>"
  response.write "<TITLE>"
  if not isblank(Media_Title) then response.write Media_Title else response.write "Fluke Video"
  response.write "</TITLE>"
  response.write "<LINK REL=STYLESHEET HREF=""/include/SW-Style.css"">"
  response.write "</HEAD>"  
  response.write "<BODY BGCOLOR=""white"" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 LINK=""#000000"" VLINK=""#000000"" ALINK=""#000000"">"

  select case Media_Type

    case "ASF", "MP3", "WAV"

      response.write "<TABLE WIDTH=""220"" BORDER=0 CELLPADDING=0 CELLSPACING=0>" & vbCRLF
      response.write "  <TR>" & vbCRLF
      response.write "    <TD WIDTH=220 VALIGN=TOP CLASS=Normal>" & vbCRLF
      response.write "<IMG SRC=""/images/tv_top.jpg"" width=220 height=10><BR>"
      response.write "<IMG SRC=""/images/tv_left.jpg"" width=10 height=200>"
      if not isblank(Media_File) then
        response.write "<EMBED SRC=""" & Media_File & """ HEIGHT=200 WIDTH=200 TITLE=""" & Media_Title & """ AUTOSTART=True MASTERSOUND>"
      else
        response.write "<IMG SRC=""/images/TV.jpg"" WIDTH=200 HEIGHT=200>"
      end if
      response.write "<IMG SRC=""/images/TV_Right.jpg"" width=10 height=200><BR>"
      if not isblank(Media_File) then
        response.write "<IMG SRC=""/images/TV_Bottom.jpg"" width=220 height=49><BR>"
        if not isblank(Media_Title) then
          response.write "<BR><B>" & Media_Title & "</B><BR>"
        end if
        if not isblank(Media_Description) then
          response.write "<BR>" & Media_Description
        end if  
      else
        response.write "<IMG SRC=""images/TV_Bottom_Off.jpg"" width=220 height=49>"
      end if        
    
      response.write "    </TD>" & vbCRLF
      response.write "  </TR>" & vbCRLF
      response.write "</TABLE>" & vbCRLF
        
    case "SVJ", "SVO"
  
    case else
      Response.write "The Media File: " & Media_File & " is not supported by this viewer."
  end select

  response.write "</BODY>"
  response.write "</HTML>"
  
%>
