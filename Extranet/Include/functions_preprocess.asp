<%
' --------------------------------------------------------------------------------------
' Author: K. D. Whitlock
' Date:   June 1, 2000
'
' Uses Translation Table in SiteWide DB
' This routine only searches any paired occurance of the ## begin and end translation tags
' --------------------------------------------------------------------------------------

Function Translate_Embedded(MyString, Login_Language, conn)

  TempString = MyString
  
  do while instr(1,TempString,"##") > 0

    PreIndex = instr(1,TempString,"##")
    EndIndex = instr(PreIndex + 1,TempString,"##")
  
    if PreIndex > 0 and EndIndex > 0 then

      PreTranslate = Mid(TempString,1,PreIndex - 1)   
      MidTranslate = Mid(TempString,PreIndex + 2, (EndIndex - PreIndex) - 2)
      EndTranslate = Mid(TempString,EndIndex + 2)

      TempString = PreTranslate & Translate(MidTranslate, Login_Language,conn) & EndTranslate

    else
      TempString = TempString
      exit do
    end if
      
  loop
    
  Translate_Embedded = TempString
  
End Function

' --------------------------------------------------------------------------------------
' Translate_Include()
' This function searches for ][Name of Include File][ and does the substitution of an
' include file.  There has got to be a better more robust way to do this, but I haven't
' thought about that yet.  Clugie but it works.
' --------------------------------------------------------------------------------------

Function Translate_Include(TempString)

  if not isnull(TempString) then
    PreIndex = instr(1,TempString,"][")
    EndIndex = instr(PreIndex + 1,TempString,"][")
    
    if PreIndex > 0 and EndIndex > 0 then
      PreTranslate = Mid(TempString,1,PreIndex - 1)   
      MidTranslate = Mid(TempString,PreIndex + 2, (EndIndex - PreIndex) - 2)
      EndTranslate = Mid(TempString,EndIndex + 2)
        
      response.write PreTranslate
      
      select case LCase(MidTranslate)
        case "headlines"
          %>
          <!--#include virtual="/sw-common/sw-headlines.asp"-->
          <%
        case else  
          ' No Match -- Do not write out delimeters or key word
      end select
      
      response.write EndTranslate
      
    else  
      
      response.write TempString
       
    end if

  else  
      
    response.write TempString
    
  end if
  
End Function

' --------------------------------------------------------------------------------------
%>