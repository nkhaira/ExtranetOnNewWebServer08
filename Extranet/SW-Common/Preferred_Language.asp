<%

' Author: K. D. Whitlock
' Sets Preferred Language code. Alt is used for areas like window title bars that do not display DBCS
' 02/01/2000

Dim Login_Language
Dim Alt_Language
Dim Language_Filter

if not isblank(request.QueryString("Language")) and request.QueryString("Language") <> "XON" and request.QueryString("Language") <> "XOF" then
  Login_Language = LCase(request.QueryString("Language"))
  Session("Language") = Login_Language
elseif not isblank(Session("Language")) then
  Login_Language = Session("Language")
elseif not isblank(Login_Language) then
  Session("Language") = Login_Language
elseif isblank(Login_Language) then
  Login_Language = "eng"
end if

' Used with Quick Find Results to filter language independant of Login_Language setting.

if not isblank(request.QueryString("Language_Filter")) then
  Language_Filter = LCase(request.QueryString("Language_Filter"))
else
  Language_Filter = ""
end if  

' Alternate Language is set to ENG for UTF-8. This parameter is typically used for
' Window Headers and Alert Boxes that cannot show UTF-8 extended character sets without special configuration settings.

select case LCase(Login_Language)
  case "chi", "zho", "thi", "jpn", "kor"
    Alt_Language = "eng"
  case else
    Alt_Language = LCase(Login_Language)
end select      

%>