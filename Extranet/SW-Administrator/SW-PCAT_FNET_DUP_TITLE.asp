<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<HTML>
<HEAD>
<TITLE>Duplicate Title </TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
</HEAD>
<%

on error resume next
strmessage="Duplicate title found.<br>Record Id prefixed to title while saving the record.<br>Title modified from " & _
Request.QueryString("DupTitle") & _
" to " & "[" & Request.QueryString("RecId") & "] " & Request.QueryString("DupTitle")

response.write "<HTML>" & vbCrLf
response.write "<HEAD>" & vbCrLf
response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
response.write "<TITLE>Error</TITLE>" & vbCrLf
response.write "</HEAD>"
response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
response.write "<DIV ALIGN=CENTER>"
Call Nav_Border_Begin
response.write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
Response.Write strmessage & "<br><br>"
response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & "javascript:window.close();" & """>Close</A>&nbsp;&nbsp;</SPAN>"
response.write "</TD></TR></TABLE>" & vbCrLf
Call Nav_Border_End
response.write "</DIV>"
response.write "</BODY>"
response.write "</HTML>"
on error goto 0
Response.End

if err.number <> 0 then
Response.Write err.Description
end if
%> 