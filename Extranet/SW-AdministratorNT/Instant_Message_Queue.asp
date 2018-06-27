<HTML>
<HEAD>
<TITLE>Instant Message Queue</TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
</HEAD>

<BODY BGCOLOR="White">

<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
Call Connect_SiteWide

NTLogin        = request("NTLogin")
Login_Language = request("Language")

SQL       = "SELECT * FROM Messages WHERE NTLogin='" & NTLogin & "' ORDER BY Message_Date"
Set rsMessage = Server.CreateObject("ADODB.Recordset")
rsMessage.Open SQL, conn, 3, 3
  
response.write "<SPAN CLASS=HEADING4>" & "Message Queue" & "</SPAN><BR>"
  
response.write "<TABLE BGCOLOR=Black BORDER=0 CELLPADDING=4 WIDTH=""100%"">"
  
do while not rsMessage.EOF

    response.write "<TR>"
    response.write "<TD CLASS=Small BGCOLOR=""White"" WIDTH=""100%"">"
    response.write "Message To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & rsMessage("To_tName") & "<BR>"
    response.write "Message From:&nbsp;&nbsp;" & rsMessage("Fm_Name") & "<BR>"
    response.write "Message Date:&nbsp;&nbsp;" & rsMessage("Message_Date") & "<BR>"
    response.write "Status:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    select case CInt(rsMessage("Status"))
      case 0
        response.write "<SPAN CLASS=SmallBoldRed>Not Read</SPAN>"
      case -1
        response.write "Read by user on: "        
    end select
    response.write "<P>"    
      
    response.write "Message: " & rsMessage("Message") & "<BR>"
    response.write "</TD>"
    response.write "</TR>"
    
    rsMessage.MoveNext
    
loop

rsMessage.close
set rsMessage = nothing
  
response.write "</TABLE>"

response.write "</BODY>"
response.write "</HTML>"
  
Call Disconnect_SiteWide

%> 