
<%@Language="VBScript" %>

<%
    Response.Write("Send Email using CDOSYS")
%>


<%
Set msg = Server.CreateObject("CDO.Message")
Set conf= Server.CreateObject("CDO.Configuration")
%>

<%
confURL = "http://schemas.microsoft.com/cdo/configuration/"
with conf
    .Fields.Item(confURL & "sendusing") = 2
    .Fields.Item(confURL & "smtpserverport") = 25
    .Fields.Item(confURL & "smtpserver") = "mail.evt.danahertm.com"
    .Fields.Update
end with
%>

<%
with msg
    .From = "nitin.khaira@fluke.com"
    .To = "nitin.khaira@fluke.com"
    .Cc = "ankur.jana@fluke.com"
    .Bcc = "santosh.tembhare@fluke.com"
end with
%>

<%
with msg
    .Subject = "Send Email using CDOSYS: Testing ASP code on WIN2012 - Testing"
    .TextBody = "This is a plain text email"
    .HTMLBody = "<em>Hi, This is a testing HTML email using CDOSYS</b>"
end with
%>

<%
msg.Configuration = conf
msg.Send
Set conf = Nothing
Set msg = Nothing
%>

</br>

<%

    Response.Write("Email sent using CDOSYS")
%>
