<%@Language="VBScript" %>

<%
    Response.Write("Send Email using CDOSYS :: Production Server TEST, Host - smtp.fortivemail.com")
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
    .Fields.Item(confURL & "smtpserver") = "smtp.fortivemail.com"    
    .Fields.Update
end with
%>

<%
with msg
    .Fields.Item("urn:schemas:mailheader:X-Priority") = 2
    .Fields.Update
end with
%>

<%
with msg
    .From = "nitin.khaira@fluke.com"
    .To = "nitin.khaira@fluke.com"
    .Cc = "santosh.tembhare@fluke.com"
    '.Bcc = "nitin.khaira@fluke.com"
end with
%>

<%
with msg
    .Subject = "Send Email using CDOSYS"
    .TextBody = "This is a plain text email. Production Server, Host - smtp.fortivemail.com"
    .HTMLBody = "<em>Hi, This is a testing HTML email using CDOSYS. VM - Production Server, Host - smtp.fortivemail.com</b>"
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
