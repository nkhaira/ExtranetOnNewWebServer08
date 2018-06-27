<script language="JScript" runat="Server">
var datNow = new Date();
var strDate = datNow.toGMTString();
</script>

<%
Mailer.RemoteHost  = GetEmailServer()
Mailer.TimeOut     = 180
Mailer.WordWrap    = False
Mailer.WordWrapLen = 150
Mailer.AddExtraHeader "X-MimeOLE: Produced By Microsoft Exchange V6.5"
Mailer.DateTime    = strDate
%>