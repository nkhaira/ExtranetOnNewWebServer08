
<script language="JScript" runat="Server">
var datNow = new Date();
var strDate = datNow.toGMTString();
</script>

<%
'Now we can use the variables above in VBscript
response.write(strDate)

response.write "<P>"
response.write request.servervariables("SCRIPT_NAME")
%>

