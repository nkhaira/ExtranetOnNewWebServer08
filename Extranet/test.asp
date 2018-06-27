<%@LANGUAGE="VBSCRIPT"%>
<%
function TestCaptcha(byval valSession, byval valCaptcha)
	dim tmpSession
	valSession = Trim(valSession)
	valCaptcha = Trim(valCaptcha)
	if (valSession = vbNullString) or (valCaptcha = vbNullString) then
		TestCaptcha = false
	else
		tmpSession = valSession
		valSession = Trim(Session(valSession))
		Session(tmpSession) = vbNullString
		if valSession = vbNullString then
			TestCaptcha = false
		else
			valCaptcha = Replace(valCaptcha,"i","I")
			if StrComp(valSession,valCaptcha,1) = 0 then
				TestCaptcha = true
			else
				TestCaptcha = false
			end if
		end if		
	end if
end function
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>CAPTCHA Example</title>
<style type="text/css">
body {
	font-size:12px;
	font-family:Verdana, Arial, Helvetica, sans-serif;
}
</style>
<script language="javascript">
function RefreshImage(valImageId) {
	var objImage = document.images[valImageId];
	if (objImage == undefined) {
		return;
	}
	var now = new Date();
	objImage.src = objImage.src.split('?')[0] + '?x=' + now.toUTCString();
}
</script>
</head>
<body>
<form id="form1" name="form1" method="post" action="">
  <table width="400" border="1" align="center">
    <tr>
      <td colspan="2" align="center"><strong>CAPTCHA Example</strong></td>
    </tr>
    <tr>
      <td width="261">CAPTCHA Image</td>
      <td width="123"><img id="imgCaptcha" src="captcha.asp" /><br /><a href="javascript:RefreshImage('imgCaptcha')" onclick="RefreshImage('imgCaptcha')">Change Image</a></td>
    </tr>
    <tr>
      <td>Write the characters in the image above</td>
      <td><input name="captchacode" type="text" id="captchacode" size="10" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="btnTest" id="btnTest" value="Test Input" /></td>
    </tr>
    <%
	if not IsEmpty(Request.Form("btnTest")) then
		Response.Write("<tr><td colspan=""2"" align=""center"">")
		if TestCaptcha("ASPCAPTCHA", Request.Form("captchacode")) then
			Response.Write("<b style=""color:#00CC33"">The code you enter verified.</b>")
		else
			Response.Write("<b style=""color:#FF0000"">You entered the wrong code.</b>")
		end if
		Response.Write("</td></tr>" & vbCrLf)
	end if
	%>
  </table>
</form>
</body>
</html>
