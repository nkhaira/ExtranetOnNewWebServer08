<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML>
<HEAD>
	<TITLE>Go To Registration Page</TITLE>
</HEAD>

<BODY>
Go Directly to Registration Page for Specific Site ID
<FORM NAME="Logon" ACTION="https://support.fluke.com/register/register.asp" METHOD="Post">

<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="3">
<INPUT TYPE="Hidden" NAME="Account_ID"  VALUE="new">
<INPUT TYPE="Hidden" NAME="Language"  VALUE="dut">

<%

' If core_email is supplied,
' an attempt will be done to look up user information in Fluke WWW.Fluke.COM
' Core Table and pre-populate the form.  If the email is not found, then user
' will be shown standard form.
  
  Test_Core_Email = False
  if Test_Core_Email then
    response.write "<INPUT TYPE=""Hidden"" NAME=""Core_Email""  VALUE=""j.schoenmakers@privatedomain.nl"">" & vbCrLf
  end if
%>
    
<INPUT TYPE="Submit" NAME="Test" VALUE="Test">
</FORM>

</BODY>
</HTML>
