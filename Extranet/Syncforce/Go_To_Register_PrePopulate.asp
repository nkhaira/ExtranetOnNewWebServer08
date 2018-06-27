<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML>
<HEAD>
	<TITLE>Go To Registration Page</TITLE>
</HEAD>

<BODY>
Go Directly to Registration Page for Specific Site ID

<FORM NAME="Logon" ACTION="https://support.fluke.com/register/register.asp" METHOD="Post">

<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="44">
<INPUT TYPE="Hidden" NAME="Account_ID"  VALUE="new">

<%
' Prepopulate Registration Page, supply values for these fields:
'
'  <INPUT TYPE="Hidden" NAME="CoreID" VALUE="">  ' www.fluke.com Core ID
'  <INPUT TYPE="Hidden" NAME="Prefix" VALUE="">
'  <INPUT TYPE="Hidden" NAME="FirstName_MI" VALUE="">
'  <INPUT TYPE="Hidden" NAME="LastName" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Suffix" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Title" VALUE="">
'  <INPUT TYPE="Hidden" NAME="MailStop" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Company" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Address1" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Address2" VALUE="">
'  <INPUT TYPE="Hidden" NAME="City" VALUE="">
'  <INPUT TYPE="Hidden" NAME="State_Province" VALUE="">
'  <INPUT TYPE="Hidden" NAME="State_Other" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Zip" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Country" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Email" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Phone" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Extension" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Fax" VALUE="">
'  <INPUT TYPE="Hidden" NAME="Title" VALUE="">  
'  <INPUT TYPE="Hidden" NAME="Language" VALUE="">
'  <INPUT TYPE="Hidden" NAME="msscid" VALUE="">  ' eStore Shopper ID

' or if core_email is supplied,
' an attempt will be done to look up user information in Fluke WWW.Fluke.COM
' Core Table and pre-populate the form.  If the email is not found, then user
' will be shown standard form.
  
  Test_Core_Email = False
  if Test_Core_Email then
    response.write "<INPUT TYPE=""Hidden"" NAME=""Core_Email""  VALUE="""">" & vbCrLf
  end if
%>
    
<INPUT TYPE="Submit" NAME="Pre-Register" VALUE="Pre-Register">
</FORM>

</BODY>
</HTML>
