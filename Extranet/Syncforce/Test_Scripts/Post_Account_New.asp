<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML>
<HEAD>
	<TITLE>Test DCM Record</TITLE>
</HEAD>

<BODY>

<FONT FACE="Arial" SIZE=2>

<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
' Test routine to check db for test record, not part of standard code for DCM interface scripts.

NTLogin = "DCM_Test_Record"
Password= "mystifyed"
ID = 0

response.write "<FONT COLOR=""red"">"

Call Connect_SiteWide

SQL = "SELECT ID, Site_ID, NTLogin, Password from UserData WHERE NTLOGIN='" & NTLogin & "' ORDER BY Site_ID"
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open SQL, conn, 3, 3

if rsUser.EOF then
  response.write "No Records in DB for NTLogin: " & NTLogin & "<P>"
else  
  response.write "Records in DB for NTLogin: " & NTLogin & "<P>"
  do while not rsUser.EOF
    response.write "Site_ID: " & rsUser("Site_ID") & "<BR>"
    response.write "ID: " & rsUser("ID") & "<P>"
    if CInt(rsUser("Site_ID")) = 3 then
      ID = rsUser("ID") ' Set Default ID for Site 3 FIND-Sales
    end if  
    rsUser.MoveNext
  loop
end if
rsUser.close
set rsUser = nothing

Call Disconnect_SiteWide

response.write "</FONT>"
%>

<!-------------------------------------------------------------------------->

Verify NTLogin: "<%=NTLogin%>", ID: "<%=ID%>" (Variable User Name)
<P>
<FORM NAME="Verify_Account" ACTION="/sw-administrator/account_admin.asp" METHOD="POST">

<!-- Account Administrator Credentials -->

<INPUT TYPE="Hidden" NAME="Logon_User" VALUE="EuroDCM">
<INPUT TYPE="Hidden" NAME="Logon_Password" VALUE="!SyncForce#">
<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="3">  <!-- Force Admin Site ID to trigger DCM transfer-->

<!-- Action to Perform-->
<INPUT TYPE="Hidden" NAME="Verify" VALUE="Verify">
<INPUT TYPE="Text" NAME="NTLogin" VALUE="<%=NTLogin%>">

<INPUT TYPE="Submit" NAME="Submit" VALUE="Verify NTLogin">

</FORM>

<!-------------------------------------------------------------------------->

New Record (Add)
<P>
<FORM NAME="Add_Account" ACTION="/sw-administrator/account_admin.asp" METHOD="POST">

<!-- Account Administrator Credentials -->

<INPUT TYPE="Hidden" NAME="Logon_User" VALUE="EuroDCM">
<INPUT TYPE="Hidden" NAME="Logon_Password" VALUE="!SyncForce#">

<!-- Action to Perform-->

<INPUT TYPE="Hidden" NAME="Update" VALUE="Update">

<!-- Required For Adding New Records -->

<INPUT TYPE="Hidden" NAME="ID" VALUE="add">
<INPUT TYPE="Hidden" NAME="NewFlag" VALUE="on">

<!-- User Specific Information for all Sites-->

<INPUT TYPE="Hidden" NAME="NTLogin" VALUE="<%=NTLogin%>">
<INPUT TYPE="Hidden" NAME="Password" VALUE="<%=Password%>">

<INPUT TYPE="Hidden" NAME="Type_Code_Required" VALUE="on">
<INPUT TYPE="Hidden" NAME="Type_Code" VALUE="1">
<INPUT TYPE="Hidden" NAME="Prefix" VALUE="">
<INPUT TYPE="Hidden" NAME="FirstName" VALUE="Sample">
<INPUT TYPE="Hidden" NAME="MiddleName" VALUE="Euro">
<INPUT TYPE="Hidden" NAME="LastName" VALUE="Record">
<INPUT TYPE="Hidden" NAME="Suffix" VALUE="">
<INPUT TYPE="Hidden" NAME="Company" VALUE="Sample Euro Company">
<INPUT TYPE="Hidden" NAME="Job_Title" VALUE="Sample Euro Job">
<INPUT TYPE="Hidden" NAME="Region" VALUE="2"-->
<INPUT TYPE="Hidden" NAME="Account_Region" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Address" VALUE="Sample Address">
<INPUT TYPE="Hidden" NAME="Business_Address_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_MailStop" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_City" VALUE="Sample City">
<INPUT TYPE="Hidden" NAME="Business_State" VALUE="ZZ">
<INPUT TYPE="Hidden" NAME="Business_State_Other" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Postal_Code" VALUE="Sample Postal Code">
<INPUT TYPE="Hidden" NAME="Business_Country" VALUE="NL">
<INPUT TYPE="Hidden" NAME="Shipping_Address" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_Address_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_MailStop" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_City" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_State" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_State_Other" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_Postal_Code" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_Country" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Fax" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Phone" VALUE="123456789">
<INPUT TYPE="Hidden" NAME="Business_Phone_Extension" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Phone_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Phone_2_Extension" VALUE="">
<INPUT TYPE="Hidden" NAME="Mobile_Phone" VALUE="">
<INPUT TYPE="Hidden" NAME="Pager" VALUE="">
<INPUT TYPE="Hidden" NAME="Email" VALUE="test@fluke.com">
<INPUT TYPE="Hidden" NAME="Email_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Email_Method" VALUE="">
<INPUT TYPE="Hidden" NAME="Connection_Speed" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Method" VALUE="">
<INPUT TYPE="Hidden" NAME="Account_Language" VALUE="eng">
<INPUT TYPE="Hidden" NAME="Fluke_ID" VALUE="">
<INPUT TYPE="Hidden" NAME="Core_ID" VALUE="">
<INPUT TYPE="Hidden" NAME="eStore_ID" VALUE="">
<INPUT TYPE="Hidden" NAME="Lit_Store" VALUE="">
<INPUT TYPE="Hidden" NAME="Image_Store" VALUE="">
<INPUT TYPE="Hidden" NAME="Pub_Store" VALUE="">
<INPUT TYPE="Hidden" NAME="Comment" VALUE="">
<INPUT TYPE="Hidden" NAME="Logon" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_0" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_1" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_3" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_4" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_5" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_6" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_7" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_8" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_9" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_System" VALUE="">

<!-- Site Specific Information for User-->

<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="3">
<INPUT TYPE="Hidden" NAME="Fcm" VALUE="off">
<INPUT TYPE="Hidden" NAME="Fcm_ID" VALUE="0">
<INPUT TYPE="Hidden" NAME="Groups" VALUE="Find-Sales">
<INPUT TYPE="Hidden" NAME="SubGroups" VALUE="eura, euda">
<INPUT TYPE="Hidden" NAME="Groups_Aux" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Sections" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Options" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Date" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Frequency" VALUE="">
<INPUT TYPE="Hidden" NAME="Broadcast_Date" VALUE="">
<INPUT TYPE="Hidden" NAME="ExpirationDate" VALUE="12/31/2002">
<INPUT TYPE="Hidden" NAME="ChangeDate" VALUE="<%=Date()%>">
<INPUT TYPE="Hidden" NAME="ChangeID" VALUE="5145">
<INPUT TYPE="Hidden" NAME="RTE_Enabled" VALUE="">
<INPUT TYPE="Hidden" NAME="Reg_Request_Date" VALUE="<%=Date()%>">
<INPUT TYPE="Hidden" NAME="Reg_Approval_Date" VALUE="<%=Date()%>">
<INPUT TYPE="Hidden" NAME="Reg_Admin" VALUE="5145">
<INPUT TYPE="Hidden" NAME="Reg_Site_Code" VALUE="Find-Sales">

<INPUT TYPE="Submit" NAME="Submit" VALUE="Add New Account">

</FORM>

<!-------------------------------------------------------------------------->

Update Record NTLogin: "<%=NTLogin%>", ID: "<%=ID%>"

<P>
<FORM NAME="Update_Account" ACTION="/sw-administrator/account_admin.asp" METHOD="POST">

<!-- Account Administrator Credentials -->

<INPUT TYPE="Hidden" NAME="Logon_User" VALUE="EuroDCM">
<INPUT TYPE="Hidden" NAME="Logon_Password" VALUE="!SyncForce#">

<!-- Action to Perform-->

<INPUT TYPE="Hidden" NAME="Update" VALUE="Update">

<!-- Required For Updating Record -->

<INPUT TYPE="Hidden" NAME="ID" VALUE="<%=ID%>">

<!-- User Specific Information for all Sites-->

<INPUT TYPE="Hidden" NAME="NTLogin" VALUE="<%=NTLogin%>">
<INPUT TYPE="Hidden" NAME="Password" VALUE="<%=Password%>">

<INPUT TYPE="Hidden" NAME="Change_Password" VALUE="">

<INPUT TYPE="Hidden" NAME="Type_Code_Required" VALUE="on">
<INPUT TYPE="Hidden" NAME="Type_Code" VALUE="1">
<INPUT TYPE="Hidden" NAME="Prefix" VALUE="">
<INPUT TYPE="Hidden" NAME="FirstName" VALUE="Sample">
<INPUT TYPE="Hidden" NAME="MiddleName" VALUE="Euro">
<INPUT TYPE="Hidden" NAME="LastName" VALUE="Record">
<INPUT TYPE="Hidden" NAME="Suffix" VALUE="">
<INPUT TYPE="Hidden" NAME="Company" VALUE="Sample Euro Company">
<INPUT TYPE="Hidden" NAME="Job_Title" VALUE="Sample Euro Job">
<INPUT TYPE="Hidden" NAME="Region" VALUE="2"-->
<INPUT TYPE="Hidden" NAME="Account_Region" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Address" VALUE="Sample Address">
<INPUT TYPE="Hidden" NAME="Business_Address_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_MailStop" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_City" VALUE="Sample City">
<INPUT TYPE="Hidden" NAME="Business_State" VALUE="ZZ">
<INPUT TYPE="Hidden" NAME="Business_State_Other" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Postal_Code" VALUE="Sample Postal Code">
<INPUT TYPE="Hidden" NAME="Business_Country" VALUE="NL">
<INPUT TYPE="Hidden" NAME="Shipping_Address" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_Address_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_MailStop" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_City" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_State" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_State_Other" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_Postal_Code" VALUE="">
<INPUT TYPE="Hidden" NAME="Shipping_Country" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Fax" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Phone" VALUE="123456789">
<INPUT TYPE="Hidden" NAME="Business_Phone_Extension" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Phone_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_Phone_2_Extension" VALUE="">
<INPUT TYPE="Hidden" NAME="Mobile_Phone" VALUE="">
<INPUT TYPE="Hidden" NAME="Pager" VALUE="">
<INPUT TYPE="Hidden" NAME="Email" VALUE="test@fluke.com">
<INPUT TYPE="Hidden" NAME="Email_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Email_Method" VALUE="">
<INPUT TYPE="Hidden" NAME="Connection_Speed" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Method" VALUE="">
<INPUT TYPE="Hidden" NAME="Account_Language" VALUE="eng">
<INPUT TYPE="Hidden" NAME="Fluke_ID" VALUE="">
<INPUT TYPE="Hidden" NAME="Core_ID" VALUE="">
<INPUT TYPE="Hidden" NAME="eStore_ID" VALUE="">
<INPUT TYPE="Hidden" NAME="Lit_Store" VALUE="">
<INPUT TYPE="Hidden" NAME="Image_Store" VALUE="">
<INPUT TYPE="Hidden" NAME="Pub_Store" VALUE="">
<INPUT TYPE="Hidden" NAME="Comment" VALUE="">
<INPUT TYPE="Hidden" NAME="Logon" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_0" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_1" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_2" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_3" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_4" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_5" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_6" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_7" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_8" VALUE="">
<INPUT TYPE="Hidden" NAME="Aux_9" VALUE="">
<INPUT TYPE="Hidden" NAME="Business_System" VALUE="">

<!-- Site Specific Information for User-->

<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="3">
<INPUT TYPE="Hidden" NAME="Fcm" VALUE="off">
<INPUT TYPE="Hidden" NAME="Fcm_ID" VALUE="0">
<INPUT TYPE="Hidden" NAME="Groups" VALUE="Find-Sales">
<INPUT TYPE="Hidden" NAME="SubGroups" VALUE="eura, euda">
<INPUT TYPE="Hidden" NAME="Groups_Aux" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Sections" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Options" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Date" VALUE="">
<INPUT TYPE="Hidden" NAME="Subscription_Frequency" VALUE="">
<INPUT TYPE="Hidden" NAME="Broadcast_Date" VALUE="">
<INPUT TYPE="Hidden" NAME="ExpirationDate" VALUE="12/31/2002">
<INPUT TYPE="Hidden" NAME="ChangeDate" VALUE="<%=Date()%>">
<INPUT TYPE="Hidden" NAME="ChangeID" VALUE="5145">
<INPUT TYPE="Hidden" NAME="RTE_Enabled" VALUE="">
<INPUT TYPE="Hidden" NAME="Reg_Request_Date" VALUE="">
<INPUT TYPE="Hidden" NAME="Reg_Approval_Date" VALUE="">
<INPUT TYPE="Hidden" NAME="Reg_Admin" VALUE="5145">
<INPUT TYPE="Hidden" NAME="Reg_Site_Code" VALUE="Find-Sales">

<INPUT TYPE="Submit" NAME="Submit" VALUE="Update Account" <%if ID = 0 then response.write "DISABLED"%>>

</FORM>

Retrieve Account Data: "<%=NTLogin%>", ID: "<%=ID%>"
<P>
<FORM NAME="Retrieve_Account" ACTION="/sw-administrator/account_admin.asp" METHOD="POST">

<!-- Account Administrator Credentials -->

<INPUT TYPE="Hidden" NAME="Logon_User" VALUE="EuroDCM">
<INPUT TYPE="Hidden" NAME="Logon_Password" VALUE="!SyncForce#">
<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="3">  <!-- Force Admin Site ID to trigger DCM transfer-->

<!-- Action to Perform-->
<INPUT TYPE="Hidden" NAME="Retrieve" VALUE="Retrieve">
<INPUT TYPE="Hidden" NAME="ID" VALUE="<%=ID%>">

<INPUT TYPE="Submit" NAME="Submit" VALUE="Retrieve Account" <%if ID = 0 then response.write "DISABLED"%>>

</FORM>

Delete NTLogin: "<%=NTLogin%>", ID: "<%=ID%>" (User Name)
<P>
<FORM NAME="Delete_Account" ACTION="/sw-administrator/account_admin.asp" METHOD="POST">

<!-- Account Administrator Credentials -->

<INPUT TYPE="Hidden" NAME="Logon_User" VALUE="EuroDCM">
<INPUT TYPE="Hidden" NAME="Logon_Password" VALUE="!SyncForce#">

<!-- Action to Perform-->

<INPUT TYPE="Hidden" NAME="Delete" VALUE="Delete">
<INPUT TYPE="Hidden" NAME="ID" VALUE="<%=ID%>">

<INPUT TYPE="Submit" NAME="Submit" VALUE="Delete Account" <%if ID = 0 then response.write "DISABLED"%>>

</FORM>

<P>
After clicking on Action Button, use ,[BACK] button to return to this form and [REFRESH] page.


</FONT>

</BODY>
</HTML>
