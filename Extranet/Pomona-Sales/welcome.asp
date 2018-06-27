<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
'             Sandbox
' --------------------------------------------------------------------------------------

response.buffer = true

Dim ErrMessage, DoTransfer
ErrMessage = ""
DoTransfer = True

Call Connect_SiteWide
site_id = 14

%>
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

' --------------------------------------------------------------------------------------
' Start building the page
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent

Screen_Title    = Site_Description & " - " & Translate("Welcome",Login_Language,conn)
Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & Translate("Welcome",Login_Language,conn) & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

<form name="welcome" method="POST" action="/register/register.asp">
<%

with response
    .write "<INPUT TYPE=""Hidden"" NAME=""ID"" VALUE=""new"">" & vbcrlf
    .write "<INPUT TYPE=""Hidden"" NAME=""Site_ID"" VALUE=""14"">" & vbcrlf
    .write "<INPUT TYPE=""Hidden"" NAME=""Account_ID"" VALUE=""new"">" & vbcrlf
	.write "<SPAN CLASS=Heading3>"
	.write Translate("Distributor Extranet Upgrade",Login_Language,conn) & "</SPAN><P>" & vbcrlf
end with

strStd = "<P>Click on the button below to start the registration process.  Your account approval " &_
    "should take about 2 days, at which time you will receive a welcoming email."
    
strC = "If you're coming from pomonatest.com and you do not want to register now you can close this window."

' header is built - determine if they're coming from 
' get additional user data
Pom_userid = ""
Pom_userid = Request("puserid")

if Pom_userid = "" then
	with Response
		.write "<SPAN CLASS=""NormalBoldRed"">The registration process will be enhanced if you "
		.write "first log into www.pomonatest.com. </span><p>" & vbcrlf
	end with
else
	sql = "select * from PomonaCustomers where login = " & Pom_userid
	set dbRS = conn.Execute(sql)
	
	if dbRS.EOF then
		response.write "Hmmm, couldn't find the user....<P>"
	else
		'a few of the users already have accounts on other Fluke sites, this is noted in 
        ' the Active column of PomonaCustomers...
        if not IsNull(dbRS("Active")) then
            DoTransfer = False
            strA = "Since you have an account on another support.fluke.com site we have already" &_
                " created your account on pomona-sales. You will soon (if you haven't already)" &_
                " receive an email welcoming you and giving you more information."
            strStd = ""
            strC = "If you're coming from pomonatest.com you can close this window."
        else
    		with Response
    			.write "<input type=""hidden"" name=""Email"" value=""" & dbRS("Email") & """>" & vbcrlf
    			.write "<input type=""hidden"" name=""FirstName_MI"" value=""" & dbRS("FirstName") & """>" & vbcrlf
    			.write "<input type=""hidden"" name=""LastName"" value=""" & dbRS("LastName") & """>" & vbcrlf
    			.write "<input type=""hidden"" name=""Company"" value=""" & dbRS("Customer") & """>" & vbcrlf
    			.write "<input type=""hidden"" name=""Phone"" value=""" & dbRS("Phone") & """>" & vbcrlf
    			.write "<input type=""hidden"" name=""Fax"" value=""" & dbRS("Fax") & """>" & vbcrlf
    		end with
            strA = "We will transfer your contact information to the new registration screen."
        end if
		
		dbRS.Close
		' log the activity
		
		SQL = "update PomonaCustomers" & vbcrlf &_
			"set Net = 'welcome'" & vbcrlf &_
			"where login = " & Pom_userid
		conn.Execute (SQL)
	end if
	set dbRS = Nothing
end if

' tell the user what's going on
with response
	.write "<span class=""Normal"">Welcome to the new Pomoma Electronics Partner Portal (distributor"
	.write " extranet site).  In support of the new (and future) features and functionality it is"
	.write " necessary for you to go through the registration process again.  We will be asking for"
	.write " significantly more information than in the past and our intent is that the site "
	.write "delivers more as well." & vbcrlf
	.write strStd & vbcrlf
    .write "<P>" & strA & vbcrlf
end with

if DoTransfer then
    Response.write "<P><input type=""SUBMIT"" value=""Register Now"" class=""NavLeftHighLight1"">" & vbcrlf
end if

with response
	.write "<P>" & strC & vbcrlf
	.write "</span>" & vbcrlf
end with

%>
</form>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

' ------------- end of main --------------------------------------------------------------------
sub Notused
sql = "select * from userdata where id = " & dbRS("Active")
            set dbRS= Nothing
            set dbRS = conn.Execute(sql)
            if dbRS.EOF then
                Response.write "Wow, records indicate there is a support user...<P>" & vbcrlf
            else
                ' there's a whole lot of fields to copy over...
        		with Response
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""NTLogin"" VALUE="""
                    .write dbRS("NTLogin") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Email"" VALUE="""
                    .write dbRS("Email") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""FirstName_MI"" VALUE="""
                    .write dbRS("FirstName") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""MiddleName"" VALUE="""
                    .write dbRS("MiddleName") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""LastName"" VALUE="""
                    .write dbRS("LastName") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Suffix"" VALUE="""
                    .write dbRS("Suffix") & """>"                                 & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Company"" VALUE="""
                    .write dbRS("Company") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Job_Title"" VALUE="""
                    .write dbRS("Job_Title") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""MailStop"" VALUE="""
                    .write dbRS("Business_MailStop") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Address1"" VALUE="""
                    .write dbRS("Business_Address") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Address2"" VALUE="""
                    .write dbRS("Business_Address_2") & """>"                 & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""City"" VALUE="""
                    .write dbRS("Business_City") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""State_Province"" VALUE="""
                    .write dbRS("Business_State") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""State_Other"" VALUE="""
                    .write dbRS("Business_State_Other") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Country"" VALUE="""
                    .write dbRS("Business_Country") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Zip"" VALUE="""
                    .write dbRS("Business_Postal_Code") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Phone"" VALUE="""
                    .write dbRS("Business_Phone") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Extension"" VALUE="""
                    .write dbRS("Business_Phone_Extension") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Fax"" VALUE="""
                    .write dbRS("Business_Fax") & """>" & vbcrlf
                    
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Business_Phone_2"" VALUE="""
                    .write dbRS("Business_Phone_2") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Business_Phone_2_Extension"" VALUE="""
                    .write dbRS("Business_Phone_2_Extension") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Mobil_Phone"" VALUE="""
                    .write dbRS("Mobile_Phone") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Pager"" VALUE="""
                    .write dbRS("Pager") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Email_2"" VALUE="""
                    .write dbRS("Email_2") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_MailStop"" VALUE="""
                    .write dbRS("Shipping_MailStop") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_Address"" VALUE="""
                    .write dbRS("Shipping_Address") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_Address_2"" VALUE="""
                    .write dbRS("Shipping_Address_2") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_City"" VALUE="""
                    .write dbRS("Shipping_City") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_State_Other"" VALUE="""
                    .write dbRS("Shipping_State_Other") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_Postal_Code"" VALUE="""
                    .write dbRS("Shipping_Postal_Code") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_State"" VALUE="""
                    .write dbRS("Shipping_State") & """>" & vbcrlf
                    .write "<INPUT TYPE=""HIDDEN"" NAME=""Shipping_Country"" VALUE="""
                    .write dbRS("Shipping_Country") & """>" & vbcrlf
        		end with
                strA = "Since you have an account on another support.fluke.com site this will be easy."
            end if
end sub
%>
