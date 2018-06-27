<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/connections/connection_EMail.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
' --------------------------------------------------------------------------------------
' Author:     P. Barbee
' Date:       11/15/2000
'             Dev
'				
' NOTE: this page does not require credentials nor display any real nav.  It's only
' purpose is to capture an email address and send an invitation to the user
' --------------------------------------------------------------------------------------

response.buffer = true

Call Connect_SiteWide
bSentMail = False
site_id = 3

%>
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Site_Description & " - Order Inquiry - Invitation"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>Order Inquiry - Invitation</FONT>"
Top_Navigation  = False 
Side_Navigation = True
Content_Width   = 95  ' Percent

email_address = ""
email_address = trim(request.form("em_address"))

if isblank(email_address) then
	email_address = trim(request.querystring("em_address"))
end if

strText = "Fluke Order Inquiry is available on the Fluke Partner Portal." & vbcrlf & vbcrlf &_
	"To register for access to the Fluke Partner Portal, please go to" & vbcrlf &_
	"http://support.fluke.com" & vbcrlf &_
	"- choose ""Partner Portal - Fluke Industrial Tools"" from the selection for the Name of " & _
  "the Site where you want to go" & vbcrlf &_
	"- click register" & vbcrlf &_
	"- complete the form" & vbcrlf &_
	"- click submit" & vbcrlf & vbcrlf &_
	"It will take up to 48 hours for your submittal to be approved, at which time you" & vbcrlf &_
	"will receive an e-mail that reminds you of the user name and password you have selected," & vbcrlf &_
	"and the web address of the Fluke Partner Portal."
        
if not isblank(email_address) then
  if lcase(request.form("company")) = "fnet" then
    strText = Replace(strText,"Partner Portal - Fluke Industrial Tools","ChannelVision")
    strText = Replace(strText,"Fluke","Fluke Networks")
    strText = Replace(strText,"Partner Portal","ChannelVision")
  elseif lcase(request.form("company")) = "pom" then
    strText = Replace(strText,"Fluke Industrial Tools","Pomona Electronics")
    strText = Replace(strText,"Fluke","Pomona")
  end if

	Send_invite_mail
	bSentMail = True
end if

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<SPAN CLASS=Heading3>Order Inquiry - Invitation</SPAN><BR>"
response.write "<BR>"


if bSentMail then
	
	Response.write "<P class=""medium"">An invitation has been sent to " & email_address & "." & vbcrlf
	Response.write "<P class=""medium"">You can close this window" & vbcrlf
	
else
	
	with response
		.write "<FORM name=""oi_mail"" METHOD=""POST"" "
		.write "onsubmit=""return My_submit('rtn');"">" & vbcrlf
        .write "<input type=""hidden"" name=""company"" value=""find"">" & vbcrlf
		.write "<TABLE BORDER=1 WIDTH=""90%"" BORDERCOLOR=""#666666"" BGCOLOR="""
		.write Contrast & """ CELLPADDING=0 CELLSPACING=0>" & vbcrlf
		.write "  <TR>" & vbcrlf
		.write "    <TD align=""center"">" & vbcrlf
		.write "      <TABLE CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR="""
		.write Contrast & """ WIDTH=""80%"">" & vbcrlf
		.write "        <TR>" & vbcrlf
 		.write "         <TD CLASS=""SmallBold"" nowrap>"
		.write "Email Address:</TD>" & vbcrlf
		.write "         <TD CLASS=""SmallBold"">"
		.write "          	<INPUT CLASS=""Medium"" TYPE=""Text"" NAME=""em_address"" SIZE=45>" & vbcrlf
		.write "         </TD>" & vbcrlf
		.write "        </TR>" & vbcrlf
		.write "        <TR>" & vbcrlf
 		.write "         <TD CLASS=""SmallBold"" nowrap>"
		.write "Optional Text:</TD>" & vbcrlf
		.write "         <TD CLASS=""SmallBold"">"
		.write "          <Textarea class=""medium"" name=""otext"" cols=60 rows=7></textarea>" & vbcrlf
		.write "         </TD>" & vbcrlf
		.write "        </TR>" & vbcrlf
		.write "        <TR>" & vbcrlf
		.write "		 <TD bgcolor=""black"" align=""center"" colspan=3 class=""SmallBoldGold"">" & vbcrlf
        .write "            Send Email for:&nbsp;&nbsp;" & vbcrlf
		.write "		 <input type=""button"" value=""Fluke Industrial"" onclick=""My_submit('find');"""
		.write " class=""NavLeftHighlight1"">" & vbcrlf
		.write "		 <input type=""button"" value=""Fluke Networks"" onclick=""My_submit('fnet');"""
		.write " class=""NavLeftHighlight1"">" & vbcrlf
        .write "            <br>" & vbcrlf
		.write "		 <input type=""button"" value=""Pomona Sales"" onclick=""My_submit('pom');"""
		.write " class=""NavLeftHighlight1"">" & vbcrlf
		.write "         </TD>" & vbcrlf
		.write "        </TR>" & vbcrlf
		.write "      </TABLE>" & vbcrlf
		.write "    </TD>" & vbcrlf
		.write "  </TR>" & vbcrlf
		.write "</TABLE>" & vbcrlf
		.write "</form>" & vbcrlf
		.write "<P class=""SmallBold"">Text automatically included: <span class=""MediumItalic"">"
		.write "optional text appears above this</span></p>" & vbcrlf
		.write "<P class=""Medium"">" & replace(strText,vbcrlf,"<BR>"&vbcrlf) & vbcrlf
	end with
	%>

<script language="Javascript">
	
	function My_submit(wh_from) {
		var df = document.oi_mail;
		var GoodtoGo = 0;
		
		if (df.em_address.value.length) {
			GoodtoGo = 1;
		}
		
		if (GoodtoGo) {
			if (wh_from == 'rtn') {
				return true;
			}
			else {
                df.company.value = wh_from;
				// alert('This is a submit');
				df.submit();
			}
		}
		else {
			alert('Please enter an email address');
			return false;
		}
	}
</script>

<%
end if

%>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
response.end

sub Send_invite_mail
	
	'Set Mailer = CreateObject("SMTPsvg.Mailer")
	'adding new email method
	%>
	<!--#include virtual="/connections/connection_email_new.asp"-->
	<%
	
	'Mailer.ReturnReceipt = false
	'Mailer.ConfirmRead = false
	'Mailer.WordWrapLen = 80
  'QMessage = True > use AspQMail --- QMessage = False > don't use it
	'Mailer.QMessage = False
	'Mailer.ClearAttachments
	'Mailer.RemoteHost = GetEmailServer()
	' send mail to Peter
	
	'Mailer.ClearRecipients
	''Mailer.AddRecipient "David Whitlock","David.Whitlock@fluke.com"
	
  select case lcase(request.form("company"))
    case "fnet"
    	'Mailer.AddRecipient "ChannelVision User",email_address
    	'Mailer.FromName = "Fluke Networks"
    	'Mailer.Subject = "Fluke Networks ChannelVision Request"

		msg.To = """ChannelVision User""" & email_address
		msg.From = """Fluke Networks""" & "webmail@fluke.com"
		msg.Subject = "Fluke Networks ChannelVision Request"

    case "pom"
    	'Mailer.AddRecipient "Pomona Extranet User",email_address
    	'Mailer.FromName = "Pomona Partner Portal"
    	'Mailer.Subject = "Pomona Partner Portal Request"

		
		msg.To = """Pomona Extranet User""" & email_address
		msg.From = """Pomona Partner Portal""" & "webmail@fluke.com"
		msg.Subject = "Pomona Partner Portal Request" 

    case else
    	'Mailer.AddRecipient "Extranet User",email_address
    	'Mailer.FromName = "Fluke Partner Portal"
    	'Mailer.Subject = "Fluke Partner Portal Request"

		msg.To = """Extranet User""" & email_address
		msg.From = """Fluke Partner Portal""" & "webmail@fluke.com"
		msg.Subject = "Fluke Partner Portal Request" 

  end select
  
	'Mailer.FromAddress = "webmail@fluke.com"
	'Mailer.ReplyTo = "webmaster@fluke.com"
	'Mailer.WordWrap = True
	'Mailer.ContentType = "text/plain"

	msg.ReplyTo = "webmaster@fluke.com"
	
	'Mailer.ClearBodyText
	
	if request.form("otext") <> "" then
		strText = request.form("otext") & vbcrlf & vbcrlf & strText
	end if
	'Mailer.BodyText = strText
	msg.TextBody = strText
	
	'Mailer.SendMail
	msg.Configuration = conf
	On Error Resume Next
	msg.Send
	If Err.Number = 0 then
		'Success
	Else
		'Fail
	End If

	'set Mailer = Nothing
end sub
%>