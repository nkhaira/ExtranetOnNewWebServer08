<%

' --------------------------------------------------------------------------------------
' Decode Sequence
' --------------------------------------------------------------------------------------

Dim sequence
sequence = 0

if request.querystring("seq") = "" and request.form("seq") = "" then
  sequence = 0
elseif request.querystring("seq") <> "" then
  sequence = request.querystring("seq")
elseif request.form("seq") <> "" then
  sequence = request.form("seq")
end if  

' --------------------------------------------------------------------------------------
' Default Variables
' --------------------------------------------------------------------------------------

Logo_Path    = "http://www.seattlefirstbaptist.org/Common/Images/Logo/SFBC_Quatrefoil_Logo_PayPal.jpg"
Logo_Path    = "https://support.fluke.com/Images/Spacer_SFBC.jpg"
Return_Path  = "http://www.seattlefirstbaptist.org/donate/default.asp"
Cancel_Path  = "http://www.seattlefirstbaptist.org/Default.asp?Header=Task%20Force,%20Comissions%20and%20Groups%20Directory&Locator=Church_Groups"

FormName_1 = "OneTime"
FormName_2 = "Monthly"

Dim Menu_Group
Menu_Group = "Main" 

%>
<!--#INCLUDE virtual="/Common/Include/functions_string.asp" -->
<!--#INCLUDE virtual="/Common/include/functions_DB.asp"-->
<%

' --------------------------------------------------------------------------------------
' Debug
' --------------------------------------------------------------------------------------

Script_Debug = false

if Script_Debug = true then
  
  with response
    .write "<B>QueryString Collection:</B><BR>"
    for each item in request.querystring
      .write Item & ": " & request.querystring(item) & "<BR>"
    next

    .write "<BR>"
  
    .write "<B>Form Collection:</B><BR>"
    for each item in request.form
      .write Item & ": " & request.form(item) & "<BR>"
    next

    .flush
'    .end
    
  end with
    
end if 

select case sequence
    case 0, 2

    if sequence = 0 then
        Page_Name="Donation Form"
    else
        Page_Name="Donation Form - Thank You"
    end if    

    %>
    <!--#INCLUDE virtual="/Common/Include/Header.asp" -->
    <%

end select

select case sequence
    case 0, 2

    if sequence = 0 then

        Page_Intro = "<SPAN CLASS=HEADING3>""It Takes Time and Energy to be a Giver.""</SPAN><BR>" &_
                               "<SPAN CLASS=HEADING4>&nbsp;&nbsp;Do you have what it takes?</SPAN><P>" &_
                               "<SPAN CLASS=HEADING6>Those who give of themselves change the world.  There are so many positive ways to give that can make a world-changing difference. "  &_
                               "Seattle First Baptist Church is a community of values and commitments that truly makes a difference.  It isn't a church that only " &_
                               "serves its own members, but has a deep passion to change the world with tolerance, respect, celebration of diversity, and " &_
                               "advocacy for those at the margins of our society.</SPAN><P>" &_
                               "If you would like to be a giver and make a financial contribution to Seattle First Baptist Church, there are a number  of ways to do so.  You can also designate how you would like your financial contribution to be used." &_
                               "<UL>" &_
                               "<LI>If you would like to make a one-time financial contribution to Seattle First Baptist Church, <B>use option A or C</B>.</LI>" &_
                               "<LI>If you would like to make a weekly or monthly (same-amount contribution over a period of time) financial contribution to Seattle First Baptist Church, <B>use option B</B>.</LI>" &_
                               "<LI>If you would like to offer your time and energy to Seattle First Baptist Church, you will find many ways to do so throughout our website.  " &_
                               "Some of the ways have to do with serving the needs of the wider community.  To find out more, <B>use option D</B>.</LI>" &_
                               "</UL>" &_
                               "Just in case you have questions about how the on-line contributions work, we have prepared a " &_
                               "<A HREF="" LANGUAGE=""JavaScript"" onclick=""Publication = window.open('/Donate/SFBC_Giving_Online_QnA.pdf', 'Publication', 'fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=600,height=400,left=50,top=50');return false;"" TITLE=""Click to Read"">" &_
                               "<B>Question and Answer</B></A> brochure that you can read.<P>"

        response.write Page_Intro
                                       
    end if    

 end select
 
' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

select case sequence

  case 0  ' Donation Form
 
    with response
    
      .write "<TABLE BORDER=""8"" BORDERCOLOR=""#3F6359"">" & vbCrLf
      .write "<TR>" & vbCrLf
      .write "<TD BGCOLOR=""#E9E3DB"">" & vbCrLf
    
      ' Form 1 - On-Time Donation
      
      .write "<FORM NAME=""" & FormName_1 & """ ACTION=""Default.asp"" METHOD=""post"" onsubmit=""return CheckRequiredFields(this.form);"">" & vbCrLf
      .write "<TABLE BORDER=0 CELLSPACING=4 CELLPADDING=4 WIDTH=""100%"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD BGCOLOR=""#3F6359"" NOWRAP WIDTH=""1%"" CLASS=SmallWhite>" & vbCrLf
      .write "      <B>&nbsp;Option A&nbsp;&nbsp;</B>"
      .write "    </TD>" & vbCrLf
      .write "    <TD Class=Medium>" & vbCrLf
      .write "      <INPUT TYPE=""hidden"" NAME=""seq"" VALUE=""" & sequence + 1 & """>" & vbCrLf
      .write "      <INPUT TYPE=""hidden"" NAME=""method"" VALUE=""0"">" & vbCrLf
      .write "      I would like to give a <B>one-time contribution</B> in the amount of:&nbsp;&nbsp;$"
      .write "      <INPUT TYPE=""text"" NAME=""amount"" SIZE=""6"" CLASS=SmallBOLD VALUE="""" onClick=""" & FormName_1 & ".amount.style.backgroundColor='#FFFFFF';" & FormName_2 & ".a3.value='';" & FormName_2 & ".srt.options[0].selected = true;" & FormName_1 & ".submit1.disabled=false;" & FormName_2 & ".submit2.disabled=true;"">."
      .write "    </TD>" & vbCrLf
      .write "    <TD WIDTH=""1%"" CLASS=SMALL>" & vbCrLf
      .write "        <INPUT TYPE=""SUBMIT"" NAME=""submit1"" CLASS=SMALL VALUE=""Next Step"" Title=""Click here to continue to PayPal."">"
      .write "    </TD>" & vbCrLf
      .write "  </TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf
      .write "</FORM>" & vbCrLf & vbCrLf
  
      ' Form 2 - Monthly Donation
      
      .write "<FORM NAME=""" & FormName_2 & """ ACTION=""Default.asp"" METHOD=""post"" onsubmit=""return CheckRequiredFields(this.form);"">" & vbCrLf
      .write "<TABLE BORDER=0 CELLSPACING=4 CELLPADDING=4 WIDTH=""100%"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD BGCOLOR=""#3F6359"" NOWRAP WIDTH=""1%"" CLASS=SMALLWHITE>" & vbCrLf
      .write "      <B>&nbsp;Option B&nbsp;&nbsp;</B>"
      .write "    </TD>" & vbCrLf
      .write "    <TD CLASS=MEDIUM>" & vbCrLf
      .write "      <INPUT TYPE=""hidden"" NAME=""seq"" VALUE=""" & sequence + 1 & """>" & vbCrLf
      .write "      <INPUT TYPE=""hidden"" NAME=""method"" VALUE=""1"">" & vbCrLf
      .write "      I would like to give a&nbsp;"
      .write "      <SELECT NAME=""period"" CLASS=SMALLBOLD onClick=""" & FormName_2 & ".period.style.backgroundColor='#FFFFFF';"" onChange=""UpdateDuration()"">" & vbCrLf
      .write "        <OPTION VALUE="""">Select monthly or weekly</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""M"">monthly </OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""W"">weekly</OPTION>" & vbCrLf
      .write "      </SELECT>" & vbCrLf
      .write "      contribution in the amount of:&nbsp;&nbsp;$"
      .write "      <INPUT TYPE=""text"" NAME=""a3"" CLASS=SmallBOLD VALUE="""" SIZE=""6"" onClick=""" & FormName_2 & ".a3.style.backgroundColor='#FFFFFF';" & FormName_1 & ".amount.value='';" & FormName_1 & ".submit1.disabled=true;" & FormName_2 & ".submit2.disabled=false;"">"
      .write "      <BR>for&nbsp;" & vbCrLf
      .write "      <SELECT NAME=""srt"" CLASS=SMALLBOLD onClick=""" & FormName_2 & ".srt.style.backgroundColor='#FFFFFF'"";>" & vbCrLf
      .write "        <OPTION VALUE="""">Select how many times</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""1"">1 month </OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""2"">2 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""3"">3 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""4"">4 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""5"">5 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""6"">6 months</OPTION>" & vbCrLf 
      .write "        <OPTION VALUE=""7"">7 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""8"">8 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""9"">9 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""10"">10 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""11"">11 months</OPTION>" & vbCrLf
      .write "        <OPTION VALUE=""12"">12 months (1-year )</OPTION>" & vbCrLf
      .write "      </SELECT>." & vbCrLf
      .write "    </TD>" & vbCrLf
      .write "    <TD WIDTH=""1%"" CLASS=SMALL>" & vbCrLf
      .write "      <INPUT TYPE=""SUBMIT"" NAME=""submit2"" CLASS=Small VALUE=""Next Step"" Title=""Click here to continue to PayPal."">" & vbCrLf
      .write "    </TD>" & vbCrLf
      .write "  </TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf
      .write "</FORM>" & vbCrLf & vbCrLf
    
      ' Send Check to SFBC
      
      .write "<TABLE BORDER=0 CELLSPACING=4 CELLPADDING=4 WIDTH=""100%"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD BGCOLOR=""#3F6359"" NOWRAP WIDTH=""1%"" CLASS=SMALLWHITE>" & vbCrLf
      .write "      <B>&nbsp;Option C&nbsp;</B>"
      .write "    </TD>" & vbCrLf
      .write "    <TD CLASS=Small>" & vbCrLf
      .write "      <B>Or you can mail a check to</B>:<P>"
      .write "      <SPAN CLASS=Medium>"
      .write "      Seattle First Baptist Church<BR>"
      .write "      	1111 Harvard Ave.<BR>"
      .write "      	Seattle, WA&nbsp;&nbsp;98122<P>"
      .write "        Phone: 206.325.6051"
      .write "        </SPAN>"
      .write "    </TD>" & vbCrLf
      .write "    <TD WIDTH=""1%"" CLASS=SMALL>" & vbCrLf
      .write "      &nbsp;"
      .write "    </TD>" & vbCrLf
      .write "  </TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf & vbCrLf
      
      ' Donate Time and Energy to SFBC
      
      .write "<TABLE BORDER=0 CELLSPACING=4 CELLPADDING=4 WIDTH=""100%"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD BGCOLOR=""#3F6359"" NOWRAP WIDTH=""1%"" CLASS=SMALLWHITE>" & vbCrLf
      .write "      <B>&nbsp;Option D&nbsp;</B>"
      .write "    </TD>" & vbCrLf
      .write "    <TD CLASS=MEDIUM>" & vbCrLf
      .write "      I would like to donate my time and energy by serving on a church group, commissions, committee or task force:" 
      .write "    </TD>" & vbCrLf
      .write "    <TD WIDTH=""1%"" CLASS=SMALL>" & vbCrLf
      .write "    <INPUT TYPE=""Button"" NAME=""submit3"" CLASS=Small VALUE=""Next Step"" Title=""Click here to find out how to donate your time and energy."" LANGUAGE=""JavaScript"" onclick=""window.location.href='" & Cancel_Path & "'"">" & vbCrLf
      .write "    </TD>" & vbCrLf
      .write "  </TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf & vbCrLf
  
      ' Using Paypal
      
      .write "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=4 WIDTH=""100%"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD BGCOLOR=""#3F6359"" CLASS=SMALLWHITE>" & vbCrLf
      .write "      Secure contribution transactions to Seattle First Baptist Church are processed through "
      .write "      PayPal.  PayPal is a trusted leader in online payments services and has over 96 "
      .write "      million member accounts in 55 countries and regions.  PayPay allows you to make contributions from your Visa, Master Card, Discover, American Express, or from your bank account using EFT (electronic funds transfer) with just a few easy steps. "
      .write "    </TD>" & vbCrLf
      .write "    <TD BGCOLOR=""#3F6359"" CLASS=SMALLWHITE>" & vbCrLf
      .write "      <A HREF=""https://www.paypal.com/us/cgi-bin/?cmd=_login-run"">"   
      .write "      <IMG SRC=""images/My_Account_PayPal.jpg"" BORDER=0 ALIGN=""absmiddle"" ALT=""Click to Login to your PayPal Account"">"
      .write "      </A>"
      .write "    </TD>" & vbCrLf    
      .write "  </TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf & vbCrLf
  
      
      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf & vbCrLf
      
    end with
    
  case 1  ' Send to PayPal
  
      Call Connect_SFBC
  
      Add_ID = Get_New_Record_ID ("PayPal", "Sequence", sequence, conn)
      
      SQL = "UPDATE PayPal SET "
      
      with response
      
          .write "<HTML>" & vbCrLf
          .write "<HEAD>" & vbCrLf
          .write "<TITLE>SFBC - Contribution Form</TITLE>" & vbCrLf
          .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
          .write "</HEAD>" & vbCrLf
          .write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""Black"">"
 
        select case request.form("method")
        
          case 0

            ' Sequence
            SQL = SQL & "Sequence=" & sequence & ","
            ' Frequency
            SQL = SQL & "Frequency='O'" & ","
            ' Duration
            SQL = SQL & "Duration=1" & ","
            ' Amount
            SQL = SQL & "Amount=" & request.form("Amount") & ","
            ' BDate
           SQL = SQL & "BDate=#" & now() & "#,"
            ' Session_ID
            SQL = SQL & "Session_ID='" & Session.SessionID & "'"
            SQL = SQL & " WHERE ID=" & Add_ID & ";"
            
            '.write SQL & "<P>"
            conn.execute (SQL)

            .write "<BODY BGCOLOR=""White"" onLoad='document.FORM1.submit()'>" & vbCrLf    
            .write "<FORM NAME=""FORM1"" ACTION=""https://www.paypal.com/cgi-bin/webscr"" METHOD=""post"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cmd"" VALUE=""_xclick"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""business"" VALUE=""paypal@seattlefirstbaptist.org"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cbt"" VALUE=""Return to Seattle First Baptist Church Website"">" & vbCrLf  
            .write "  <INPUT TYPE=""hidden"" NAME=""return"" VALUE=""" & Return_Path & "?seq=" & sequence + 1 & "&sid=" & Add_ID & """>" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cancel_return"" VALUE=""http://www.seattlefirstbaptist.org/donate/default.asp?seq=3&sid=" & Add_ID & """>" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""item_name"" VALUE=""One-Time Contribution to Seattle First Baptist Church - Online"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""item_number"" VALUE=""SFBC-OTC"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""no_shipping"" VALUE=""1"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""no_note"" VALUE=""0"">" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""rm"" VALUE=""POST"">" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""cn"" Value=""My contribution is to be designated for:"">" & vbCrLf                  ' Label above the note field.
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_header_image"" Value=""" & Server.URLEncode(Logo_Path) & """>" & vbCrLf    ' SFBC Logo
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_headerback_color"" Value=""FFFFFF"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_headerborder_color"" Value=""FFFFFF"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_payflow_color"" Value=""D0DCD8"">" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""amount"" VALUE=""" & request.form("amount") & """>" & vbCrLf
            .write "</FORM>" & vbCrLf & vbCrLf
    
          case 1

            ' Sequence
            SQL = SQL & "Sequence=" & sequence & ", "
            ' Frequency
            SQL = SQL & "Frequency='" & request.form("period") & "', "
            ' Duration
            SQL = SQL & "Duration='" & request.form("srt") & "', "
            ' Amount
            SQL = SQL & "Amount=" & request.form("a3") & ", "
            ' BDate
           SQL = SQL & "BDate=#" & now() & "#," 
            ' Session_ID
            SQL = SQL & "Session_ID='" & Session.SessionID & "'"
            
            SQL = SQL & " WHERE ID=" & Add_ID & ";"
            
            '.write SQL & "<P>"
            conn.execute (SQL)

            .write "<BODY BGCOLOR=""White"" onLoad='document.FORM2.submit()'>" & vbCrLf    
            .write "<FORM NAME=""FORM2"" ACTION=""https://www.paypal.com/cgi-bin/webscr"" METHOD=""post"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""p3"" VALUE=""1"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""t3"" VALUE=""" & request.form("period") & """>" & vbCrLf        'Weekly or Monthly
            .write "  <INPUT TYPE=""hidden"" NAME=""src"" VALUE=""1"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""sra"" VALUE=""1"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cmd"" VALUE=""_xclick-subscriptions"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""business"" VALUE=""paypal@seattlefirstbaptist.org"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cbt"" VALUE=""Return to Seattle First Baptist Church Website"">" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""return"" VALUE=""" & Return_Path & "?seq=" & sequence + 1 & "&sid=" & Add_ID & """>" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cancel_return"" VALUE=""http://www.seattlefirstbaptist.org/donate/default.asp?seq=3&sid=" & Add_ID & """>" & vbCrLf        
            
            select case request.form("period")
              case "W"
                .write "  <INPUT TYPE=""hidden"" NAME=""item_name"" VALUE=""Weekly Contribution to Seattle First Baptist Church - Online"">" & vbCrLf
                .write "  <INPUT TYPE=""hidden"" NAME=""item_number"" VALUE=""SFBC-MWC"">" & vbCrLf

              case "M"
                .write "  <INPUT TYPE=""hidden"" NAME=""item_name"" VALUE=""Monthly Contribution to Seattle First Baptist Church - Online"">" & vbCrLf
                .write "  <INPUT TYPE=""hidden"" NAME=""item_number"" VALUE=""SFBC-MMC"">" & vbCrLf
            end select    
                
            .write "  <INPUT TYPE=""hidden"" NAME=""no_shipping"" VALUE=""1"">" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""cn"" Value=""My contribution is to be used for:"">" & vbCrLf                  ' Label above the note field.
            .write "  <INPUT TYPE=""hidden"" NAME=""no_note"" VALUE=""0"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""rm"" VALUE=""POST"">" & vbCrLf        
            .write "  <INPUT TYPE=""hidden"" NAME=""currency_code"" VALUE=""USD"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""a3"" VALUE="""  & request.form("a3")  & """>" & vbCrLf     ' Amount
            .write "  <INPUT TYPE=""hidden"" NAME=""srt"" VALUE=""" & request.form("srt") & """>" & vbCrLf     ' Duration
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_header_image"" Value=""" & Server.URLEncode(Logo_Path) & """>" & vbCrLf    ' SFBC Logo
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_headerback_color"" Value=""FFFFFF"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_headerborder_color"" Value=""FFFFFF"">" & vbCrLf
            .write "  <INPUT TYPE=""hidden"" NAME=""cpp_payflow_color"" Value=""D0DCD8"">" & vbCrLf        
    
            .write "</FORM>" & vbCrLf & vbCrLf
            
            .write "</FONT>"
            .write "</BODY>"
            .write "</HTML>"  
            
        end select
        
      end with
      
      Call Disconnect_SFBC
    
  case 2  ' Thank you
  
        with response
        
          .write "Thank you for your generous contribution. Your contribution transaction to Seattle First Baptist Church has been completed, and a receipt for your contribution "
          .write "has been emailed to you. You can log into your account at <A HREF=""www.paypal.com/us"">PayPal</A> to view details of your contribution transactions at any time."
  
          ' --------------------------------------------------------------------------------------
          ' Debug
          ' --------------------------------------------------------------------------------------
          
          Script_Debug = true
          if Script_Debug = true then
            
            .write "<P>"
            .write "Debug Information<P>"
            .write "<B>QueryString Collection:</B><BR>"
            for each item in request.querystring
              .write Item & ": " & request.querystring(item) & "<BR>"
            next
          
            .write "<BR>"
           
            .write "<B>Form Collection:</B><BR>"
            for each item in request.form
              .write Item & ": " & request.form(item) & "<BR>"
            next
          
            .flush
          end if  
            
          SQL = "UPDATE PayPal SET "
            
          ' Sequence
          SQL = SQL & "Sequence=" & sequence & ","
          ' Transaction ID
          SQL = SQL & "Transaction='" & request.querystring("tx") & "',"
          ' Status
          SQL = SQL & "Status='" & request.querystring("status") & "',"
          ' Amount
          SQL = SQL & "Amount=" & request.querystring("amt") & ","
          ' EDate
          SQL = SQL & "EDate=#" & now() & "#,"
          ' Currency
          SQL = SQL & "Currency='" & request.querystring("cc") & "',"
          ' Session_ID
          SQL = SQL & "Session_ID='" & Session.SessionID & "',"
          ' Message
          SQL = SQL & "Memo='" & request.querystring("cm") & "'"

          SQL = SQL & " WHERE ID='" & request.querystring("sid") & ";"
            
          Call Connect_SFBC
          conn.execute (SQL)
          Call Disconnect_SFBC
            
          ' seq: 2
          ' tx: 94S617205X478184S
          ' st: Completed
          ' amt: 0.05
          ' cc: USD
          ' cm: 
          ' sig: omlTjvDUoeIoMY6K1JGDa9kRvYpMx+c96c9bdt3/2Y3cmXN6z4noRIKM9aUV/xvwQ/MmFJBZ4GkxcFVXAchuzOfV5uWNiFp0JKHHwhQQk/bilN/ygmWE0uzt+VQ3tGU8jebT3GbAx+NCVHUePB7NbzgtKg5heXfYp79giXnbO8g=
          
        end with
        
  case 3 ' Cancel Return
  
    response.redirect Cancel_Return & vbCrLf          

  case else
  
end select

select case sequence

    case 0,2
        %>
        <!--#INCLUDE virtual="/Common/Include/Footer.asp" -->
        <%
        
end select        
    

%>

<SCRIPT LANGUAGE="JAVASCRIPT">
<!--//

var FormName_1 = document.<%=FormName_1%>;
var FormName_2 = document.<%=FormName_2%>;
var ErrorMsg       = "";
var strVal              = "";
var strChk            = "";

var Flag_Period  = true;
var Flag_Times   = true;
var Flag_Amount = true;

function CheckRequiredFields() {
  
  ErrorMsg = "";
  
  if (FormName_1.amount.value.length == 0 && FormName_2.a3.value.length == 0) {
      ErrorMsg = ErrorMsg + "Please enter the amount of your contribution for either Option A or B.\n";
  }

  if (FormName_1.amount.value.length != 0) {
    if (! IsNumeric(FormName_1.amount.value)) {
      FormName_1.amount.style.backgroundColor = "#FFB9B9";
      FormName_1.amount.focus();
      ErrorMsg = ErrorMsg + 'Option A contribution amount is in US dollars using the format of "0.00".  The contribution amount that you have entered contains non-numeric or invalid characters.\n';
    }
    else {
      if (FormName_1.amount.value <= 0) {
        FormName_1.amount.style.backgroundColor = "#FFB9B9";
        FormName_1.amount.focus();    
        ErrorMsg = ErrorMsg + 'Option A contribution amount cannot be zero or a negative dollar amount.\n';
      }  
    }
  }                

  // Check Amount Numeric
  if (FormName_2.a3.value.length != 0) {
    if (! IsNumeric(FormName_2.a3.value)) {
      FormName_2.a3.style.backgroundColor = "#FFB9B9";
      FormName_2.srt.style.backgroundColor = "#FFFFFF";              
      FormName_2.a3.focus();    
      ErrorMsg = ErrorMsg + 'Option B contribution amount is in US dollars using the format of "0.00".  The contribution amount that you have entered contains non-numeric or invalid characters.\n';
    }
    
    // Check Amount > 0
    if (FormName_2.a3.value <=0) {
        FormName_2.a3.style.backgroundColor = "#FFB9B9";
        FormName_2.srt.style.backgroundColor = "#FFFFFF";        
        FormName_2.a3.focus();    
        ErrorMsg = ErrorMsg + 'Option B contribution amount cannot be zero or a negative dollar amount.\n';
    }
    
    // Check Period if Amount > 0
    if (IsNumeric(FormName_2.a3.value) && FormName_2.a3.value > 0) {
       strVal = FormName_2.period.value;
       strChk = strVal.substring(0, 7);
       if ((strVal == "") || (strChk.toLowerCase() == "select")) {
           FormName_2.period.style.backgroundColor = "#FFB9B9";
           FormName_2.period.focus();        
           ErrorMsg = ErrorMsg + "\n" + "Select a monthly or a weekly contribution period.";
        }
    }

    // Check Times if Amount > 0
    if (IsNumeric(FormName_2.a3.value) && FormName_2.a3.value > 0) {
      strVal = FormName_2.srt.value;
      strChk = strVal.substring(0, 7);
      if ((strVal == "") || (strChk.toLowerCase() == "select")) {
        FormName_2.srt.style.backgroundColor = "#FFB9B9";
        FormName_2.srt.focus();
        if (FormName_2.period.value == "M") {
          ErrorMsg = ErrorMsg + "\n" + "Select how many months you would like your monthly contribution ";
        }
        else if (FormName_2.period.value == "W") {
          ErrorMsg = ErrorMsg + "\n" + "Select how many weeks you would like your weekly contribution ";        
        }
        else {
          ErrorMsg = ErrorMsg + "\n" + "Select how many times you would like your contribution "; 
        }                 
        if (IsNumeric(FormName_2.a3.value) && FormName_2.a3.value >0) {
          ErrorMsg = ErrorMsg + "of $" + FormName_2.a3.value
        }
        ErrorMsg = ErrorMsg + " to renew.\n";
      }
    }
  }
  
  if (ErrorMsg.length) {
    alert (ErrorMsg);
    return (false);
  }
  else {
    return (true);
  }
}

function IsNumeric(sText) {
  var ValidChars = "0123456789.-";
  var IsNumber = true;
  var Char;
  
  if (sText == null) sText = "";
   
  for (i = 0; i < sText.length && IsNumber == true; i++) { 
    Char = sText.charAt(i); 
    if (ValidChars.indexOf(Char) == -1) {
      IsNumber = false;
    }
  }
  return IsNumber;
}

function UpdateDuration() {

var store = new Array();

    store[0] = new Array(
	                'Select how many weeks','','2 weeks','3 weeks','4 weeks','5 weeks','6 weeks','7 weeks','8 weeks','9 weeks','10 weeks','11 weeks','12 weeks');

    store[1] = new Array(
	                'Select how many months','','2 months','3 months','4 months','5 months','6 months','7 months','8 months','9 months','10 months','11 months','12 months');
   
var list_label = new Array();
    list_label = new Array('','','2','3','4','5','6','7','8','9','10','11','12');

 	var box = FormName_2.period;
	var number = box.options[box.selectedIndex].value;
	if (number == "") return;
    if (number == "W") {
 	    var list = store[0];
    }
    if (number == "M") {
 	    var list = store[1];
    }

	var box2 = FormName_2.srt;
	box2.options.length = 0;
	for(i = 0; i < list.length; i += 2)
	{
		box2.options[i/2] = new Option(list[i],list_label[i]);
	}
}


// -->
</SCRIPT>
