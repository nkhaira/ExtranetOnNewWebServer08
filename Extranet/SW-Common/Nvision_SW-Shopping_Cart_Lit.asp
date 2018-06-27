<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       6/2/2002
'             Printed Literature Fulfillment Shopping Cart
' --------------------------------------------------------------------------------------

Dim Transfer_Debug
Transfer_Debug = true

Dim Script_Debug
Script_Debug = false

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_FormData.asp"-->


<STYLE>
.ZeroValue  {font-size:8.5pt;font-weight:Bold;color:Black;background:#FFFF99;text-decoration:none;font-family:Arial,Verdana;}
</STYLE>
<%

response.buffer=true

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Border_Toggle
Border_Toggle = 0

Dim x, s

Dim Logon_Name
Logon_Name = Session("Logon_User")
'Login_User = Logon_Name

Dim ErrMessage, PostErrMessage, ErrCode, ErrorMesg
ErrMessage     = ""
PostErrMessage = ""

Dim Cart_Mode             ' Limits Edit information in SW-Profile_Edit.asp and No-Cache in Header.asp
Cart_Mode = True

if not isblank(request("Cart_Type")) then
  Cart_Type = LCase(request("Cart_Type"))
else
  Cart_Type = "lit_us"
end if
    
SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Logon_Name & "' AND Site_ID=" & Site_ID & " AND NewFlag=0"

Set rsLogin = Server.CreateObject("ADODB.Recordset")
rsLogin.Open SQL, conn, 3, 3

Dim Login_ID, Common_ID, Login_Region, Login_Country, Login_Type_Code
Login_ID             = rsLogin("ID")
Common_ID            = rsLogin("ID")
Login_Region         = rsLogin("Region")
Login_Country        = rsLogin("Business_Country")
Login_Type_Code      = rsLogin("Type_Code")
Login_Language       = rsLogin("Language")

rsLogin.close
set rsLogin = nothing

%>
<!--#include virtual="/sw-Common/Preferred_Language.asp"-->
<%

' --------------------------------------------------------------------------------------

Dim Action_Sequence       ' Sequencer
Dim Review                ' Final Review View
Review = False

Dim Account_Info_ID, Shipping_Address_ID, Address_Text
Dim FirstName, MiddleName, LastName, Company, Job_Title
Dim Business_Address, Business_Address_2, Business_City, Business_City_Other, Business_State
Dim Business_Postal_Code, Business_Country, Business_Email, Business_Fax, Business_Phone
Dim Shipping_Address, Shipping_Address_2, Shipping_City, Shipping_City_Other, Shipping_State
Dim Shipping_Postal_Code, Shipping_Country, Shipping_Comment
Dim	strLocalHostName, strLocalHostIP, strProtocol, strMethod, strRemoteHostName, strRemoteHostIP, iRemoteHostPort, strRemoteHostTargetFile
Dim strKeyValueDelimiter, strPairDelimiter, cProtocol, cMethod, strLocalHostReferrerFile
Dim strResponse, strPost_QueryString
Dim Form_Name, Action_Count
Dim Ship_To_Alternates, Save_Cart, Order_Number, Order_Number_Results, Order_Status_Info, Order_Status_Info_Alt
Dim Screen_Width, Screen_Height
Dim MailSubject, MailMessage, DotLine
Dim ErrorFlag

ErrorFlag = false

DotLine = "----------------------------------------------------------" & vbCrLf

Screen_Width  = Session("Screen_Width")
Screen_Height = Session("Screen_Height")

Dim Account_Reviewed
if isblank(session("Account_Reviewed")) then
  Account_Reviewed = False
else
  Account_Reviewed = True
  session("Account_Reviewed") = Account_Reviewed
end if  

Dim Cart_Width
Cart_Width = 100

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = "Shopping Cart"
Bar_Title       = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Literature Order",Login_Language,conn) & " - " & Translate("Shopping Cart",Login_Language,conn) & "</FONT>"
Top_Navigation  = False 
Side_Navigation = False
Content_Width   = 100  ' Percent
BackURL = Session("BackURL")

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->

<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
<!--
var Shopping_Cart = this.window;
var Ship_To = 0;
//-->
</SCRIPT>
<%

Action_Count = 0

response.write "<SPAN CLASS=Small>"

' --------------------------------------------------------------------------------------
' Action Sequencer
' --------------------------------------------------------------------------------------

Action_Sequence = LCase(request("Action"))

if Action_Sequence <> "account" then
  Form_Name = "Cart_Data"
  response.write "<FORM METHOD=""GET"" NAME=""" & Form_Name & """ onsubmit=""return(CheckRequiredFields(this.form));"">" & vbCrLf
end if  

if Action_Sequence = "transfer" then
  if instr(1,request("Ship_To"),",") > 0 then
    Ship_To_Addresses = split(request("Ship_To"),", ")
    Ship_To_Count = UBound(Ship_To_Addresses)
  else
    REDim Ship_To_Addresses(0)
    Ship_To_Addresses(0) = request("Ship_To")
    Ship_To_Count = 0
  end if
end if

Select case Action_Sequence

  case "add"

    ' Determine if item is already in Shopping Cart if not then add Item

    if not isblank(request("Cart_ID")) then    ' Use Calendar DB
      SQL = "SELECT Account_NTLogin, Asset_ID, Submit_Date, Item_Number " &_
            "FROM   Shopping_Cart_Lit " &_
            "WHERE Account_NTLogin='" & Logon_Name & "' " &_
            "   AND Asset_ID=" & Request("Cart_ID") & " " &_
            "   AND Submit_Date IS NULL"
    else                                      ' Use Literature_Items_US DB
      SQL = "SELECT Account_NTLogin, Asset_ID, Submit_Date, Item_Number " &_
            "FROM   Shopping_Cart_Lit " &_
            "WHERE Account_NTLogin='" & Logon_Name & "' " &_
            "   AND Item_Number='" & Request("Lit_ID") & "' " &_
            "   AND Submit_Date IS NULL"
    end if 
    
    Set rsCartCK = Server.CreateObject("ADODB.Recordset")
    rsCartCK.Open SQL, conn, 3, 3

    if rsCartCK.EOF then

      ' Determine if item is Orderable by 7-digit Item Number and SubGroup Code "user"
      
      if not isblank(request("Cart_ID")) then    ' Use Calendar DB
        SQL = "SELECT * FROM Calendar WHERE ID=" & request("Cart_ID")
      else
        SQL = "SELECT Item FROM Literature_Items_US WHERE ITEM=" & request("Lit_ID")
      end if  
      Set rsItemCK = Server.CreateObject("ADODB.Recordset")
      rsItemCK.Open SQL, conn, 3, 3

      if not rsItemCK.EOF then

        if not isblank(request("Cart_ID")) then    ' Use Calendar DB      
          if not isblank(rsItemCk("Item_Number")) and isnumeric(rsItemCk("Item_Number")) then
            SQL = "INSERT INTO Shopping_Cart_Lit " &_
                  "( Site_ID, Account_ID, Account_NTLogin, Asset_ID, Quantity, Max_Limit, Cart_Date, Cart_Type, Item_Number, Region, Country ) " &_
                  "VALUES ( " & Site_ID & ", " & Common_ID & ", '" & Logon_Name & "', " & Request("Cart_ID") & ", " & "0" & ", " & Request("Max_Limit") & ", '" & Date() & "', '" & Cart_Type & "', '" & rsItemCk("Item_Number") & "', " & Login_Region & ", '" & Login_Country & "' )"

            conn.execute(SQL)
            ErrMessage = ErrMessage &_
                         "<LI>" & Translate("The top item listed was just added to your shopping cart.",Login_Language,conn) &_
                         "<LI>" & Translate("Please update the Quantity for this item.",Login_Language,conn)
          end if
        elseif not isblank(request("Lit_ID")) and isnumeric(request("Lit_ID")) then

            SQL = "INSERT INTO Shopping_Cart_Lit " &_
                  "( Site_ID, Account_ID, Account_NTLogin, Asset_ID, Quantity, Max_Limit, Cart_Date, Cart_Type, Item_Number ) " &_
                  "VALUES ( " & Site_ID & ", " & Common_ID & ", '" & Logon_Name & "', " & "0" & ", " & "0" & ", " & Request("Max_Limit") & ", '" & Date() & "', '" & Cart_Type & "', '" & request("Lit_ID") & "', " & Login_Region & ", '" & Login_Country & "' )"

            conn.execute(SQL)
            ErrMessage = ErrMessage &_
                         "<LI>" & Translate("The top item listed was just added to your shopping cart.",Login_Language,conn) &_
                         "<LI>" & Translate("Please update the Quantity for this item.",Login_Language,conn)
        else
          ErrMessage = ErrMessage & "<LI>" & Translate("The item that you have selected was not added to your shopping cart because of an invalid Item Number ID or SubGroup Code.",Login_Language,conn)        
        end if  
      else
        ErrMessage = ErrMessage & "<LI>" & Translate("The item that you have selected was not added to your shopping cart because of an invalid Item Number ID or SubGroup Code.",Login_Language,conn)
      end if

      rsItemCk.close
      set rsItemCk = nothing

    else
      ErrMessage = ErrMessage & "<LI>" & Translate("The item that you have selected is already in your shopping cart.",Login_Language,conn)
      Action_Sequence = ""
    end if

    rsCartCk.close
    set rsCartCk = nothing
    
    Session("Cart_Active") = True

    %>
    <SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
    <!--
    if (document.all) {
      Shopping_Cart.window.resizeTo((screen.width / 2) - 19 ,screen.height - 44);
    }
    else if (document.layers || document.getElementById) {
      if (Shopping_Cart.window.outerHeight < screen.height || Shopping_Cart.window.outerWidth < screen.width){
        Shopping_Cart.window.outerWidth = (screen.width / 2) - 19;
        Shopping_Cart.window.outerHeight = screen.height - 44;        
      }
    }
    Shopping_Cart.window.moveTo((screen.width / 2),24);
    //-->
    </SCRIPT>
    <%

  case "update"
  
    if not isblank(request("Quantity")) and isnumeric(request("Quantity")) then

      if request("Quantity") > 0 then
      
        if not isblank(request("Cart_ID")) then
          SQL = "SELECT ID, Max_Limit FROM Shopping_Cart_Lit WHERE ID=" & request("Cart_ID")
        else
          SQL = "SELECT ID, Max_Limit FROM Shopping_Cart_Lit WHERE Item_Number='" & request("Lit_ID") & "'"
        end if  

        Set rsLimit = Server.CreateObject("ADODB.Recordset")
        rsLimit.Open SQL, conn, 3, 3
        Quantity  = request("Quantity")
        Max_Limit = rsLimit("Max_Limit")
        rsLimit.close
        set rsLimit = nothing
        
        if Quantity  = 0 then Quantity  = 1
        if Max_Limit = 0 then Max_Limit = 1
        
        if CDbl(Quantity) <= CDbl(Max_Limit) or Quantity = 1 then
          if not isblank(request("Cart_ID")) then         
            SQL = "UPDATE Shopping_Cart_Lit SET Quantity=" & Quantity & " WHERE ID=" & request("Cart_ID")
          else
            SQL = "UPDATE Shopping_Cart_Lit SET Quantity=" & Quantity & " WHERE Item_Number='" & request("Lit_ID") & "'"
          end if  
          conn.execute(SQL)
          Session("Cart_Active") = True
        else
          if not isblank(request("Cart_ID")) then
            SQL = "UPDATE Shopping_Cart_Lit SET Quantity=" & Max_Limit & " WHERE ID=" & request("Cart_ID")
          else
            SQL = "UPDATE Shopping_Cart_Lit SET Quantity=" & Max_Limit & " WHERE Item_Number='" & request("Lit_ID") & "'"
          end if              
          conn.execute(SQL)
          ErrMessage = ErrMessage & "<LI>" & Translate("The quantity that you have requested exceeded the maximum quantity limit of for this item of:",Login_Language,conn) & " " & Max_Limit & "</LI>" &_
                                    "<LI>" & Translate("The quantity requested has been adjusted not to exceed this limit.",Login_Language,conn) & "</LI>" &_
                                    "<LI>" & Translate("If you require additional quantities above the maximum quantity limit please phone or email the contact person listed below:",Login_Language,conn) & "<BR>" &_
                                    "<SPAN CLASS=Small>Marty Jezek - " & Translate("Phone",Login_Language,conn) & ": 425.446.6207" & ", " & Translate("Email",Login_Language,conn) & ": <A HREF=""mailto:Marty.Jezek@Fluke.com"">Marty.Jezek@Fluke.com</A></SPAN>" &_
                                    "</LI>"          
          Session("Cart_Active") = True
        end if  
      else
        ErrMessage = ErrMessage & "<LI>" & Translate("The quantity value for this item cannot be blank or less than 0 - Update Failed.",Login_Language,conn)
      end if
    else
      ErrMessage = ErrMessage & "<LI>" & Translate("The quantity value for this item cannot be blank or less than 0 or non-numeric - Update Failed.",Login_Language,conn)
    end if
      
  case "delete"

    if not isblank(request("Cart_ID")) and isnumeric(request("Cart_ID")) then
      SQL = "DELETE FROM Shopping_Cart_Lit WHERE ID=" & request("Cart_ID")
      conn.execute(SQL)
    end if

  case "account"
    %>
    <SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
    <!--
    Shopping_Cart.window.name = "Shipping_Information";
    //-->
    </SCRIPT>  
    <%
    
    Session("Account_Reviewed") = True
    
    Cart_Width = 95
    response.write "<DIV ALIGN=CENTER>" & vbCrLf
    response.write "<TABLE WIDTH=""" & Cart_Width & "%"" CELLSPACING=0 CELLPADDING=0 BORDER=""" & Border_Toggle & """>"
    response.write "<TR>" & vbCrLf
    response.write "<TD WIDTH=""50%"" CLASS=Small>" & vbCrLf
    response.write "<SPAN CLASS=HEADING3>" & Translate("Literature Order",Login_Language,conn) & " - " & Translate("Shopping Cart",Login_Language,conn) & "</SPAN><BR><SPAN CLASS=HEADING6> " & Translate("Account Information Review",Login_Language,conn) & "</SPAN><BR>"
    response.write "<UL><LI>" & Translate("Review your Account and Shipping Information below.",Login_Language,conn) & "<BR>"
    response.write "<LI>" & Translate("After your review, if there are no changes or if you have made changes to your Account Information, click on the [ Update ] button at the bottom of the form to continue.",Login_Language,conn)
    response.write "</UL>"
    response.write "</TD>" & vbCrLf
    response.write "<TD WIDTH=""50%"">" & vbCrLf
    response.write "&nbsp;"
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf

    response.write "<TR>" & vbCrLf
    response.write "<TD COLSPAN=2 CLASS=Small>" & vbCrLf
    
    BackURLSecure = "http://" & request.ServerVariables("Server_Name") & request.ServerVariables("SCRIPT_Name") & "?Action=Review"
    %>
    <!--#include Virtual="/sw-common/SW-Profile_Edit.asp"-->
    <%
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf
    response.write "</DIV>" & vbCrLf
         
  case "review"

    Cart_Width = 95
    Review = True
       
  case "transfer"

    Cart_Width = 95
    ' Check to see if there are any 0 Quantity items before committing order

    SQL = "SELECT Quantity, Account_NTLogin FROM Shopping_Cart_Lit WHERE Account_NTLogin='" & Logon_Name & "' AND (Quantity=NULL OR Quantity=0)"
    Set rsItemCK = Server.CreateObject("ADODB.Recordset")
    rsItemCK.Open SQL, conn, 3, 3

    if not rsItemCK.EOF then
      Action_Sequence = "review"
      Review = True
      ErrMessage = "<LI>" & Translate("Your order could not be submitted because there are items with 0 quantity (shown in yellow).",Login_Language,conn) &_
                   "<LI>" & Translate("Please correct the items that you are ordering with 0 quantity, or delete these items from your shopping cart.",Login_Language,conn)
    else
      Review = True
    end if

    rsItemCK.Close
    set rsItemCK = nothing

    if Review = True and isblank(ErrMessage) then
    
      'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
 
      'adding new email method
      %>
      <!--#include virtual="/connections/connection_email_new.asp"-->
      <%

      'Mailer.ClearAllRecipients
        
      'Mailer.ReturnReceipt  = False
      'Mailer.Priority       = 3
      
      'Mailer.FromName       = Translate(Site_Description,Alt_Language,conn)
      'Mailer.FromAddress    = "WebMail@Fluke.com"
      'Mailer.ReplyTo        = Site_Admin_Email     
      ''Mailer.AddBCC           "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"  ' Domain Administrator
      'Mailer.AddBCC           "Extranet Group", "ExtranetAlerts@Fluke.com"  ' Domain Administrator
      
      msg.From = """" & Translate(Site_Description,Alt_Language,conn) & """" & "webmail@fluke.com"
      msg.ReplyTo = Site_Admin_Email
      msg.Bcc = """Extranet Group""" & "ExtranetAlerts@Fluke.com"
      
      MailSubject           = Translate("Literature Order Advisory",Alt_Language,conn)     
  
      MailMessage = Translate("This is an automated email advisory from the",Alt_Language,conn) & " " & Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Literature Ordering System",Alt_Language,conn) & ". " & vbCrLf & vbCrLf
      Call Shipping_Cart_Contact_Info_Text
      if Login_Language <> Alt_Language then
        MailMessage = MailMessage & Replace(Replace(Order_Status_Info_Alt,"<LI>","- "),"</LI>",vbCrLf) & vbCrLf
      else
        MailMessage = MailMessage & Replace(Replace(Order_Status_Info,"<LI>","- "),"</LI>",vbCrLf) & vbCrLf
      end if
      MailMessage = MailMessage & Translate("Literature Order Summary",Alt_Language,conn) & ":" & vbCrLf & vbCrLf

      Account_Info_ID = 0
      Call Get_Account_Information

      MailMessage = MailMessage & "=======================================" & vbCrLf
      MailMessage = MailMessage & Translate("Ordered by",Alt_Language,conn) & vbCrLf
      MailMessage = MailMessage & "=======================================" & vbCrLf & vbCrLf  
      Call Build_MailMessage_Order_By
      
      'Mailer.AddRecipient FormatFullName(FirstName, MiddleName, LastName), Business_Email
      msg.To = """" & FormatFullName(FirstName, MiddleName, LastName) & """" & Business_Email

      Order_Date           = Date()
      Order_Time           = Time()
      PostErrMessage       = ""
      Order_Number_Results = ""

      ' Loop through each Ship To and submit separate orders for each
         
      for stc = 0 to Ship_To_Count

        ' Action     

        strPost_QueryString = "Action=Order_Post"
        
        ' Get Order By Profile Info
        Account_Info_ID = 0
        
        Call Get_Account_Information        
        
        strPost_QueryString = strPost_QueryString &_
          "&SORDNH=" & Server.URLEncode("") &_
          "&SRDSTS=" & Server.URLEncode("04") &_
          
          "&OUSERN=" &_
          "&OPSSWD=" &_
          
          "&OCUNAM=" & Server.URLEncode(FormatFullName(FirstName, MiddleName, LastName)) &_
          "&OCMPNM=" & Server.URLEncode(Company) &_
          "&OTITL="  & Server.URLEncode(Job_Title) &_
          
          "&OADRS1=" & Server.URLEncode(Business_Address) &_
          "&OADRS2=" & Server.URLEncode(Business_Address_2) &_
          "&OCITY="  & Server.URLEncode(Business_City) &_
          "&OSTATE=" & Server.URLEncode(Business_State) &_
          "&OZIP="   & Server.URLEncode(Business_Postal_Code) &_
          "&OCNTRY=" & Server.URLEncode(Business_Country) &_
          
          "&OPHONE=" & Server.URLEncode(FormatPhone(Business_Phone)) &_
          "&OFAX="   & Server.URLEncode(FormatPhone(Business_Fax)) &_
          "&OEMAIL=" & Server.URLEncode(Business_Email)
                    
        ' Get Ship To Profile Info
        
        Account_Info_ID = Ship_To_Addresses(stc)
        select case Account_Info_ID
          case 0
            Call Get_Account_Information
          case else                           ' Drop Ship To
            Call Get_Account_Information
        end select
        
        ' DCG ORDER SYSTEM
										strPost_QueryString = strPost_QueryString      &_
          "&SCUSID=" & Server.URLEncode("FLUKEPP:" & Common_ID & ":" & Account_Info_ID) &_
          "&SCUNAM=" & Server.URLEncode(FormatFullName(FirstName, MiddleName, LastName)) &_
          "&STITL="  & Server.URLEncode(Job_Title) &_
          "&SCMPNM=" & Server.URLEncode(Company) &_
          "&SADRS1=" & Server.URLEncode(Shipping_Address) &_
          "&SADRS2=" & Server.URLEncode(Shipping_Address_2) &_
          "&SCITY="  & Server.URLEncode(Shipping_City) &_
          "&SSTATE=" & Server.URLEncode(Shipping_State) &_
          "&SZIP="   & Server.URLEncode(Shipping_Postal_Code) &_
          "&SCNTRY=" & Server.URLEncode(Shipping_Country) &_
          "&SPHONE=" & Server.URLEncode(FormatPhone(Business_Phone)) &_
          "&SFAX="   & Server.URLEncode(FormatPhone(Business_Fax)) &_
          "&SEMAIL=" & Server.URLEncode(Business_Email) &_
         
          "&SRDDT="  & Server.URLEncode(Order_Date) &_
          "&SRDTM="  & Server.URLEncode(Order_Time) &_
          "&SDTAP="  & Server.URLEncode(Order_Date) &_
          "&SDTSP="  & Server.URLEncode("") &_
         
          "&SSHPV="  & Server.URLEncode("UPS") &_
          "&SSHPIN=" & Server.URLEncode(Shipping_Comment) &_
          "&SPOST="  & Server.URLEncode("") &_
          "&SMATL="  & Server.URLEncode("") &_
          "&SHAND="  & Server.URLEncode("") &_
          "&SDTIV="  & Server.URLEncode("") &_
          "&SPONUM=" & Server.URLEncode("") &_
          "&SCCNUM=" & Server.URLEncode("") &_
          "&SDTDU="  & Server.URLEncode("") &_
          "&SCOMM="  & Server.URLEncode("") &_
          "&SPPKG="  & Server.URLEncode("") &_
          "&SPCAR="  & Server.URLEncode("") &_
          "&SPSRV="  & Server.URLEncode("") &_
          "&SSHPI2=" & Server.URLEncode("")
         
        MailMessage = MailMessage & DotLine
        MailMessage = MailMessage & Translate("Ship to Address",Alt_Language,conn) & " " & (stc + 1) & vbCrLf
        MailMessage = MailMessage & DotLine & vbCrLf  
        Call Build_MailMessage_Ship_To
        MailMessage = MailMessage & Translate("Order Date",Alt_Language,conn) & ": " & Order_Date & "  " & Order_Time & " (PST)" & vbCrLf

        ' Order Items
        
        SQLOrder = "SELECT " &_
                     "Shopping_Cart_Lit.Account_NTLogin, " &_
                     "Shopping_Cart_Lit.Quantity AS Quantity, " &_
                     "Shopping_Cart_Lit.Item_Number AS Item_Number, " &_
                     "Shopping_Cart_Lit.Submit_Date " &_
                   "FROM Shopping_Cart_Lit " &_
                   "WHERE Account_NTLogin='" & Logon_Name & "' AND Submit_Date IS NULL " &_
                   "ORDER BY Item_Number"

        'response.write SQLOrder & "<P>"    

        Set rsOrder = Server.CreateObject("ADODB.Recordset")
        rsOrder.Open SQLOrder, conn, 3, 3

        ' Build Item/Quantity Array
        
        qsItems   = "&SITSNA="
        qsRev     = "&REVID="
        MailItems = ""

        do while not rsOrder.EOF
          qsItems = qsItems &  rsOrder("Item_Number") & "," & rsOrder("Quantity") & ","
          
          ' DCG System Needs Latest Revision Code
          SQLRevision = "SELECT Item, Revision " &_
                        "FROM dbo.Literature_Items_US " &_
                        "WHERE STATUS='Active' AND [ACTION]='Complete' AND Item=" & rsOrder("Item_Number") & " " &_
                        "ORDER BY Revision DESC"
          
         'response.write SQLRevision & "<P>"
          
          Set rsRevision = Server.CreateObject("ADODB.Recordset")
          rsRevision.Open SQLRevision, conn, 3, 3
          
          if not rsRevision.EOF then
            qsRev = qsRev & rsRevision("Revision") & ","
          else  
            qsRev = qsRev & ","
          end if
          
          rsRevision.close
          set rsRevision = nothing
                                    
          MailItems = MailItems & Translate("Item Number",Alt_Language,conn) & ": " & rsOrder("Item_Number") & "      " & Translate("Quantity",Alt_Language,conn) & ": " & rsOrder("Quantity") & vbCrLf
          rsOrder.MoveNext
        loop
        
        if mid(qsItems,len(qsItems),1) <> "=" then
          qsItems = mid(qsItems,1,len(qsItems)-1)
        end if
        if mid(qsRev,len(qsRev),1) <> "=" then          
          qsRev = mid(qsRev,1,len(qsRev)-1)        
        end if
        
        rsOrder.close
        set rsOrder = nothing
        
        ' --------------------------------------------------------------------------------------
        ' Specific IP Address - If Failing to connect with LOS, check this script
        ' --------------------------------------------------------------------------------------

        %><!--#include virtual="/connections/connection_Literature_Order_System.asp"--><%
  
        strRemoteHostURL = "http://" & strRemoteHostName & strRemoteHostTargetFile
								response.Write strRemoteHostURL & "<BR>"
        strLocalHostReferrerFile = request.ServerVariables("SCRIPT_Name")
        strResponse              = ""
  
        strPost_QueryString      = strPost_QueryString & qsItems
        strPost_QueryString      = strPost_QueryString & qsRev
        
        
        Dim HTTPRequest

        Set HTTPRequest = Server.CreateObject("Msxml2.XMLHTTP.3.0") 
        
        if not Script_Debug then
          HTTPRequest.Open "POST", strRemoteHostURL , False 
          HTTPRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
          on error resume next          
          HTTPRequest.Send strPost_QueryString
          
          if err.Number <> 0 then
            strResponse = err.Description & "<BR>Order not received by remote server."
            bResponse   = err.Number
          else  
            strResponse = HTTPRequest.responseText
            bResponse   = HTTPRequest.status
          end if
          
          on error goto 0  
          
        else
          response.write "URL: " & strRemoteHostURL & "<P>"
          response.write "Form Data: " & Replace(strPost_QueryString,"&","<BR>&") & "<P>"
          response.flush
          response.end        
        end if  
        
       ' response.Write strResponse & "<BR>"
       ' response.End
        
        if bResponse = 200 then
          if Transfer_Debug then
            PostErrMessage = PostErrMessage & "<LI>Order Post Successful:<BR><BR>"
            PostErrMessage = PostErrMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR><BR>"
            PostErrMessage = PostErrMessage & "<P>" & strResponse & "<P>"
          end if
          
          ' Decode Response and Extract Order Number
  
          xmlResponse = ""
          xmlText     = ""
          if instr(1,LCase(strResponse),"<dataset>") > 0 then
            xmlText = Mid(strResponse,instr(1,LCase(strResponse),"<dataset>")+9)
          else  
            if CInt(Transfer_Debug) = CInt(True) then
              response.write "<SPAN CLASS=SMALLBOLDRED>Error: No &lt;dataset&gt; Opening Tag Found</SPAN><BR>"
            end if
            xmlResponse = strResponse
            ErrorFlag = true
          end if  
          if instr(1,LCase(strResponse),"</dataset>") > 0 then
            xmlText = Mid(xmlText,1,instr(1,LCase(xmlText),"</dataset>") - 1)
          else
            if CInt(Transfer_Debug) = CInt(True) then
              response.write "<SPAN CLASS=SMALLBOLDRED>Error: No &lt;/dataset&gt; Ending Tag Found</SPAN><BR>"
            end if
            xmlResponse = strResponse
            ErrorFlag = true
          end if
          
          if CInt(Transfer_Debug) = CInt(True) and not isblank(xmlResponse) then
            response.write "<BR><SPAN CLASS=SMALLBOLDRED>Returned Value From DCG (View Source to see tags):</SPAN><SPAN CLASS=SMALL>" & xmlResponse & "</SPAN><P>"            
          end if
  
          if CInt(ErrorFlag) = CInt(false) then
          
            xmlTags = "ACTION,ERRORMESG,ERRCODE,SCUSID,SORDNH,SIT,SITQTY,SPPKG,SPCAR"
            xmlTag  = Split(xmlTags,",")
            xmlData = Split(xmlTags,",")
            xmlTagCount = UBound(xmlTag)
    
            for TagCount = 0 to xmlTagCount
              if instr(1,LCase(xmlText),"<" & LCase(xmlTag(TagCount)) & ">") > 0 then
                xmlTemp = xmlText
                xmlTemp = Mid(xmlTemp,instr(1,LCase(xmlTemp),"<" & LCase(xmlTag(TagCount)) & ">") + Len(xmlTag(TagCount)) + 2 )
                xmlTemp = Mid(xmlTemp,1,instr(1,LCase(xmlTemp),"</" & LCase(xmlTag(TagCount)) & ">") - 1)
                xmlData(TagCount) = xmlTemp
              else
                xmlData(TagCount) = ""
              end if
            next    
              
            ' Grab Returning Data
            
            ErrCode      = GetXMLData(xmlTag, xmlData, "ERRCODE")
            ErrorMesg    = GetXMLData(xmlTag, xmlData, "ERRORMESG")
            Order_Number = GetXMLData(xmlTag, xmlData, "SORDNH")
                   
          end if
          
          if (isblank(ErrCode) or ErrCode = "0") and CInt(ErrorFlag) = CInt(false) then
  
            if not isblank(Order_Number_Results) then Order_Number_Results = Order_Number_Results & ","
            Order_Number_Results = Order_Number_Results & Account_Info_ID & "," & Order_Number
    
            ' Update Shoping Cart Data with Order Number
          
            SQLOrder = "SELECT * FROM Shopping_Cart_Lit WHERE Account_NTLogin='" & Logon_Name & "' AND Submit_Date is NULL"
            Set rsOrder = Server.CreateObject("ADODB.Recordset")
            rsOrder.Open SQLOrder, conn, 3, 3
    
            ' Clone Record and Update
            do while not rsOrder.EOF
              SQLKey = "(Site_ID, Account_ID, Account_NTLogin, Item_Number, Asset_ID, Quantity, Max_Limit, Shipping_Address_ID, Cart_Type, Cart_Date, Submit_Date, Order_Number, Region, Country) "
              SQLVal = "(" & rsOrder("Site_ID") & ", " & rsOrder("Account_ID") & ", '" & rsOrder("Account_NTLogin") & "', '" & rsOrder("Item_Number") & "', " & rsOrder("Asset_ID") & ", " & rsOrder("Quantity") & ", " & rsOrder("Max_Limit") & ", " & Account_Info_ID & ", '" & rsOrder("Cart_Type") & "', '" & rsOrder("Cart_Date") & "', '" & Order_Date & "', '" & Order_Number & "', " & Login_Region & ", '" & Login_Country & "' )"
              SQL = "INSERT INTO Shopping_Cart_Lit " & SQLKey & " VALUES " & SQLVal
              conn.execute SQL
              rsOrder.MoveNext
            loop
    
            rsOrder.close
            set rsOrder = nothing
            set SQL     = nothing
            set SQLKey  = nothing
            set SQLVal  = nothing
                               
            MailMessage = MailMessage & Translate("Order Number",Alt_Language,conn) & ": " & Replace(Order_Number,"FLUKECO","") & vbCrLf & vbCrLf
            MailMessage = MailMessage & MailItems & vbCrLf

          else
          
            'Mailer.Priority       = 1
            'Mailer.FromName       = Site_Description
            'Mailer.FromAddress    = "WebMail@Fluke.com"
            'Mailer.ReplyTo        = "WebMail@Fluke.com"

            msg.From = """" & Site_Description & """" & "WebMail@Fluke.com"
            msg.ReplyTo = "WebMail@Fluke.com"
    
            'Mailer.ClearAllRecipients
            'Mailer.AddRecipient "Kelly.Whitlock", "Kelly.Whitlock@Fluke.com"
            'Mailer.AddRecipient "Extranet Group", "ExtranetAlerts@Fluke.com"
            'Mailer.AddRecipient "David Vick", "dvick@hkm.dcgcentral.com"

            msg.To = """Extranet Group""" & "ExtranetAlerts@Fluke.com"
            msg.To = msg.To & ";" & """David Vick""" & "dvick@hkm.dcgcentral.com"
        
            PostErrMessage = "Order Post Error: " & bResponse & vbCrLf &_
                           "Calling Script: " & request.ServerVariables("Script_Name") & vbCrLf & vbCrLf &_
                           "The following debugging information was returned by the receiving Literature Order System:" & vbCrLf & vbCrLf &_
                           strResponse & vbCrLf
                           
            MailSubject = "Literature Order System - Posting Error"
            MailMessage = PostErrMessage
            MailMessage = MailMessage & chr(13) & chr(10)
            MailMessage = MailMessage & strPost_QueryString                
            Call Send_EMail
            
          end if
            
        else

          'Mailer.Priority       = 1
          'Mailer.FromName       = Site_Description
          'Mailer.FromAddress    = "WebMail@Fluke.com"
          'Mailer.ReplyTo        = "WebMail@Fluke.com"

          msg.From = """" & Site_Description & """" & "WebMail@Fluke.com"
          msg.ReplyTo = "WebMail@Fluke.com"
  
          'Mailer.ClearAllRecipients
'         ' Mailer.AddRecipient "Kelly.Whitlock", "Kelly.Whitlock@Fluke.com"
          'Mailer.AddRecipient "Extranet Group", "ExtranetAlerts@Fluke.com"
          'Mailer.AddRecipient "David Vick", "dvick@hkm.dcgcentral.com"
          
          msg.To = """Extranet Group""" & "ExtranetAlerts@Fluke.com"
          msg.To = msg.To & ";" & """David Vick""" & "dvick@hkm.dcgcentral.com"
      
          PostErrMessage = "Order Post Error: " & bResponse & vbCrLf &_
                           "Calling Script: " & request.ServerVariables("Script_Name") & vbCrLf & vbCrLf &_
                           "The following debugging information was returned by the receiving Literature Order System:" & vbCrLf & vbCrLf &_
                           strResponse & vbCrLf
                         
          MailSubject = "Literature Order System - Posting Error"
          MailMessage = PostErrMessage
       	  MailMessage = MailMessage & chr(13) & chr(10)
          MailMessage = MailMessage & strPost_QueryString               
          Call Send_EMail
  
        end if

      next

      MailMessage = MailMessage & DotLine & vbCrLf
      MailMessage = MailMessage & Translate("Sincerely,",Alt_Language,conn) & vbCrLf & vbCrLf & Translate(Site_Description,Alt_Language,conn) & " " & Translate("Support Team",Login_Language,conn)
      
      if bResponse = 200 then
        Call Send_EMail
      end if  
      
    end if
    
  case "noop"  
    
  case else

end select

Select case Action_Sequence
  case "transfer", "review", "account", "noop"
    %>
    <SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
      <!--
      if (document.all) {
        Shopping_Cart.window.resizeTo((screen.width) - 19 ,screen.height - 44);
      }
      else if (document.layers || document.getElementById) {
        if (Shopping_Cart.window.outerHeight < screen.height || Shopping_Cart.window.outerWidth < screen.width){
          Shopping_Cart.window.outerWidth = screen.width - 19;
          Shopping_Cart.window.outerHeight = screen.height - 44;        
        }
      }
      Shopping_Cart.window.moveTo((0),24);
      //-->
    </SCRIPT>
    <%
end select

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

if Action_Sequence <> "account" then

  SQL = "SELECT  Shopping_Cart_Lit.*, " &_
                "Calendar.Title AS Title, " &_
                "Calendar.Description AS Description, " &_
                "Calendar.Product AS Product, " &_
                "Calendar.Sub_Category AS Sub_Category, " &_
                "Calendar.Thumbnail AS Thumbnail, " &_
                "Site.Site_Code AS Site_Code, " &_
                "Language.Description AS Language, " &_
                "Calendar_Category.Category AS Category, " &_
                "Literature_Items_US.EFULFILLMENT AS Lit_Description, " &_
                "Literature_Items_US.[LANGUAGE] AS Lit_Language " &_
        "FROM   Literature_Items_US " &_
                "LEFT OUTER JOIN Shopping_Cart_Lit ON Literature_Items_US.ITEM = Shopping_Cart_Lit.Item_Number " &_
                "LEFT OUTER JOIN Language " &_
                "INNER JOIN Calendar ON Language.Code = Calendar.[Language] " &_
                "INNER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID ON Shopping_Cart_Lit.Asset_ID = Calendar.ID " &_
                "LEFT OUTER JOIN Site ON Shopping_Cart_Lit.Site_ID = Site.ID " &_
        "WHERE  (Shopping_Cart_Lit.Account_NTLogin = '" & Logon_Name & "') AND (Shopping_Cart_Lit.Submit_Date IS NULL) AND (dbo.Language.Enable = - 1) " &_
                "AND (dbo.Literature_Items_US.ACTIVE_FLAG = - 1) " &_
        "ORDER BY Shopping_Cart_Lit.ID DESC"   

  'response.write SQL
        
  Set rsCart = Server.CreateObject("ADODB.Recordset")
  rsCart.Open SQL, conn, 3, 3

  if not rsCart.EOF then

    if Review = False and Action_Sequence <> "transfer" then
      Call Shopping_Cart_Notes
      Call Shopping_Cart_Header
      Call Shopping_Cart_Data
      Call Shopping_Cart_Footer
    elseif Review = True and Action_Sequence <> "transfer" then
      Call Account_Information
      Call Shopping_Cart_Notes
      Call Shopping_Cart_Header
      Call Shopping_Cart_Data
      Call Shopping_Cart_Footer    
    elseif Review = True and Action_Sequence = "transfer" then
      Call Account_Information
      Call Shipping_Cart_Contact_Info
      Call Shopping_Cart_Notes
      Call Shopping_Cart_Header
      Call Shopping_Cart_Data
      Call Shopping_Cart_Footer

      ' Check if User is saving cart for another order or delete it
       
      if bResponse = 200 and request("Save_Cart") <> "yes" then
        SQL = "DELETE FROM Shopping_Cart_Lit WHERE Account_NTLogin='" & Logon_Name & "' AND Submit_Date IS NULL" 
        conn.execute SQL
        Session("Cart_Active") = False      
      end if  

    end if
    
  else
    response.write "<SPAN CLASS=SmallBoldRed><UL><LI>" & Translate("You have no items in your Shopping Cart",Login_Language,conn) & "</UL></SPAN><P>" & vbCrLf
    Session("Cart_Active") = False
  end if

  rsCart.close
  set rsCart = nothing
  
end if

response.write "</SPAN>" & vbCrLf

if Action_Sequence <> "account" then
  response.write "</FORM>" & vbCrLf
end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%response.flush%>

<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>
<!--
Shopping_Cart.window.focus();
//-->
</SCRIPT>
<%

' --------------------------------------------------------------------------------------
' Debug
' --------------------------------------------------------------------------------------

if CInt(Script_Debug) = CInt(True) or CInt(Transfer_Debug) = CInt(true) then

  with response
    .write "<P>------------------------------------------------<BR>" & vbCrLf
    .write "Debug - Key=Value Pairs<BR>" & vbCrLf
    .write "------------------------------------------------<P>" & vbCrLf
    .write "<SPAN CLASS=SmallBoldRed>QueryString Object</SPAN><BR>" & vbCrLf
    for each item in request.querystring
      .write item & "=" & request.querystring(item) & "<BR>" & vbCrLf
    next  
    .write "<BR>------------------------------------------------<P>" & vbCrLf
    .write "<SPAN CLASS=SmallBoldRed>Form Object</SPAN><BR>" & vbCrLf
    for each item in request.form
      .write item & "=" & request.form(item) & "<BR>" & vbCrLf
    next  
    .write "<BR>------------------------------------------------<P>" & vbCrLf
    .write "<SPAN CLASS=SmallBoldRed>Session Object</SPAN><BR>" & vbCrLf
    
    for each item in Session.Contents
      if IsObject(Session.Contents(item)) then
        if lcase(TypeName(Session.Contents(item))) = "dictionary" then
          dumpDictionary(Session.Contents(item)) & "<BR>"
        else
          .write TypeName(Session.Contents(item)) & "<BR>"
        end if
      else
        if IsArray(Session.Contents(item)) then
          for each n in Session.Contents(item)
            .write Session.Contents(item)(n) & "<BR>"
          next
        else
          .write Session.Contents(item) & "<BR>"
        end if
      end if  
    next
    
    .write "<BR>------------------------------------------------<P>" & vbCrLf

  end with

end if  

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

sub Menu_Bar

  if review = False and Action_Sequence <> "transfer" then

    Call Nav_Border_Begin
    response.write "<TABLE BORDER=""" & Border_Toggle & """ WIDTH=""" & Cart_Width & "%"" CELLPADDING=0 CELLSPACING=0>" & vbCrLf
    response.write "<TR>" & vbCrLf
    response.write "<TD CLASS=SmallBold ALIGN=RIGHT VALIGN=TOP NOWRAP>" & vbCrLf
  
    if Action_Sequence <> "noop" then
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Shop"" Title=""Continue Shopping"" VALUE=""" & " " & Translate("Continue Shopping",Login_Language,conn) & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""Shopping_Cart.window.blur(); window.opener.focus()"">"
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Account"" TITLE=""Checkout"" VALUE=""" & Translate("Checkout",Login_Language,conn) & " " & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action="
      if Account_Reviewed = True then
        response.write "Review"
      else
        response.write "Account"
      end if
      response.write "';"">"
    else
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Shop"" TITLE=""Continue Shopping"" VALUE=""" & " " & Translate("Continue Shopping",Login_Language,conn) & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""Shopping_Cart.window.blur(); window.opener.focus();"">"
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Account"" TITLE=""Edit Account"" VALUE=""" & Translate("Edit Account",Login_Language,conn) & "" & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Account';"">"
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Review"" TITLE=""Checkout"" VALUE=""" & Translate("Checkout",Login_Language,conn) & " " & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Review';"">"
    end if

    response.write "</TD>"    & vbCrLf
    response.write "</TR>"    & vbCrLf
    response.write "</TABLE>" & vbCrLf

    Call Nav_Border_End

  elseif Action_Sequence <> "transfer" then

    Call Nav_Border_Begin
    response.write "<TABLE BORDER=""" & Border_Toggle & """ WIDTH=""" & Cart_Width & "%"" CELLPADDING=0 CELLSPACING=0>" & vbCrLf
    response.write "<TR>" & vbCrLf
    response.write "<TD CLASS=SmallBold ALIGN=RIGHT VALIGN=TOP NOWRAP>" & vbCrLf

    if Action_Sequence <> "noop" and not isblank(Action_Sequence) then
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" TITLE=""Continue Shopping"" NAME=""Shop"" VALUE=""" & " " & Translate("Continue Shopping",Login_Language,conn) & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""Shopping_Cart.window.blur(); window.opener.focus();"">"
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" Title=""Edit Shopping Cart"" NAME=""Update"" VALUE=""" & Translate("Edit Cart",Login_Language,conn) & "" & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=NoOp';"">"
      response.write "&nbsp;&nbsp;"
    end if
    response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" TITLE=""Edit Account"" NAME=""Account"" VALUE=""" & Translate("Edit Account",Login_Language,conn) & "" & """ "
    response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Account';"">"
    if isblank(ErrMessage) then
      response.write "&nbsp;&nbsp;"
      response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" TITLE=""Submit Order"" NAME=""Submit"" VALUE=""" & Translate("Submit Order",Login_Language,conn) & """>"    
      if Action_Count = 0 then
        response.write "<INPUT TYPE=""HIDDEN"" NAME=""ACTION"" VALUE=""Transfer"">"
        Action_Count = 1
      end if    
    end if
    
    response.write "</TD>"    & vbCrLf
    response.write "</TR>"    & vbCrLf
    response.write "</TABLE>" & vbCrLf

    Call Nav_Border_End
          
  end if
  
end sub

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Notes

  response.write "<DIV ALIGN=CENTER>" & vbCrLf
  response.write "<TABLE BORDER=""" & Border_Toggle & """ WIDTH=""" & Cart_Width & "%"" CELLPADDING=0 CELLSPACING=0 BGCOLOR=""Black"">" & vbCrLf
  response.write "<TR>" & vbCrLf

  response.write "<TD CLASS=Small WIDTH=""100%"" BGCOLOR=""White"">"

  if review = False and Action_Sequence <> "transfer" then
    response.write "<SPAN CLASS=HEADING3>" & Translate("Literature Order",Login_Language,conn) & " - " & Translate("Shopping Cart",Login_Language,conn) & "</SPAN><BR><SPAN CLASS=HEADING6>" & Translate("Review Order",Login_Language,conn) & "</SPAN><P>"
    response.write "<UL><LI>" & Replace(Replace(Translate("Review your Shopping Cart items below.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;") & "<BR>"
    response.write "<LI>" & Replace(Replace(Translate("To change the quantity of an item, delete the current quantity, enter a new quantity, then click on the [ Update ] button.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;")
    response.write "<LI>" & Replace(Replace(Translate("To delete an item from your shopping cart, click on the [ Delete ] button.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;")
    if Action_Sequence <> "noop" then
      response.write "<LI>" & Replace(Replace(Translate("After you are done updating your Shopping Cart, click on the [ Checkout ] button to continue.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;")
    else
      response.write "<LI>" & Replace(Replace(Translate("After you are done updating your Shopping Cart, click on the [ Checkout ] button to continue.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;")
    end if
    response.write "</UL>"
  end if
  
  ' Standard Error Messages
  if not isblank(ErrMessage) then
    response.write "<A NAME=""Cart_List""></A>" & vbCrLf
    response.write "<SPAN CLASS=SmallBoldRed><BR><UL>" & ErrMessage & "</UL></SPAN>" & vbCrLf
    'ErrMessage = ""
  end if
  
  ' Post Error Messages
  if not isblank(PostErrMessage) then
    if isblank(ErrMessage) then
      response.write "<A NAME=""Cart_List""></A>" & vbCrLf
    end if  
    response.write "<SPAN CLASS=SmallBoldRed><BR><UL>" & PostErrMessage & "</UL></SPAN>" & vbCrLf
    PostErrMessage = ""
  end if

  if Ship_To_Alternates > 0 and Action_Sequence <> "transfer" then
    response.write "<BR><B>" & Translate("Note",Login_Language,conn) & "</B>: " & Translate("Selection of one or more &quot;ship to&quot; address will result in separate orders for each of the &quot;ship to&quot; addresses containing the line items and quantities contained in your cart as shown below.",Login_Language,conn) & "<P>"
  end if  
  
  response.write "</TD>"
  response.write "</TR>"
  
  response.write "<TR>" & vbCrLf

  response.write "<TD CLASS=Small WIDTH=""100%"" BGCOLOR=""White"">"

  response.write "<B>" & Translate("Total Number of Line Items",Login_Language,conn) & ": " & rsCart.recordcount & "</B><P>"
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf
  
  response.write "</TABLE>"

end sub

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Header

  Call Menu_Bar
  response.write "<BR>" & vbCrLf

  if isblank(ErrMessage) and isBlank(PostErrMessage) then
    response.write "<A NAME=""Cart_List""></A>" & vbCrLf
  end if
  response.write "<DIV ALIGN=CENTER>" & vbCrLf
   
  response.write "<TABLE BORDER=""" & Border_Toggle & """ WIDTH=""" & Cart_Width & "%"" CELLPADDING=0 CELLSPACING=0 BGCOLOR=""Black"">" & vbCrLf
  response.write "<TD CLASS=Small BGCOLOR=""white"" ALIGN=RIGHT VALIGN=TOP WIDTH=""100%"">"

  Call Table_Begin

  response.write "<TABLE BGCOLOR=Black BORDER=""" & Border_Toggle & """ WIDTH=""100%"" CELLPADDING=4 CELLSPACING=1>" & vbCrLf
  response.write "<TR>" & vbCrLf
  if Action_Sequence = "transfer" then
    response.write "<TD BGCOLOR=""Black"" CLASS=SmallBoldGold WIDTH=""5%"">&nbsp;" & Translate("Item",Login_Language,conn) & "</TD>" & vbCrLf
  end if  
  response.write "<TD BGCOLOR=""Black"" CLASS=SmallBoldGold WIDTH=""5%"">&nbsp;" & Translate("Quantity",Login_Language,conn) & "</TD>" & vbCrLf
  if Action_Sequence <> "transfer" and Review = False then
    response.write "<TD BGCOLOR=""Black"" CLASS=SmallBoldGold WIDTH=""5%"">" & Translate("Thumbnail",Login_Language,conn) & "</TD>" & vbCrLf
  end if
  response.write "<TD BGCOLOR=""Black"" CLASS=SmallBoldGold WIDTH=""90%"">" & Translate("Description",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf
  
end sub  

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Data

  rsCart.MoveFirst
  Line_Number = 0

  do while not rsCart.EOF
    Line_Number = Line_Number + 1

    response.write "<TR>" & vbCrLf
    response.write "<A NAME=""Q" & rsCart("ID") & """></A>"                             ' Index Anchor
    ' Line Number
    if Action_Sequence = "transfer" then
      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""Silver"">" & Line_Number & "</TD>"
    end if  

    ' Quantity / Update
    response.write "<TD CLASS=Small ALIGN=CENTER"
    if not isblank(request("Cart_ID")) then
      if CLng(request("Cart_ID")) = CLng(rsCart("ID")) or (Action_Sequence = "add" and Line_Number = 1) then
        response.write " BGCOLOR=""#F5DEB3"""
      else
        response.write " BGCOLOR=""Silver"""
      end if
    elseif Action_Sequence = "transfer" or Action_Sequence = "review" then
      if isblank(rsCart("Quantity")) or rsCart("Quantity") = 0 then
        response.write " BGCOLOR=""#FFFF99"""
      else  
        response.write " BGCOLOR=""White"""
      end if  
    else
      response.write " BGCOLOR=""Silver"""  
    end if
    response.write ">"
    
    if Action_Sequence <> "transfer" and Action_Sequence <> "review" then
      response.write "<INPUT SIZE=""6"" MAXLENGTH=""4"" Title=""Order Quantity"" TYPE=""Text"" NAME=""Q" & rsCart("ID") & """ VALUE=""" & rsCart("Quantity") & """"
      if Review = True or Action_Sequence = "transfer" then
        response.write " DISABLED"
      end if
      if isblank(rsCart("Quantity")) or rsCart("Quantity") = 0 then
        response.write " CLASS=ZeroValue"
      else
        response.write " CLASS=SmallBold"
      end if
      response.write ">"
      if Review = False and Action_Sequence <> "transfer" then
        response.write "<BR>"
'        response.write "&nbsp;&nbsp;"
        response.write "<INPUT TYPE=""Button"" CLASS=NavLeftHighlight1 TITLE=""Update Quantity"" NAME=""Update"" VALUE=""" & Translate("Update",Login_Language,conn) & """ "
        response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Update&Cart_ID=" & rsCart("ID") & "&Quantity=' + document.Cart_Data.Q" & rsCart("ID") & ".value + '#Q" & rsCart("ID") & "';"">"
      end if
      
      ' Delete Item
      if Review = False then
        response.write "<P>"
        response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" TITLE=""Delete Item"" NAME=""Delete"" VALUE=""" & Translate("Delete",Login_Language,conn) & """ "
        response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Delete&Cart_ID=" & rsCart("ID") & "'"">"
      end if
      
    else
      response.write rsCart("Quantity")
    end if  
    response.write "</TD>" & vbCrLf

    ' Thumbnail
    
    if Review = False then
      response.write "<TD BGCOLOR=""White"" CLASS=Small>"
      if not isblank(rsCart("Thumbnail")) then
        response.write "<IMG SRC=""" & "/" & rsCart("Site_Code") & "/" & rsCart("Thumbnail") & """ WIDTH=""60"" BORDER=1>"
      else
        response.write "&nbsp;"
      end if
      response.write "</TD>" & vbCrLf
    end if  
   
    ' Description

    response.write "<TD BGCOLOR=""White"" CLASS=Small VALIGN=TOP>"

    ' Item number
    response.write Translate("Item Number",Login_Language,conn) & ": <SPAN CLASS=SmallBoldRed>" & rsCart("Item_Number") & "</SPAN>"
    if Review = False then
      response.write "<P>"
    else
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("Title",Login_Language,conn) & ": "  
    end if

    if not isblank(rsCart("Title")) then      
      response.write "<B>" & rsCart("Title") & "</B>"
    else
      response.write "<B>" & ProperCase(rsCart("Lit_Description")) & "</B>"
    end if  

    if Review = False then
      response.write "<BR>"
      if not isblank(rsCart("Product")) then
        response.write rsCart("Product") & "<BR>"
        response.write Translate("Langugage",Login_Language,conn) & ": " & rsCart("Language")
      else
        response.write Translate("Langugage",Login_Language,conn) & ": " & ProperCase(rsCart("Lit_Language"))
      end if  
    end if  
    response.write "</TD>" & vbCrLf
      
    response.write "</TR>" & vbCrLf
    rsCart.MoveNext
  loop

end sub

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Footer
 
  response.write "</Table>" & vbCrLf

  Call Table_End
  
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf

  response.write "</Table>" & vbCrLf
  
  response.write "<BR>" & vbCrLf
    
  Call Menu_Bar  

  response.write "</DIV>" & vbCrLf
  
end sub

' --------------------------------------------------------------------------------------

sub Shipping_Cart_Contact_Info_Text

  Order_Status_Info = "<LI>" & Translate("All orders received will be shipped the following business day.",Login_Langugage,conn) & "</LI>" &_
                      "<LI>" & Translate("You will receive an order acknowledge email with instructions for reviewing the status of your order.",Login_Language,conn) & "</LI>" 
                      
  Order_Status_Info_Alt = "<LI>" & Translate("All orders received will be shipped the following business day.",Alt_Langugage,conn) & "</LI>" &_
                          "<LI>" & Translate("You will receive an order acknowledge email with instructions for reviewing the status of your order.",Alt_Language,conn) & "</LI>" 

  if Ship_to_Count > 0 then
    Order_Status_Info = Order_Status_Info & "<LI>" & Translate("Each 'ship to' address will also receive an order acknowledge email with instructions for reviewing the status of the order you have placed on their behalf.",Login_Language,conn) & "</LI>"
    Order_Status_Info_Alt = Order_Status_Info_Alt & "<LI>" & Translate("Each 'ship to' address will also receive an order acknowledge email with instructions for reviewing the status of the order you have placed on their behalf.",Alt_Language,conn) & "</LI>"    
  end if
    
end sub

' --------------------------------------------------------------------------------------

sub Shipping_Cart_Contact_Info

  response.write "<BR>" & vbCrLf
  response.write "<TABLE ALIGN=CENTER WIDTH=""" & Cart_Width & "%"" BORDER=""" & Border_Toggle & """ BGCOLOR=""Silver"" CELLSPACING=""0"" CELLPADDING=""4"">" & vbCrLf
  response.write "<TR>" & vbCrLf
  response.write "<TD CLASS=Small BGCOLOR=WHITE WIDTH=""100%"">" & vbCrLf
  response.write "<UL>"
  Call Shipping_Cart_Contact_Info_Text
  response.write Replace(Order_Status_Info,"</LI>","</LI>" & vbCrLf)
  response.write "</UL>"
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf & vbCrLf

end sub

' --------------------------------------------------------------------------------------
  
sub Account_Information

  response.write "<DIV ALIGN=CENTER>" & vbCrLf
  response.write "<TABLE WIDTH=""" & Cart_Width & "%"" CELLPADDING=4 CELLSPACING=0 BGCOLOR=""WHITE"" BORDER=""" & Border_Toggle & """>" & vbCrLf
  response.write "<TR>" & vbCrLf
  response.write "<TD CLASS=Small COLSPAN=2 BGCOLOR=""White"">" & vbCrLf
  response.write "<P>" & vbCrLf
  if Action_Sequence <> "transfer" then
    response.write "<SPAN CLASS=HEADING3>" & Translate("Literature Order",Login_Language,conn) & " - " & Translate("Shopping Cart",Login_Language,conn) & "</SPAN><BR><SPAN CLASS=HEADING6>" & Translate("Checkout",Login_Language,conn) & "</SPAN>" & vbCrLf
  else
    response.write "<SPAN CLASS=HEADING3>" & Translate("Literature Order",Login_Language,conn) & "</SPAN><BR><SPAN CLASS=HEADING6>" & Translate("Your order has been submitted.",Login_Language,conn) & "</SPAN>" & vbCrLf
  end if
  response.write "</TD>"  
  response.write "</TR>" & vbCrLf
      
  if Action_Sequence <> "transfer" then
    response.write "<TR>" & vbCrLf
    response.write "<TD CLASS=Small NOWRAP BGCOLOR=""White"">" & vbCrLf
    response.write "&nbsp;"
    response.write "</TD>"
    response.write "<TD WIDTH=""80%"">&nbsp;</TD>"
  else 
    response.write "<TR>" & vbCrLf
    response.write "<TD CLASS=Small NOWRAP>" & vbCrLf
'    response.write "<FORM NAME=""Print_It"">"
'    Call Nav_Border_Begin
'    response.write "<INPUT TYPE=""BUTTON"" TITLE=""Print Order"" CLASS=NavLeftHighlight1 VALUE=""" & Translate("Print",Login_Language,conn) & """ NAME=""Print"" LANGUAGE=""JavaScript"" onClick=""PrintIt()"">"
'    Call Nav_Border_End
'    response.write "</FORM>"
    response.write "</TD>"
    response.write "<TD WIDTH=""80%"" ALIGN=RIGHT VALIGN=TOP NOWRAP>"
    response.write "<FORM NAME=""Close_It"">"
    Call Nav_Border_Begin
    response.write "<INPUT TYPE=""BUTTON"" CLASS=NavLeftHighlight1 TITLE=""Print Order Information"" VALUE=""" & Translate("Print",Login_Language,conn) & """ NAME=""Print"" LANGUAGE=""JavaScript"" onClick=""PrintIt()"">&nbsp;&nbsp;&nbsp;"
    response.write "<INPUT ALIGN=RIGHT CLASS=NavLeftHighlight1 TITLE=""Close Shopping Cart Window"" TYPE=""BUTTON"" NAME=""Portal"" VALUE=""" & Translate("Close Window",Login_Language,conn) & "" & """ LANGUAGE=""Javascript"" ONCLICK=""window.close();"">"
    Call Nav_Border_End
    response.write "</FORM>"
    response.write "</TD>"
  end if

  response.write "</TR>" & vbCrLf
  
  response.write "<TR>" & vbCrLf
  response.write "<TD CLASS=SMALL BGCOLOR=WHITE VALIGN=TOP>" & vbCrLf
  
  call table_begin
  
  response.write "<TABLE bgcolor=WHITE WIDTH=""100%"" BORDER=""" & Border_Toggle & """ CELLPADDING=4>" & vbCrLf
  response.write "<TR>" & vbCrLf
  
  response.write "<TD WIDTH=""8%"" NOWRAP CLASS=SmallBold BGCOLOR=""#CECECE"" VALIGN=TOP>" & vbCrLf
  response.write Translate("Ordered by",Login_Language,conn) & ":" & vbCrLf
  response.write "</TD>" & vbCrLf
 
  response.write "<TD WIDTH=""92%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP>" & vbCrLf
  
  Account_Info_ID = 0
  Call Get_Account_Information

  response.write FormatFullName(FirstName, MiddleName, LastName) & "<BR>" & vbCrLf
  response.write Company & "<BR>" & vbCrLf
  response.write Business_Address & "<BR>" & vbCrLf
  if not isblank(Business_Address_2) then
    response.write Business_Address_2 & "<BR>" & vbCrLf
  end if
  response.write Business_City
  if not isblank(Business_State) then
    response.write ", " & Business_State
  elseif not isblank(Business_State_Other) then
    response.write ", " & Business_State_Other
  end if
  response.write "&nbsp;&nbsp;" & Business_Postal_Code & "<BR>" & vbCrLf
  response.write Business_Country & "<P>" & vbCrLf
  response.write Translate("Phone",Login_Language,conn) & ": " & Business_Phone & "<BR>" & vbCrLf
  response.write Translate("Email",Login_Language,conn) & ": " & Business_Email & vbCrLf
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf

  call Table_End
  response.write "</TD>" & vbCrLf

  ' Shipping Information Account Holder

  response.write "<TD CLASS=Small BGCOLOR=WHITE VALIGN=TOP ALIGN=RIGHT>" & vbCrLf

  Call Table_Begin
  
  response.write "<TABLE BGCOLOR=WHITE WIDTH=""100%"" BORDER=""" & Border_Toggle & """ CELLPADDING=4 VALIGN=TOP>" & vbCrLf
  response.write "<TR>" & vbCrLf
  
  response.write "<TD WIDTH=""8%"" NOWRAP CLASS=SmallBold BGCOLOR=""#CECECE"" ALIGN=RIGHT VALIGN=TOP>" & vbCrLf
  response.write Translate("Ship to",Login_Language,conn) & ":" & vbCrLf
  response.write "</TD>" & vbCrLf
  
  response.write "<TD WIDTH=""92%"" NOWRAP CLASS=SmallBold BGCOLOR=""#F6F6F6"" VALIGN=TOP>" & vbCrLf

  response.write "<TABLE WIDTH=""100%"" BGCOLOR=WHITE CELLPADDING=0 CELLSPACING=1 BORDER=""" & Border_Toggle & """>" & vbCrLf
  
  ' Display Account Ship To
  if Action_Sequence <> "transfer" or (Action_Sequence = "transfer" and mid(request("Ship_To"),1,1) = "0") then
  
    Address_Text = FormatFullName(FirstName, MiddleName, LastName) & "   [ " &_
                   Company & ", " &_
                   Shipping_Address & " " &_
                   Shipping_Address_2 & ", " &_
                   Shipping_City & ", " &_
                   Shipping_State & " " &_
                   Shipping_State_Other & ", " &_
                   Shipping_Postal_Code & ", " &_
                   Shipping_Country & " ]"
    
    response.write "<TR>" & vbCrLf
  
    if Action_Sequence <> "transfer" then
      response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP ALIGN=CENTER>" & vbCrLf
      response.write "<INPUT ALIGN=""TEXTTOP"" TITLE=""Ship To"" TYPE=""CHECKBOX"" NAME=""Ship_To"" VALUE=""0"">"      
      response.write "</TD>" & vbCrLf 
    end if    
 
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"
    response.write "&nbsp;&nbsp;"
    response.write "<SPAN TITLE=""" & Address_Text & """>"
    response.write FormatFullName(FirstName, MiddleName, LastName)
    response.write "</SPAN>"
    response.write "</TD>" & vbCrLf
    

    if Screen_Width >= 1024 then    
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"
      response.write "&nbsp;&nbsp;" & Company & vbCrLf
      response.write "</TD>" & vbCrLf
    end if  
    
    if Action_Sequence <> "transfer" and Screen_Width >= 1024 then
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
      response.write "&nbsp;&nbsp;" & Shipping_Address & vbCrLf
      response.write "</TD>" & vbCrLf
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
      if not isblank(Shipping_Address_2) then
        response.write "&nbsp;&nbsp;" & Shipping_Address_2 & vbCrLf
      else
        response.write "&nbsp;&nbsp;"  
      end if
      response.write "</TD>" & vbCrLf
    end if
      
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
    response.write "&nbsp;&nbsp;" & Shipping_City & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
    if not isblank(Shipping_State) then
      response.write "&nbsp;&nbsp;" & Shipping_State
    elseif not isblank(Shipping_State_Other) then
      response.write "&nbsp;&nbsp;" & Shipping_State_Other
    else
      response.write "&nbsp;&nbsp;"
    end if
    response.write "</TD>" & vbCrLf
  
    if Action_Sequence <> "transfer"  and Screen_Width >= 1024 then
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE ALIGN=RIGHT>"  
      response.write "&nbsp;&nbsp;" & Shipping_Postal_Code & vbCrLf
      response.write "</TD>" & vbCrLf
    end if
      
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
    response.write "&nbsp;&nbsp;" & Shipping_Country & vbCrLf
    response.write "</TD>" & vbCrLf & vbCrLf
  
    if Action_Sequence <> "transfer" then
      response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" ALIGN=RIGHT VALIGN=MIDDLE>"
      response.write "&nbsp;&nbsp" & vbCrLf
      response.write "<A HREF=""JavaScript:void(0);"" TITLE=""Edit"" LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Account'; return false;""><SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;</SPAN></A>"
    else
      response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#CECECE"" ALIGN=RIGHT VALIGN=RIGHT>"
      ThisOrder     = Split(Order_Number_Results,",")
      ThisOrderFlag = false
      for n = 0 to UBound(ThisOrder) step 2
        if ThisOrder(n) = 0 then
          response.write "&nbsp;&nbsp;" & ThisOrder(n+1)
          ThisOrderFlag = true
          exit for
        end if
      next
      if CInt(ThisOrderFlag) = CInt(false) then response.write "&nbsp;"
    end if    
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
  end if  

  SQLShip = "SELECT * FROM Shopping_Cart_Ship_To WHERE NTLogin='" & Session("LOGON_USER") & "' AND Disabled=" & CInt(False) & " ORDER BY Lastname, FirstName"
  Set rsShip = Server.CreateObject("ADODB.Recordset")
  rsShip.Open SQLShip, conn, 3, 3
  
  Ship_To_Alternates = rsShip.RecordCount
   
  do while not rsShip.EOF

    ' Shipping Information Account Holder
    
    if Action_Sequence = "transfer" then
      Ship_To_Display = false
      for s = 0 to Ship_To_Count
        if CLng(rsShip("ID")) = CLng(Ship_To_Addresses(s)) then
          Ship_To_Display = true
          exit for
        end if
      next
    else
      Ship_To_Display = true  
    end if
    
    if Ship_To_Display = true then
      response.write "<TR>" & vbCrLf

      if Action_Sequence <> "transfer" then
        response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP ALIGN=CENTER>"
        response.write "<INPUT ALIGN=""TEXTTOP"" TITLE=""Select Ship To Address For This Order"" TYPE=""CHECKBOX"" NAME=""Ship_To"" VALUE=""" & rsShip("ID") &""">"
        response.write "</TD>" & vbCrLf
      end if    

      Address_Text = FormatFullName(rsShip("FirstName"), rsShip("MiddleName"), rsShip("LastName")) & "   [ " &_
                     rsShip("Company") & ", " &_
                     rsShip("Shipping_Address") & " " &_
                     rsShip("Shipping_Address_2") & ", " &_
                     rsShip("Shipping_City") & ", " &_
                     rsShip("Shipping_State") & " " &_
                     rsShip("Shipping_State_Other") & ", " &_
                     rsShip("Shipping_Postal_Code") & ", " &_
                     rsShip("Shipping_Country") & " ]"
                                      
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"
      response.write "&nbsp;&nbsp;"
      response.write "<SPAN TITLE=""" & Address_Text & """>"
      
      
      response.write FormatFullName(rsShip("FirstName"), rsShip("MiddleName"), rsShip("LastName"))
      response.write "</SPAN>" & vbCrLf
      response.write "</TD>" & vbCrLf

      if Screen_Width >= 1024 then      
        response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"
        response.write "&nbsp;&nbsp;" & rsShip("Company") & vbCrLf
        response.write "</TD>" & vbCrLf
      end if  
 
      if Action_Sequence <> "transfer" and Screen_Width >= 1024 then
        response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
        response.write "&nbsp;&nbsp;" & rsShip("Shipping_Address") & vbCrLf
        response.write "</TD>" & vbCrLf
        response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
        if not isblank(rsShip("Shipping_Address_2")) then
          response.write "&nbsp;&nbsp;" & rsShip("Shipping_Address_2") & vbCrLf
        else
          response.write "&nbsp;"  
        end if
        response.write "</TD>" & vbCrLf
      end if
        
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
      response.write "&nbsp;&nbsp;" & rsShip("Shipping_City") & vbCrLf
      response.write "</TD>" & vbCrLf
      
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
      if not isblank(rsShip("Shipping_State")) then
        response.write "&nbsp;&nbsp;" & rsShip("Shipping_State")
      elseif not isblank(rsShip("Shipping_State_Other")) then
        response.write "&nbsp;&nbsp;" & rsShip("Shipping_State_Other")
      else
        response.write "&nbsp;&nbsp;"
      end if
      response.write "</TD>" & vbCrLf

      if Action_Sequence <> "transfer" and Screen_Width >= 1024 then    
        response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE ALIGN=RIGHT>"  
        response.write "&nbsp;&nbsp;" & rsShip("Shipping_Postal_Code") & vbCrLf
        response.write "</TD>" & vbCrLf
      end if
        
      response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=MIDDLE>"  
      response.write "&nbsp;&nbsp;" & rsShip("Shipping_Country") & vbCrLf
      response.write "</TD>" & vbCrLf
    

      if Action_Sequence <> "transfer" then  
        response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" ALIGN=RIGHT VALIGN=MIDDLE>"
        response.write "<A HREF=""JavaScript:void(0);"" ONCLICK=""var ShipTo = window.open('/sw-common/SW-Ship_To_Edit.asp?Language=" & Login_Language & "&Site_ID=" & Site_ID & "&Ship_To_ID=" & rsShip("ID") &  "&Sequence=1','ShipTo','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=380,height=494,left=350,top=100'); ShipTo.focus(); return false;"" TITLE="" " & Address_Text & " ""><SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;</SPAN></A>"
      else
        response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#CECECE"" ALIGN=RIGHT VALIGN=RIGHT>"
        ThisOrder     = Split(Order_Number_Results,",")
        ThisOrderFlag = false
        for n = 0 to UBound(ThisOrder) step 2
          if CLng(ThisOrder(n)) = CLng(rsShip("ID")) then
            response.write "&nbsp;&nbsp;" & ThisOrder(n+1)
            ThisOrderFlag = true
            exit for
          end if
        next
        if CInt(ThisOrderFlag) = CInt(false) then response.write "&nbsp;"
      end if    
      response.write "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
      
    end if  
    
    rsShip.MoveNext
  
  loop
  
  rsShip.close
  set rsShip = nothing
  
  if Screen_Width < 1024 then
    Num_Cols = 5
  else
    Num_Cols = 9  
  end if  
  
  if Action_Sequence <> "transfer" then  
    response.write "<TR>" & vbCrLf
    response.write "<TD COLSPAN=" & Num_Cols & " CLASS=Small BGCOLOR=""#F6F6F6"" ALIGN=Left VALIGN=MIDDLE>"
    response.write "<TABLE BGCOLOR=""#CECECE"" CELLPADDING=2 WIDTH=""100%"">"
    response.write "<TR>"
    response.write "<TD BGCOLOR=""F6F6F6"" CLASS=SMALL>"
    response.write Translate("The top &quot;ship to&quot; address corresponds to the &quot;ordered by&quot; account holder and can be edited by clicking on the [Edit Account] button.",Login_Language,conn)
    response.write "&nbsp;&nbsp;"
    response.write Translate("Additional &quot;ship to&quot; address can be added, edited or deleted by using the corresponding [Add] or [Edit] buttons.",Login_Languages,conn)
    if Screen_Width < 1024 then
      response.write "<P>" & Translate("Hold your mouse over the &quot;ship to&quot; name to display expanded address information.",Login_Language,conn)
    end if  
    response.write "</TD>"
    response.write "</TR>"
    response.write "</TABLE>"

    response.write "</TD>"

    response.write "<TD COLSPAN=1 NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" ALIGN=RIGHT VALIGN=MIDDLE>"  
    response.write "<A HREF=""JavaScript:void(0);"" TITLE=""Add"" ONCLICK=""var ShipTo = window.open('/sw-common/SW-Ship_To_Edit.asp?Ship_To_ID=0&Sequence=0','ShipTo','fullscreen=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=380,height=494,left=350,top=100'); ShipTo.focus(); return false;""><SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Add",Login_Language,conn) & "&nbsp;</SPAN></A>"
    response.write "</TD>" & vbCrLf

    response.write "</TR>" & vbCrLf

    response.write "<TR>" & vbCrLf
    response.write "<TD WIDTH=""1%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP ALIGN=CENTER>"
    response.write "<INPUT TYPE=""CHECKBOX"" TITLE=""Save Shopping Cart Item for Next Order"" ALIGN=""TEXTTOP"" NAME=""Save_Cart"" VALUE=""yes"">"
    response.write "</TD>"
    response.write "<TD COLSPAN=" & Num_Cols & " CLASS=Small BGCOLOR=""#F6F6F6"" ALIGN=Left VALIGN=MIDDLE>"
    response.write Translate("Click this checkbox to save your shopping cart items for future ordering after you submit this order.",Login_Language,conn)
    response.write "</TD>"
    response.write "</TR>"

  end if  

  response.write "</TABLE>" & vbCrLf

  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf

  Call Table_End
  
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf
    
  response.write "</DIV>" & vbCrLf
  
end sub

' --------------------------------------------------------------------------------------

sub Get_Account_Information

  select case Account_Info_ID
    case 0
      SQL =  "SELECT * FROM UserData WHERE NTLogin='" & Logon_Name & "' AND Site_ID=" & Site_ID & " AND NewFlag=0"
    case else
      SQL =  "SELECT * FROM Shopping_Cart_Ship_To WHERE ID=" & Account_Info_ID
  end select    

  Set rsLogin = Server.CreateObject("ADODB.Recordset")
  rsLogin.Open SQL, conn, 3, 3
  
  if not rsLogin.EOF then
  
    if isblank(rsLogin("Shipping_Address")) or isblank(rsLogin("Shipping_City")) or isblank(rsLogin("Shipping_Postal_Code")) or isblank(rsLogin("Shipping_Country")) then
      ErrMessage = ErrMessage &_
                   "<LI>" & Translate("We are unable to process your order, missing shipping information.",Login_Language,conn) &_
                   "<LI>" & Replace(Replace(Translate("Click on the [ Edit Account ] button to provide this information.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;")
    else
      ' User Information
      if not isblank(rsLogin("FirstName"))  then FirstName = rsLogin("FirstName") else FirstName = ""
      if not isblank(rsLogin("MiddleName")) then MiddleName = rsLogin("MiddleName") else MiddleName = ""
      if not isblank(rsLogin("LastName"))   then LastName = rsLogin("LastName") else LastName = ""
      if not isblank(rsLogin("Company"))    then Company = rsLogin("Company") else Company = ""
      if not isblank(rsLogin("Job_Title"))  then Job_Title = rsLogin("Job_Title") else Job_Title = ""
      PPAccount_ID = rsLogin("ID")
      
      Business_Phone = rsLogin("Business_Phone")
      Business_Email = rsLogin("Email")
            
      if Account_Info_ID = 0 then
        ' Business Address Information
        Business_Address = rsLogin("Business_Address")
        if not isblank(rsLogin("Business_Address_2")) then
          Business_Address_2 = rsLogin("Business_Address_2")
        else
          Business_Address_2 = ""  
        end if
        Business_City = rsLogin("Business_City")
        if rsLogin("Business_State") <> "ZZ" then
          Business_State = rsLogin("Business_State")
        else
          Business_State = ""
        end if    
        if not isblank(rsLogin("Business_State_Other")) then
          Business_State_Other = rsLogin("Business_State_Other")
        else
          Business_State_Other = ""  
        end if  
        Business_Postal_Code = rsLogin("Business_Postal_Code")
        Business_Country = rsLogin("Business_Country")
      end if
      
      ' Shipping Address Information
      Shipping_Address = rsLogin("Shipping_Address")
      if not isblank(rsLogin("Shipping_Address_2")) then
        Shipping_Address_2 = rsLogin("Shipping_Address_2")
      else
        Shipping_Address_2 = ""  
      end if
      Shipping_City = rsLogin("Shipping_City")
      if rsLogin("Shipping_State") <> "ZZ" then
        Shipping_State = rsLogin("Shipping_State")
      else
        Shipping_State = ""
      end if    
      if not isblank(rsLogin("Shipping_State_Other")) then
        Shipping_State_Other = rsLogin("Shipping_State_Other")
      else
        Shipping_State_Other = ""  
      end if  
      Shipping_Postal_Code = rsLogin("Shipping_Postal_Code")
      Shipping_Country = rsLogin("Shipping_Country")
      
      if Account_Info_ID <> 0 then
        if not isblank(rsLogin("Comment")) then
          Shipping_Comment = rsLogin("Comment")
        else
          Shipping_Comment = ""
        end if  
      else
        Shipping_Comment = ""
      end if    
    end if
  end if
  
  rsLogin.close
  set rsLogin = Nothing

end sub

'--------------------------------------------------------------------------------------

sub Build_MailMessage_Order_By

    MailMessage = MailMessage & FormatFullName(FirstName, MiddleName, LastName) & vbCrLf
          
    if not isblank(Job_Title) then
      MailMessage = MailMessage & Job_Title & vbCrLf & vbCrLf         
    end if
    if not isblank(Company) then  
      MailMessage = MailMessage & Company & vbCrLf
    end if  
    if not isblank(Business_Address) then
      MailMessage = MailMessage & Business_Address & vbCrLf
    end if  
    if not isblank(Business_Address_2) then
      MailMessage = MailMessage & Business_Address_2 & vbCrLf                    
    end if  
    MailMessage = MailMessage & Business_City
    if Business_State <> "ZZ" then
      MailMessage = MailMessage & ", " & Business_State & "  "
    end if
    if not isblank(Business_State_Other) then
      MailMessage = MailMessage & ", " & Business_State_Other & "  "
    end if  
    MailMessage = MailMessage & Business_Postal_Code & vbCrLf
    MailMessage = MailMessage & Business_Country & vbCrLf & vbCrLf
    if not isblank(Business_Phone) then
      MailMessage = MailMessage & Translate("Phone",Alt_Language,conn) & ": " & FormatPhone(Business_Phone) & vbCrLf
    end if  
    if not isblank(Business_Fax) then
      MailMessage = MailMessage & Translate("Fax",Alt_Language,conn) & ": " & FormatPhone(Business_Fax) & vbCrLf
    end if
    if not isblank(Business_Email) then
      MailMessage = MailMessage & Translate("Email",Alt_Language,conn) & ": " & Business_Email & vbCrLf
    end if  
    
    MailMessage = MailMessage & vbCrLf
        
end sub

'--------------------------------------------------------------------------------------

sub Build_MailMessage_Ship_To
    MailMessage = MailMessage & FormatFullName(FirstName, MiddleName, LastName) & vbCrLf
          
    if not isblank(Job_Title) then
      MailMessage = MailMessage & Job_Title & vbCrLf & vbCrLf         
    end if
    if not isblank(Company) then  
      MailMessage = MailMessage & Company & vbCrLf
    end if  
    if not isblank(Shipping_Address) then
      MailMessage = MailMessage & Shipping_Address & vbCrLf
    end if  
    if not isblank(Shipping_Address_2) then
      MailMessage = MailMessage & Shipping_Address_2 & vbCrLf                    
    end if  
    MailMessage = MailMessage & Shipping_City
    if Shipping_State <> "ZZ" then
      MailMessage = MailMessage & ", " & Shipping_State & "  "
    end if
    if not isblank(Shipping_State_Other) then
      MailMessage = MailMessage & ", " & Shipping_State_Other & "  "
    end if  
    MailMessage = MailMessage & Shipping_Postal_Code & vbCrLf
    MailMessage = MailMessage & Shipping_Country & vbCrLf & vbCrLf
    if not isblank(Business_Phone) then
      MailMessage = MailMessage & Translate("Phone",Alt_Language,conn) & ": " & FormatPhone(Business_Phone) & vbCrLf
    end if  
    if not isblank(Business_Fax) then
      MailMessage = MailMessage & Translate("Fax",Alt_Language,conn) & ": " & FormatPhone(Business_Fax) & vbCrLf
    end if
    if not isblank(Business_Email) then
      MailMessage = MailMessage & Translate("Email",Alt_Language,conn) & ": " & Business_Email & vbCrLf
    end if  
    
    MailMessage = MailMessage & vbCrLf
        
end sub

'--------------------------------------------------------------------------------------

sub Send_EMail

  ' --------------------------------------------------------------------------------------
  ' Configure EMail Header Information
  ' --------------------------------------------------------------------------------------
  
  'Mailer.QMessage = False
  'Mailer.Subject  = MailSubject
  'Mailer.BodyText = MailMessage

  msg.Subject = MailSubject
  msg.TextBody = MailMessage

  'if Mailer.SendMail then
'
  'else
  '  ErrMessage = ErrMessage & vbCrLf & "<LI>" & Translate("Send Email Failure.",Login_Language,conn) & "<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Mailer.Response & ". " & Translate("Report this error to the Site Administrator.",Login_Language,conn) & "</LI>"
  'end if   
  msg.Configuration = conf
  On Error Resume Next
  msg.Send
  If Err.Number = 0 then
    'Success
  Else
    ErrMessage = ErrMessage & vbCrLf & "<LI>" & Translate("Send Email Failure.",Login_Language,conn) & "<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Err.Description & ". " & Translate("Report this error to the Site Administrator.",Login_Language,conn) & "</LI>"
  End If

	Set conf = Nothing
	Set msg = Nothing

end sub

'--------------------------------------------------------------------------------------

function GetXMLData(xmlTag, xmlData, Pointer)

  temp = ""
  for v = 0 to UBound(xmlTag)
    if LCase(xmlTag(v)) = LCase(Pointer) then
      temp = xmlData(v)
      exit for
    end if
  next

  GetXMLData = temp
  
end function  
  
'--------------------------------------------------------------------------------------

if Action_Sequence = "review" then
  %>
  <SCRIPT LANGUAGE=JAVASCRIPT>
  <!--
  function CheckRequiredFields() {
  
    var ErrorMsg = "";
    var RadioChecked = 0;
    var dNT = document.<%=Form_Name%>;
    var Error;
    var ctr;

    if(dNT.elements['Ship_To'].length) { 
      for(ctr=0;ctr < dNT.Ship_To.length;ctr++) {
        if(dNT.Ship_To[ctr].checked) {
          RadioChecked++;
          break;
        }
      }
    }
    else {
      if(dNT.elements['Ship_To'].checked) {
        RadioChecked = 1;
      }
    }  
    
    if (RadioChecked == 0) {
      ErrorMsg = "<%=Translate("Please select one or more Ship To: address for this order.",Login_Language,conn)%>\r\n";
    }
    
    if (ErrorMsg.length) {
      //ErrorMsg = "<%=Translate("Please enter the missing information for following REQUIRED fields (or use N/A)",Alt_Language,conn)%>:\r\n\n" + ErrorMsg;
      alert (ErrorMsg);
      return (false);
    }
    else {
      if (dNT.all || dNT.getElementById) {
        for (i = 0; i < dNT.length; i++) {
          var tempobj = dNT.elements[i];
          if (tempobj.type.toLowerCase() == "submit" || tempobj.type.toLowerCase() == "reset")
            tempobj.disabled = true;
          }
        }
        setTimeout('alert("<%=Translate("Your literature order has been submitted.",Alt_Language,conn)%>")', 2000);
      }
      return(true);
    }
//-->
</SCRIPT>
  
<%
end if
%>

<SCRIPT TYPE="text/javascript" LANGUAGE="JavaScript">
function PrintIt(){

  var NS = (navigator.appName == "Netscape");
  var VERSION = parseInt(navigator.appVersion);
  
  if (window.print) {
    window.print() ;  
  }
  else {
    var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
    document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
    WebBrowser1.ExecWB(6, 2);   //Use a 1 vs. a 2 for a prompting dialog box    WebBrowser1.outerHTML = "";  
  }
}

//-->
</SCRIPT>

<%
'--------------------------------------------------------------------------------------
Call Disconnect_SiteWide
'--------------------------------------------------------------------------------------
%>
      