<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_db.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/SW-Common/SW-Order_Inquiry_Literature_OStatus.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       6/15/2003
'             Sandbox
' --------------------------------------------------------------------------------------
Response.write "The literature order inquiry feature has been discontinued from this site.<br>"
Response.write "Please check your literature order confirmation email for a link to the fulfillment vendor order status website."
Response.End 

response.buffer = true

Dim ErrMessage
ErrMessage = ""

Dim Border_Toggle
Border_Toggle = 0

Dim Show_Order_Detail
Show_Order_Detail = false   ' not implemented on DCG System

Call Connect_SiteWide

Dim sDomain, sNodes, sMyEnvironment
sDomain        = UCase(request.ServerVariables("SERVER_NAME"))
sNodes         = split(sDomain,".")
sMyEnvironment = ""

for each node in sNodes
 	select case node
		case "DEV", "TST", "TEST", "PRD", "DTMEVTVSDV15", "DTMEVTVSDV18"
			sMyEnvironment = node
			exit for
	end select
next
 
Dim Logon_User

if not isblank(request("Logon_User")) then
  Logon_User = request("Logon_User")
  Session("Logon_User") = Logon_User
  Session("No_Nav") = True
else  
  Logon_User = Session("Logon_User")
end if

if not isblank(request("Logon_Cart")) then
  Logon_Cart = request("Logon_Cart")
  Session("Logon_Cart") = Logon_Cart
  Session("No_Nav") = True
elseif not isblank(Session("Logon_Cart")) then
  Logon_Cart = Session("Logon_Cart")
else  
  Logon_Cart = Session("Logon_User")
end if

if not isblank(Logon_User) then
	%>
	<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
	<%
else
  response.redirect "/register/default.asp"
'	site_id = 3
end if
%>
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
<%



' get additional user data
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Order_GetUserInfo"
' create all the parameters we want
cmd.Parameters.Append cmd.CreateParameter("@login",adVarChar,adParamInput,50,Logon_User)
cmd.Parameters.Append cmd.CreateParameter("@site",adInteger,adParamInput,,Site_Id)
set dbRS = cmd.Execute
set cmd = nothing

if dbRS.EOF then
	response.write "Something is seriously wrong "
	disconnect_SiteWide
	response.end
else
	Login_ID = dbRS("ID")
	Login_Type_Code = cInt(dbRS("Type_Code"))
  Login_Region    = CInt(dbRS("Region"))
end if
set dbRS = Nothing

' --------------------------------------------------------------------------------------
' Start of Main
' --------------------------------------------------------------------------------------

Dim Action_Sequence, Action_Title, Sync
Dim Order_Number, Order_Status, Order_Prev, Order_Next
Dim Show_All

Dim BackURL

Dim Account_Info_ID, Shipping_Address_ID, Address_Text
Dim FirstName, MiddleName, LastName, Company, Job_Title
Dim Business_Address, Business_Address_2, Business_City, Business_City_Other, Business_State
Dim Business_Postal_Code, Business_Country, Business_Email, Business_Fax, Business_Phone
Dim Shipping_Address, Shipping_Address_2, Shipping_City, Shipping_City_Other, Shipping_State
Dim Shipping_Postal_Code, Shipping_Country, Order_Ship_Date, Ship_Carrier, Ship_Tracking_No, Lit_Pod

Dim Cart_Width

Cart_Width = 100

Order_Ship_Date  = ""
Ship_Carrier     = ""
Ship_Tracking_No = ""
Lit_Pod          = False

Sync = false
if not isblank(request("Sync")) then                  ' Set once in Order_Inquiry_Redirect.asp
  Sync = true                                         ' Update Shopping Cart with Order Status info from DCG
end if

if not isblank(request("Action")) then
  Action_Sequence = LCase(request("Action"))
else
  Action_Sequence = ""
end if  
Show_All        = request("Show")
Order_Number    = request("Order_Number")

Order_Prev = ""
Order_Next = ""

if CInt(Sync) = CInt(True) or Action_Sequence = "detail" then

  SQL = "SELECT   DISTINCT Order_Number, Submit_Date, Order_Status " &_
        "FROM     Shopping_Cart_Lit " &_
        "WHERE   (Submit_Date IS NOT NULL) AND (Order_Number IS NOT NULL) AND (Account_NTLogin = '" & Logon_Cart & "') " &_
        "ORDER BY Submit_Date DESC, Order_Number DESC"

  Set rsOrders = Server.CreateObject("ADODB.Recordset")
  rsOrders.Open SQL, conn, 3, 3

  if not rsOrders.EOF then
    
    if CInt(Sync) = CInt(True) then
      Do while not rsOrders.EOF
        if CInt(rsOrders("Order_Status")) <> 10 and CInt(rsOrders("Order_Status")) <> 15 and CInt(rsOrders("Order_Status")) <> 16 and CInt(rsOrders("Order_Status")) <> 20 and CInt(rsOrders("Order_Status")) <> 99 then
          Call DCG_Status("Order_Status", rsOrders("Order_Number"), conn)
        end if
        rsOrders.MoveNext      
      loop
      Sync = False    
    
    elseif Action_Sequence = "detail" then
    
      if not rsOrders.EOF then    ' Find Previous and Next Orders
        do while not rsOrders.EOF
          if Order_Number = rsOrders("Order_Number") then
            rsOrders.MovePrevious
            if not rsOrders.BOF then
              if  not isblank(rsOrders("Order_Number")) then
                Order_Prev = rsOrders("Order_Number")
              else
                Order_Prev = ""
              end if  
            end if
            rsOrders.MoveNext
            rsOrders.MoveNext
            if not rsOrders.EOF then
              if not isblank(rsOrders("Order_Number")) then
                Order_Next = rsOrders("Order_Number")
              else
                Order_Next = ""
              end if  
            end if
            exit do
          end if
          rsOrders.MoveNext
        loop
      end if
      
      select case LCase(request("Order_Sequence"))
        case "prev"
          if not isblank(Order_Prev) then Order_Number = Order_Prev
        case "next"
          if not isblank(Order_Next) then Order_Number = Order_Next
      end select  
    end if
  else
  
    Action_Sequence = ""
    Order_Number = ""
    Show_All = 0   
  
  end if

  rsOrders.close
  set rsOrders = nothing
  set SQL      = nothing

end if    
       
select case Action_Sequence
  case "detail"
    Action_Title = "Literature Order Detail Information"
  case else  
    Action_Title = "Literature Order Summary Information"
end select
   
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate(Action_Title,Login_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) &_
                  "<BR><SPAN CLASS=MediumBoldGold>" & _
                  Translate(Action_Title,Login_Language,conn) & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

if not isblank(request("Admin_URL")) then
  BackURL = request("Admin_URL")
  Session("BackURL") = BackURL

else
  BackURL = Replace(Session("BackURL"),"=9007","=9000")
end if  

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
<%

' Log the Activity

if isblank(Session("Logon_Cart")) then

  ActivitySQL = "INSERT INTO Activity (Account_ID, Site_ID, Session_ID, View_Time, CID, SCID, Language, Method, Calendar_ID, Region, Country)" & vbcrlf &_
            		"Values(" &_
            		Login_ID & "," &_
            		Site_ID & "," &_
            		Session("Session_ID") & ",'" &_
            		Date() & "'," &_
            		"9007," &_
            		"1," &_
            		"'" & Login_Language & "'," & _
                "0,"

                select case Action_Sequence
                  case "detail"
                    ActivitySQL = ActivitySQL & "104,"     ' Literature Order Detail
                  case else  
                    ActivitySQL = ActivitySQL & "103,"     ' Literature Order List
                end select          
            
                ActivitySQL = ActivitySQL & Login_Region & "," &_
                              "'" & Session("Login_Country") & "')"
                 
  	conn.Execute (ActivitySQL)
    
end if
'--------------------------------------------------------------------------------------

select case Action_Sequence

  case "detail"
  
    response.write "<SPAN CLASS=Heading3>" & Translate(Action_Title,Login_Language,conn) & "</SPAN><BR>"
    response.write "<BR>"

    Call Menu_Bar
    
    if Action_Sequence = "detail" then
    
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
                "Literature_Items_US.POD AS Lit_POD, " &_                
                "Literature_Items_US.[LANGUAGE] AS Lit_Language " &_
        "FROM   Literature_Items_US " &_
                "LEFT OUTER JOIN Shopping_Cart_Lit ON Literature_Items_US.ITEM = Shopping_Cart_Lit.Item_Number " &_
                "LEFT OUTER JOIN Language " &_
                "INNER JOIN Calendar ON Language.Code = Calendar.[Language] " &_
                "INNER JOIN Calendar_Category ON Calendar.Category_ID = Calendar_Category.ID ON Shopping_Cart_Lit.Asset_ID = Calendar.ID " &_
                "LEFT OUTER JOIN Site ON Shopping_Cart_Lit.Site_ID = Site.ID " &_
        "WHERE  (Shopping_Cart_Lit.Account_NTLogin = '" & Logon_Cart & "') AND (Shopping_Cart_Lit.Order_Number='" & Order_Number & "') " &_
        "ORDER BY Shopping_Cart_Lit.ID DESC"   

      'response.write SQL & "<P>"
      
      Set rsCart = Server.CreateObject("ADODB.Recordset")
      rsCart.Open SQL, conn, 3, 3
      
      if not rsCart.EOF then
        Call Account_Information
        Call Shopping_Cart_Header
        Call Shopping_Cart_Data
        Call Shopping_Cart_Footer
        Call Shopping_Cart_Notes
        Call Menu_Bar
      else
        response.write Translate("Invalid Literature Order Number",Login_Language,conn)  
      end if
      rsCart.Close
      set rsCart = nothing
    end if      
  
  case else

    response.write "<SPAN CLASS=Heading3>" & Translate(Action_Title,Login_Language,conn) & "</SPAN><BR>"
    response.write "<BR>"
    
    Call Menu_Bar
    
    SQLOrders = "SELECT    Shopping_Cart_Ship_To.FirstName, Shopping_Cart_Ship_To.MiddleName, Shopping_Cart_Ship_To.LastName, Shopping_Cart_Ship_To.Company, " &_
                "          Shopping_Cart_Ship_To.Shipping_State, Shopping_Cart_Ship_To.Shipping_State_Other, Shopping_Cart_Ship_To.Shipping_Postal_Code, " &_
                "          Shopping_Cart_Ship_To.Shipping_Country, Shopping_Cart_Ship_To.Shipping_City, Shopping_Cart_Lit.Submit_Date, Shopping_Cart_Lit.Order_Number, " &_
                "          Shopping_Cart_Ship_To.ID, Shopping_Cart_Lit.Shipping_Address_ID, Shopping_Cart_Lit.Order_Status, Shopping_Cart_Lit.Order_Ship_Date, Shopping_Cart_Lit.Ship_Carrier, Shopping_Cart_Lit.Ship_Tracking_No " &_
                "FROM      Shopping_Cart_Ship_To RIGHT OUTER JOIN " &_
                "          Shopping_Cart_Lit ON Shopping_Cart_Ship_To.ID = Shopping_Cart_Lit.Shipping_Address_ID " &_
                "WHERE    (Shopping_Cart_Lit.Order_Number IN " &_
                "            (SELECT DISTINCT Shopping_Cart_Lit.Order_Number " &_
                "               FROM   Shopping_Cart_Lit " &_
                "               WHERE (Shopping_Cart_Lit.Submit_Date IS NOT NULL) AND (Shopping_Cart_Lit.Account_NTLogin = '" & Logon_Cart & "'))) " &_
                "ORDER BY Shopping_Cart_Lit.Submit_Date DESC, Shopping_Cart_Lit.Order_Number DESC"

    'response.write SQLOrders
    'response.end
    
    Set rsOrders = Server.CreateObject("ADODB.Recordset")
    rsOrders.Open SQLOrders, conn, 3, 3
               
    if not rsOrders.EOF then
    
      Call Order_Summary_Header
    
      Order_Number = ""
      do while not rsOrders.EOF
        if Order_Number <> rsOrders("Order_Number") then
          Call Order_Summary_Data
        end if
        Order_Number = rsOrders("Order_Number")  
        rsOrders.MoveNext
      loop
      
      Call Order_Summary_Footer
      Call Menu_Bar   
     
    else
    
      response.write Translate("There are no Literature Orders to display.",Login_Language,conn)
    
    end if
    
    rsOrders.close
    set rsOrders = nothing
   
end select  
 
'--------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide

'--------------------------------------------------------------------------------------
' Sub Routines and Functions
'--------------------------------------------------------------------------------------

sub Menu_Bar

  response.write "<FORM NAME=""Menu_Bar"">" & vbCrLf

  Call Nav_Border_Begin
  response.write "<TABLE BORDER=""" & Border_Toggle & """ WIDTH=""" & Cart_Width & "%"" CELLPADDING=0 CELLSPACING=0>" & vbCrLf
  response.write "<TR>" & vbCrLf
  response.write "<TD CLASS=SmallBold ALIGN=RIGHT VALIGN=TOP>" & vbCrLf

  response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Home"" VALUE=""" & " " & Translate("Home",Login_Language,conn) & """ "
  response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & BackURL & "'; return false;"" TITLE=""Return to Order Inquiry Select"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"

  if Action_Sequence = "detail" then
    response.write "&nbsp;&nbsp;"
    response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Order_List"" VALUE=""" & " " & Translate("Order List",Login_Language,conn) & """ "
    response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "'; return false;"" TITLE=""Return to Order List"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
    response.write "&nbsp;&nbsp;"
    response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" NAME=""Show"" VALUE="""
    if CInt(Show_All) = CInt(True) then
      response.write" " & Translate("Show Less",Login_Language,conn) & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Detail&Show=0&Order_Number=" & Order_Number & "'; return false;"" TITLE=""Hide Line Item Extended Information"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
    else  
      response.write" " & Translate("Show More",Login_Language,conn) & """ "
      response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Detail&Show=-1&Order_Number=" & Order_Number & "'; return false;""TITLE=""Show Line Item Extended Information"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
    end if
    
    if not isblank(Order_Prev) or not isblank(Order_Next) then
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
      if not isblank(Order_Prev) then  
        response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""IMAGE"" NAME=""Order_Sequence"" SRC=""/Images/Button_Up.gif"" ALT=""View Up in Order List"" VALUE="" > "" "
        response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Detail&Show=0&Order_Number=" & Order_Prev & "&Sequence=Prev'; return false;"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
        response.write "&nbsp;&nbsp;"
      end if
      if not isblank(Order_Next) then
        response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""IMAGE"" NAME=""Order_Sequence"" SRC=""/Images/Button_Down.gif"" ALT=""View Down in Order List"" VALUE="" < "" "
        response.write "LANGUAGE=""Javascript"" ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Detail&Show=0&Order_Number=" & Order_Next & "&Sequence=Next'; return false;"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
        response.write "&nbsp;&nbsp;"
      end if
    end if  
  end if
  
  response.write "&nbsp;&nbsp;"
  response.write "<INPUT TYPE=""BUTTON"" CLASS=NavLeftHighlight1 VALUE=""" & Translate("Print",Login_Language,conn) & """ NAME=""Print"" LANGUAGE=""JavaScript"" onClick=""PrintIt()"" TITLE=""Print Page"" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"

  response.write "</TD>"    & vbCrLf
  response.write "</TR>"    & vbCrLf
  response.write "</TABLE>" & vbCrLf

  Call Nav_Border_End

  response.write "</FORM>" & vbCrLf
  
  response.write "<P>" & vbCrLf

end sub

'--------------------------------------------------------------------------------------

sub Account_Information

  Call Table_Begin
  
  response.write "<TABLE bgcolor=WHITE WIDTH=""" & Cart_Width & "%"" BORDER=""" & Border_Toggle & """ CELLPADDING=4>" & vbCrLf

  response.write "<TR>" & vbCrLf
  
  ' Order Detail
  response.write "<TD WIDTH=""8%"" NOWRAP CLASS=SmallBold BGCOLOR=""#CECECE"" VALIGN=TOP ALIGN=RIGHT>" & vbCrLf
  response.write Translate("Order Number",Login_Language,conn) & ":<BR>" & vbCrLf
  response.write Translate("Order Date",Login_Language,conn) & ":<BR>" & vbCrLf
  response.write Translate("Order Status",Login_Language,conn) & ":<P>" & vbCrLf
  if not isblank(rsCart("Order_Ship_Date")) then  
    response.write Translate("Shipment Date",Login_Language,conn) & ":<BR>" & vbCrLf
  end if  
  if not isblank(rsCart("Ship_Carrier")) and rsCart("Ship_Carrier") > 0 then        
    response.write Translate("Carrier",Login_Language,conn) & ":<BR>" & vbCrLf
  end if
  if not isblank(rsCart("Ship_Tracking_No")) and not isblank(rsCart("Ship_Carrier")) and rsCart("Ship_Carrier") > 0 then    
    response.write Translate("Tracking Number",Login_Language,conn) & ":" & vbCrLf    
  end if  
  response.write "</TD>" & vbCrLf
 
  response.write "<TD WIDTH=""8%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP ALIGN=RIGHT>" & vbCrLf

  ' Order Number
  
  response.write "<SPAN CLASS=SmallRed>" & Replace(rsCart("Order_Number"),"FLUKECO","") & "</SPAN><BR>"
  
  ' Order Submit Date
  
  response.write "<SPAN CLASS=SmallRed>" & FormatDate(1,rsCart("Submit_Date")) & "</SPAN><BR>"
  
  ' Order / Shippment Status
  
  ' 01 To Be Approved
  ' 02 Order Received
  ' 03 In Process
  ' 04 Order Accepted
  ' 05 In Process
  ' 10 Shipped
  ' 11 In Shipping
  ' 15 Shipped
  ' 16 POD Order
  ' 20 Invoiced
  ' 80 Back-Ordered
  ' 99 Cancelled
  
  select case rsCart("Order_Status")
    case 10, 15, 20
      response.write "<IMG SRC=""/Images/Check_Green.gif"" BORDER=0>&nbsp;&nbsp;<SPAN CLASS=SmallRed>" & Translate("Shipped",Login_Language,conn) & "</SPAN><P>" & vbCrLf

      ' Shipping Date
      if not isblank(rsCart("Order_Ship_Date")) then
        response.write FormatDate(1,rsCart("Order_Ship_Date")) & "<BR>" & vbCrLf
      end if
      
      ' Shipping Carrier
      if not isblank(rsCart("Ship_Carrier")) and rsCart("Ship_Carrier") > 0 then        
        SQLCarrier = "SELECT * FROM Shopping_Cart_Ship_Tracking WHERE ID=" & rsCart("Ship_Carrier")
        Set rsCarrier = Server.CreateObject("ADODB.Recordset")
        rsCarrier.Open SQLCarrier, conn, 3, 3
        if not rsCarrier.EOF then
          response.write rsCarrier("Name")
          Carrier_URL = rsCarrier("URL")
        else
          response.write Translate("TBD",Login_language,conn)
          Carrier_URL = ""
        end if
        rsCarrier.close
        set rsCarrier = nothing
        response.write "<BR>" & vbCrLf
      end if

      ' Tracking Number
      if not isblank(rsCart("Ship_Tracking_No")) and not isblank(Carrier_URL) then
        response.write "<A HREF=""Javascript:openit_maxi('"
        response.write Replace(Carrier_URL,"[Number]",rsCart("Ship_Tracking_No"))
        response.write "', 'horizontal')"" ALT=""View Shipment Tracking Information"">"
        response.write "<SPAN CLASS=SmallBoldRed>" & rsCart("Ship_Tracking_No") & "</SPAN>"
        response.write "</A><BR>" & vbCrLf
      end if        
        
    case 16
      response.write "<IMG SRC=""/Images/Check_Green.gif"" BORDER=0>&nbsp;&nbsp;<SPAN CLASS=SmallRed>" & Translate("See Note 1",Login_Language,conn) & "</SPAN><P>" & vbCrLf
    case 80
      response.write Translate("Back-Ordered",Login_Language,conn) & "<P>" & vbCrLf
    case 99
      response.write Translate("Cancelled",Login_Language,conn) & "<P>" & vbCrLf
    case else
      response.write Translate("In Process",Login_Language,conn) & "<P>" & vbCrLf
   end select
  
  response.write "</TD>"

  ' Ordered by
  response.write "<TD WIDTH=""10%"" NOWRAP CLASS=SmallBold BGCOLOR=""#CECECE"" VALIGN=TOP ALIGN=RIGHT>" & vbCrLf
  response.write Translate("Ordered by",Login_Language,conn) & ":" & vbCrLf
  response.write "</TD>" & vbCrLf
 
  response.write "<TD WIDTH=""20%"" NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP>" & vbCrLf
  
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
  response.write Translate("Email",Login_Language,conn) & ": " & Business_Email & "<BR>" & vbCrLf
  if Account_Info_ID = 0 then
    response.write Translate("Phone",Login_Language,conn) & ": " & Business_Phone & vbCrLf
  end if  
  response.write "</TD>" & vbCrLf

  ' Ship To
  response.write "<TD WIDTH=""10%"" NOWRAP CLASS=SmallBold BGCOLOR=""#CECECE"" VALIGN=TOP ALIGN=RIGHT>" & vbCrLf
  response.write Translate("Ship to",Login_Language,conn) & ":" & vbCrLf
  response.write "</TD>" & vbCrLf
 
  response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#F6F6F6"" VALIGN=TOP>" & vbCrLf
  
  Account_Info_ID = rsCart("Shipping_Address_ID")
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
  
  response.write Translate("Email",Login_Language,conn) & ": " & Business_Email & "<BR>" & vbCrLf
  if Account_Info_ID = 0 then
    response.write Translate("Phone",Login_Language,conn) & ": " & Business_Phone & vbCrLf
  end if
    
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf

  response.write "</TABLE>" & vbCrLf

  Call Table_End
  
end sub

'--------------------------------------------------------------------------------------

sub Get_Account_Information

  select case Account_Info_ID
    case 0
      SQL =  "SELECT * FROM UserData WHERE NTLogin='" & Logon_Cart & "' AND NewFlag=0"
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
    end if
  end if
  
  rsLogin.close
  set rsLogin = Nothing

end sub

'--------------------------------------------------------------------------------------

sub Order_Summary_Header

  response.write "<A NAME=Detail></A>" & vbCrLf
   
  Call Table_Begin

  response.write "<TABLE BGCOLOR=SILVER BORDER=""" & Border_Toggle & """ WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0>" & vbCrLf
  response.write "<TR>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold NOWRAP WIDTH=""1%"">" & Translate("Order Number",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold NOWRAP WIDTH=""1%"" ALIGN=RIGHT>&nbsp;&nbsp;" & Translate("Order Date",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold>&nbsp;&nbsp;" & Translate("Ship to",Login_Language,conn) & ": " & Translate("Name",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold>&nbsp;&nbsp;" & Translate("Company",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold>&nbsp;&nbsp;" & Translate("City",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold>&nbsp;&nbsp;" & Translate("Postal Code",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold ALIGN=CENTER>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold ALIGN=CENTER>" & Translate("Ship",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold ALIGN=CENTER>" & Translate("Detail",Login_Language,conn) & "</TD>" & vbCrLf
  if CInt(Show_Order_Detail) = CInt(True) then
    response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold ALIGN=CENTER>" & Translate("Detail",Login_Language,conn) & "</TD>" & vbCrLf            
  end if  
  response.write "</TR>" & vbCrLf
  
end sub  

'--------------------------------------------------------------------------------------

sub Order_Summary_Data

    if rsOrders("Shipping_Address_ID") = 0 then
      Call Get_Account_Information
      Mask_Color = "#E6E6E6"
    else
      FirstName  = rsOrders("FirstName")
      MiddleName = rsOrders("MiddleName")
      LastName  = rsOrders("LastName")
      Company   = rsOrders("Company")
      Shipping_City = rsOrders("Shipping_City")        
      Shipping_Postal_Code = rsOrders("Shipping_Postal_Code")
      Shipping_Country = rsOrders("Shipping_Country")
      Mask_Color = "#F6F6F6"
    end if  

    response.write "<TR>" & vbCrLf
    
    ' Order Number
    response.write "<TD CLASS=Small BGCOLOR=""White"" VALIGN=MIDDLE ALIGN=CENTER>"
    response.write Replace(rsOrders("Order_Number"),"FLUKECO","")
    response.write "</TD>" & vbCrLf

    ' Order Date
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""White"" VALIGN=MIDDLE ALIGN=RIGHT>"
    response.write FormatDate(1,rsOrders("Submit_Date")) & "&nbsp;&nbsp;&nbsp;"
    response.write "</TD>" & vbCrLf
 
    ' Ship To Name
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""" & Mask_Color & """ VALIGN=MIDDLE>"
    response.write "&nbsp;&nbsp;"
    response.write FormatFullName(FirstName, MiddleName, LastName)
    response.write "</TD>" & vbCrLf
    
    ' Company
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""" & Mask_Color & """ VALIGN=MIDDLE>"
    response.write "&nbsp;&nbsp;" & Company
    response.write "</TD>" & vbCrLf

    ' City
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""" & Mask_Color & """ VALIGN=MIDDLE>"
    response.write "&nbsp;&nbsp;" & Shipping_City
    response.write "</TD>" & vbCrLf

    ' Postal_Code
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""" & Mask_Color & """ VALIGN=MIDDLE>"
    response.write "&nbsp;&nbsp;" & Shipping_Postal_Code
    response.write "</TD>" & vbCrLf
    
    ' Country
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""" & Mask_Color & """ VALIGN=MIDDLE ALIGN=CENTER>"
    response.write Shipping_Country
    response.write "</TD>" & vbCrLf

    ' Shipped
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""White"" VALIGN=MIDDLE ALIGN=CENTER>"
    
    ' 01 To Be Approved
    ' 02 Order Received
    ' 03 In Process
    ' 04 Order Accepted
    ' 05 In Process
    ' 10 Shipped
    ' 11 In Shipping
    ' 15 Shipped
    ' 16 POD Order
    ' 20 Invoiced
    ' 80 Back-Ordered
    ' 99 Cancelled
    
    select case CInt(rsOrders("Order_Status"))
      case 10, 15, 16, 20
        response.write "<IMG SRC=""/images/Check_Green.gif"" BORDER=0>"
      case 80
        response.write "<IMG SRC=""/images/Check_Yellow.gif"" BORDER=0>"
      case 99
        response.write "<IMG SRC=""/images/Check_Red.gif"" BORDER=0>"
      case else
        response.write Translate("In-Process",Login_Language,conn)
    end select      
    response.write "</TD>" & vbCrLf

    ' Summary
    response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#666666"" VALIGN=MIDDLE ALIGN=CENTER>"
    response.write "&nbsp;&nbsp;"
    response.write "<A HREF=""JavaScript:void(0);"" CLASS=NavLeftHighlight1 ONCLICK=""location.href='" & request.ServerVariables("SCRIPT_Name") & "?Action=Detail&Order_Number=" & rsOrders("Order_Number") & "#Detail'; return false;"")>"
    response.write "&nbsp;" & Translate("View",Login_Language,conn) & "&nbsp;"
    response.write "</A>"
    response.write "</TD>" & vbCrLf

    ' Tracking Information
    if CInt(Show_Order_Detail) = CInt(True) then    
      response.write "<TD NOWRAP CLASS=Small CLASS=NavLeftHighlight1 BGCOLOR=""#666666"" VALIGN=MIDDLE ALIGN=CENTER>"
      response.write "<A HREF=""JavaScript:void(0);"" CLASS=NavLeftHighlight1>"
      response.write "&nbsp;" & Translate("View",Login_Language,conn) & "&nbsp;"
      response.write "</A>"
      response.write "</TD>" & vbCrLf
    end if  

    response.write "<TR>" & vbCrLf

end sub

'--------------------------------------------------------------------------------------

sub Order_Summary_Footer
 
  response.write "</Table>" & vbCrLf

  Call Table_End
  
  response.write "<BR>" & vbCrLf
  
end sub

'--------------------------------------------------------------------------------------

sub Shopping_Cart_Header

'  Call Menu_Bar
  response.write "<BR>" & vbCrLf

  if isblank(ErrMessage) and isBlank(PostErrMessage) then
    response.write "<A NAME=""Cart_List""></A>" & vbCrLf
  end if

  Call Table_Begin

  response.write "<TABLE BGCOLOR=#666666 BORDER=""" & Border_Toggle & """ WIDTH=""100%"" CELLPADDING=4 CELLSPACING=1>" & vbCrLf
  response.write "<TR>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold WIDTH=""2%"" ALIGN=CENTER>" & Translate("Line",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold WIDTH=""4%"" ALIGN=CENTER>" & Translate("Item Number",Login_Language,conn) & "</TD>" & vbCrLf  
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold WIDTH=""4%"" ALIGN=CENTER>" & Translate("Quantity Ordered",Login_Language,conn) & "</TD>" & vbCrLf
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold WIDTH=""2%"" ALIGN=CENTER>" & Translate("Note",Login_Language,conn) & "</TD>" & vbCrLf  
  if CInt(Show_All) = CInt(True) then
    response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold WIDTH=""6%"" ALIGN=CENTER>" & Translate("Thumbnail",Login_Language,conn) & "</TD>" & vbCrLf
  end if  
  response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold>" & Translate("Description",Login_Language,conn) & "</TD>" & vbCrLf
  if CInt(Show_Order_Detail) = CInt(True) then
    response.write "<TD BGCOLOR=""#666666"" CLASS=SmallBoldGold WIDTH=""5%"" ALIGN=CENTER>" & Translate("Detail",Login_Language,conn) & "</TD>" & vbCrLf
  end if    
  response.write "</TR>" & vbCrLf
  
end sub  

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Data

  rsCart.MoveFirst
  Line_Number = 0
  Last_ID     = 0

  do while not rsCart.EOF
  
    if Last_ID <> rsCart("ID") then
      Line_Number = Line_Number + 1
  
      response.write "<TR>" & vbCrLf
  
      ' Line Number
      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""Silver"">"
      response.write "<A NAME=""Q" & rsCart("ID") & """></A>"                             ' Index Anchor
      response.write Line_Number & "</TD>"
  
      ' Item Number
      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""White"">"
      response.write rsCart("Item_Number")
      response.write "</TD>" & vbCrLf
  
  
      ' Quantity
      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""White"">"
      response.write rsCart("Quantity")
      response.write "</TD>" & vbCrLf
  
      ' POD
      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""White"">"
'      if CInt(rsCart("Lit_POD")) = CInt(True) then
'        Lit_POD = True
'        response.write "1"
'      else
'        response.write "&nbsp;"
'      end if  
      if CInt(rsCart("Order_Status")) = 16 then
        Lit_POD = True
        response.write "1"
      else
        response.write "&nbsp;"
      end if  

      response.write "</TD>" & vbCrLf
      
      
      ' Thumbnail
      
      if CInt(Show_All) = CInt(True) then
        response.write "<TD BGCOLOR=""White"" CLASS=Small>"
        if not isblank(rsCart("Thumbnail")) then
          response.write "<A HREF=""javascript:void(0);"" onclick=""openit('http://" & Request("SERVER_NAME") & "/Find_It.asp?Document=" & rsCart("Item_Number") & "','Vertical');return false;"">"
          response.write "<IMG SRC=""" & "/" & rsCart("Site_Code") & "/" & rsCart("Thumbnail") & """ WIDTH=""60"" BORDER=1 ALT=""Click to View"">"
          response.write "</A>"
        else
          response.write "&nbsp;"
        end if
        response.write "</TD>" & vbCrLf
      end if  
     
      ' Title
      response.write "<TD BGCOLOR=""White"" CLASS=Small VALIGN=TOP>"
  
      response.write Translate("Title",Login_Language,conn) & ": "  
  
      ' Description
      if not isblank(rsCart("Title")) then      
        response.write "<B>" & rsCart("Title") & "</B>"
      else
        response.write "<B>" & ProperCase(rsCart("Lit_Description")) & "</B>"
      end if  
      response.write "<BR>"
  
      if CInt(Show_All) = CInt(True) then
        if not isblank(rsCart("Product")) then
          response.write Translate("Product",Login_Language,conn) & ": " & rsCart("Product") & "<BR>"
          response.write Translate("Langugage",Login_Language,conn) & ": " & rsCart("Language")
        else
          response.write Translate("Langugage",Login_Language,conn) & ": " & ProperCase(rsCart("Lit_Language"))
        end if  
        response.write "</TD>" & vbCrLf
      end if  
      
      ' Tracking Information
      if CInt(Show_Order_Detail) = CInt(True) then
        response.write "<TD NOWRAP CLASS=Small BGCOLOR=""#666666"" VALIGN=MIDDLE ALIGN=CENTER>"
        response.write "<A HREF=""JavaScript:void(0);"" CLASS=NavLeftHighlight1>"
        response.write "&nbsp;" & Translate("View",Login_Language,conn) & "&nbsp;"
        response.write "</A>"
        response.write "</TD>" & vbCrLf 
      end if  
        
      response.write "</TR>" & vbCrLf
      
    end if
    
    Last_ID = rsCart("ID")
    rsCart.MoveNext
  loop

end sub

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Footer
 
  response.write "</Table>" & vbCrLf

  Call Table_End
  
  response.write "<P>"
   
end sub

' --------------------------------------------------------------------------------------

sub Shopping_Cart_Notes

  if CInt(Lit_Pod) = CInt(True) then
    response.write "<P>"
    response.write "<SPAN CLASS=SMALLBOLD>" & Translate("Note",Login_Language,conn) & " 1: </SPAN><SPAN CLASS=Small>" & Translate("&quot;Print on Demand&quot; literature items are shipped separately from this order.", Login_Language,conn) & " "
    select case Login_Region
      case 1        ' US
        response.write Translate("Typical lead-time for printing in addition to transit time from origin is 3-4 business days.",Login_Language,conn)
      case 2        ' Europe 
        response.write Translate("Typical lead-time for printing in addition to transit time from origin is 3-4 business days.",Login_Language,conn) & " "
        response.write Translate("Literature orders are consolidated for International shipments outside of the U.S.",Login_Language,conn) & " "
        response.write Translate("Consolidated country shipments are made once a week and may add up to 7 additional business days.",Login_Language,conn)
      case else     ' Intercon
        response.write Translate("Typical lead-time for printing in addition to transit time from origin is 3-4 business days.",Login_Language,conn) & " "
        response.write Translate("Literature orders are consolidated for International shipments outside of the U.S.",Login_Language,conn) & " "
        response.write Translate("Consolidated country shipments are made once a week and may add up to 7 additional business days.",Login_Language,conn)
    end select
    response.write "</SPAN><P>"
  end if      
end sub

'--------------------------------------------------------------------------------------

if CInt(Sync) = CInt(True) then
  %><!--#Include virtual="/SW-Common/SW-Order_Inquiry_Literature_OStatus.asp"--><%
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
</SCRIPT>

<!--#include virtual="/include/Pop-Up.asp"-->


