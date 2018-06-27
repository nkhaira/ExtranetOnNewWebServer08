<%
if not isblank(Site_ID) then

  SQLInfo = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQLInfo, conn, 3, 3
  Site_Code         = rsSite("Site_Code")
  Site_Timeout      = CInt(rsSite("Site_Timeout"))
  if not isnull(rsSite("URL")) then
    Site_URL          = Replace(LCase(rsSite("URL")), "support.fluke.com", LCase(Request("SERVER_NAME")))
  end if  
  Site_Description  = rsSite("Site_Description")
  Site_Logon_Method = rsSite("Login_Method")
  Logo              = rsSite("Logo")
  Logo_Left         = rsSite("Logo_Left")
  Footer_Disabled   = rsSite("Footer_Disabled")
  Business          = CInt(rsSite("Business"))
  Privacy_Statement = rsSite("Privacy_Statement_Link")
  Site_Admin_Name   = rsSite("MailToName")
  Site_Admin_Email  = rsSite("MailTo")
  Contrast          = rsSite("Contrast")
  Site_Company      = rsSite("Company")
  Shopping_Cart     = CInt(rsSite("Shopping_Cart"))
  Shopping_Cart_R1  = CInt(rsSite("Shopping_Cart_R1"))
  Shopping_Cart_R2  = CInt(rsSite("Shopping_Cart_R2"))
  Shopping_Cart_R3  = CInt(rsSite("Shopping_Cart_R3"))
  Shopping_Cart_Country = rsSite("Shopping_Cart_Country")
  Order_Inquiry     = CInt(rsSite("Order_Inquiry"))
  Order_Entry       = CInt(rsSite("Order_Entry"))
  Price_Delivery    = CInt(rsSite("Price_Delivery"))
  Path_Site_Secure  = CInt(rsSite("Secure_Stream"))
  
  if CInt(rsSite("Closed")) = True then Site_Closed = True else Site_Closed = False  

  rsSite.close
  set rsSite=nothing

else

  Session("ErrorString") = "<LI>" & Translate("Your session has expired.",Login_Language,conn) & " " & Translate("For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please type in the site's code in the &quot;Name of the Site where you want to go&quot;, then click on [ Login ] or",Login_Language,conn) & "</LI><LI>" & Translate("Use the Site Search feature below.",Login_Language,conn) & "</LI>"
  Call Disconnect_SiteWide
  response.redirect "/register/default.asp"

end if
%>