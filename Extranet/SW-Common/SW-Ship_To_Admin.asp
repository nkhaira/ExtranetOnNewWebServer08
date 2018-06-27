<%
' --------------------------------------------------------------------------------------
' Author: K. D. Whitlock
' Date:   02/01/2000
' Update Account Profile - Limited Data Version
' 06/19/2002 - Added Fields and Re-Ordered Form to work with Euro DCM
' --------------------------------------------------------------------------------------

Dim SQL
Dim BackURL
Dim Site_ID
Dim Send_2_CMS_Debug
Dim Action
Dim ErrorMessage

Dim Account_ID
Dim New_Account_ID
Dim User_Region

Send_2_CMS_Debug = False

BackURL = request.form("BackURL")
Site_ID = request.form("Site_ID")
Action  = "Save / Update"
User_Region = CInt(request.form("Region"))

if isblank(Site_ID) then
  site_ID = 0
end if

if CInt(request("Cart_Mode")) = CInt(True) then
  Cart_Mode = True
else
  Cart_Mode = False
end if

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_DB.asp"-->  
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/DCMHTTP_DataTransfer.asp"-->
<%

Call Connect_SiteWide

if isblank(Session("Logon_User")) or CInt(request.form("Permit_Update")) <> CInt(True) then

  ' Security Breech Redirect

  if not isblank(request.form("Login_Language")) then
    Login_Language = LCase(request.form("Login_Language"))
  else
    Login_Language = "eng"
  end if
      
  Session("ErrorString") = "<LI>" & Translate("Your session has expired.",Login_Language,conn) & " " & Translate("For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please logon below.",Login_Language,conn) & "</LI>"
  Call Disconnect_SiteWide
  response.redirect "/register/Login.asp?Site_ID=" & request.form("Site_ID") & "&Language=" & request.form("Language")
  
else  

  ' --------------------------------------------------------------------------------------
  ' Connect to SiteWide DB
  ' --------------------------------------------------------------------------------------

  SQL_Site = "SELECT * FROM Site Where Site.ID=" & Site_ID
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL_Site, conn, 3, 3

  Dim CMS_Region(3)

  CMS_Region(1) = rsSite("CMS_Region_1")
  CMS_Region(2) = rsSite("CMS_Region_2")
  CMS_Region(3) = rsSite("CMS_Region_3")

  rsSite.close
  set rsSite = nothing

  ' Build SQL Update String
   
  SQL = ""
    
    ' Gender / Prefix
    if not isblank(request.form("Gender")) then
      select case request.form("Gender")
        case 0    ' Male
          SQL = SQL & "Prefix='Mr'"
          SQL = SQL & ",Gender=0"       
        case 1    ' Female
          SQL = SQL & "Prefix='Ms'"
          SQL = SQL & ",Gender=1"       
        case else ' Unknown
          SQL = SQL & "Prefix=NULL"
          SQL = SQL & ",Gender=NULL"       
      end select       
    else
      SQL = SQL & "Prefix=NULL"
      SQL = SQL & ",Gender=NULL"       
    end if
  
    ' First Name
    if not isblank(request.form("FirstName")) then 
      SQL = SQL & ",FirstName="& "'" & replacequote(request.form("FirstName")) & "'"
    else
      SQL = SQL & ",FirstName=NULL"    
    end if
  
    ' Middle Name
    if not isblank(request.form("MiddleName")) then 
      SQL = SQL & ",MiddleName="& "'" & replacequote(request.form("MiddleName")) & "'"         
    else
      SQL = SQL & ",MiddleName=NULL"    
    end if
  
    ' Last Name
    if not isblank(request.form("LastName")) then 
      SQL = SQL & ",LastName="& "'" & replacequote(request.form("LastName")) & "'"         
    else
      SQL = SQL & ",LastName=NULL"   
    end if
  
    ' Suffix
    if not isblank(request.form("Suffix")) then 
      SQL = SQL & ",Suffix="& "'" & replacequote(request.form("Suffix")) & "'"
    else
      SQL = SQL & ",Suffix=NULL"             
    end if
          
    ' Initials
    if not isblank(request.form("Initials")) then 
      SQL = SQL & ",Initials="& "'" & replacequote(request.form("Initials")) & "'"
    else
      SQL = SQL & ",Initials=NULL"             
    end if

    ' Company
    if not isblank(request.form("Company")) then 
      SQL = SQL & ",Company="& "'" & replacequote(request.form("Company")) & "'"
    else
      SQL = SQL & ",Company=NULL"             
    end if

    ' Company_Website
    if not isblank(request.form("Company_Website")) then 
      SQL = SQL & ",Company_Website="& "'" & killquote(request.form("Company_Website")) & "'"
    else
      SQL = SQL & ",Company_Website=NULL"             
    end if
 
    ' Job Title
    if not isblank(request.form("Job_Title")) then 
      SQL = SQL & ",Job_Title="& "'" & replacequote(request.form("Job_Title")) & "'"         
    else
      SQL = SQL & ",Job_Title=NULL"    
    end if
    
    ' Business Mail Stop
    if not isblank(request.form("Business_MailStop")) then 
      SQL = SQL & ",Business_MailStop="& "'" & replacequote(request.form("Business_MailStop")) & "'"
    else
      SQL = SQL & ",Business_MailStop=NULL"
      
    end if
    
    ' Business Address
    if not isblank(request.form("Business_Address")) then 
      SQL = SQL & ",Business_Address="& "'" & replacequote(request.form("Business_Address")) & "'"         
    else
      SQL = SQL & ",Business_Address=NULL"
    end if
    
    ' Business Address 2
    if not isblank(request.form("Business_Address_2")) then 
      SQL = SQL & ",Business_Address_2="& "'" & replacequote(request.form("Business_Address_2")) & "'"
    else
      SQL = SQL & ",Business_Address_2=NULL"    
    end if
  
    ' Business City
    if not isblank(request.form("Business_City")) then 
      SQL = SQL & ",Business_City="& "'" & replacequote(request.form("Business_City")) & "'"         
    else
      SQL = SQL & ",Business_City=NULL"    
    end if
  
    ' Business State
    if not isblank(request.form("Business_State")) or not isblank(request.form("Business_State_Other")) then 
      if not isblank(request.form("Business_State")) then
        SQL = SQL & ",Business_State="& "'" & replacequote(request.form("Business_State")) & "'"
      else
        SQL = SQL & ",Business_State=NULL"               
      end if              
      if not isblank(request.form("Business_State_Other")) then
        SQL = SQL & ",Business_State_Other="& "'" & replacequote(request.form("Business_State_Other")) & "'"         
      else
        SQL = SQL & ",Business_State_Other=NULL"      
      end if              
    end if
  
    ' Business Postal Code
    if not isblank(request.form("Business_Postal_Code")) then 
      SQL = SQL & ",Business_Postal_Code="& "'" & replacequote(request.form("Business_Postal_Code")) & "'"         
    else
      SQL = SQL & ",Business_Postal_Code=NULL"
    end if
  
    ' Business_Country
    if not isblank(request.form("Business_Country")) then 
      SQL = SQL & ",Business_Country="& "'" & replacequote(request.form("Business_Country")) & "'"         
    else
      SQL = SQL & ",Business_Country=NULL"
    end if
  
    ' Postal Address
    if not isblank(request.form("Postal_Address")) then 
      SQL = SQL & ",Postal_Address="& "'" & replacequote(request.form("Postal_Address")) & "'"
    else
      SQL = SQL & ",Postal_Address=NULL"
    end if
    
    ' Postal City
    if not isblank(request.form("Postal_City")) then 
      SQL = SQL & ",Postal_City="& "'" & replacequote(request.form("Postal_City")) & "'"
    else
      SQL = SQL & ",Postal_City=NULL"    
    end if
  
    ' Postal State  
    if not isblank(request.form("Postal_State")) > 0 or not isblank(request.form("Postal_State_Other")) then 
      if not isblank(request.form("Postal_State")) then
        SQL = SQL & ",Postal_State="& "'" & replacequote(request.form("Postal_State")) & "'"
      else
        SQL = SQL & ",Postal_State=NULL"               
      end if              
      if not isblank(request.form("Postal_State_Other")) then
        SQL = SQL & ",Postal_State_Other="& "'" & replacequote(request.form("Postal_State_Other")) & "'"         
      else
        SQL = SQL & ",Postal_State_Other=NULL"      
      end if              
    end if
  
    ' Postal Postal Code
    if not isblank(request.form("Postal_Postal_Code")) then 
      SQL = SQL & ",Postal_Postal_Code="& "'" & replacequote(request.form("Postal_Postal_Code")) & "'"         
    else
      SQL = SQL & ",Postal_Postal_Code=NULL"
    end if
  
    ' Postal_Country
    if not isblank(request.form("Postal_Country")) then 
      SQL = SQL & ",Postal_Country="& "'" & replacequote(request.form("Postal_Country")) & "'"         
    else
      SQL = SQL & ",Postal_Country=NULL"
    end if
    
    ' Shipping Mail Stop
    if not isblank(request.form("Shipping_MailStop")) then 
      SQLCart = SQLCart & ",Shipping_MailStop="& "'" & replacequote(request.form("Shipping_MailStop")) & "'"
    else
      SQLCart = SQLCart & ",Shipping_MailStop=NULL"
    end if
  
    ' Shipping Address
    if not isblank(request.form("Shipping_Address")) then 
      SQLCart = SQLCart & ",Shipping_Address="& "'" & replacequote(request.form("Shipping_Address")) & "'"
    else
      SQLCart = SQLCart & ",Shipping_Address=NULL"
    end if
    
    ' Shipping Address 2
    if not isblank(request.form("Shipping_Address_2")) then 
      SQLCart = SQLCart & ",Shipping_Address_2="& "'" & replacequote(request.form("Shipping_Address_2")) & "'"
    else
      SQLCart = SQLCart & ",Shipping_Address_2=NULL"
    end if
  
    ' Shipping City
    if not isblank(request.form("Shipping_City")) then 
      SQLCart = SQLCart & ",Shipping_City="& "'" & replacequote(request.form("Shipping_City")) & "'"
    else
      SQLCart = SQLCart & ",Shipping_City=NULL"    
    end if
  
    ' Shipping State  
    if not isblank(request.form("Shipping_State")) > 0 or not isblank(request.form("Shipping_State_Other")) then 
      if not isblank(request.form("Shipping_State")) then
        SQLCart = SQLCart & ",Shipping_State="& "'" & replacequote(request.form("Shipping_State")) & "'"
      else
        SQLCart = SQLCart & ",Shipping_State=NULL"               
      end if              
      if not isblank(request.form("Shipping_State_Other")) then
        SQLCart = SQLCart & ",Shipping_State_Other="& "'" & replacequote(request.form("Shipping_State_Other")) & "'"         
      else
        SQLCart = SQLCart & ",Shipping_State_Other=NULL"      
      end if              
    end if
  
    ' Shipping Postal Code
    if not isblank(request.form("Shipping_Postal_Code")) then 
      SQLCart = SQLCart & ",Shipping_Postal_Code="& "'" & replacequote(request.form("Shipping_Postal_Code")) & "'"         
    else
      SQLCart = SQLCart & ",Shipping_Postal_Code=NULL"
    end if
  
    ' Shipping_Country
    if not isblank(request.form("Shipping_Country")) then 
      SQLCart = SQLCart & ",Shipping_Country="& "'" & replacequote(request.form("Shipping_Country")) & "'"         
    else
      SQLCart = SQLCart & ",Shipping_Country=NULL"
    end if
    
    ' --------------------------------------------------------------------------------------
    ' If not in Shopping Cart Mode, then Changes to Shipping Information are updated in User's Profile
    ' otherwise they are posted to SiteWide.Shopping_Cart_Ship_To
    ' --------------------------------------------------------------------------------------
    
    if CInt(Cart_Mode) = CInt(False) then      
      SQL = SQL & SQLCart
    end if
  
    ' --------------------------------------------------------------------------------------  
    ' Business Phone
    if not isblank(request.form("Business_Phone")) then
      if User_Region = 1 or User_Region = 3 then 
        SQL = SQL & ",Business_Phone="& "'" & FormatPhone(replacequote(request.form("Business_Phone"))) & "'"         
      else
        SQL = SQL & ",Business_Phone="& "'" & replacequote(request.form("Business_Phone")) & "'"         
      end if  
    else
      SQL = SQL & ",Business_Phone=NULL"
    end if
    
    ' Business Phone Extension
    if not isblank(request.form("Business_Phone_Extension")) then 
      SQL = SQL & ",Business_Phone_Extension="& "'" & replacequote(request.form("Business_Phone_Extension")) & "'"
    else
      SQL = SQL & ",Business_Phone_Extension=NULL"
    end if
  
    ' Business_Phone_2
    if not isblank(request.form("Business_Phone_2")) then 
      if User_Region = 1 or User_Region = 3 then 
        SQL = SQL & ",Business_Phone_2="& "'" & FormatPhone(replacequote(request.form("Business_Phone_2"))) & "'"
      else
        SQL = SQL & ",Business_Phone_2="& "'" & replacequote(request.form("Business_Phone_2")) & "'"
      end if
    else
      SQL = SQL & ",Business_Phone_2=NULL"    
    end if
  
    ' Business Phone Extension
    if not isblank(request.form("Business_Phone_2_Extension")) then 
      SQL = SQL & ",Business_Phone_2_Extension="& "'" & replacequote(request.form("Business_Phone_2_Extension")) & "'"
    else
      SQL = SQL & ",Business_Phone_2_Extension=NULL"             
    end if  
    
    ' Business_Fax
    if not isblank(request.form("Business_Fax")) then 
      if User_Region = 1 or User_Region = 3 then
        SQL = SQL & ",Business_Fax="& "'" & FormatPhone(replacequote(request.form("Business_Fax"))) & "'"         
      else
        SQL = SQL & ",Business_Fax="& "'" & replacequote(request.form("Business_Fax")) & "'"         
      end if  
    else
      SQL = SQL & ",Business_Fax=NULL"    
    end if
    
    ' Mobile_Phone
    if not isblank(request.form("Mobile_Phone")) then 
      if User_Region = 1 or User_Region = 3 then
        SQL = SQL & ",Mobile_Phone="& "'" & FormatPhone(replacequote(request.form("Mobile_Phone"))) & "'"         
      else
        SQL = SQL & ",Mobile_Phone="& "'" & replacequote(request.form("Mobile_Phone")) & "'"      
      end if      
    else
      SQL = SQL & ",Mobile_Phone=NULL"    
    end if
    
    ' Pager
    if not isblank(request.form("Pager")) then 
      if User_Region = 1 or User_Region = 3 then
        SQL = SQL & ",Pager="& "'" & FormatPhone(replacequote(request.form("Pager"))) & "'"         
      else
        SQL = SQL & ",Pager="& "'" & replacequote(request.form("Pager")) & "'"
      end if  
    else
      SQL = SQL & ",Pager=NULL"    
    end if
  
    ' Email
    if not isblank(request.form("Email")) then 
      SQL = SQL & ",Email="& "'" & replacequote(request.form("Email")) & "'"         
    else
      SQL = SQL & ",Email=NULL"    
    end if
   
    ' Email Method
    if not isblank(request.form("Email_Method")) then 
      SQL = SQL & ",Email_Method="& "" & replacequote(request.form("Email_Method")) & ""
    else
      SQL = SQL & ",Email_Method="& "0"    
    end if
  
    ' Email_2
    if not isblank(request.form("Email_2")) then 
      SQL = SQL & ",Email_2="& "'" & replacequote(request.form("Email_2")) & "'"         
    else
      SQL = SQL & ",Email_2=NULL"    
    end if
    
    ' Connection_Speed
    if not isblank(request.form("Connection_Speed")) then 
      SQL = SQL & ",Connection_Speed="& "" & replacequote(request.form("Connection_Speed")) & ""
    else
      SQL = SQL & ",Connection_Speed=0"
    end if
   
    ' Language
    if not isblank(request.form("Language")) then 
      SQL = SQL & ",Language="& "'" & replacequote(request.form("Language")) & "'"
    else
      SQL = SQL & ",Language="& "'eng'"    
    end if
  
    ' Subscription 
    if request.form("Subscription") = "on" then           
      SQL = SQL & ",Subscription=-1"
    else
      SQL = SQL & ",Subscription=0"
    end if        
  
    ' Change ID   
    if not isblank(request("Account_ID")) and isnumeric(request.form("Account_ID")) then               
      SQL = SQL & ",ChangeID="& killquote(request.form("Account_ID"))
    end if
  
    ' Change Date
    if not isblank(request.form("ChangeDate")) then
      SQL = SQL & ",ChangeDate="& "'" & request.form("ChangeDate") & "'"         
    end if
    
    ' Auxiliary Fields  (Do not update for User initiated Profile Update)
    for aux = 0 to 9
      if not isblank(request.form("Aux_" & Trim(aux))) then 
        SQL = SQL & ",Aux_" & Trim(aux) & "="& "'" & replacequote(request.form("Aux_" & Trim(aux))) & "'"
      else
        SQL = SQL & ",Aux_" & Trim(aux) & "=NULL"  
      end if
    next
    
    ' Build Ending SQL Statement - Use Session NTLogin Name to prevent user from Modifying ASP script to update other accounts.
    
    strPost_QueryString = SQL
    
    SQLU = "UPDATE UserData SET " & SQL & " WHERE UserData.NTLogin='" & Session("Logon_User") & "'"

    'response.write replace(SQLU,",",",<BR>")
    'response.end

    conn.Execute (SQLU)

    New_Account_ID = CInt(request.form("Account_ID"))
    
    ' Complete Key=Value set by filling in exsisting data from record for missing Keys
    
    SQL_Fields = "SELECT * FROM UserData Where ID=" & New_Account_ID
    Set rsFields = Server.CreateObject("ADODB.Recordset")
    rsFields.Open SQL_Fields, conn, 3, 3
    
    for each field In rsFields.Fields
    
      select case UCase(field.Name)
        case "BROADCAST_DATE", "SUBSCRIPTION_DATE", "SUBSCRIPTION_FREQUENCY", _
             "SUBSCRIPTION_METHOD", "SUBSCRIPTION_OPTIONS", _
             "SUBSCRIPTION_SECTIONS", "RTE_ENABLED", "REG_SITE_CODE"
        case else
          if instr(1,UCase(strPost_QueryString),"," & UCase(field.Name) & "=") = 0 and _
             instr(1,Mid(UCase(strPost_QueryString),1,instr(1,UCase(strPost_QueryString),"=")),UCase(field.Name) & "=") = 0 then
            strPost_QueryString = strPost_QueryString & "," & Construct_SQLKeyValueSet(field.Name, field.Type, field.Value, True)
          end if
      end select    
      
    next
    
    rsFields.close
    set rsFields = nothing
    
    Action = "Update"
  
    ' Determine if Posting is required to a Regional (CM) Contact Management System,
    ' if false notify admin default email method.

    select case CInt(CMS_Region(User_Region))
      case CInt(True)       ' CM Sytem Only
        Call Send_2_CMS     ' Send Data to CM Sytem
      case -2               ' CM System and Email
        Call Send_2_CMS     ' Send Data to CM Sytem        
    end select
    
    if CInt(Send_2_CMS_Debug) = CInt(True) and not isblank(ErrorMessage) then
      response.write ErrorMessage & "<P>"
      response.flush
      response.end
    end if

  end if
  
  ' Auto Renew Expiration Date
  if CInt(CMS_Region(User_Region)) = CInt(False) then   ' Update Non CMS Systems Only

    ' Update Logon Date for all accounts for user that have been logged
    Last_Logon = Now()
    SQLUser = "UPDATE UserData SET Logon='" & Last_Logon & "' WHERE Logon IS NOT NULL AND NTLogin='" & Session("Logon_User") & "' AND NewFlag=0"
    conn.Execute (SQLUser)

    SQLUser = "SELECT NTLogin, ExpirationDate, Site_ID, Logon FROM UserData WHERE UserData.NTLogin='" & Session("Logon_User") & "' AND NewFlag=0"
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQLUser, conn, 3, 3
    
    do while not rsUser.EOF
    
      if CDate(rsUser("ExpirationDate")) < CDate("9/9/9999") then

        ' Get Auto Renew Info from Site
        SQLSite = "SELECT ID, Renew_Days FROM Site WHERE ID=" & rsUser("Site_ID")    
        Set rsSite = Server.CreateObject("ADODB.Recordset")
        rsSite.Open SQLSite, conn, 3, 3
   
        if not rsSite.EOF then

          ' If New Date > Existing Account Date, Push Date out for Recriprical Accounts
          if CDate(request.form("ExpirationDate")) > CDate(rsUser("ExpirationDate")) then
            SQLU = "UPDATE UserData SET ExpirationDate='" & CDate(request.form("ExpirationDate")) & "' " &_
                   "WHERE NTLogin='" & Session("Logon_User") & "' AND Site_ID=" & rsUser("Site_ID")          
            conn.Execute (SQLU)

          ' Auto Renew - if greater than current expiration Date, then Push Date out for Recriprical Accounts.
          elseif CInt(rsSite("Renew_Days")) > 0 and CDate(DateAdd("d",CInt(rsSite("Renew_Days")),Date())) > CDate(Last_Logon) and CDate(DateAdd("d",CInt(rsSite("Renew_Days")),Date())) > CDate(rsUser("ExpirationDate")) then
            SQLU = "UPDATE UserData SET ExpirationDate='" & CDate(DateAdd("d",CInt(rsSite("Renew_Days")),Last_Logon)) & "' " &_
                   "WHERE NTLogin='" & Session("Logon_User") & "' AND Site_ID=" & rsUser("Site_ID")          
            conn.Execute (SQLU)
          end if

        end if
        rsSite.close
        set rsSite = Nothing
        
      end if
      
      rsUser.MoveNext
    loop
    
    rsUser.close
    set rsUser = nothing
    
  end if      
  
  Call Disconnect_SiteWide
       
  BackURL = replace(BackURL,"CINN=0","CINN=1")
  
  if CInt(Send_2_CMS_Debug) = CInt(False) then
    response.redirect BackURL
  else
    response.flush
    response.end
  end if  

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

' Send_2_CMS
%>
<!--#include virtual="/Include/Function_DCMDataTransfer.asp"-->

