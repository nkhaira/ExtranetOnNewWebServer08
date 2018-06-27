<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author: K. D. Whitlock
' Date:   02/01/2000
' 06/19/2002 - Added Fields and Re-Ordered Form to work with Euro DCM
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_DB.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/DCMHTTP_DataTransfer.asp"-->
<%
' on error resume next
Call Connect_SiteWide
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim DebugFlag
DebugFlag = false

Dim MailTo
Dim MailFrom
Dim ReplyToName
Dim ReplyTo
Dim MailCCName
Dim MailCC
Dim MailBCCName
Dim MailBCC
Dim MailSubject
Dim MailMessage

Dim User_Name
Dim User_Login
Dim User_Password
Dim User_Email
Dim User_Region
Dim Account_ID
Dim New_Account_ID
Dim Action
Dim Groups
Dim SubGroups
Dim Admin_Access
Dim ErrorMessage
Dim ErrorPriority
Dim strPost_QueryString

ErrorPriority   = 0
Account_ID      = request("ID")
Site_ID         = request("Site_ID")
Login_Language  = request("Language")
User_Email      = request("EMail")

if request("Business") = "-1" then Business = True else Business = False

' --------------------------------------------------------------------------------------
' Open Connection to Mail Server
' --------------------------------------------------------------------------------------

Set Mailer = Server.CreateObject("SMTPsvg.Mailer") 

%>
<!--#include virtual="/connections/connection_EMail.asp"-->
<!--#include virtual="/connections/connection_EMail_Timeout.asp"--> 
<%
    
' --------------------------------------------------------------------------------------
' Get Site Information
' --------------------------------------------------------------------------------------

SQL = "SELECT * FROM Site Where Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Groups = rsSite("Site_Code")
Site_Description = rsSite("Site_Description")
Logo = rsSite("Logo")
Footer_Disabled = rsSite("Footer_Disabled")
RFT = rsSite("Business")

MailFromName    = rsSite("FromName")
MailFromAddress = rsSite("FromAddress")
MailReplyToName = rsSite("ReplyToName")
MailReplyTo     = rsSite("ReplyTo")
MailCCName      = rsSite("MailCCName")
MailCC          = rsSite("MailCC")
MailBCCName     = rsSite("MailBCCName")
MailBCC         = rsSite("MailBCC")

Dim CMS_Region(3)

CMS_Region(1) = CInt(rsSite("CMS_Region_1"))
CMS_Region(2) = CInt(rsSite("CMS_Region_2"))
CMS_Region(3) = CInt(rsSite("CMS_Region_3"))

rsSite.close
set rsSite = nothing

' --------------------------------------------------------------------------------------
' Configure EMail Header Information
' --------------------------------------------------------------------------------------

Mailer.ClearAllRecipients

' --------------------------------------------------------------------------------------
' Send Login and Password Information
' --------------------------------------------------------------------------------------

if instr(1,lcase(request("Action")),"send password") > 0 then

  Call Get_UserData
  
  Mailer.ReturnReceipt = False  
  Mailer.Priority      = 3
  Mailer.QMessage      = False
  ' Use Primary Site Administrator's Info
  Mailer.FromName      = MailFromName
  Mailer.FromAddress   = MailFromAddress 

  Mailer.ReplyTo       = MailReplyTo
  Mailer.AddRecipient    User_Name, User_Email
  
  if not isblank(MailBCC) then
    Mailer.AddBCC        MailBCCName, MailBCC
  end if
  
  MailSubject          = Translate("Information Requested",Alt_Language,conn)
    
  MailMessage = Translate("This is an automated notification message from the",Alt_Language,conn) & vbCrLf
  MailMessage = MailMessage & MailFromName & " " & Translate("Extranet Support Site",Alt_Language,conn) & " :" & vbCrLf & vbCrLf
  MailMessage = MailMessage & "http://support.fluke.com/" & lcase(groups) & vbCrLf & vbCrLf

  MailMessage = MailMessage & vbCrLf & Translate("Here is the information that you have requested",Alt_Language,conn) & ":" & vbCrLf & vbCrLf
  MailMessage = MailMessage & Translate("User Name",Alt_Language,conn) & " : " & User_Login & vbCrLf
  MailMessage = MailMessage & Translate("Password",Alt_Language,conn)  & "  : " & User_Password & vbCrLf & vbCrLf
  MailMessage = MailMessage & Translate("Sincerely",Alt_Language,conn) & "," & vbCrLf & vbCrLf & MailFromName & " " & Translate("Support Team",Alt_Language,conn)
  
  Call Send_EMail

  if not isblank(ErrorMessage) then
    Call ErrorHandler
  else
    Call Disconnect_SiteWide
    response.redirect "/" & Groups       
'    response.redirect "/register/default.asp"    
  end if

' --------------------------------------------------------------------------------------
' New 
' --------------------------------------------------------------------------------------

elseif instr(1,lcase(request("Action")),"registration") > 0 then
    
  if lcase(Account_ID) = "new" then
  
    Mailer.ReturnReceipt  = False
    Mailer.Priority       = 3
  
    ' Use Primary Site Administrator's Info
    Mailer.FromName       = MailFromName
    Mailer.FromAddress    = MailFromAddress
    ' Domain Administrator - Testing Purposes
    'Modified by zensar for replacing Kelly's mail id.below statement is commented out as there is no need to send
    'an email to other person's except the site administrator.
    '----------------
    'Mailer.AddBCC        "K. David Whitlock", "David.Whitlock@Fluke.com"
    'Mailer.AddBCC         "Extranet Admin Group","extranetalerts@fluke.com"
    '>>>>>>>>>>>>>>>>>>>>>>>>>
    ' Get New Account's Region
    
    SQL       = "SELECT Country.* "
    SQL = SQL & "FROM Country "
    SQL = SQL & "WHERE Country.Abbrev='" & Request("Business_Country") & "'"
    
    Set rsCountry = Server.CreateObject("ADODB.Recordset")
    rsCountry.Open SQL, conn, 3, 3
    
    if not rsCountry.EOF then
      User_Region = CInt(rsCountry("Region"))
    else
      User_Region = 1         ' Default to US
    end if
    
    rsCountry.close
    set rsCountry = nothing    

    ' Check to see if there is an Account Administrator Assigned  

    SQL =       "SELECT Approvers_Account.* "
    SQL = SQL & "FROM Approvers_Account "
    SQL = SQL & "WHERE Approvers_Account.Site_ID=" & request("Site_ID") & " "
    SQL = SQL & "AND Approvers_Account.Region=" & User_Region & " "
    SQL = SQL & "AND Approvers_Account.Approver_ID<>0"

    Set rsApprovers = Server.CreateObject("ADODB.Recordset")
    rsApprovers.Open SQL, conn, 3, 3

    Admin_Language = "eng"
    Approver_Flag  = False    

    if not rsApprovers.EOF then
    
      if CInt(rsApprovers("Email_Site_Admin")) = CInt(True) then
        Mailer.AddBCC           MailFromName, MailFromAddress
      end if
    
      SQL       = "SELECT UserData.* "
      SQL = SQL & "FROM UserData "
      SQL = SQL & "WHERE ID=" & rsApprovers("Approver_ID")

      Set rsUser = Server.CreateObject("ADODB.Recordset")
      rsUser.Open SQL, conn, 3, 3

      if not rsUser.EOF then
        if Approver_Flag = False then
          Mailer.ReplyTo      = rsUser("EMail")
          Mailer.AddRecipient   rsUser("FirstName") & " " & rsUser("LastName"), rsUser("EMail")
          Approver_Flag       = True
          Admin_Language      = rsUser("Language")
        else
          Mailer.AddCC        rsUser("FirstName") & " " & rsUser("LastName"), rsUser("EMail")
        end if    
      end if

      rsUser.Close
      set rsUser = nothing
    
    end if
    
    rsApprovers.Close
    set rsApprovers = nothing
    
    ' No Approvers then Default to Site Admin

    if not Approver_Flag then 
      Mailer.ReplyTo        = MailReplyTo    
      Mailer.AddRecipient     MailFromName, MailFromAddress

      if not isblank(MailCC) then
        Mailer.AddCC          MailCCName, MailCC
      end if
    end if
         
    MailSubject = Translate("New Account Request Notice",Admin_Language,conn)
      
    MailMessage = Translate("This is an automated notification message from the",Admin_Language,conn) & vbCrLf
    MailMessage = MailMessage & Translate(MailFromName,Admin_Language,conn) & " " & Translate("Extranet Support Server",Admin_Language,conn) & ":" & vbCrLf
    MailMessage = MailMessage & "http://support.fluke.com/" & lcase(groups) & vbCrLf & vbCrLf & vbCrLf
    MailMessage = MailMessage & Translate("To Approve or Disapprove this New Account Request, click on the following URL:",Admin_Language,conn) & vbCrLf & vbCrLf
    MailMessage = MailMessage & "http://support.fluke.com/" & lcase(groups) & vbCrLf & vbCrLf
    MailMessage = MailMessage & Translate("Logon to the site then click on the Account Administrators navigation button, then select the 'Account Administrators - Tool Kit' link. Check your 'Approve - New Users Account Profile Request' queue.",Admin_Language,conn) & vbCrLf & vbCrLf & vbCrLf

    MailMessage = MailMessage & Translate("Here is the New Account Profile Information as Supplied by the Requestor",Admin_Language,conn) & ":" & vbCrLf & vbCrLf

    MailMessage = MailMessage & "--------------------------------------------" & vbCrLf
    MailMessage = MailMessage & Translate("User Information",Admin_Language,conn) & vbCrLf
    MailMessage = MailMessage & "--------------------------------------------" & vbCrLf & vbCrLf  
    
  end if
  
  ' Begin Building SQL Statement
  
  sqlf = ""
  sql  = ""
  sqlu = ""

  if not isblank(Site_ID) then
    sqlf = sqlf & "Site_ID"
    sql  = sql  & "" & killquote(request("Site_ID")) & ""
    sqlu = sqlu & "Site_ID=" & killquote(request("Site_ID"))
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Internal Site ID, report this error to the site Webmaster.",Login_Language,conn) & "</LI>"
    ErrorPriority = 1
  end if
  
  ' ExpirationDate
  DateString = "12/31/" & DatePart("yyyy",Date)                   
  ''>>>>>>>
  sqlf = sqlf & ",ExpirationDate"
  sql  = sql  & ",'" & DateString & "'"
  sqlu = sqlu & ",ExpirationDate=" & "'" & DateString & "'"

  ' Type Code

  if not isblank(request("Type_Code")) then 
    sqlf = sqlf & ",Type_Code"       
    sql  = sql  & "," & request("Type_Code")
    sqlu = sqlu & ",Type_Code=" & request("Type_Code")
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Customer Type (Relationship).",Login_Language,conn) & "</LI>"
  end if
  
  ' Gender / Prefix
  if not isblank(request("Gender")) then
    select case request("Gender")
      case 0    ' Male
        sqlf = sqlf & ",Prefix"       
        sql  = sql  & ",'Mr'"
        sqlu = sqlu & ",Prefix='Mr'"
        sqlf = sqlf & ",Gender"       
        sql  = sql  & ",0"
        sqlu = sqlu & ",Gender=0"
      case 1    ' Female
        sqlf = sqlf & ",Prefix"       
        sql  = sql  & ",'Ms'"
        sqlu = sqlu & ",Prefix='Ms'"
        sqlf = sqlf & ",Gender"       
        sql  = sql  & ",1"
        sqlu = sqlu & ",Gender=1"
      case else ' Unknown
        sqlf = sqlf & ",Prefix"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Prefix=NULL"
        sqlf = sqlf & ",Gender"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Gender=NULL" 
    end select       
  else
    sqlf = sqlf & ",Prefix"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Prefix=NULL"
    sqlf = sqlf & ",Gender"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Gender=NULL"    
  end if

  ' First Name
  if not isblank(request("FirstName")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",FirstName"       
      sql  = sql  & ",N'" & ProperCase(replacequote(request("FirstName"))) & "'"
      sqlu = sqlu & ",FirstName=" & "N'" & ProperCase(replacequote(request("FirstName"))) & "'"
    else  
      sqlf = sqlf & ",FirstName"       
      sql  = sql  & ",N'" & replacequote(request("FirstName")) & "'"
      sqlu = sqlu & ",FirstName=" & "N'" & replacequote(request("FirstName")) & "'"
    end if   
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing First Name",Login_Language,conn) & "</LI>"         
  end if

  ' Middle Name
  if not isblank(request("MiddleName")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",MiddleName"       
      sql  = sql  & ",N'" & ProperCase(ReplaceQuote(request("MiddleName"))) & "'"
      sqlu = sqlu & ",MiddleName=" & "N'" & ProperCase(replacequote(request("MiddleName"))) & "'"
    else
      sqlf = sqlf & ",MiddleName"       
      sql  = sql  & ",N'" & ReplaceQuote(request("MiddleName")) & "'"
      sqlu = sqlu & ",MiddleName=" & "N'" & replacequote(request("MiddleName")) & "'"
    end if  
  else
    sqlf = sqlf & ",MiddleName"
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",MiddleName=NULL"
  end if

  ' Last Name
  if not isblank(request("LastName")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",LastName"       
      sql  = sql  & ",N" & ProperCase(ReplaceQuote(request("LastName"))) & "'"
      sqlu = sqlu & ",LastName=" & "N'" & ProperCase(replacequote(request("LastName"))) & "'"
    else
      sqlf = sqlf & ",LastName"       
      sql  = sql  & ",N'" & ReplaceQuote(request("LastName")) & "'"
      sqlu = sqlu & ",LastName=" & "N'" & replacequote(request("LastName")) & "'"
    end if  
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Last Name",Login_Language,conn) & "</LI>"         
  end if
    
  ' Suffix
  if not isblank(request("Suffix")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Suffix"       
      sql  = sql  & ",N'" & ProperCase(replacequote(request("Suffix"))) & "'"
      sqlu = sqlu & ",Suffix=" & "N'" & ProperCase(replacequote(request("Suffix"))) & "'"
    else
      sqlf = sqlf & ",Suffix"       
      sql  = sql  & ",N'" & replacequote(request("Suffix")) & "'"
      sqlu = sqlu & ",Suffix=" & "N'" & replacequote(request("Suffix")) & "'"
    end if  
  else
    sqlf = sqlf & ",Suffix"
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Suffix=NULL"
  end if 

  ' Initials
  if not isblank(request("Initials")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Initials"       
      sql  = sql  & ",N'" & UCase(ReplaceQuote(request("Initials"))) & "'"
      sqlu = sqlu & ",Initials=" & "N'" & UCase(replacequote(request("Initials"))) & "'"
    else
      sqlf = sqlf & ",Initials"       
      sql  = sql  & ",N'" & ReplaceQuote(request("Initials")) & "'"
      sqlu = sqlu & ",Initials=" & "N'" & replacequote(request("Initials")) & "'"
    end if  
  else
    sqlf = sqlf & ",Initials"
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Initials=NULL"
  end if

  MailMessage     = MailMessage & Translate("Name",Admin_Language,conn) & ": "  
  MailMessage   = MailMessage & request("FirstName")

  if not isblank(request("MiddleName")) then
    MailMessage = MailMessage & " " & request("MiddleName")
  end if
  
  MailMessage   = MailMessage & " " & request("LastName")

  if not isblank(request("Suffix")) then
    MailMessage = MailMessage     & ", " & request("Suffix") & vbCrLf
  else
    MailMessage = MailMessage     & vbCrLf
  end if

  MailMessage   = MailMessage & Translate("Email Address",Admin_Language,conn) & ": " & request("EMail") & vbCrLf & vbCrLf
    
  ' Company
  if not isblank(request("Company")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Company"       
      sql  = sql  & ",N'" & ProperCase(ReplaceQuote(request("Company"))) & "'"
      sqlu = sqlu & ",Company=" & "N'" & ProperCase(replacequote(request("Company"))) & "'"
      MailMessage = MailMessage & Translate("Company",Admin_Language,conn) & ": " & ProperCase(request("Company")) & vbCrLf        
    else
      sqlf = sqlf & ",Company"       
      sql  = sql  & ",N'" & ReplaceQuote(request("Company")) & "'"
      sqlu = sqlu & ",Company=" & "N'" & replacequote(request("Company")) & "'"
      MailMessage = MailMessage & Translate("Company",Admin_Language,conn) & ": " & request("Company") & vbCrLf        
    end if  

  elseif business then
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Company",Login_Language,conn) & "</LI>"         
  end if

    ' Company Website
  if not isblank(request("Company_Website")) then
      sqlf = sqlf & ",Company_Website"       
      sql  = sql  & ",N'" & replace(replace(KillQuote(LCase(request("Company_Website"))),"http://",""),"https://","") & "'"
      sqlu = sqlu & ",Company_Website=" & "N'" & replace(replace(KillQuote(LCase(request("Company_Website"))),"http://",""),"https://","") & "'"
  else
    sqlf = sqlf & ",Company_Website"
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Company_Website=NULL"
  end if

  
  ' Job Title
  if not isblank(request("Job_Title")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Job_Title"       
      sql  = sql  & ",N'" & ProperCase(ReplaceQuote(request("Job_Title"))) & "'"
      sqlu = sqlu & ",Job_Title=" & "N'" & ProperCase(replacequote(request("Job_Title"))) & "'"
      MailMessage = MailMessage & Translate("Job Title",Admin_Language,conn) & ": " & ProperCase(request("Job_Title")) & vbCrLf & vbCrLf
    else
      sqlf = sqlf & ",Job_Title"       
      sql  = sql  & ",N'" & ReplaceQuote(request("Job_Title")) & "'"
      sqlu = sqlu & ",Job_Title=" & "N'" & replacequote(request("Job_Title")) & "'"
      MailMessage = MailMessage & Translate("Job Title",Admin_Language,conn) & ": " & request("Job_Title") & vbCrLf & vbCrLf
    end if  
  elseif not business then
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Job Title",Login_Language,conn) & "</LI>"         
  end if
  
  ' NT Login
  if not isblank(request("NTLogin")) and (len(request("NTLogin")) >= 7 and len(request("NTLogin")) <=20) then            
    ' Check SiteWide to be unique
    SQL_Login = "SELECT UserData.NTLogin FROM UserData WHERE UserData.NTLogin='" & request("NTLogin") & "'"
    Set rsCkLogin = Server.CreateObject("ADODB.Recordset")
    rsCkLogin.Open SQL_Login, conn, 3, 3
    If not rsCkLogin.EOF then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("The New Logon User Name that you have requested is not available. Please select an alternate Logon User Name.",Login_Language,conn) & "</LI>"
    else
      sqlf = sqlf & ",NTLogin"       
      sql  = sql  & ",N'" & Trim(killquote(request("NTLogin"))) & "'"
      sqlu = sqlu & ",NTLogin=" & "N'" & Trim(Killquote(request("NTLogin"))) & "'"
    end if
    rsCkLogin.close
    set rsCkLogin = nothing
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Logon User Name or the length of your Logon User Name is less than 7 characers",Login_Language,conn) & ", " & Translate("maximum",Login_Language,conn) & "20</LI>"         
  end if
  
  ' Core ID
  if not isblank(request("Core_ID")) then            
    sqlf = sqlf & ",Core_ID"       
    sql  = sql  & ",'" & killquote(request("Core_ID")) & "'"
    sqlu = sqlu & ",Core_ID=" & "'" & Killquote(request("Core_ID")) & "'"
  else
    sqlf = sqlf & ",Core_ID"
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Core_ID=NULL"
  end if

  ' eStore ID
  if not isblank(request("MSCSSID")) then            
    sqlf = sqlf & ",eStore_ID"       
    sql  = sql  & ",'" & killquote(request("eStore_ID")) & "'"
    sqlu = sqlu & ",eStore_ID=" & "'" & Killquote(request("eStore_ID")) & "'"
  else
    sqlf = sqlf & ",eStore_ID"
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",eStore_ID=NULL"
  end if

  ' Password
  if not isblank(request("Password")) and (len(request("Password")) >= 7 and len(request("Password")) <= 14) then
    sqlf = sqlf & ",Password"       
    sql  = sql  & ",N'" & Trim(replacequote(request("Password"))) & "'"
    sqlu = sqlu & ",Password=" & "N'" & Trim(replacequote(request("Password"))) & "'"
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Password or the length of your Password is less than 7 characters",Login_Language,conn) & ", " & Translate("maximum",Login_Language,conn) & "14</LI>"         
  end if

  ' Password Confirm
  if isblank(request("Password_Confirm")) then 
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Password Confirm",Login_Language,conn) & "</LI>"         
  end if

  ' Password and Password Confirm Match
  if not isblank(request("Password")) and not isblank(request("Password_Confirm")) then
    if request("Password") <> request("Password_Confirm") then 
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Password and Password Confirm do not match",Login_Language,conn) & "</LI>"
    end if
  end if
  
  ' EMail Format
  if not isblank(request("EMail_Method")) then            
    sqlf = sqlf & ",EMail_Method"       
    sql  = sql  & "," & killquote(request("EMail_Method")) & ""
    sqlu = sqlu & ",Email_Method=" & "" & Killquote(request("EMail_Method")) & ""
  end if 

  ' Subscription 
  if request("Subscription") = "on" or request("Subscription") = "-1" then           
    sqlf = sqlf & ",Subscription"
    sql  = sql  & "," & CInt(True)     
    sqlu = sqlu & ",Subscription=" & CInt(True)
    MailMessage = MailMessage & "Subscription Service: Enabled" & vbCrLf    
  else
    sqlf = sqlf & ",Subscription"
    sql  = sql  & "," & CInt(False)
    sqlu = sqlu & ",Subscription=" & CInt(False)
    MailMessage = MailMessage & "Subscription Service: Disabled" & vbCrLf    
  end if        

  ' Connection Speed
  if not isblank(request("Connection_Speed")) then            
    sqlf = sqlf & ",Connection_Speed"       
    sql  = sql  & "," & killquote(request("Connection_Speed")) & ""
    sqlu = sqlu & ",Connection_Speed=" & "" & Killquote(request("Connection_Speed")) & ""
  end if 
  
  ' Region
  sqlf = sqlf & ",Region"       
  sql  = sql  & "," & User_Region
  sqlu = sqlu & ",Region=" & User_Region
    
  MailMessage = MailMessage & "--------------------------------------------" & vbCrLf
  MailMessage = MailMessage & Translate("Office Address",Admin_Language,conn) & vbCrLf
  MailMessage = MailMessage & "--------------------------------------------" & vbCrLf & vbCrLf  
  
  ' Business Address
  if not isblank(request("Business_Address")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Business_Address"       
      sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_Address"))) & "'"
      sqlu = sqlu & ",Business_Address=" & "N'" & ProperCase(replacequote(request("Business_Address"))) & "'"
      MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 1: " & ProperCase(request("Business_Address")) & vbCrLf
    else
      sqlf = sqlf & ",Business_Address"       
      sql  = sql  & ",N'" & replacequote(request("Business_Address")) & "'"
      sqlu = sqlu & ",Business_Address=" & "N'" & replacequote(request("Business_Address")) & "'"
      MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 1: " & request("Business_Address") & vbCrLf
    end if  
  else
    if business then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Address",Login_Language,conn) & "</LI>"
    else
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Address",Login_Language,conn) & "</LI>"
    end if        
  end if
  
  ' Business Address 2
  if not isblank(request("Business_Address_2")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Business_Address_2"       
      sql  = sql  & ",N'" & ProperCase(ReplaceQuote(request("Business_Address_2"))) & "'"
      sqlu = sqlu & ",Business_Address_2=" & "N'" & ProperCase(replacequote(request("Business_Address_2"))) & "'"
      MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 2: " & ProperCase(request("Business_Address_2")) & vbCrLf        
    else
      sqlf = sqlf & ",Business_Address_2"       
      sql  = sql  & ",N'" & ReplaceQuote(request("Business_Address_2")) & "'"
      sqlu = sqlu & ",Business_Address_2=" & "N'" & replacequote(request("Business_Address_2")) & "'"
      MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 2: " & request("Business_Address_2") & vbCrLf        
    end if  
  else
    sqlf = sqlf & ",Business_Address_2"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Business_Address_2=NULL"
  end if

  ' Business Mail Stop
  if not isblank(request("Business_MailStop")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Business_MailStop"       
      sql  = sql  & ",N'" & Killquote(request("Business_MailStop")) & "'"
      sqlu = sqlu & ",Business_MailStop=" & "N'" & UCase(replacequote(request("Business_MailStop"))) & "'"
      MailMessage = MailMessage & Translate("Mail Stop",Admin_Language,conn) & ": " & UCase(request("Business_MailStop")) & vbCrLf        
    else
      sqlf = sqlf & ",Business_MailStop"       
      sql  = sql  & ",N'" & Killquote(request("Business_MailStop")) & "'"
      sqlu = sqlu & ",Business_MailStop=" & "N'" & replacequote(request("Business_MailStop")) & "'"
      MailMessage = MailMessage & Translate("Mail Stop",Admin_Language,conn) & ": " & request("Business_MailStop") & vbCrLf        
    end if  
  else
    sqlf = sqlf & ",Business_MailStop"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Business_MailStop=NULL"
  end if

  ' Business City
  if not isblank(request("Business_City")) then 
    if user_region = 1 or user_region = 3 then
      sqlf = sqlf & ",Business_City"       
      sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_City"))) & "'"
      sqlu = sqlu & ",Business_City=" & "N'" & ProperCase(replacequote(request("Business_City"))) & "'"
      MailMessage = MailMessage & Translate("City",Admin_Language,conn) & ": " & ProperCase(request("Business_City")) & vbCrLf
    else
      sqlf = sqlf & ",Business_City"       
      sql  = sql  & ",N'" & replacequote(request("Business_City")) & "'"
      sqlu = sqlu & ",Business_City=" & "N'" & replacequote(request("Business_City")) & "'"
      MailMessage = MailMessage & Translate("City",Admin_Language,conn) & ": " & request("Business_City") & vbCrLf
    end if  
  else
    if business then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office City",Login_Language,conn) & "</LI>"         
    else
      ErrorMessage = ErrorMessage & "<LI>Missing City</LI>"               
    end if  
  end if

  ' Business State
  if not isblank(request("Business_State")) or not isblank(request("Business_State_Other")) then 
    if not isblank(request("Business_State")) then
      sqlf = sqlf & ",Business_State"       
      sql  = sql  & ",N'" & replacequote(request("Business_State")) & "'"
      sqlu = sqlu & ",Business_State=" & "N'" & replacequote(request("Business_State")) & "'"
      if request("Business_State") <> "ZZ" then
        MailMessage = MailMessage & Translate("USA State or Canadian Province",Admin_Language,conn) & ": " & request("Business_State") & vbCrLf
      end if
    end if
    if not isblank(request("Business_State_Other")) then
      if user_region = 1 or user_region = 3 then
        sqlf = sqlf & ",Business_State_Other"       
        sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_State_Other"))) & "'"
        sqlu = sqlu & ",Business_State_Other=" & "N'" & ProperCase(replacequote(request("Business_State_Other"))) & "'"
        MailMessage = MailMessage & Translate("Other State, Province or Local",Admin_Language,conn) & ": " & ProperCase(request("Business_State_Other")) & vbCrLf
      else
        sqlf = sqlf & ",Business_State_Other"       
        sql  = sql  & ",N'" & replacequote(request("Business_State_Other")) & "'"
        sqlu = sqlu & ",Business_State_Other=" & "N'" & replacequote(request("Business_State_Other")) & "'"
        MailMessage = MailMessage & Translate("Other State, Province or Local",Admin_Language,conn) & ": " & request("Business_State_Other") & vbCrLf
      end if  
    end if              
  else
    if business then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office State / Province (drop-down) or Other Office State / Province (free-form)",Login_Language,conn) & "</LI>"
    else
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing State / Province (drop-down) or Other State / Province (free-form)",Login_Language,conn) & "</LI>"               
    end if  
  end if

  ' Business Postal Code
  if not isblank(request("Business_Postal_Code")) then
    if UCase(request("Business_Country")) = "US" then
      if isnumeric(replace(request("Business_Postal_Code"),"-","")) then
        sqlf = sqlf & ",Business_Postal_Code"       
        sql  = sql  & ",N'" & UCase(replacequote(request("Business_Postal_Code"))) & "'"
        sqlu = sqlu & ",Business_Postal_Code=" & "N'" & UCase(replacequote(request("Business_Postal_Code"))) & "'"
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Invalid Office Postal Code",Login_Language,conn) & "</LI>"
      end if
    else
        sqlf = sqlf & ",Business_Postal_Code"       
        sql  = sql  & ",N'" & replacequote(request("Business_Postal_Code")) & "'"
        sqlu = sqlu & ",Business_Postal_Code=" & "N'" & replacequote(request("Business_Postal_Code")) & "'"
    end if    
    MailMessage = MailMessage & Translate("Postal Code",Admin_Language,conn) & ": " & request("Business_Postal_Code") & vbCrLf        
  else
    if business then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Postal Code",Login_Language,conn) & "</LI>"         
    else
      ErrorMessage = ErrorMessage & "<LI>Missing Postal Code</LI>"             
    end if
  end if

  ' Business_Country
  if not isblank(request("Business_Country")) then 
    sqlf = sqlf & ",Business_Country"       
    sql  = sql  & ",N'" & replacequote(request("Business_Country")) & "'"
    sqlu = sqlu & ",Business_Country=" & "N'" & replacequote(request("Business_Country")) & "'"
    MailMessage = MailMessage & Translate("Country",Admin_Language,conn) & ": " & request("Business_Country") & vbCrLf &vbCrLf       
  else
    if business then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Country",Login_Language,conn) & "</LI>"         
    else
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Country",Login_Language,conn) & "</LI>"
    end if  
  end if

  if business then
  
    MailMessage = MailMessage & "--------------------------------------------" & vbCrLf
    MailMessage = MailMessage & Translate("Postal Address",Admin_Language,conn) & vbCrLf
    MailMessage = MailMessage & "--------------------------------------------" & vbCrLf & vbCrLf    
 
    ' Postal Address
    if not isblank(request("Postal_Address")) then 
      if user_region = 1 or user_region = 3 then
        sqlf = sqlf & ",Postal_Address"       
        sql  = sql  & ",N'" & ProperCase(replacequote(request("Postal_Address"))) & "'"
        sqlu = sqlu & ",Postal_Address=" & "N'" & ProperCase(replacequote(request("Postal_Address"))) & "'"
        MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & "(" & Translate("PO Box",Admin_Language,conn) & ") : " & ProperCase(request("Postal_Address")) & vbCrLf        
      else
        sqlf = sqlf & ",Postal_Address"       
        sql  = sql  & ",N'" & replacequote(request("Postal_Address")) & "'"
        sqlu = sqlu & ",Postal_Address=" & "N'" & replacequote(request("Postal_Address")) & "'"
        MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & "(" & Translate("PO Box",Admin_Language,conn) & ") : " & request("Postal_Address") & vbCrLf        
      end if  
    else
      sqlf = sqlf & ",Postal_Address"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Postal_Address=NULL"
    end if
     
    ' Postal City
    if not isblank(request("Postal_City")) then 
      if user_region = 1 or user_region = 3 then
        sqlf = sqlf & ",Postal_City"       
        sql  = sql  & ",N'" & ProperCase(replacequote(request("Postal_City"))) & "'"
        sqlu = sqlu & ",Postal_City=" & "N'" & ProperCase(replacequote(request("Postal_City"))) & "'"
        MailMessage = MailMessage & Translate("City",Admin_Language,conn) & ": " & ProperCase(request("Postal_City")) & vbCrLf                      
      else
        sqlf = sqlf & ",Postal_City"       
        sql  = sql  & ",N'" & replacequote(request("Postal_City")) & "'"
        sqlu = sqlu & ",Postal_City=" & "N'" & replacequote(request("Postal_City")) & "'"
        MailMessage = MailMessage & Translate("City",Admin_Language,conn) & ": " & request("Postal_City") & vbCrLf                      
      end if  
    else
      sqlf = sqlf & ",Postal_City"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Postal_City=NULL"
    end if
  
    ' Postal State
    
    if not isblank(request("Postal_State")) or not isblank(request("Postal_State_Other")) then 
      if not isblank(request("Postal_State")) and request("Postal_State") <> "ZZ" then
        sqlf = sqlf & ",Postal_State"
        sql  = sql  & ",N'" & replacequote(request("Postal_State")) & "'"
        sqlu = sqlu & ",Postal_State=" & "N'" & replacequote(request("Postal_State")) & "'"
        if request("Postal_State") <> "ZZ" then
          MailMessage = MailMessage & Translate("USA State or Canadian Province",Admin_Language,conn) & ": " & request("Postal_State") & vbCrLf
        end if
      else
        sqlf = sqlf & ",Postal_State"
        sql  = sql  & ",'ZZ'"
        sqlu = sqlu & ",Postal_State=" & "N'" & "ZZ" & "'"
      end if
      if not isblank(request("Postal_State_Other")) then
        if user_region = 1 or user_region = 3 then
          sqlf = sqlf & ",Postal_State_Other"
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Postal_State_Other"))) & "'"
          sqlu = sqlu & ",Postal_State_Other=" & "N'" & ProperCase(replacequote(request("Postal_State_Other"))) & "'"
          MailMessage = MailMessage & Translate("Other State, Province or County",Admin_Language,conn) & ": " & ProperCase(request("Postal_State_Other")) & vbCrLf
        else
          sqlf = sqlf & ",Postal_State_Other"
          sql  = sql  & ",N'" & replacequote(request("Postal_State_Other")) & "'"
          sqlu = sqlu & ",Postal_State_Other=" & "N'" & replacequote(request("Postal_State_Other")) & "'"
          MailMessage = MailMessage & Translate("Other State, Province or County",Admin_Language,conn) & ": " & request("Postal_State_Other") & vbCrLf
        end if  
      end if              
    end if
  
    ' Postal Postal Code
    if not isblank(request("Postal_Postal_Code")) then
      if UCase(request("Postal_Country")) = "US" then
        if isnumeric(replace(request("Postal_Postal_Code"),"-","")) then
          sqlf = sqlf & ",Postal_Postal_Code"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Postal_Postal_Code"))) & "'"
          sqlu = sqlu & ",Postal_Postal_Code=" & "N'" & UCase(replacequote(request("Postal_Postal_Code"))) & "'"
        end if
      else
        sqlf = sqlf & ",Postal_Postal_Code"       
        sql  = sql  & ",N'" & replacequote(request("Postal_Postal_Code")) & "'"
        sqlu = sqlu & ",Postal_Postal_Code=" & "N'" & replacequote(request("Postal_Postal_Code")) & "'"
      end if
      MailMessage = MailMessage & Translate("Postal Code",Admin_Language,conn) & ": " & request("Postal_Postal_Code") & vbCrLf     
    else
      sqlf = sqlf & ",Postal_Postal_Code"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Postal_Postal_Code=NULL"
    end if
  
    ' Postal_Country
    if not isblank(request("Postal_Country")) then 
      sqlf = sqlf & ",Postal_Country"       
      sql  = sql  & ",N'" & replacequote(request("Postal_Country")) & "'"
      sqlu = sqlu & ",Postal_Country=" & "N'" & replacequote(request("Postal_Country")) & "'"
      MailMessage = MailMessage & Translate("Country",Admin_Language,conn) & ": " & request("Postal_Country") & vbCrLf & vbCrLf
    end if

    MailMessage = MailMessage & "--------------------------------------------" & vbCrLf
    MailMessage = MailMessage & Translate("Shipping Address",Admin_Language,conn) & vbCrLf
    MailMessage = MailMessage & "--------------------------------------------" & vbCrLf & vbCrLf    

    ' Shipping Mail Stop
    if not isblank(request("Shipping_MailStop")) then 
      sqlf = sqlf & ",Shipping_MailStop"       
      sql  = sql  & ",N'" & UCase(replacequote(request("Shipping_MailStop"))) & "'"
      sqlu = sqlu & ",Shipping_MailStop=" & "N'" & UCase(replacequote(request("Shipping_MailStop"))) & "'"
      MailMessage = MailMessage & Translate("Mail Stop",Admin_Language,conn) & ": " & UCase(request("Shipping_MailStop")) & vbCrLf        
    else
      sqlf = sqlf & ",Shipping_MailStop"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Shipping_MailStop=NULL"
    end if
  
    ' Shipping Address
    if not isblank(request("Shipping_Address")) then 
      if user_region = 1 or user_region = 3 then
        sqlf = sqlf & ",Shipping_Address"       
        sql  = sql  & ",N'" & ProperCase(replacequote(request("Shipping_Address"))) & "'"
        sqlu = sqlu & ",Shipping_Address=" & "N'" & ProperCase(replacequote(request("Shipping_Address"))) & "'"
        MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 1: " & ProperCase(request("Shipping_Address")) & vbCrLf
      else
        sqlf = sqlf & ",Shipping_Address"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_Address")) & "'"
        sqlu = sqlu & ",Shipping_Address=" & "N'" & replacequote(request("Shipping_Address")) & "'"
        MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 1: " & request("Shipping_Address") & vbCrLf
      end if  
    else
      sqlf = sqlf & ",Shipping_Address"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Shipping_Address=NULL"
    end if
    
    ' Shipping Address 2
    if not isblank(request("Shipping_Address_2")) then 
      if user_region = 1 or user_region = 3 then
        sqlf = sqlf & ",Shipping_Address_2"       
        sql  = sql  & ",N'" & ProperCase(replacequote(request("Shipping_Address_2"))) & "'"
        sqlu = sqlu & ",Shipping_Address_2=" & "N'" & ProperCase(replacequote(request("Shipping_Address_2"))) & "'"
        MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 2: " & ProperCase(request("Shipping_Address_2")) & vbCrLf             
      else
        sqlf = sqlf & ",Shipping_Address_2"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_Address_2")) & "'"
        sqlu = sqlu & ",Shipping_Address_2=" & "N'" & replacequote(request("Shipping_Address_2")) & "'"
        MailMessage = MailMessage & Translate("Address",Admin_Language,conn) & " 2: " & request("Shipping_Address_2") & vbCrLf             
      end if  
    else
      sqlf = sqlf & ",Shipping_Address_2"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Shipping_Address_2=NULL"
    end if
  
    ' Shipping City
    if not isblank(request("Shipping_City")) then 
      if user_region = 1 or user_region = 3 then
        sqlf = sqlf & ",Shipping_City"       
        sql  = sql  & ",N'" & ProperCase(replacequote(request("Shipping_City"))) & "'"
        sqlu = sqlu & ",Shipping_City=" & "N'" & ProperCase(replacequote(request("Shipping_City"))) & "'"
        MailMessage = MailMessage & Translate("City",Admin_Language,conn) & ": " & ProperCase(request("Shipping_City")) & vbCrLf                      
      else
        sqlf = sqlf & ",Shipping_City"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_City")) & "'"
        sqlu = sqlu & ",Shipping_City=" & "N'" & replacequote(request("Shipping_City")) & "'"
        MailMessage = MailMessage & Translate("City",Admin_Language,conn) & ": " & request("Shipping_City") & vbCrLf                      
      end if  
    else
      sqlf = sqlf & ",Shipping_City"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Shipping_City=NULL"
    end if
  
    ' Shipping State
    
    if not isblank(request("Shipping_State")) or not isblank(request("Shipping_State_Other")) then 
      if not isblank(request("Shipping_State")) and request("Shipping_State") <> "ZZ" then
        sqlf = sqlf & ",Shipping_State"
        sql  = sql  & ",N'" & replacequote(request("Shipping_State")) & "'"
        sqlu = sqlu & ",Shipping_State=" & "N'" & replacequote(request("Shipping_State")) & "'"
        if request("Shipping_State") <> "ZZ" then
          MailMessage = MailMessage & Translate("USA State or Canadian Province",Admin_Language,conn) & ": " & request("Shipping_State") & vbCrLf
        end if
      else
        sqlf = sqlf & ",Shipping_State"
        sql  = sql  & ",N'ZZ'"
        sqlu = sqlu & ",Shipping_State=" & "'" & "ZZ" & "'"
      end if
      if not isblank(request("Shipping_State_Other")) then
        if user_region = 1 or user_region = 3 then
          sqlf = sqlf & ",Shipping_State_Other"
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Shipping_State_Other"))) & "'"
          sqlu = sqlu & ",Shipping_State_Other=" & "N'" & ProperCase(replacequote(request("Shipping_State_Other"))) & "'"
          MailMessage = MailMessage & Translate("Other State, Province or County",Admin_Language,conn) & ": " & ProperCase(request("Shipping_State_Other")) & vbCrLf
        else
          sqlf = sqlf & ",Shipping_State_Other"
          sql  = sql  & ",N'" & replacequote(request("Shipping_State_Other")) & "'"
          sqlu = sqlu & ",Shipping_State_Other=" & "N'" & replacequote(request("Shipping_State_Other")) & "'"
          MailMessage = MailMessage & Translate("Other State, Province or County",Admin_Language,conn) & ": " & request("Shipping_State_Other") & vbCrLf
        end if  
      end if              
    end if
  
    ' Shipping Postal Code
    if not isblank(request("Shipping_Postal_Code")) then
      if UCase(request("Shipping_Country")) = "US" then
        if isnumeric(replace(request("Business_Postal_Code"),"-","")) then
          sqlf = sqlf & ",Shipping_Postal_Code"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Shipping_Postal_Code"))) & "'"
          sqlu = sqlu & ",Shipping_Postal_Code=" & "N'" & UCase(replacequote(request("Shipping_Postal_Code"))) & "'"
        end if
      else
        sqlf = sqlf & ",Shipping_Postal_Code"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_Postal_Code")) & "'"
        sqlu = sqlu & ",Shipping_Postal_Code=" & "N'" & replacequote(request("Shipping_Postal_Code")) & "'"
      end if
      MailMessage = MailMessage & Translate("Postal Code",Admin_Language,conn) & ": " & request("Shipping_Postal_Code") & vbCrLf     
    else
      sqlf = sqlf & ",Shipping_Postal_Code"       
      sql  = sql  & ",NULL"
      sqlu = sqlu & ",Shipping_Postal_Code=NULL"
    end if
  
    ' Shipping_Country
    if not isblank(request("Shipping_Country")) then 
      sqlf = sqlf & ",Shipping_Country"       
      sql  = sql  & ",N'" & replacequote(request("Shipping_Country")) & "'"
      sqlu = sqlu & ",Shipping_Country=" & "N'" & replacequote(request("Shipping_Country")) & "'"
      MailMessage = MailMessage & Translate("Country",Admin_Language,conn) & ": " & request("Shipping_Country") & vbCrLf & vbCrLf
    end if
  end if

  MailMessage = MailMessage & "--------------------------------------------" & vbCrLf
  MailMessage = MailMessage & Translate("Contact Information",Admin_Language,conn) & vbCrLf
  MailMessage = MailMessage & "--------------------------------------------" & vbCrLf & vbCrLf      
  
  ' Business Phone
  if not isblank(request("Business_Phone")) then
    if User_Region = 1 or User_Region = 3 then
      sqlf = sqlf & ",Business_Phone"       
      sql  = sql  & ",N'" & FormatPhone(replacequote(request("Business_Phone"))) & "'"
      sqlu = sqlu & ",Business_Phone=" & "N'" & FormatPhone(replacequote(request("Business_Phone"))) & "'"
      MailMessage = MailMessage & Translate("Phone",Admin_Language,conn) & " 1: " & FormatPhone(request("Business_Phone"))
    else  
      sqlf = sqlf & ",Business_Phone"       
      sql  = sql  & ",N'" & replacequote(request("Business_Phone")) & "'"
      sqlu = sqlu & ",Business_Phone=" & "N'" & replacequote(request("Business_Phone")) & "'"
      MailMessage = MailMessage & Translate("Phone",Admin_Language,conn) & " 1: " & request("Business_Phone")
    end if
  else
    if business then
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Business Phone",Login_Language,conn) & "</LI>"         
    else
      ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Phone",Login_Language,conn) & "</LI>"             
    end if  
  end if
  
  ' Business Phone Extension
  if not isblank(request("Business_Phone_Extension")) then 
    sqlf = sqlf & ",Business_Phone_Extension"       
    sql  = sql  & ",N'" & replacequote(request("Business_Phone_Extension")) & "'"
    sqlu = sqlu & ",Business_Phone_Extension=" & "N'" & replacequote(request("Business_Phone_Extension")) & "'"
    MailMessage = MailMessage & "  (" & Translate("Extension",Admin_Language,conn) & "): " & request("Business_Phone_Extension") & vbCrLf
  else
    sqlf = sqlf & ",Business_Phone_Extension"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Business_Phone_Extension=NULL"
    MailMessage = MailMessage & vbCrLf
  end if

  ' Business_Phone_2
  if not isblank(request("Business_Phone_2")) then 
    if User_Region = 1 or User_Region = 3 then
      sqlf = sqlf & ",Business_Phone_2"       
      sql  = sql  & ",N'" & FormatPhone(replacequote(request("Business_Phone_2"))) & "'"
      sqlu = sqlu & ",Business_Phone_2=" & "N'" & FormatPhone(replacequote(request("Business_Phone_2"))) & "'"
      MailMessage = MailMessage & Translate("Phone",Admin_Language,conn) & " 2: " & FormatPhone(request("Business_Phone_2"))
    else  
      sqlf = sqlf & ",Business_Phone_2"       
      sql  = sql  & ",N'" & replacequote(request("Business_Phone_2")) & "'"
      sqlu = sqlu & ",Business_Phone_2=" & "N'" & replacequote(request("Business_Phone_2")) & "'"
      MailMessage = MailMessage & Translate("Phone",Admin_Language,conn) & " 2: " & request("Business_Phone_2")
    end if
  else
    sqlf = sqlf & ",Business_Phone_2"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Business_Phone_2=NULL"
  end if

  ' Business Phone Extension
  if not isblank(request("Business_Phone_2_Extension")) then 
    sqlf = sqlf & ",Business_Phone_2_Extension"       
    sql  = sql  & ",N'" & replacequote(request("Business_Phone_2_Extension")) & "'"
    sqlu = sqlu & ",Business_Phone_2_Extension=" & "N'" & replacequote(request("Business_Phone_2_Extension")) & "'"
    MailMessage = MailMessage & "  (" & Translate("Extension",Admin_Language,conn) & "): " & request("Business_Phone2_Extension") & vbCrLf
  else
    sqlf = sqlf & ",Business_Phone_2_Extension"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Business_Phone_2_Extension=NULL"
    MailMessage = MailMessage & vbCrLf
  end if  
  
  ' Business_Fax
  if not isblank(request("Business_Fax")) then
    if User_Region = 1 or User_Region = 3 then
      sqlf = sqlf & ",Business_Fax"       
      sql  = sql  & ",N'" & FormatPhone(replacequote(request("Business_Fax"))) & "'"
      sqlu = sqlu & ",Business_Fax=" & "N'" & FormatPhone(replacequote(request("Business_Fax"))) & "'"
      MailMessage = MailMessage & Translate("Fax",Admin_Language,conn) & ": " & FormatPhone(request("Business_Fax")) & vbCrLf
    else  
      sqlf = sqlf & ",Business_Fax"       
      sql  = sql  & ",N'" & replacequote(request("Business_Fax")) & "'"
      sqlu = sqlu & ",Business_Fax=" & "N'" & replacequote(request("Business_Fax")) & "'"
      MailMessage = MailMessage & Translate("Fax",Admin_Language,conn) & ": " & request("Business_Fax") & vbCrLf
    end if
  else
    sqlf = sqlf & ",Business_Fax"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Business_Fax=NULL"
  end if
  
  ' Mobile_Phone
  if not isblank(request("Mobile_Phone")) then 
    if User_Region = 1 or User_Region = 3 then
      sqlf = sqlf & ",Mobile_Phone"       
      sql  = sql  & ",N'" & FormatPhone(replacequote(request("Mobile_Phone"))) & "'"
      sqlu = sqlu & ",Mobile_Phone=" & "N'" & FormatPhone(replacequote(request("Mobile_Phone"))) & "'"
      MailMessage = MailMessage & Translate("Mobile Phone",Admin_Language,conn) & ": " & FormatPhone(request("Mobile_Phone")) & vbCrLf
    else
      sqlf = sqlf & ",Mobile_Phone"       
      sql  = sql  & ",N'" & replacequote(request("Mobile_Phone")) & "'"
      sqlu = sqlu & ",Mobile_Phone=" & "N'" & replacequote(request("Mobile_Phone")) & "'"
      MailMessage = MailMessage & Translate("Mobile Phone",Admin_Language,conn) & ": " & request("Mobile_Phone") & vbCrLf     
    end if  
  else
    sqlf = sqlf & ",Mobile_Phone"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Mobile_Phone=NULL"
  end if
  
  ' Pager
  if not isblank(request("Pager")) then 
    if User_Region = 1 or User_Region = 3 then
      sqlf = sqlf & ",Pager"       
      sql  = sql  & ",N'" & FormatPhone(replacequote(request("Pager"))) & "'"
      sqlu = sqlu & ",Pager=" & "N'" & FormatPhone(replacequote(request("Pager"))) & "'"
      MailMessage = MailMessage & Translate("Pager",Admin_Language,conn) & ": " & FormatPhone(request("Pager")) & vbCrLf
    else
      sqlf = sqlf & ",Pager"       
      sql  = sql  & ",N'" & replacequote(request("Pager")) & "'"
      sqlu = sqlu & ",Pager=" & "N'" & replacequote(request("Pager")) & "'"
      MailMessage = MailMessage & Translate("Pager",Admin_Language,conn) & ": " & request("Pager") & vbCrLf
    end if  
  else
    sqlf = sqlf & ",Pager"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Pager=NULL"
  end if

  ' Email
  if not isblank(request("Email")) then 
    sqlf = sqlf & ",Email"       
    sql  = sql  & ",N'" & replacequote(request("Email")) & "'"
    sqlu = sqlu & ",Email=" & "N'" & replacequote(request("Email")) & "'"
    MailMessage = MailMessage & Translate("EMail",Admin_Language,conn) & " 1: " & request("EMail") & vbCrLf
  else
    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Email (Primary)",Login_Language,conn) & "</LI>"         
  end if
 
  ' Email_2
  if not isblank(request("Email_2")) then 
    sqlf = sqlf & ",Email_2"       
    sql  = sql  & ",N'" & replacequote(request("Email_2")) & "'"
    sqlu = sqlu & ",Email_2=" & "N'" & replacequote(request("Email_2")) & "'"
    MailMessage = MailMessage & Translate("EMail",Admin_Language,conn) & " 2: " & request("EMail_2") & vbCrLf
  else
    sqlf = sqlf & ",Email_2"       
    sql  = sql  & ",NULL"
    sqlu = sqlu & ",Email_2=NULL"
  end if

  MailMessage = MailMessage & vbCrLf & "--------------------------------------------" & vbCrLf
  MailMessage = MailMessage & Translate("Other Information",Admin_Language,conn) & vbCrLf
  MailMessage = MailMessage & "--------------------------------------------" & vbCrLf & vbCrLf  
   
  ' Language
  if not isblank(request("Language")) then 
    sqlf = sqlf & ",Language"       
    sql  = sql  & ",N'" & replacequote(request("Language")) & "'"
    sqlu = sqlu & ",Language=" & "N'" & replacequote(request("Language")) & "'"
    
    LangSQL = "SELECT Language.* FROM Language WHERE Language.Code='" & request("Language") & "'"
    Set rsLanguage = Server.CreateObject("ADODB.Recordset")
    rsLanguage.Open LangSQL, conn, 3, 3
                          
    if not rsLanguage.EOF then
      MailMessage = MailMessage & Translate("Preferred Language",Admin_Language,conn) & ": " & rsLanguage("Description") & vbCrLf
    else
      MailMessage = MailMessage & Translate("Preferred Language",Admin_Language,conn) & ": " & request("Language") & vbCrLf
    end if              
    rsLanguage.close
    set rsLanguage=nothing
  end if
  
    ' Auxiliary Fields
    
  AuxiliarySQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.Site_ID=" & request("Site_ID") & " ORDER BY Auxiliary.Order_Num"
  Set rsAuxiliary = Server.CreateObject("ADODB.Recordset")
  rsAuxiliary.Open AuxiliarySQL, conn, 3, 3

  if not rsAuxiliary.EOF then 
	for aux = 0 to 9 
  
	  if not isblank(request("Aux_" & Trim(aux))) then 
	    sqlf = sqlf & ",Aux_" & Trim(aux)       
	    sql  = sql  & ",N'" & replacequote(request("Aux_" & Trim(aux))) & "'"
      sqlu = sqlu & ",Aux_" & Trim(aux) & "=" & "N'" & replacequote(request("Aux_" & Trim(aux))) & "'"

	    MailMessage = MailMessage & Translate(rsAuxiliary("Description"),Admin_Language,conn) & ": " & request("Aux_" & Trim(aux)) & vbCrLf
	 
	 elseif isblank(request("Aux_" & Trim(aux))) and rsAuxiliary("Required") then
	    ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing answer to the following question",Login_Language,conn) & " :<BR><FONT COLOR=""Black"">" & Translate(rsAuxiliary("Description"),Login_Language,conn) & "</FONT></LI>"                   
	 else
	    sqlf = sqlf & ",Aux_" & Trim(aux)
	    sql  = sql  & ",NULL"
	    sqlu = sqlu & ",Aux_" & Trim(aux) & "=NULL"
	  end if

		rsAuxiliary.MoveNext    
	next
  end if 
  
  rsAuxiliary.close
  set rsAuxiliary = nothing
     
  MailMessage = MailMessage & vbCrLf & "--------------------------------------------" & vbCrLf & vbCrLf        
  MailMessage = MailMessage & vbCrLf & Translate("Sincerely",Admin_Language,conn) & "," & vbCrLf & vbCrLf & Translate(MailFromName,Admin_Language,conn) & " " & Translate("Support Team",Admin_Language,conn)

  sqlf = sqlf & ",NewFlag"
  sql  = sql  & "," & CInt(True)
  sqlu = sqlu & ",NewFlag=" & "" & CInt(True) & ""

  sqlf = sqlf & ",Reg_Request_Date"
  sql  = sql  & ",'" & Now & "'"
  sqlu = sqlu & ",Reg_Request_Date=" & "'" & Now & "'"
  
  ' Build Ending SQL Statement
  
  Session("Reg_Fields") = Replace(sqlf,"'","")          ' Field Names - Used By Thank-You.asp to forward on Promotional Fulfillment
  Session("Reg_Values") = Replace(sql,  "'","")         ' Field Data  - Used By Thank-You.asp to forward on Promotional Fulfillment
  
  sqlf = "INSERT INTO UserData (" & sqlf & ") "
  sql  = "VALUES (" & sql & ")"
  sql  = sqlf & sql

' --------------------------------------------------------------------------------------
' Add Record / EMail Add / Redirect
' --------------------------------------------------------------------------------------             

  if isblank(ErrorMessage) then

    ' Create New Account using Add/Update Method to get New_Account_ID
    New_Account_ID = Get_New_Record_ID ("UserData", "NewFlag", 0, conn)
    Action = "Register"
    strPost_QueryString = sqlu
    sqlu = "UPDATE UserData SET " & sqlu & " WHERE UserData.ID=" & New_Account_ID
   
    if DebugFlag = True then
      response.write Replace(SQLU,",",",<BR>")
      response.flush
      response.end
    else  
      conn.execute (SQLU)
    end if  
    
    ' Determine if Posting is required to a Regional (CM) Contact Management System,
    ' if false notify admin default email method.

    select case CInt(CMS_Region(User_Region))
      case CInt(True)       ' CM Sytem Only
        Call Send_2_CMS     ' Send Data to CM Sytem
      case -2               ' CM System and Email
        Call Send_2_CMS     ' Send Data to CM Sytem        
        Call Send_EMail     ' Send Email Notification to Account Administrator
      case else             ' Default
        Call Send_EMail
    end select

  end if
  
  if not isblank(ErrorMessage) then
  
    Call ErrorHandler
    
  else

    with response
      ' Need to hide Site ID so use the onLoad Method with POST
      .write "<HTML>" & vbCrLf
      .write "<HEAD>" & vbCrLf
      .write "<TITLE>Thank You</TITLE>" & vbCrLf
      .write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf
      .write "</HEAD>" & vbCrLf
      .write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
      .write "<FORM NAME=""FORM1"" ACTION=""Thank-You.asp"" METHOD=""POST"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Region"" VALUE=""" & User_Region & """>" & vbCrLf      
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
      .write "</FORM>" & vbCrLf
      .write "</BODY>" & vbCrLf
      .write "</HTML>" & vbCrLf
    end with

  end if
  
end if

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/Include/Function_DCMDataTransfer.asp"-->
<%

' --------------------------------------------------------------------------------------

sub ErrorHandler

  Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Registration Request",Alt_Language,conn) & " - " & Translate("Error Report Screen",Alt_Language,conn)
  Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=NormalBoldGold>" & Translate("Registration Request",Alt_Language,conn) & "<BR>" & Translate("Error Report Screen",Login_Language,conn) & "</FONT>"
  Navigation      = false
  Side_Navigation = false
  Content_Width   = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%

  response.write "<DIV ALIGN=CENTER>"
  response.write "<TABLE WIDTH=""" & Content_Width & "%"">"
  response.write "<TR>"
  response.write "<TD WIDTH=""100%"" CLASS=Medium>"

  response.write "<B>" & Translate("Use the <FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;[Back]&nbsp;</FONT> button of your browser to return to the previous screen to correct the error(s) listed below",Login_Language,conn) & ".</B><BR><BR>"
  response.write "<FONT COLOR=""Red""><UL>" & ErrorMessage & "</UL></FONT>"
  response.write Translate("Please correct any errors above.",Login_Language,conn) & "&nbsp;" & Translate("Send any errors noted to be reported to the Webmaster",Login_Language,conn) & ", "
  response.write Translate("or if you have questions regarding this site, to",Login_Language,conn) & ": "
  'Modified kelly's mail id. 23-10-2007
  response.write "<A HREF=""mailto:extranetalerts@fluke.com"">Extranet Admin Group</A>, " & Translate("Webmaster",Login_Language,conn) & ".<BR>"
  '>>>>>>>>>>>>>>>>  
  response.write "</TD>"
  response.write "</TR>"
  response.write "</TABLE>"
  response.write "</DIV>"
    
  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%
  
end sub    

' --------------------------------------------------------------------------------------
sub Send_EMail

  Mailer.QMessage = False
  Mailer.Subject  = MailSubject
  Mailer.BodyText = MailMessage

  if Mailer.SendMail then
    ' Success 
    if Mailer.Response <> "" then
      strError = Mailer.Response
      Mailer.ClearAllRecipients
      Mailer.FromName    = "Support Postmaster"
      'Modified kelly's mail id. 23-10-2007
      Mailer.FromAddress = "extranetalerts@fluke.com"
      Mailer.AddRecipient  "Extranet Admin Group", "extranetalerts@fluke.com"
      Mailer.Subject     = "Send Email Failure"
      Mailer.BodyText    = strError
      Mailer.SendMail
    end if
  ' Success
  else
    if Mailer.Response <> "" then
      strError = Mailer.Response
      Mailer.ClearAllRecipients
      Mailer.FromName    = "Support Postmaster"
      'Modified kelly's mail id. 23-10-2007
      Mailer.FromAddress = "extranetalerts@fluke.com"
      Mailer.AddRecipient  "Extranet Admin Group", "extranetalerts@fluke.com"
      Mailer.Subject     = "Send Email Failure"
      Mailer.BodyText    = strError
      Mailer.SendMail
    end if
    ErrorMessage = ErrorMessage & vbCrLf & "<LI>" & Translate("Send email failure",Login_Language,conn) & ".<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Mailer.Response & ". " & Translate("Send any errors noted to be reported to the Webmaster",Login_Language,conn) & ". </LI>"   
  end if   

end sub

' --------------------------------------------------------------------------------------

sub Get_UserData

  SQL = "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND UserData.EMail='" & User_Email & "'"
  Set rsUser = Server.CreateObject("ADODB.Recordset")
  rsUser.Open SQL, conn, 3, 3

  User_Name       = ""
  User_Login      = Translate("The information you requested was not found for this site.",Alt_Language,conn)
  User_Password   = Translate("The information you requested was not found for this site.",Alt_Language,conn)
  
  if not rsUser.EOF then
    User_Name     = rsUser("FirstName") & " " & rsUser("LastName")
    User_Login    = rsUser("NTLogin")
    User_Password = rsUser("Password")
  end if
  
  rsUser.close
  Set rsUser = nothing  

end sub
  
%>

