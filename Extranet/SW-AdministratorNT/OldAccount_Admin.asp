<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   01/26/2001
' --------------------------------------------------------------------------------------
' 01/26/2001 - Major Revision - Add Translations
' 07/04/2001 - Removed NT Account Creation
' 07/04/2001 - Added Account Recprocity
' 06/19/2002 - Added Fields and Re-Ordered Form to work with Euro DCM
' 11/28/2008 - Modified for Field Marketing Site.
' --------------------------------------------------------------------------------------

Dim Script_Debug
Dim Send_2_CMS_Debug

Script_Debug     = false

Send_2_CMS_Debug = false
if LCase(request("DCM_Debug")) = "true" then
  Send_2_CMS_Debug = true
elseif LCase(request.form("DCM_Debug")) = "true" then
  Send_2_CMS_Debug = true
end if

if Script_Debug then
  response.write "<HTML><BODY><TABLE Border=1>"
  for each item in request.querystring
    response.write "<TR><TD>" & item & "</TD><TD>" & request.querystring(item) & "</TD></TR>"
  next
  response.write "</TABLE></BODY></HTML>"
  response.flush
  response.end
end if    

if Script_Debug = false then

  Dim Site_ID
  
  %>
  <!--#include virtual="/include/functions_string.asp"-->
  <!--#include virtual="/include/functions_file.asp"-->
  <!--#include virtual="/include/functions_date_formatting.asp"-->
  <!--#include virtual="/include/functions_translate.asp"-->
  <!--#include virtual="/include/functions_DB.asp"-->  
  <!--#include virtual="/SW-Common/Preferred_Language.asp"-->
  <!--#include virtual="/connections/connection_SiteWide.asp"-->
  <!--#include virtual="/include/DCMHTTP_DataTransfer.asp"-->
  <%
  
  Call Connect_SiteWide
  
  %>
  <!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
  <%
  
  ' --------------------------------------------------------------------------------------
  ' Declarations
  ' --------------------------------------------------------------------------------------
  
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
  Dim MailMessageUser
  
  Dim Account_ID
  Dim New_Account_ID
  Dim Action
  
  Dim Groups
  Dim SubGroups
  Dim User_Region
  Dim NewFlag
  Dim Admin_Access
  Dim ErrorMessage

  Account_ID    = request("ID")

  Dim BackURL
  Dim HomeURL
  
  Dim sqlf, sql, sqlu, sqlfR, sqlR, sqluR, strPost_QueryString
  
  BackURL       = request("BackURL")
  HomeURL       = request("HomeURL")
  
  Dim No_Post_Back
  if request("No_Post_Back") = "-1" then
    No_Post_Back = True
  else
    No_Post_Back = False
  end if
    
  ' --------------------------------------------------------------------------------------
  ' Verify Admin Credentials
  ' --------------------------------------------------------------------------------------
  
  select case Admin_Access
    case 6, 8, 9
      'Admin_Access = true
    case else
      ErrorMessage = ErrorMessage & "<LI>" & Translate("You do not have administration previlages to perform this action.",Login_Language,conn) & "</LI>"          
  end select
  
  ' --------------------------------------------------------------------------------------
  ' Get Site Information
  ' --------------------------------------------------------------------------------------
  
  SQL_Site = "SELECT * FROM Site Where Site.ID=" & Site_ID
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL_Site, conn, 3, 3
  
  Groups = rsSite("Site_Code")
  Site_Description = rsSite("Site_Description")
  
  Login_Method      = UCase(rsSite("Login_Method"))
  
  MailFromName      = rsSite("FromName")
  MailFromAddress   = rsSite("FromAddress")
  MailReplyToName   = rsSite("ReplyToName")
  MailReplyTo       = rsSite("ReplyTo")
  MailCCName        = rsSite("MailCCName")
  MailCC            = rsSite("MailCC")
  MailBCCName       = rsSite("MailBCCName")
  MailBCC           = rsSite("MailBCC")
  
  if request("Send_Email_Domain") = "on" or request("Send_Email_Domain") = "-1" then Send_Email_Domain = true
  Send_Email_Domain = true
  if request("Send_Email_Admin")  = "on" or request("Send_Email_Admin")  = "-1" then Send_Email_Admin  = true
  if request("Send_Email_Fcm")    = "on" or request("Send_Email_Fcm")    = "-1" then Send_Email_Fcm    = true
  if request("Send_Email_User")   = "on" or request("Send_Email_User")   = "-1" then Send_Email_User   = true

  Dim CMS_Region(3)

  CMS_Region(1) = rsSite("CMS_Region_1")
  CMS_Region(2) = rsSite("CMS_Region_2")
  CMS_Region(3) = rsSite("CMS_Region_3")

  rsSite.close
  set rsSite = nothing

  ' --------------------------------------------------------------------------------------
  ' Open Connection to Mail Server
  ' --------------------------------------------------------------------------------------
  
  set Mailer = Server.CreateObject("SMTPsvg.Mailer") 
  
  ' --------------------------------------------------------------------------------------
  ' Configure EMail Header Information
  ' --------------------------------------------------------------------------------------
  
  Mailer.ClearAllRecipients
  
  %>
  <!--#include virtual="/connections/connection_EMail.asp"-->
  <!--#include virtual="/connections/connection_EMail_Timeout.asp"-->  
  <%
  
  Mailer.ReturnReceipt  = False
  Mailer.Priority       = 3
  
  Mailer.FromName       = MailFromName
  Mailer.FromAddress    = MailFromAddress
  Mailer.ReplyTo        = MailReplyTo
  
  If Send_Email_Admin = true then
    Mailer.AddRecipient MailReplyToName, MailReplyTo      ' Sub-Site Administrator
    if not isblank(Admin_EMail) then
      Mailer.AddCC         Admin_Name, Admin_Email        ' Sub-Site Administrator  
    end if  
    if not isblank(MailCC) then
      Mailer.AddCC         MailCCName, MailCC             ' Sub-Site Administrator
    end if
  end if  
  
  if Send_Email_Domain = true and not isblank(MailBCC) then
    Mailer.AddBCC MailBCCName, MailBCC                    ' Domain Administrator
  end if  
  
  MailMessage     = "This is an automated notification message from the" & vbCrLf
  MailMessage     = MailMessage & MailFromName & " Extranet Server." & vbCrLf & vbCrLf
  
  MailMessageUser = MailMessage
    
  ' --------------------------------------------------------------------------------------
  ' Verify NT Account
  ' --------------------------------------------------------------------------------------

  if len(request("Verify")) > 0 then
    
    SQL_Verify = "Select ID, NTLogin, Site_ID, NewFlag, Region FROM UserData WHERE UserData.NTLogin='" & request("NTLogin") & "' ORDER BY Site_ID"
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL_Verify, conn, 3, 3
    
    ' Determine if Posting is required to a Regional (CM) Contact Management System,
    ' if false just show as debug text.
   
    if not rsUser.EOF then

      Action = "Verify"
      User_Region = Admin_Region
      New_Account_ID = ""
      strPost_QueryString = "NTLogin='" & rsUser("NTLogin") & "',Status=" & CInt(True) & ",Site_ID="
    
      ' Build Array of Site_ID's
      rsUser.MoveFirst
      Delimeter_Flag = False
      do while not rsUser.EOF
        if Delimeter_Flag = True then
          strPost_QueryString = strPost_QueryString & "|"
        end if  
        strPost_QueryString = strPost_QueryString & rsUser("Site_ID")
        Delimeter_Flag = True
        rsUser.MoveNext
      loop

      ' Build Array of NewFlag Status
      rsUser.MoveFirst
      Delimeter_Flag = False
      strPost_QueryString = strPost_QueryString & ",NewFlag="
      do while not rsUser.EOF
        if Delimeter_Flag = True then
          strPost_QueryString = strPost_QueryString & "|"
        end if  
        strPost_QueryString = strPost_QueryString & rsUser("NewFlag")
        Delimeter_Flag = True
        rsUser.MoveNext
      loop

      ' Build Array of Account_ID's
      rsUser.MoveFirst
      Delimeter_Flag = False
      do while not rsUser.EOF
        if Delimeter_Flag = True then
          New_Account_ID = New_Account_ID & "|"
        end if  
        New_Account_ID = New_Account_ID & rsUser("ID")
        Delimeter_Flag = True
        rsUser.MoveNext
      loop
    else
      Action = "Verify"
      User_Region = Admin_Region
      New_Account_ID = "0"
      strPost_QueryString = "NTLogin='" & request("NTLogin") & "',Site_ID=" & request("Site_ID") & ",Status=" & CInt(False)
    end if  

    rsUser.close
    set rsUser=nothing

    if CInt(CMS_Region(Admin_Region)) = CInt(True) and CInt(No_Post_Back) = CInt(False) then
        Call Send_2_CMS     ' Send Data to CM Sytem
        Call ErrorHandler   ' Display Transfer Data if Send_2_CMS_Debug = True
    end if
        
    if not isblank(BackURL) then
      response.redirect BackURL
    else
      response.end  
    end if   
  
  end if
  
  ' --------------------------------------------------------------------------------------
  ' Retrieve NTAccount UserData
  ' --------------------------------------------------------------------------------------

  if len(request("Retrieve")) > 0 then

    if isnumeric(request("ID")) then
      SQL = "SELECT * FROM UserData WHERE ID=" & request("ID")
      Set rsUser = Server.CreateObject("ADODB.Recordset")
      rsUser.Open SQL, conn, 3, 3

      if not rsUser.EOF then
        if instr(1,LCase(rsUser("Subgroups")),"domain") = 0 and _
           instr(1,LCase(rsUser("Subgroups")),"administrator") = 0 and _
           instr(1,LCase(rsUser("Subgroups")),"system") = 0 and _
           instr(1,LCase(rsUser("Subgroups")),"account") = 0 then
          
          Action = "Retrieve"
          
          For field_num = 0 To rsUser.Fields.Count - 1
            strPost_QueryString = strPost_QueryString & "," & rsUser.Fields(field_num).Name & "="
            if isblank(rsUser(rsUser.Fields(field_num).Name)) then
              strPost_QueryString = strPost_QueryString & "NULL"
            else
              strPost_QueryString = strPost_QueryString & "'" & rsUser(rsUser.Fields(field_num).Name) & "'"
            end if  
          Next
          strPost_QueryString = Mid(strPost_QueryString,2) ' Trim off leading comma
        else
          ErrorMessage = "<LI>" & Translate("You do not have administration previlages to view this account.",Login_Language,conn) & "</LI>"
        end if  
      end if

      rsUser.close
      set rsUser = nothing
      
    end if
    
    User_Region = Admin_Region
    New_Account_ID = request("ID")

    if not isblank(ErrorMessage) then  
      Call ErrorHandler
    else
      if CInt(CMS_Region(Admin_Region)) = CInt(True) and CInt(No_Post_Back) = CInt(False) then
          Call Send_2_CMS     ' Send Data to CM Sytem
          Call ErrorHandler   ' Display Transfer Data if Send_2_CMS_Debug = True
      end if
    end if  
        
    if not isblank(BackURL) then
      response.redirect BackURL
    else
      response.end  
    end if   

  end if

  ' --------------------------------------------------------------------------------------
  ' Delete NT Account
  ' --------------------------------------------------------------------------------------

  if Len(request("Delete")) > 0 then
    
    SQL_Delete = "Select UserData.* FROM UserData WHERE UserData.ID=" & Account_ID & " AND UserData.Site_ID=" & Site_ID
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL_Delete, conn, 3, 3
     
    if not rsUser.EOF then
    
      if isblank(rsUser("SubGroups")) _
      or instr(1,lcase(rsUser("SubGroups")), "domain") = 0 _
      or instr(1,lcase(rsUser("SubGroups")), "administrator") = 0 _
      or instr(1,lcase(rsUser("SubGroups")), "system") = 0 _
      and (Admin_Account_Region = 4 or Admin_Account_Region = rsUser("Region")) then   ' Check to see that User is not an Admin Account
  
        MailSubject = "Deleted Account Profile Notice " & User_Region
  
        MailMessage = MailMessage & "The following Extranet Account has been deleted:" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Name: " & rsUser("FirstName") &  " " & rsUser("LastName") & vbCrLf
        MailMessage = MailMessage & "Company Name: " & rsUser("Company") & vbCrLf
        MailMessage = MailMessage & "Email Address: " & rsUser("EMail") & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Account ID Number: " & rsUser("ID") & vbCrLf
        MailMessage = MailMessage & "Account User Login Name: " & rsUser("NTLogin") & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Primary Site: " & UCase(rsUser("Groups")) & vbCrLf
        if not isblank(rsUser("Groups_Aux")) then
          MailMessage = MailMessage & "Reciprocal Site(s): " & UCase(rsUser("Groups_Aux")) & vbCrLf
        end if  
        MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Sincerely," & vbCrLf & vbCrLf & "The " & MailFromName & " Support Team"
        
        NewFlag = rsUser("NewFlag")
        
        ' Delete Account from SITEWIDE DB      
  
        if not isblank(rsUser("Groups_Aux")) then
  
          SQL_Site_Aux =                "SELECT Site_Aux.Site_ID, Site_Aux.Site_ID_Aux, Site.Site_Code, Site.Enabled, Site.Site_Description "
          SQL_Site_Aux = SQL_Site_Aux & "FROM Site_Aux LEFT JOIN Site ON Site_Aux.Site_ID_Aux = Site.ID "
          SQL_Site_Aux = SQL_Site_Aux & "WHERE Site_Aux.Site_ID=" & Site_ID & " ORDER BY Site.Site_Description"
    
          Set rsSite_Aux = Server.CreateObject("ADODB.Recordset")
          rsSite_Aux.Open SQL_Site_Aux, conn, 3, 3
  
          if not rsSite_Aux.EOF then
  
            do while not rsSite_Aux.EOF
  
              SQL_Login = "SELECT Site_ID, NTLogin FROM UserData WHERE Site_ID=" & CInt(rsSite_Aux("Site_ID_Aux")) & " AND NTLogin='" & request("NTLogin") & "'"
              Set rsNTLogin = Server.CreateObject("ADODB.Recordset")
              rsNTLogin.Open SQL_Login, conn, 3, 3
  
              if not rsNTLogin.EOF then
        
                ' Delete Accounts where Receprocity has been revoked
  
                SQL_Delete = "DELETE FROM UserData WHERE Site_ID=" & CInt(rsSite_Aux("Site_ID_Aux")) & " AND NTLogin='" & request("NTLogin") & "'"
                conn.execute (SQL_Delete)           
  
              end if
        
              rsNTLogin.close
              set rsNTLogin = nothing
              
              rsSite_Aux.MoveNext
  
            loop
          
          end if
        
          rsSite_Aux.Close
          set rsSite_Aux = nothing
          
        end if  
        
        ' Delete Master Account      
                
        SQL_Delete = "DELETE FROM UserData WHERE UserData.ID=" & Account_ID  
  
        conn.execute (SQL_Delete)
        '----------------------------------------------------------------------------------------
        if site_id = 3 then
													set cmd = Server.CreateObject("ADODB.Command")
													cmd.ActiveConnection = conn
													cmd.CommandType = adCmdStoredProc
													cmd.CommandText = "FMDeleteUser"
													cmd.Parameters.Append cmd.CreateParameter("@AccountID",adBigInt,adParamInput,Account_ID)
													cmd.Parameters(0).Value= Account_ID
													cmd.execute
													set cmd = nothing        
								end if
        '----------------------------------------------------------------------------------------
        
        User_Region = Admin_Region
        Action = "Delete"
        New_Account_ID = Account_ID
        strPost_QueryString = ""
        
        if NewFlag <> True and (Send_Email_Domain = True or Send_Email_Admin = True or Send_Email_Fcm = True) then
  
          ' Get Channel Manager's Info
                  
          if Send_Email_Fcm = true then
            SQL_FCM = "Select * FROM Manager  WHERE Manager.ID=" & rsUser("Fcm_ID")
            Set rsManager = Server.CreateObject("ADODB.Recordset")
            rsManager.Open SQL_FCM, conn, 3, 3
          
            if not rsManager.EOF then
              ManagerName = rsManager("FirstName") & " " & rsManager("LastName")
              Mailer.AddCC ManagerName, rsManager("EMail")
            end if
          
            rsManager.Close
            Set rsManager = nothing
          end if          
        
          Call Send_EMail      
  
        end if    
  
      else
    
        ErrorMessage = ErrorMessage & "<LI>" & Translate("You do not have administrative previlages to delete this account.",Login_Language,conn) & "</LI>"
  
      end if
        
    end if
    
    rsUser.close
    set rsUser=nothing
      
    if not isblank(ErrorMessage) then
      Call ErrorHandler
    else
      if CInt(CMS_Region(User_Region)) = CInt(True) and CInt(No_Post_Back) = CInt(False) then
          Call Send_2_CMS     ' Send Data to CM Sytem
          Call ErrorHandler   ' Display Transfer Data if Send_2_CMS_Debug = True
      end if
        
      if not isblank(BackURL) then
        response.redirect BackURL
      else
        response.end  
      end if   
    end if    
  
  ' --------------------------------------------------------------------------------------
  ' Add / Update Record
  ' --------------------------------------------------------------------------------------
  
  elseif len(request("Update")) > 0 then
  
    if lcase(Account_ID) = "add" then

      MailSubject = "New Account Profile Notice " & User_Region
      
      MailMessage = MailMessage & "The following Extranet Account has been created:" & vbCrLf & vbCrLf
  
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Authorized Extranet Site(s)" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf  
      MailMessage = MailMessage & "Primary Site URL:" & vbCrLf
      MailMessage = MailMessage & "http://" & Request("SERVER_NAME") & "/" & lcase(groups) & vbCrLf & vbCrLf
  
      if instr(1,request("SubGroups"),"domain") > 0 or _
         instr(1,request("SubGroups"),"administrator") > 0 or _
         instr(1,request("SubGroups"),"account") > 0 or _
         instr(1,request("SubGroups"),"content") > 0 or _
         instr(1,request("SubGroups"),"branch") > 0 or _
         instr(1,request("SubGroups"),"submitter") > 0 or _
         instr(1,request("SubGroups"),"ordad") > 0 or _
         instr(1,request("SubGroups"),"order") > 0 or _
         instr(1,request("SubGroups"),"literature") > 0 or _         
         instr(1,request("SubGroups"),"forum") > 0 then
  
        MailMessage   = MailMessage & "This account has Administrative previlages for the following site URL:" & vbcrLf
        MailMessage   = MailMessage & "http://" & Request("SERVER_NAME") & "/SW-Administrator" & vbCrLf & vbCrLf
        MailMessage   = MailMessage & "The following authorizations will allow this account to perform the following Administrative function(s):" & vbCrLf & vbCrLf
  
        if instr(1,request("SubGroups"),"domain") > 0 then
          MailMessage = MailMessage & "+ Domain Administration"
        elseif instr(1,request("SubGroups"),"administrator") > 0 then
          MailMessage = MailMessage & "+ Site Administration"
        elseif instr(1,request("SubGroups"),"account") > 0 then
          MailMessage = MailMessage & "+ User Account Administration"
        elseif instr(1,request("SubGroups"),"content") > 0 then
          MailMessage = MailMessage & "+ Content / Event Administration"
        elseif instr(1,request("SubGroups"),"submitter") > 0 then
          MailMessage = MailMessage & "+ Content / Event Submitter"
        elseif instr(1,request("SubGroups"),"literature") > 0 then
          MailMessage = MailMessage & "+ Literature Order Administrator"        
        elseif instr(1,request("SubGroups"),"branch") > 0 then
          MailMessage = MailMessage & "+ Branch Location Administrator"        
        end if
  
        if instr(1,request("SubGroups"),"forum") > 0 or _
           instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"content") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Forum or Discussion Group Moderator"
        end if

        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"ordad") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry Super User"
         end if  
  
        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"order") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry with Search"
         end if  

      end if
    
      ' --------------------------------------------------------------------------------------
      ' User Message
      ' --------------------------------------------------------------------------------------      
  
      MailMessageUser = MailMessageUser & "The following Extranet Account has been created:" & vbCrLf & vbCrLf
      
      MailMessageUser = MailMessageUser & "----------------------------------------------------------" & vbCrLf
      MailMessageUser = MailMessageUser & "Authorized Extranet Site(s)" & vbCrLf
      MailMessageUser = MailMessageUser & "----------------------------------------------------------" & vbCrLf & vbCrLf  
  
      MailMessageUser = MailMessageUser & "Primary Site URL:" & vbCrLf
      MailMessageUser = MailMessageUser & "http://" & Request("SERVER_NAME") & "/" & lcase(groups) & vbCrLf & vbCrLf
  
      if instr(1,request("SubGroups"),"domain") > 0 or _
         instr(1,request("SubGroups"),"administrator") > 0 or _
         instr(1,request("SubGroups"),"account") > 0 or _
         instr(1,request("SubGroups"),"content") > 0 or _
         instr(1,request("SubGroups"),"branch") > 0 or _
         instr(1,request("SubGroups"),"submitter") > 0 or _
         instr(1,request("SubGroups"),"ordad") > 0 or _
         instr(1,request("SubGroups"),"order") > 0 or _
         instr(1,request("SubGroups"),"literature") > 0 or _         
         instr(1,request("SubGroups"),"forum") > 0 then
  
        MailMessageUser   = MailMessageUser & "This account has Administrative previlages for the following site URL:" & vbcrLf
        MailMessageUser   = MailMessageUser & "http://" & Request("SERVER_NAME") & "/SW-Administrator" & vbCrLf & vbCrLf
        MailMessageUser   = MailMessageUser & "The following authorizations will allow this account to perform the following Administrative function(s):" & vbCrLf & vbCrLf
        
        if instr(1,request("SubGroups"),"domain") > 0 then
          MailMessageUser = MailMessageUser & "+ Domain Administration"
        elseif instr(1,request("SubGroups"),"administrator") > 0 then
          MailMessageUser = MailMessageUser & "+ Site Administration"
        elseif instr(1,request("SubGroups"),"account") > 0 then
          MailMessageUser = MailMessageUser & "+ Account Administration"
        elseif instr(1,request("SubGroups"),"content") > 0 then
          MailMessageUser = MailMessageUser & "+ Content / Event Administration"
        elseif instr(1,request("SubGroups"),"submitter") > 0 then
          MailMessageUser = MailMessageUser & "+ Content / Event Submitter"
        elseif instr(1,request("SubGroups"),"literature") > 0 then
          MailMessage = MailMessage & "+ Literature Order Administrator"        
        elseif instr(1,request("SubGroups"),"branch") > 0 then
          MailMessageUser = MailMessageUser & "+ Branch Location Administration"
        end if
        
        if instr(1,request("SubGroups"),"forum") > 0 or _
           instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"content") > 0 then
           MailMessageUser = MailMessageUser & vbCrLf & "+ Forum or Discussion Group Moderator"
        end if

        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"ordad") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry Super User"
         end if  
  
        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"order") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry with Search"
         end if  
  
      end if
  
    else
        
      MailSubject = "Updated Account Profile Notice " & User_Region
  
      MailMessage = MailMessage & "The following Extranet Account has been updated:" & vbCrLf & vbCrLf
  
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Authorized Extranet Site(s)" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf  
    
      MailMessage = MailMessage & "Primary Site URL:" & vbCrLf
      MailMessage = MailMessage & "http://" & Request("SERVER_NAME") & "/" & lcase(groups) & vbCrLf & vbCrLf
  
      if instr(1,request("SubGroups"),"domain") > 0 or _
         instr(1,request("SubGroups"),"administrator") > 0 or _
         instr(1,request("SubGroups"),"account") > 0 or _
         instr(1,request("SubGroups"),"content") > 0 or _
         instr(1,request("SubGroups"),"branch") > 0 or _
         instr(1,request("SubGroups"),"ordad") > 0 or _
         instr(1,request("SubGroups"),"order") > 0 or _
         instr(1,request("SubGroups"),"literature") > 0 or _                
         instr(1,request("SubGroups"),"submitter") > 0 then
                
        MailMessage   = MailMessage & "This account has Administrative previlages for the following site URL:" & vbcrLf
        MailMessage   = MailMessage & "http://" & Request("SERVER_NAME") & "/SW-Administrator" & vbCrLf & vbCrLf
        MailMessage   = MailMessage & "The following authorizations will allow this account to perform the following Administrative function(s):" & vbCrLf & vbCrLf
  
        if instr(1,request("SubGroups"),"domain") > 0 then
          MailMessage = MailMessage & "+ Domain Administration"
        elseif instr(1,request("SubGroups"),"administrator") > 0 then
          MailMessage = MailMessage & "+ Site Administration"
        elseif instr(1,request("SubGroups"),"account") > 0 then
          MailMessage = MailMessage & "+ Account Administration"
        elseif instr(1,request("SubGroups"),"content") > 0 then
          MailMessage = MailMessage & "+ Content / Event Administration"
        elseif instr(1,request("SubGroups"),"submitter") > 0 then
          MailMessage = MailMessage & "+ Content / Event Submitter"
        elseif instr(1,request("SubGroups"),"literature") > 0 then
          MailMessage = MailMessage & "+ Literature Order Administrator"        
        elseif instr(1,request("SubGroups"),"branch") > 0 then
          MailMessage = MailMessage & "+ Branch Location Administration"
        end if
  
        if instr(1,request("SubGroups"),"forum") > 0 or _
           instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"content") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Forum or Discussion Group Moderator"
        end if

        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"ordad") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry Super User"
         end if  
  
        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"order") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry with Search"
         end if  
  
      end if
        
      MailMessageUser = MailMessageUser & "Your Account Profile has been updated:" & vbCrLf & vbCrLf               
  
      MailMessageUser = MailMessageUser & "----------------------------------------------------------" & vbCrLf
      MailMessageUser = MailMessageUser & "Authorized Extranet Site(s)" & vbCrLf
      MailMessageUser = MailMessageUser & "----------------------------------------------------------" & vbCrLf & vbCrLf  
  
      MailMessageUser = MailMessageUser & "Primary Site URL:" & vbCrLf
      MailMessageUser = MailMessageUser & "http://" & Request("SERVER_NAME") & "/" & lcase(groups) & vbCrLf & vbCrLf
  
      if instr(1,request("SubGroups"),"domain") > 0 or _
         instr(1,request("SubGroups"),"administrator") > 0 or _
         instr(1,request("SubGroups"),"account") > 0 or _
         instr(1,request("SubGroups"),"content") > 0 or _
         instr(1,request("SubGroups"),"branch") > 0 or _       
         instr(1,request("SubGroups"),"submitter") > 0 or _
         instr(1,request("SubGroups"),"ordad") > 0 or _
         instr(1,request("SubGroups"),"order") > 0 or _
         instr(1,request("SubGroups"),"literature") > 0 or _         
         instr(1,request("SubGroups"),"forum") > 0 then
         
        MailMessageUser   = MailMessageUser & "This account has Administrative previlages for the following site URL:" & vbcrLf
        MailMessageUser   = MailMessageUser & "http://" & Request("SERVER_NAME") & "/SW-Administrator" & vbCrLf & vbCrLf
        MailMessageUser   = MailMessageUser & "The following authorizations will allow this account to perform the following Administrative function(s):" & vbCrLf & vbCrLf
  
        if instr(1,request("SubGroups"),"domain") > 0 then
          MailMessageUser = MailMessageUser & "+ Domain Administration"
        elseif instr(1,request("SubGroups"),"administrator") > 0 then
          MailMessageUser = MailMessageUser & "+ Site Administration"
        elseif instr(1,request("SubGroups"),"account") > 0 then
          MailMessageUser = MailMessageUser & "+ Account Administration"
        elseif instr(1,request("SubGroups"),"content") > 0 then
          MailMessageUser = MailMessageUser & "+ Content / Event Administration"
        elseif instr(1,request("SubGroups"),"branch") > 0 then
          MailMessageUser = MailMessageUser & "+ Branch Location Administration"
        elseif instr(1,request("SubGroups"),"submitter") > 0 then
          MailMessageUser = MailMessageUser & "+ Content / Event Submitter"
        elseif instr(1,request("SubGroups"),"literature") > 0 then
          MailMessage = MailMessage & "+ Literature Order Administrator"        
        end if
  
        if instr(1,request("SubGroups"),"forum") > 0 or _
           instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"content") > 0 then
           MailMessageUser = MailMessageUser & vbCrLf & "+ Forum or Discussion Group Moderator"
        end if
        
        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"ordad") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry Super User"
         end if  
  
        if instr(1,request("SubGroups"),"domain") > 0 or _
           instr(1,request("SubGroups"),"administrator") > 0 or _
           instr(1,request("SubGroups"),"account") > 0 or _
           instr(1,request("Subgroups"),"order") > 0 then
           MailMessage = MailMessage & vbCrLf & "+ Order Inquiry with Search"
         end if  
  
      end if
   
    end if  
  
  ' --------------------------------------------------------------------------------------
  '   USER SPECIFIC INFORMATION
  ' --------------------------------------------------------------------------------------
  
      MailMessage     = MailMessage     & vbCrLf & vbCrLf
      MailMessageUser = MailMessageUser & vbCrlf & vbCrLf           
  
      MailMessage     = MailMessage     & "----------------------------------------------------------" & vbCrLf
      MailMessage     = MailMessage     & "User Information" & vbCrLf
      MailMessage     = MailMessage     & "----------------------------------------------------------" & vbCrLf & vbCrLf  
  
      MailMessageUser = MailMessageUser & "----------------------------------------------------------" & vbCrLf
      MailMessageUser = MailMessageUser & "User Information" & vbCrLf
      MailMessageUser = MailMessageUser & "----------------------------------------------------------" & vbCrLf & vbCrLf  
  
  
      sqlf = "INSERT INTO UserData ("
      sql  = " VALUES ("
      sqlu = ""
         
      ' Fluke Customer Number
      if not isblank(request("Fluke_ID")) then 
        sqlf = sqlf & "Fluke_ID"       
        sql  = sql  & "'" & replacequote(request("Fluke_ID")) & "'"
        sqlu = sqlu & "Fluke_ID=" & "'" & replacequote(request("Fluke_ID")) & "'"
      else  
        sqlf = sqlf & "Fluke_ID"       
        sql  = sql  & "NULL"
        sqlu = sqlu & "Fluke_ID=NULL"
      end if
         
      ' Business System
      if not isblank(request("Business_System")) then 
        sqlf = sqlf & ",Business_System"       
        sql  = sql  & ",'" & replacequote(request("Business_System")) & "'"
        sqlu = sqlu & ",Business_System=" & "'" & replacequote(request("Business_System")) & "'"
      else  
        sqlf = sqlf & ",Business_System"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Business_System=NULL"
      end if
  
      ' Type Code
      if request("Type_Code_Required") = "on" or request("Type_Code_Required") = "-1" then
        if not isblank(request("Type_Code")) and request("Type_Code") <> "0" then
          sqlf = sqlf & ",Type_Code"       
          sql  = sql  & "," & request("Type_Code")
          sqlu = sqlu & ",Type_Code=" & request("Type_Code")
        else
          ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Customer Type",Login_Language,conn) & "</LI>"         
        end if
      else
        sqlf = sqlf & ",Type_Code"       
        sql  = sql  & ",0"
        sqlu = sqlu & ",Type_Code=0"
      end if

      ' Expiration Date
      if isdate(request("ExpirationDate")) then   
        sqlf = sqlf & ",ExpirationDate"                   
        sql  = sql  & ",'" & killquote(request("ExpirationDate")) & "'"
        sqlu = sqlu & ",ExpirationDate=" & "'" & killquote(request("ExpirationDate")) & "'"
      else
        sqlf = sqlf & ",ExpirationDate"                   
        sql  = sql  & ",'9/9/9999'"
        sqlu = sqlu & ",ExpirationDate=" & "'9/9/9999'"
      end if
  
      ' Region    
      RegionSQL = "SELECT Country.Abbrev, Country.Region FROM Country WHERE Country.Abbrev='" & UCase(Request("Business_Country")) & "'"
      Set rsRegion = Server.CreateObject("ADODB.Recordset")
      rsRegion.Open RegionSQL, conn, 3, 3

      if not rsRegion.EOF then                                       

        if (Admin_Account_Region = 4 or (CInt(Admin_Account_Region) = CInt(rsRegion("Region"))) or Admin_Access >= 8) then
          sqlf = sqlf & ",Region"
          sql  = sql  & "," & rsRegion("Region")
          sqlu = sqlu & ",Region=" & rsRegion("Region")
          User_Region = rsRegion("Region")
        else
          ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Administration authority to add/update an Account for this Region.",Login_Language,conn) & " " & Translate("report this error to the Site Administrator.",Login_Language,conn) & " " & Admin_Account_Region & "|" & rsRegion("Region") & "</LI>"
        end if  
      else  
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Country Code or Region Number.",Login_Language,conn) & " " & Translate("report this error to the Site Administrator.",Login_Language,conn) & "</LI>"             
      end if
      
      rsRegion.close
      set rsRegion=nothing

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
        sqlf = sqlf & ",FirstName"       
        if User_Region = 1 or User_Region = 3 then
          sql  = sql  & ",N'" & ProperCase(replacequote(request("FirstName"))) & "'"
          sqlu = sqlu & ",FirstName=" & "N'" & ProperCase(replacequote(request("FirstName"))) & "'"
        else  
          sql  = sql  & ",N'" & replacequote(request("FirstName")) & "'"
          sqlu = sqlu & ",FirstName=" & "N'" & replacequote(request("FirstName")) & "'"
        end if
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing First Name",Login_Language,conn) & "</LI>"
      end if
    
      ' Middle Name
      if not isblank(request("MiddleName")) then 
        sqlf = sqlf & ",MiddleName"       
        if User_Region = 1 or User_Region = 3 then  
          sql  = sql  & ",N'" & ProperCase(replacequote(request("MiddleName"))) & "'"
          sqlu = sqlu & ",MiddleName=" & "N'" & ProperCase(replacequote(request("MiddleName"))) & "'"
        else
          sql  = sql  & ",N'" & replacequote(request("MiddleName")) & "'"
          sqlu = sqlu & ",MiddleName=" & "N'" & replacequote(request("MiddleName")) & "'"
        end if
      else
        sqlf = sqlf & ",MiddleName"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",MiddleName=NULL"
      end if
    
      ' Last Name
      if not isblank(request("LastName")) then 
        sqlf = sqlf & ",LastName"       
        if User_Region = 1 or User_Region = 3 then  
          sql  = sql  & ",N'" & ProperCase(replacequote(request("LastName"))) & "'"
          sqlu = sqlu & ",LastName=" & "N'" & ProperCase(replacequote(request("LastName"))) & "'"         
        else
          sql  = sql  & ",N'" & replacequote(request("LastName")) & "'"
          sqlu = sqlu & ",LastName=" & "N'" & replacequote(request("LastName")) & "'"         
        end if  
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Last Name",Login_Language,conn) & "</LI>"         
      end if
    
      ' Suffix
      if not isblank(request("Suffix")) then 
        sqlf = sqlf & ",Suffix"       
        if User_Region = 1 or User_Region = 3 then  
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Suffix"))) & "'"
          sqlu = sqlu & ",Suffix=" & "N'" & ProperCase(replacequote(request("Suffix"))) & "'"         
        else
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
        if User_Region = 1 or User_Region = 3 then  
          sqlf = sqlf & ",Initials"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Initials"))) & "'"
          sqlu = sqlu & ",Initials=" & "N'" & UCase(replacequote(request("Initials"))) & "'"         
        else
          sqlf = sqlf & ",Initials"       
          sql  = sql  & ",N'" & replacequote(request("Initials")) & "'"
          sqlu = sqlu & ",Initials=" & "N'" & replacequote(request("Initials")) & "'"         
        end if  
      else
        sqlf = sqlf & ",Initials"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Initials=NULL"
      end if

      MailMessage     = MailMessage & "Name: "
      MailMessageUser = MailMessageUser & "Name: "
            
      MailMessage     = MailMessage & request("FirstName")
      MailMessageUser = MailMessageUser & request("FirstName")
    
      if not isblank(request("MiddleName")) then
        MailMessage = MailMessage & " " & request("MiddleName")
        MailMessageUser = MailMessageUser & " " & request("MiddleName")
      end if
      
      MailMessage = MailMessage & " " & request("LastName")
      MailMessageUser = MailMessageUser & " " & request("LastName")
    
      if not isblank(request("Suffix")) then
        MailMessage     = MailMessage     & ", " & request("Suffix")
        MailMessageUser = MailMessageUser & ", " & request("Suffix")
      end if
    
      MailMessage     = MailMessage     & vbCrLf
      MailMessageUser = MailMessageUser & vbCrLf
        
      ' Company
      if not isblank(request("Company")) then 
        sqlf = sqlf & ",Company"
        if User_Region = 1 or User_Region = 3 then          
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Company"))) & "'"
          sqlu = sqlu & ",Company=" & "N'" & ProperCase(replacequote(request("Company"))) & "'"
        else
          sql  = sql  & ",N'" & replacequote(request("Company")) & "'"
          sqlu = sqlu & ",Company=" & "N'" & replacequote(request("Company")) & "'"
        end if  
        MailMessage = MailMessage & "Company: " & request("Company") & vbCrLf        
        MailMessageUser = MailMessageUser & "Company: " & request("Company") & vbCrLf            
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Company",Login_Language,conn) & "</LI>"         
      end if
    
      ' Company Web Site
      if not isblank(request("Company_Website")) then 
        sqlf = sqlf & ",Company_Website"       
        sql  = sql  & ",N'" & replace(replace(replacequote(LCase(request("Company_Website"))),"http://",""),"https://","") & "'"
        sqlu = sqlu & ",Company_Website=" & "N'" & replace(replace(replacequote(LCase(request("Company_Website"))),"http://",""),"https://","") & "'"
      else
        sqlf = sqlf & ",Company_Website"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Company_Website=NULL"
      end if

      ' Job Title
      if not isblank(request("Job_Title")) then 
        sqlf = sqlf & ",Job_Title"
        if User_Region = 1 or User_Region = 3 then               
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Job_Title"))) & "'"
          sqlu = sqlu & ",Job_Title=" & "N'" & replacequote(request("Job_Title")) & "'"
        else
          sql  = sql  & ",N'" & replacequote(request("Job_Title")) & "'"
          sqlu = sqlu & ",Job_Title=" & "N'" & replacequote(request("Job_Title")) & "'"
        end if  
        MailMessage = MailMessage & "Job Title: " & request("Job_Title") & vbCrLf & vbCrLf
        MailMessageUser = MailMessageUser & "Job Title: " & request("Job_Title") & vbCrLf & vbCrLf
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Job Title",Login_Language,conn) & "</LI>"         
      end if
      
      ' NT Login
      if not isblank(request("NTLogin")) then            
        sqlf = sqlf & ",NTLogin"       
        sql  = sql  & ",N'" & killquote(request("NTLogin")) & "'"
        sqlu = sqlu & ",NTLogin=" & "N'" & killquote(request("NTLogin")) & "'"
        MailMessageUser = MailMessageUser & "Login / User Name: " & request("NTLogin") & vbCrLf    
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Logon / Users Name",Login_Language,conn) & "</LI>"         
      end if

      if lcase(Account_ID) = "add" and isblank(ErrorMessage) then
        SQL_Login = "SELECT UserData.NTLogin, UserData.Site_ID FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.NTLOGIN='" & request("NTLogin") & "'"
        Set rsNTLogin = Server.CreateObject("ADODB.Recordset")
        rsNTLogin.Open SQL_Login, conn, 3, 3

        if not rsNTLogin.EOF then
          ErrorMessage = ErrorMessage & "<LI>" & Translate("Logon / User Name already exists (Group Conflict), Select another Logon / User Name.",Login_Language,conn) & "</LI>"
        end if

        rsNTLogin.close
        set rsNTLogin = nothing
      end if

      ' Password
      if not isblank(request("Password")) then
        if not isblank(request("Password_Change")) then 
          sqlf = sqlf & ",Password"       
          sql  = sql  & ",N'" & replacequote(request("Password_Change")) & "'"
          sqlu = sqlu & ",Password=" & "N'" & replacequote(request("Password_Change")) & "'"
          MailMessageUser = MailMessageUser & "New Password: " & request("Password_Change") & vbCrLf & vbCrLf    
        else
          sqlf = sqlf & ",Password"       
          sql  = sql  & ",N'" & replacequote(request("Password")) & "'"
          sqlu = sqlu & ",Password=" & "N'" & replacequote(request("Password")) & "'"               
          MailMessageUser = MailMessageUser & "Password: " & request("Password") & vbCrLf & vbCrLf
        end if      
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Password",Login_Language,conn) & "</LI>"         
      end if
    
      MailMessage = MailMessage & "Email Address: " & request("EMail") & vbCrLf & vbCrLf
      
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Office Information" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf    
    
      ' Business Mail Stop
      if not isblank(request("Business_MailStop")) then 
        sqlf = sqlf & ",Business_MailStop"       
        sql  = sql  & ",N'" & replacequote(request("Business_MailStop")) & "'"
        sqlu = sqlu & ",Business_MailStop=" & "N'" & replacequote(request("Business_MailStop")) & "'"
        MailMessage = MailMessage & "Mail Stop / Building Number: " & request("Business_MailStop") & vbCrLf        
      else
        sqlf = sqlf & ",Business_MailStop"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Business_MailStop=NULL"
      end if
      
      ' Business Address
      if not isblank(request("Business_Address")) then 
        sqlf = sqlf & ",Business_Address"
        if User_Region = 1 or User_Region = 3 then                         
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_Address"))) & "'"
          sqlu = sqlu & ",Business_Address=" & "N'" & ProperCase(replacequote(request("Business_Address"))) & "'"         
        else
          sql  = sql  & ",N'" & replacequote(request("Business_Address")) & "'"
          sqlu = sqlu & ",Business_Address=" & "N'" & replacequote(request("Business_Address")) & "'"         
        end if  
        MailMessage = MailMessage & "Address 1: " & request("Business_Address") & vbCrLf        
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Address",Login_Language,conn) & "</LI>"         
      end if
      
      ' Business Address 2
      if not isblank(request("Business_Address_2")) then 
        sqlf = sqlf & ",Business_Address_2"
        if User_Region = 1 or User_Region = 3 then               
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_Address_2"))) & "'"
          sqlu = sqlu & ",Business_Address_2=" & "N'" & ProperCase(replacequote(request("Business_Address_2"))) & "'"
        else
          sql  = sql  & ",N'" & replacequote(request("Business_Address_2")) & "'"
          sqlu = sqlu & ",Business_Address_2=" & "N'" & replacequote(request("Business_Address_2")) & "'"
        end if
        MailMessage = MailMessage & "Address 2: " & request("Business_Address_2") & vbCrLf        
      else
        sqlf = sqlf & ",Business_Address_2"
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Business_Address_2=NULL"
      end if
    
      ' Business City
      if not isblank(request("Business_City")) then 
        sqlf = sqlf & ",Business_City"
        if User_Region = 1 or User_Region = 3 then
          sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_City"))) & "'"
          sqlu = sqlu & ",Business_City=" & "N'" & ProperCase(replacequote(request("Business_City"))) & "'"
        else
          sql  = sql  & ",N'" & replacequote(request("Business_City")) & "'"
          sqlu = sqlu & ",Business_City=" & "N'" & replacequote(request("Business_City")) & "'"
        end if  
        MailMessage = MailMessage & "City: " & request("Business_City") & vbCrLf
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office City",Login_Language,conn) & "</LI>"         
      end if
    
      ' Business State
      if not isblank(request("Business_State")) or not isblank(request("Business_State_Other")) then 
        if not isblank(request("Business_State")) then
          sqlf = sqlf & ",Business_State"       
          sql  = sql  & ",N'" & replacequote(request("Business_State")) & "'"
          sqlu = sqlu & ",Business_State=" & "N'" & replacequote(request("Business_State")) & "'"         
          MailMessage = MailMessage & "USA State or Canadian Province: " & request("Business_State") & vbCrLf
        else  
          sqlf = sqlf & ",Business_State"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Business_State=NULL"
        end if              
  
        if not isblank(request("Business_State_Other")) then
          sqlf = sqlf & ",Business_State_Other"
          if User_Region = 1 or User_Region = 3 then          
            sql  = sql  & ",N'" & ProperCase(replacequote(request("Business_State_Other"))) & "'"
            sqlu = sqlu & ",Business_State_Other=" & "N'" & ProperCase(replacequote(request("Business_State_Other"))) & "'"         
          else
            sql  = sql  & ",N'" & replacequote(request("Business_State_Other")) & "'"
            sqlu = sqlu & ",Business_State_Other=" & "N'" & replacequote(request("Business_State_Other")) & "'"         
          end if
          MailMessage = MailMessage & "Other State, Province or Local: " & request("Business_State_Other") & vbCrLf
        else
          sqlf = sqlf & ",Business_State_Other"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Business_State_Other=NULL"
        end if              
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office State, Province or Other Local",Login_Language,conn) & "</LI>"         
      end if

      ' Business Postal Code
      if not isblank(request("Business_Postal_Code")) then 
        if User_Region = 1 or User_Region = 3 then          
          sqlf = sqlf & ",Business_Postal_Code"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Business_Postal_Code"))) & "'"
          sqlu = sqlu & ",Business_Postal_Code=" & "N'" & replacequote(request("Business_Postal_Code")) & "'"
          MailMessage = MailMessage & "Postal Code: " & UCase(request("Business_Postal_Code")) & vbCrLf        
        else
          sqlf = sqlf & ",Business_Postal_Code"       
          sql  = sql  & ",N'" & replacequote(request("Business_Postal_Code")) & "'"
          sqlu = sqlu & ",Business_Postal_Code=" & "N'" & replacequote(request("Business_Postal_Code")) & "'"
          MailMessage = MailMessage & "Postal Code: " & request("Business_Postal_Code") & vbCrLf        
        end if
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Postal Code",Login_Language,conn) & "</LI>"         
      end if

      ' Business_Country
      if not isblank(request("Business_Country")) then 
        sqlf = sqlf & ",Business_Country"       
        sql  = sql  & ",N'" & replacequote(request("Business_Country")) & "'"
        sqlu = sqlu & ",Business_Country=" & "N'" & replacequote(request("Business_Country")) & "'"         
        MailMessage = MailMessage & "Country: " & request("Business_Country") & vbCrLf &vbCrLf       
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Country",Login_Language,conn) & "</LI>"         
      end if
    
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Postal Information" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf  
       
      ' Postal Address
      if not isblank(request("Postal_Address")) then 
        sqlf = sqlf & ",Postal_Address"       
        sql  = sql  & ",N'" & replacequote(request("Postal_Address")) & "'"
        sqlu = sqlu & ",Postal_Address=" & "N'" & replacequote(request("Postal_Address")) & "'"
        MailMessage = MailMessage & "Address (PO Box): " & request("Postal_Address") & vbCrLf        
      else
        sqlf = sqlf & ",Postal_Address"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Postal_Address=NULL"
      end if
         
      ' Postal City
      if not isblank(request("Postal_City")) then 
        sqlf = sqlf & ",Postal_City"       
        sql  = sql  & ",N'" & replacequote(request("Postal_City")) & "'"
        sqlu = sqlu & ",Postal_City=" & "N'" & replacequote(request("Postal_City")) & "'"
        MailMessage = MailMessage & "City: " & request("Postal_City") & vbCrLf
      else
        sqlf = sqlf & ",Postal_City"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Postal_City=NULL"
      end if
    
      ' Postal State    
      if not isblank(request("Postal_State")) or not isblank(request("Postal_State_Other")) then 
        if not isblank(request("Postal_State")) then
          sqlf = sqlf & ",Postal_State"       
          sql  = sql  & ",N'" & replacequote(request("Postal_State")) & "'"
          sqlu = sqlu & ",Postal_State=" & "N'" & replacequote(request("Postal_State")) & "'"         
          MailMessage = MailMessage & "USA State or Canadian Province: " & request("Postal_State") & vbCrLf
        else
          sqlf = sqlf & ",Postal_State"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Postal_State=NULL"
        end if              
        if not isblank(request("Postal_State_Other")) then
          sqlf = sqlf & ",Postal_State_Other"       
          sql  = sql  & ",N'" & replacequote(request("Postal_State_Other")) & "'"
          sqlu = sqlu & ",Postal_State_Other=" & "N'" & replacequote(request("Postal_State_Other")) & "'"         
          MailMessage = MailMessage & "Other State, Province or Local: " & request("Postal_State_Other") & vbCrLf
        else
          sqlf = sqlf & ",Postal_State_Other"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Postal_State_Other=NULL"
        end if
      else                
          sqlf = sqlf & ",Postal_State"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Postal_State=NULL"
          sqlf = sqlf & ",Postal_State_Other"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Postal_State_Other=NULL"
      end if
    
      ' Postal Postal Code
      if not isblank(request("Postal_Postal_Code")) then 
        if User_Region = 1 or User_Region = 3 then          
          sqlf = sqlf & ",Postal_Postal_Code"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Postal_Postal_Code"))) & "'"
          sqlu = sqlu & ",Postal_Postal_Code=" & "N'" & replacequote(request("Postal_Postal_Code")) & "'"         
          MailMessage = MailMessage & "Postal Code: " & UCase(request("Postal_Postal_Code")) & vbCrLf     
        else  
          sqlf = sqlf & ",Postal_Postal_Code"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Postal_Postal_Code"))) & "'"
          sqlu = sqlu & ",Postal_Postal_Code=" & "N'" & replacequote(request("Postal_Postal_Code")) & "'"         
          MailMessage = MailMessage & "Postal Code: " & UCase(request("Postal_Postal_Code")) & vbCrLf     
        end if
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
        MailMessage = MailMessage & "Country: " & request("Postal_Country") & vbCrLf & vbCrLf
      else
        sqlf = sqlf & ",Postal_Country"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Postal_Country=NULL"    
      end if

      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Shipping Information" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf  
    
      ' Shipping Mail Stop
      if not isblank(request("Shipping_MailStop")) then 
        sqlf = sqlf & ",Shipping_MailStop"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_MailStop")) & "'"
        sqlu = sqlu & ",Shipping_MailStop=" & "N'" & replacequote(request("Shipping_MailStop")) & "'"
        MailMessage = MailMessage & "Mail Stop / Building Number: " & request("Shipping_MailStop") & vbCrLf        
      else
        sqlf = sqlf & ",Shipping_MailStop"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Shipping_MailStop=NULL"
      end if
    
      ' Shipping Address
      if not isblank(request("Shipping_Address")) then 
        sqlf = sqlf & ",Shipping_Address"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_Address")) & "'"
        sqlu = sqlu & ",Shipping_Address=" & "N'" & replacequote(request("Shipping_Address")) & "'"
        MailMessage = MailMessage & "Address 1: " & request("Shipping_Address") & vbCrLf        
      else
        sqlf = sqlf & ",Shipping_Address"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Shipping_Address=NULL"
      end if
      
      ' Shipping Address 2
      if not isblank(request("Shipping_Address_2")) then 
        sqlf = sqlf & ",Shipping_Address_2"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_Address_2")) & "'"
        sqlu = sqlu & ",Shipping_Address_2=" & "N'" & replacequote(request("Shipping_Address_2")) & "'"
        MailMessage = MailMessage & "Address 2: " & request("Shipping_Address_2") & vbCrLf             
      else
        sqlf = sqlf & ",Shipping_Address_2"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Shipping_Address_2=NULL"
      end if
    
      ' Shipping City
      if not isblank(request("Shipping_City")) then 
        sqlf = sqlf & ",Shipping_City"       
        sql  = sql  & ",N'" & replacequote(request("Shipping_City")) & "'"
        sqlu = sqlu & ",Shipping_City=" & "N'" & replacequote(request("Shipping_City")) & "'"
        MailMessage = MailMessage & "City: " & request("Shipping_City") & vbCrLf
      else
        sqlf = sqlf & ",Shipping_City"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Shipping_City=NULL"
      end if
    
      ' Shipping State    
      if not isblank(request("Shipping_State")) or not isblank(request("Shipping_State_Other")) then 
        if not isblank(request("Shipping_State")) then
          sqlf = sqlf & ",Shipping_State"       
          sql  = sql  & ",N'" & replacequote(request("Shipping_State")) & "'"
          sqlu = sqlu & ",Shipping_State=" & "N'" & replacequote(request("Shipping_State")) & "'"         
          MailMessage = MailMessage & "USA State or Canadian Province: " & request("Shipping_State") & vbCrLf
        else
          sqlf = sqlf & ",Shipping_State"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Shipping_State=NULL"
        end if              
        if not isblank(request("Shipping_State_Other")) then
          sqlf = sqlf & ",Shipping_State_Other"       
          sql  = sql  & ",N'" & replacequote(request("Shipping_State_Other")) & "'"
          sqlu = sqlu & ",Shipping_State_Other=" & "N'" & replacequote(request("Shipping_State_Other")) & "'"         
          MailMessage = MailMessage & "Other State, Province or Local: " & request("Shipping_State_Other") & vbCrLf
        else
          sqlf = sqlf & ",Shipping_State_Other"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Shipping_State_Other=NULL"
        end if
      else                
          sqlf = sqlf & ",Shipping_State"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Shipping_State=NULL"
          sqlf = sqlf & ",Shipping_State_Other"       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Shipping_State_Other=NULL"
      end if
    
      ' Shipping Postal Code
      if not isblank(request("Shipping_Postal_Code")) then 
        if User_Region = 1 or User_Region = 3 then          
          sqlf = sqlf & ",Shipping_Postal_Code"       
          sql  = sql  & ",N'" & UCase(replacequote(request("Shipping_Postal_Code"))) & "'"
          sqlu = sqlu & ",Shipping_Postal_Code=" & "N'" & replacequote(request("Shipping_Postal_Code")) & "'"         
          MailMessage = MailMessage & "Postal Code: " & UCase(request("Shipping_Postal_Code")) & vbCrLf     
        else
          sqlf = sqlf & ",Shipping_Postal_Code"       
          sql  = sql  & ",N'" & replacequote(request("Shipping_Postal_Code")) & "'"
          sqlu = sqlu & ",Shipping_Postal_Code=" & "N'" & replacequote(request("Shipping_Postal_Code")) & "'"         
          MailMessage = MailMessage & "Postal Code: " & request("Shipping_Postal_Code") & vbCrLf     
        end if
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
        MailMessage = MailMessage & "Country: " & request("Shipping_Country") & vbCrLf & vbCrLf
      else
        sqlf = sqlf & ",Shipping_Country"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Shipping_Country=NULL"    
      end if
    
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Contact Information" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf      

      ' Business Phone
      if not isblank(request("Business_Phone")) then
        if User_Region = 1 then 
          sqlf = sqlf & ",Business_Phone"       
          sql  = sql  & ",N'" & FormatPhone(replacequote(request("Business_Phone"))) & "'"
          sqlu = sqlu & ",Business_Phone=" & "N'" & FormatPhone(replacequote(request("Business_Phone"))) & "'"         
          MailMessage = MailMessage & "Office Phone (Direct):" & FormatPhone(request("Business_Phone"))
        else
          sqlf = sqlf & ",Business_Phone"       
          sql  = sql  & ",N'" & replacequote(request("Business_Phone")) & "'"
          sqlu = sqlu & ",Business_Phone=" & "N'" & replacequote(request("Business_Phone")) & "'"         
          MailMessage = MailMessage & "Office Phone (Direct):" & request("Business_Phone")
        end if  
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Office Phone (Direct)",Login_Language,conn) & "</LI>"         
      end if
      
      ' Business Phone Extension
      if not isblank(request("Business_Phone_Extension")) then 
        sqlf = sqlf & ",Business_Phone_Extension"       
        sql  = sql  & ",N'" & replacequote(request("Business_Phone_Extension")) & "'"
        sqlu = sqlu & ",Business_Phone_Extension=" & "N'" & replacequote(request("Business_Phone_Extension")) & "'"         
        MailMessage = MailMessage & "  extension: " & request("Business_Phone_Extension") & vbCrLf
      else
        sqlf = sqlf & ",Business_Phone_Extension"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Business_Phone_Extension=NULL"
        MailMessage = MailMessage & vbCrLf
      end if
    
      ' Business_Phone_2
      if not isblank(request("Business_Phone_2")) then
        if User_Region = 1 then 
          sqlf = sqlf & ",Business_Phone_2"       
          sql  = sql  & ",N'" & FormatPhone(replacequote(request("Business_Phone_2"))) & "'"
          sqlu = sqlu & ",Business_Phone_2=" & "N'" & FormatPhone(replacequote(request("Business_Phone_2"))) & "'"
          MailMessage = MailMessage & "Office Phone (General):" & FormatPhone(request("Business_Phone_2"))
        else
          sqlf = sqlf & ",Business_Phone_2"       
          sql  = sql  & ",N'" & replacequote(request("Business_Phone_2")) & "'"
          sqlu = sqlu & ",Business_Phone_2=" & "N'" & replacequote(request("Business_Phone_2")) & "'"         
          MailMessage = MailMessage & "Office Phone (General):" & request("Business_Phone_2")
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
        MailMessage = MailMessage & "  extension: " & request("Business_Phone_2_Extension") & vbCrLf
      else
        sqlf = sqlf & ",Business_Phone_2_Extension"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Business_Phone_2_Extension=NULL"
        MailMessage = MailMessage & vbCrLf
      end if
      
      ' Business_Fax
      if not isblank(request("Business_Fax")) then 
        if User_Region = 1 then
          sqlf = sqlf & ",Business_Fax"       
          sql  = sql  & ",N'" & FormatPhone(replacequote(request("Business_Fax"))) & "'"
          sqlu = sqlu & ",Business_Fax=" & "N'" & FormatPhone(replacequote(request("Business_Fax"))) & "'"         
          MailMessage = MailMessage & "Office Fax: " & FormatPhone(request("Business_Fax")) & vbCrLf
        else
          sqlf = sqlf & ",Business_Fax"       
          sql  = sql  & ",N'" & replacequote(request("Business_Fax")) & "'"
          sqlu = sqlu & ",Business_Fax=" & "N'" & replacequote(request("Business_Fax")) & "'"         
          MailMessage = MailMessage & "Office Fax: " & request("Business_Fax") & vbCrLf
        end if  
      else
        sqlf = sqlf & ",Business_Fax"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Business_Fax=NULL"
      end if
      
      ' Mobile_Phone
      if not isblank(request("Mobile_Phone")) then 
        if User_Region = 1 then
          sqlf = sqlf & ",Mobile_Phone"       
          sql  = sql  & ",N'" & FormatPhone(replacequote(request("Mobile_Phone"))) & "'"
          sqlu = sqlu & ",Mobile_Phone=" & "N'" & FormatPhone(replacequote(request("Mobile_Phone"))) & "'"         
          MailMessage = MailMessage & "Mobile Phone: " & FormatPhone(request("Mobile_Phone")) & vbCrLf
        else
          sqlf = sqlf & ",Mobile_Phone"       
          sql  = sql  & ",N'" & replacequote(request("Mobile_Phone")) & "'"
          sqlu = sqlu & ",Mobile_Phone=" & "N'" & replacequote(request("Mobile_Phone")) & "'"         
          MailMessage = MailMessage & "Mobile Phone: " & request("Mobile_Phone") & vbCrLf
        end if  
      else
        sqlf = sqlf & ",Mobile_Phone"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Mobile_Phone=NULL"
      end if
      
      ' Pager
      if not isblank(request("Pager")) then 
        if User_Region = 1 then
          sqlf = sqlf & ",Pager"       
          sql  = sql  & ",N'" & FormatPhone(replacequote(request("Pager"))) & "'"
          sqlu = sqlu & ",Pager=" & "N'" & FormatPhone(replacequote(request("Pager"))) & "'"         
          MailMessage = MailMessage & "Pager: " & FormatPhone(request("Pager")) & vbCrLf
        else
          sqlf = sqlf & ",Pager"       
          sql  = sql  & ",N'" & replacequote(request("Pager")) & "'"
          sqlu = sqlu & ",Pager=" & "N'" & replacequote(request("Pager")) & "'"         
          MailMessage = MailMessage & "Pager: " & request("Pager") & vbCrLf
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
        MailMessage = MailMessage & "Email (Direct): " & request("EMail") & vbCrLf
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Email (Direct)",Login_Language,conn) & " (" & Translate("Primary",Login_Language,conn) & ")</LI>"         
      end if
     
      ' Email_2
      if not isblank(request("Email_2")) then 
        sqlf = sqlf & ",Email_2"       
        sql  = sql  & ",N'" & replacequote(request("Email_2")) & "'"
        sqlu = sqlu & ",Email_2=" & "N'" & replacequote(request("Email_2")) & "'"         
        MailMessage = MailMessage & "Email (General Office): " & request("EMail_2") & vbCrLf
      else
        sqlf = sqlf & ",Email_2"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Email_2=NULL"
      end if
  
      ' Email Format
      if not isblank(request("EMail_Method")) then 
        sqlf = sqlf & ",EMail_Method"       
        sql  = sql  & "," & replacequote(request("EMail_Method")) & ""
        sqlu = sqlu & ",EMail_Method=" & "" & replacequote(request("EMail_Method")) & ""
      else
        sqlf = sqlf & ",EMail_Method"       
        sql  = sql  & ",0"
        sqlu = sqlu & ",EMail_Method=0"
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
    
'      ' Subscription Method
'      if not isblank(request("Subscription_Method")) then 
'        sqlf = sqlf & ",Subscription_Method"       
'        sql  = sql  & "," & replacequote(request("Subscription_Method")) & ""
'        sqlu = sqlu & ",Subscription_Method=" & "" & replacequote(request("Subscription_Method")) & ""
'      else  
'        sqlf = sqlf & ",Subscription_Method"       
'        sql  = sql  & ",0"
'        sqlu = sqlu & ",Subscription_Method=0"
'      end if
    
      ' Subscription Options (Numeric Array)
'      if not isblank(request("Subscription_Options")) then 
'        sqlf = sqlf & ",Subscription_Options"
'        sql  = sql  & ",'" & replacequote(request("Subscription_Options")) & "'"
'        sqlu = sqlu & ",Subscription_Options=" & "'" & replacequote(request("Subscription_Options")) & "'"
'      else  
'        sqlf = sqlf & ",Subscription_Options"
'        sql  = sql  & ",'0'"
'        sqlu = sqlu & ",Subscription_Options='0'"
'      end if
  
      ' Connection Speed
      if not isblank(request("Connection_Speed")) and request("Connection_Speed") <> "0" then 
        sqlf = sqlf & ",Connection_Speed"       
        sql  = sql  & "," & replacequote(request("Connection_Speed")) & ""
        sqlu = sqlu & ",Connection_Speed=" & "" & replacequote(request("Connection_Speed")) & ""
      else
        sqlf = sqlf & ",Connection_Speed"       
        sql  = sql  & ",0"
        sqlu = sqlu & ",Connection_Speed=0"
      end if
  
      MailMessage = MailMessage & vbCrLf & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Other Information" & vbCrLf
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf  
     
      ' Language
      if not isblank(request("Account_Language")) then 
    
        sqlf = sqlf & ",Language"       
        sql  = sql  & ",N'" & replacequote(request("Account_Language")) & "'"
        sqlu = sqlu & ",Language=" & "N'" & replacequote(request("Account_Language")) & "'"
        
        SQL_Language = "SELECT Language.* FROM Language WHERE Language.Code='" & request("Account_Language") & "'"
        Set rsLanguage = Server.CreateObject("ADODB.Recordset")
        rsLanguage.Open SQL_Language, conn, 3, 3
    
        if not rsLanguage.EOF then
          MailMessage = MailMessage & "Preferred Language: " & rsLanguage("Description") & vbCrLf    
        else
          MailMessage = MailMessage & "Preferred Language: " & ucase(request("Account_Language")) & vbCrLf
        end if                                                
    
        rsLanguage.close
        set rsLanguage=nothing
      else
        sqlf = sqlf & ",Language"       
        sql  = sql  & ",'eng'"
        sqlu = sqlu & ",Language=" & "'eng'"
      end if

      ' --------------------------------------------------------------------------------------
      ' Reciprocal Site(s) - Copy User Specific SQL Statement up to this point
      ' --------------------------------------------------------------------------------------    
  
      sqlfR = sqlf
      sqlR  = sql
      sqluR = sqlu
  
      ' --------------------------------------------------------------------------------------
      ' Site Specific Information
      ' --------------------------------------------------------------------------------------
  
      ' Site ID  
      if not isblank(request("Site_ID")) then
        sqlf = sqlf & ",Site_ID"
        sql  = sql  & "," & killquote(request("Site_ID")) & ""       
        sqlu = sqlu & ",Site_ID=" & killquote(request("Site_ID"))
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Internal Site ID, report this error to the Site Administrator.",Login_Language,conn) & "</LI>"
      end if
  
      ' Groups
      if not isblank(Groups) then 
        sqlf = sqlf & ",Groups"       
        sql  = sql  & ",'" & replacequote(Groups) & "'"
        sqlu = sqlu & ",Groups=" & "'" & replacequote(Groups) & "'"             
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Primary Group Affiliation(s).",Login_language,conn) & "</LI>"         
      end if
    
      ' Groups_Aux
      if not isblank(request("Groups_Aux")) then 
        sqlf = sqlf & ",Groups_Aux"       
        sql  = sql  & ",'" & replacequote(request("Groups_Aux")) & "'"
        sqlu = sqlu & ",Groups_Aux=" & "'" & replacequote(request("Groups_Aux")) & "'"             
      else
        sqlf = sqlf & ",Groups_Aux"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Groups_Aux=NULL"
      end if
    
      ' SubGroups
      if not isblank(request("SubGroups")) then
        tempSubGroups = request("SubGroups")

        ' Prioritize Order Inquiry Level
        if instr(1,tempSubGroups,"ordad") > 0 and instr(1,tempSubGroups,"order") > 0 then
          tempSubGroups = replace(tempSubGroups,"order, ","")
        end if  
  
        if instr(1,tempSubGroups,"domain") > 0 then
          tempSubGroups = "domain"
          sqlf = sqlf & ",Account_Region"
          sql  = sql  & ",4"
          sqlu = sqlu & ",Account_Region=4"
        elseif instr(1,tempSubGroups,"administrator") > 0 then
          tempSubGroups = "administrator"
          sqlf = sqlf & ",Account_Region"
          sql  = sql  & ",4"
          sqlu = sqlu & ",Account_Region=4"
        elseif (instr(1,tempSubGroups,"account")   > 0 and instr(1,tempSubGroups,"content")   > 0) or _
               (instr(1,tempSubGroups,"account")   > 0 and instr(1,tempSubGroups,"submitter") > 0) or _
               (instr(1,tempSubGroups,"account")   > 0 and instr(1,tempSubGroups,"branch")  > 0) or _
               (instr(1,tempSubGroups,"content")   > 0 and instr(1,tempSubGroups,"submitter") > 0) or _
               (instr(1,tempSubGroups,"content")   > 0 and instr(1,tempSubGroups,"branch")  > 0) or _
               (instr(1,tempSubGroups,"submitter") > 0 and instr(1,tempSubGroups,"branch")  > 0) then
          ErrorMessage = ErrorMessage & "<LI>" & Translate("This Account cannot have multiple Administration or Administration / Submission previlages.",Login_Language,conn) & " " & Translate("Select only one Administration Group Option.",Login_Language,conn) & "</LI>"
        elseif instr(1,tempSubGroups,"branch")  > 0 and (isblank(request("Fluke_ID")) or isblank(request("Type_Code"))) then
          ErrorMessage = ErrorMessage & "<LI>" & Translate("Branch Location Administration privilage requires Fluke Customer Type and Fluke Customer ID Number.",Login_Language,conn) & "</LI>"
        else
        ' Account Administrator
          if instr(1,request("SubGroups"),"account") > 0 then
            if isblank(request("Account_Region")) or CInt(request("Account_Region")) = 0 then
              ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Account Administrator Region Designation.",Login_Language,conn) & "</LI>"
            else  
              sqlf = sqlf & ",Account_Region"
              sql  = sql  & "," & Request("Account_Region")
              sqlu = sqlu & ",Account_Region=" & Request("Account_Region")
            end if
          else
            sqlf = sqlf & ",Account_Region"
            sql  = sql  & ",0" 
            sqlu = sqlu & ",Account_Region=0"
          end if    
        end if
        
        sqlf = sqlf & ",SubGroups"       
        sql  = sql  & ",'" & replacequote(replace(tempSubGroups,",,",",")) & "'"
        sqlu = sqlu & ",SubGroups=" & "'" & replacequote(replace(tempSubGroups,",,",",")) & "'"
        
      else
        ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Group Affiliation(s).",Login_Language,conn) & "</LI>"
      end if
      
      ' FCM
      if request("Fcm") = "on" or request("Fcm") = "-1" then           
        sqlf = sqlf & ",Fcm"
        sql  = sql  & "," & CInt(True)     
        sqlu = sqlu & ",Fcm=" & CInt(True)
      else     
        sqlf = sqlf & ",Fcm"
        sql  = sql  & "," & CInt(False)
        sqlu = sqlu & ",Fcm=" & CInt(False)
      end if        
  
      ' FCM_ID
      if request("Fcm") <> "on" and request("Fcm") <> "-1" then
        if not isblank(request("Fcm_ID")) and isnumeric(request("Fcm_ID")) then               
          sqlf = sqlf & ",Fcm_ID"              
          sql  = sql  & "," & killquote(request("Fcm_ID"))
          sqlu = sqlu & ",Fcm_ID=" & killquote(request("Fcm_ID"))
        else
          ErrorMessage = ErrorMessage & "<LI>" & Translate("Missing Account Manager or N/A",Login_Language,conn) & "</LI>"     
        end if
      elseif request("Fcm") = "on" or request("Fcm") = "-1" then
          sqlf = sqlf & ",Fcm_ID"              
          sql  = sql  & "," & CInt(False)
          sqlu = sqlu & ",Fcm_ID=" & CInt(False)
      end if    
  
      ' Euro Contact Management System ID
      if not isblank(request("CM_ID")) and isnumeric(request("CM_ID")) then               
        sqlf = sqlf & ",CM_ID"              
        sql  = sql  & "," & killquote(request("CM_ID"))
        sqlu = sqlu & ",CM_ID=" & killquote(request("CM_ID"))
      else
        sqlf = sqlf & ",CM_ID"              
        sql  = sql  & ",0"
        sqlu = sqlu & ",CM_ID=0"
      end if

      ' Reg_Request_Date
      if not isblank(request("Reg_Request_Date")) and isnumeric(request("Reg_Request_Date")) then               
        sqlf = sqlf & ",Reg_Request_Date"              
        sql  = sql  & ",'" & killquote(request("Reg_Request_Date")) & "'"
      end if
  
      ' Change ID   
      if not isblank(request("ChangeID")) and isnumeric(request("ChangeID")) then               
        sqlf = sqlf & ",ChangeID"              
        sql  = sql  & "," & killquote(request("ChangeID"))
        sqlu = sqlu & ",ChangeID=" & killquote(request("ChangeID"))
      end if
    
      ' Change Date
      if isdate(request("ChangeDate")) then   
        sqlf = sqlf & ",ChangeDate"                   
        sql  = sql  & ",'" & killquote(request("ChangeDate")) & "'"
        sqlu = sqlu & ",ChangeDate=" & "'" & killquote(request("ChangeDate")) & "'"
      end if
      
      if ((request("RTE_Enabled") = "on" or request("RTE_Enabled") = "-1") _
         and instr(1,LCase(request("Subgroups")),"content") > 0) _           
         or instr(1,LCase(request("Subgroups")),"domain") > 0 _
         or instr(1,LCase(request("Subgroups")),"administrator") > 0 then           
        sqlf = sqlf & ",RTE_Enabled"                   
        sql  = sql  & "," & CInt(True)
        sqlu = sqlu & ",RTE_Enabled=" & CInt(True)
      else
        sqlf = sqlf & ",RTE_Enabled"                   
        sql  = sql  & "," & CInt(False)
        sqlu = sqlu & ",RTE_Enabled=" & CInt(False)
      end if
          
      ' New Account Flag - (Turn off with update)
      if request("NewFlag") = "on" or request("NewFlag") = "-1" then               
        sqlf = sqlf & ",NewFlag"              
        sql  = sql  & "," & CInt(False)
        
        sqlf = sqlf & ",Reg_Approval_Date"
        sql  = sql  & ",'" & Date & "'"
        
        sqlf = sqlf & ",Reg_Admin"
        sql  = sql  & "," & Admin_ID
        
        sqlu = sqlu & ",NewFlag=" & CInt(False)
        sqlu = sqlu & ",Reg_Approval_Date='" & Date & "'"
        sqlu = sqlu & ",Reg_Admin=" & Admin_ID           
      else
        sqlf = sqlf & ",NewFlag"
        sql  = sql  & "," & CInt(False)
        sqlu = sqlu & ",NewFlag=" & CInt(False)
      end if
             
      ' Comment
      if not isblank(request("Comment")) then 
        sqlf = sqlf & ",Comment"       
        sql  = sql  & ",N'" & replacequote(request("Comment")) & "'"
        sqlu = sqlu & ",Comment=" & "N'" & replacequote(request("Comment")) & "'"
      else
        sqlf = sqlf & ",Comment"       
        sql  = sql  & ",NULL"
        sqlu = sqlu & ",Comment=NULL"
      end if
    
      ' Auxiliary Fields
    
      for aux = 0 to 9 
      
        if not isblank(request("Aux_" & Trim(aux))) then 
          sqlf = sqlf & ",Aux_" & Trim(aux)       
          sql  = sql  & ",N'" & replacequote(request("Aux_" & Trim(aux))) & "'"
          sqlu = sqlu & ",Aux_" & Trim(aux) & "=" & "N'" & replacequote(request("Aux_" & Trim(aux))) & "'"         
        else
          sqlf = sqlf & ",Aux_" & Trim(aux)       
          sql  = sql  & ",NULL"
          sqlu = sqlu & ",Aux_" & Trim(aux) & "=NULL"
        end if
    
      next
       
      MailMessage     = MailMessage     & vbCrLf & "----------------------------------------------------------" & vbCrLf & vbCrLf          
      MailMessageUser = MailMessageUser & vbCrLf & "----------------------------------------------------------" & vbCrLf & vbCrLf          
      
      ' Build Ending SQL Statement
      
      sqlf = sqlf & ")"
      sql  = sql  & ")"
      sql= sqlf & sql

      ' --------------------------------------------------------------------------------------
      ' Add / Update  
      ' --------------------------------------------------------------------------------------
						'---------------------------------------------------------------------------------------
      New_Account_ID = Account_ID
      
      if isnumeric(Account_ID) and isblank(ErrorMessage) then
									' Update Existing Account
									Action = "Update"
									New_Account_ID = Account_ID
									strPost_QueryString = sqlu
									sqlu   = "UPDATE UserData SET " & sqlu & " WHERE UserData.ID=" & Account_ID
      elseif not isnumeric(Account_ID) and isblank(ErrorMessage) then
									' Create New Account using Add/Update Method to get New_Account_ID
									New_Account_ID = CLng(Get_New_Record_ID ("UserData", "NewFlag", 0, conn))
									Action = "Add"
									strPost_QueryString = sqlu        
									sqlu = "UPDATE UserData SET " & sqlu & " WHERE UserData.ID=" & New_Account_ID
      else
        Call ErrorHandler
      end if
      
      SQL_Fields = "SELECT * FROM UserData Where ID=" & New_Account_ID
      Set rsFields = Server.CreateObject("ADODB.Recordset")
      
      rsFields.Open SQL_Fields, conn, 3, 3
    
      for each field In rsFields.Fields
    
        select case UCase(field.Name)
          case "COMMON_ID", "BROADCAST_DATE", "SUBSCRIPTION_DATE", "SUBSCRIPTION_FREQUENCY", _
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

      if isblank(ErrorMessage) then
        ' Add - Update = UserData Table = Main Site
								' response.write Replace(sqlu,",",",<BR>") & "<P>" & vbcrlf
        conn.execute (SQLU)
        '-------------------------------------------------------------------------------------------------------
        if site_id =3 then
								  	'---------------Modified by Zensar for implementing Field Marketing Site changes--------
											set cmd = Server.CreateObject("ADODB.Command")
											cmd.ActiveConnection = conn
											cmd.CommandType = adCmdStoredProc
											if Request("CFMEnabled")= 1 then
														if CDate(request("ExpirationDate")) < CDate(Now()) then
																cmd.CommandText = "FMDeleteUser"
																cmd.Parameters.Append cmd.CreateParameter("@AccountID",adBigInt,adParamInput,Account_ID)
																cmd.Parameters(0).Value= Account_ID
														else
																	cmd.CommandText = "FMInsUpdUser"
																	cmd.Parameters.Append cmd.CreateParameter("@AccountID",adBigInt,adParamInput,New_Account_ID)
																	cmd.Parameters.Append cmd.CreateParameter("@ModifiedBy",adBigInt,adParamInput,request("ChangeID"))
																	cmd.Parameters(0).Value= New_Account_ID
																	cmd.Parameters(1).Value= request("ChangeID")
														end if	
											else
														cmd.CommandText = "FMDeleteUser"
														cmd.Parameters.Append cmd.CreateParameter("@AccountID",adBigInt,adParamInput,Account_ID)
														cmd.Parameters(0).Value= Account_ID
											end if
											cmd.execute
											set cmd = nothing  
								end if 
        '-------------------------------------------------------------------------------------------------------
        ' Update Common_ID
        SQLID = "SELECT ID, Common_ID, NTLogin, ExpirationDate, Logon FROM UserData where NTLogin='" & request("NTLogin") & "' ORDER BY ID"
        Set rsID = Server.CreateObject("ADODB.Recordset")
        rsID.Open SQLID, conn, 3, 3
  
        myID = rsID("ID")
    
        rsID.close
        set rsID = nothing
    
        SQLID = "UPDATE UserData SET Common_ID=" & myID & " WHERE NTLogin='" & request("NTLogin") & "'"
        conn.execute SQLID
        
        ' Determine if Posting is required to a Regional (CM) Contact Management System,
        ' if false notify admin default email method.

        strPost_QueryString = strPost_QueryString & ",ID=" & New_Account_ID
        
        select case CInt(CMS_Region(User_Region))
          case CInt(True)         ' CM Sytem Only
            if CInt(No_Post_Back) = CInt(false) then
              Call Send_2_CMS     ' Send Data to CM Sytem
            end if  
          case -2                 ' CM System and Email
            if CInt(No_Post_Back) = CInt(false) then          
              Call Send_2_CMS     ' Send Data to CM System        
            end if  
        end select

        SQLUser = "SELECT ID, NTLogin, ExpirationDate, Site_ID, Logon FROM UserData WHERE UserData.NTLogin='" & request("NTLogin") & "' AND NewFlag=0"
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
              if CDate(rsUser("ExpirationDate")) < CDate(Now()) then   ' Account expired - do nothing
              
              elseif trim(request("ExpirationDate")) <> "" then
                if (CDate(request("ExpirationDate")) > CDate(rsUser("ExpirationDate"))) and CInt(rsSite("Renew_Days")) = 0 then
                    SQLU = "UPDATE UserData SET ExpirationDate='" & CDate(request("ExpirationDate")) & "' " &_
                           "WHERE ID=" & rsUser("ID")
                    conn.Execute (SQLU)
                elseif (CDate(request("ExpirationDate")) > CDate(rsUser("ExpirationDate"))) and CDate(DateAdd("d",CInt(rsSite("Renew_Days")),Date())) < CDate(request("ExpirationDate")) then
                    SQLU = "UPDATE UserData SET ExpirationDate='" & CDate(request("ExpirationDate")) & "' " &_
                           "WHERE ID=" & rsUser("ID")
                    conn.Execute (SQLU)
                end if
              end if  
            end if
            rsSite.close
            set rsSite = Nothing
            
          end if
          
          rsUser.MoveNext
        loop
        
        rsUser.close
        set rsUser = nothing

        ' Add - Update - Delete UserData Table - Recprocity Site(s)

        Call Recprocity

      end if
      
      ' --------------------------------------------------------------------------------------
      ' Instant Messaging
      ' --------------------------------------------------------------------------------------             

      if not isblank(request("Message")) then
        SQL_Message = "INSERT INTO MESSAGES (NTLogin,To_tName,Fm_Name,Message_Date,Message) " &_
                      "VALUES ('" & Request("NTLogin") & "'," &_
                      "'" & request("Firstname") & " " & request("Lastname") & "'," &_
                      "'" & Admin_FirstName & " " & Admin_LastName & "'," &_
                      "'" & Now() & "'," &_
                      "N'" & replacequote(request("Message")) & "')"
        conn.execute SQL_Message
        set SQL_Message = nothing
      end if  

      ' --------------------------------------------------------------------------------------
      ' EMail Add / Update and Redirect
      ' --------------------------------------------------------------------------------------             

      if isblank(ErrorMessage) and (CInt(Send_Email_Domain) = CInt(True) or CInt(Send_Email_Admin) = CInt(True) or CInt(Send_Email_Fcm) = CInt(True)) then

        MailMessage = MailMessage & vbCrLf & "Sincerely," & vbCrLf & vbCrLf & "The " & MailFromName & " Support Team"

        ' Get Channel Manager's Info
        if Send_Email_Fcm = true then
          SQL_FCM = "Select FirstName, LastName, Email FROM UserData WHERE ID=" & request("Fcm_ID")
          Set rsManager = Server.CreateObject("ADODB.Recordset")
          rsManager.Open SQL_FCM, conn, 3, 3
        
          if not rsManager.EOF then
            ManagerName = rsManager("FirstName") & " " & rsManager("LastName")
            Mailer.AddCC ManagerName, rsManager("EMail")
          end if
        
          rsManager.Close
          Set rsManager = nothing
        end if          
        
        Call Send_EMail
      end if
      
      if not isblank(ErrorMessage) then
        Call ErrorHandler
      else
      
        if Send_Email_User = True then
          Mailer.ClearAllRecipients
          Mailer.AddRecipient request("FirstName") & " " & request("LastName"), request("Email")
          Mailer.ClearBodyText
          MailMessageUser = MailMessageUser & vbCrLf & "Sincerely," & vbCrLf & vbCrLf & "The " & MailFromName & " Support Team" & vbCrLf & vbCrLf
          MailMessageUser = MailMessageUser & "Questions or comments regarding the content of this site should be directed to: " & MailReplyToName & " email: " & MailReplyTo & vbCrLf & vbCrLf
          MailMessageUser = MailMessageUser & "Report site problems or errors to: " & MailBCCName & "  email: " & MailBCC & ".  Please copy and paste the complete URL and any error messages reported on your screen into the eMail describing the problem you are reporting." & vbCrLf & vbCrLf
          MailMessageUser = MailMessageUser & "Save this notice in a secure place for future reference.  It contains the Site URL Address, and your Login User Name and Password (case sensitive)information."
          MailMessage     = MailMessageUser
          Call Send_EMail
        end if
        
        if not isblank(ErrorMessage) then  
          Call ErrorHandler
        else
  
          if not isblank(BackURL) then
            response.redirect BackURL
          else
            response.end
          end if    
    
        end if    
                           
      end if
      
  end if
  
  ' Last bit of house cleaning.  For those records assigned a non exsistant FCM_ID, delete reference.

  SQLFCM =  "SELECT DISTINCT FCM_ID " &_
            "FROM dbo.UserData " &_
            "WHERE Fcm_ID <> 0 AND FCM_ID not in( " &_
            "SELECT ID " &_
            "FROM dbo.UserData " &_
            "WHERE ID IN(SELECT DISTINCT FCM_ID " &_
            "FROM dbo.UserData " &_
            "WHERE Fcm_ID <> 0))"
            
   conn.execute SQLFCM
  
end if  ' script debug

' --------------------------------------------------------------------------------------
' Subroutines
' --------------------------------------------------------------------------------------

' Send_2_CMS Functions
%>
<!--#include virtual="/Include/Function_DCMDataTransfer.asp"-->
<%

sub Recprocity
  
  ' Add to UserData Table - Recprociticy Site Access Requests
  
  SQL_Site_Aux =                "SELECT Site_Aux.Site_ID, Site_Aux.Site_ID_Aux, Site.Site_Code, Site.Enabled, Site.Site_Description "
  SQL_Site_Aux = SQL_Site_Aux & "FROM Site_Aux LEFT JOIN Site ON Site_Aux.Site_ID_Aux = Site.ID "
  SQL_Site_Aux = SQL_Site_Aux & "WHERE Site_Aux.Site_ID=" & Site_ID & " ORDER BY Site.Site_Description"
    
  Set rsSite_Aux = Server.CreateObject("ADODB.Recordset")
  rsSite_Aux.Open SQL_Site_Aux, conn, 3, 3

  if not rsSite_Aux.EOF then

    do while not rsSite_Aux.EOF

      if instr(1,LCase(request("Groups_Aux")),LCase(rsSite_Aux("Site_Code"))) > 0 then

        'Check to see that there is no pre-existing account for this NTLogin
        
        SQL_Login = "SELECT Site_ID, NTLogin FROM UserData WHERE Site_ID=" & CInt(rsSite_Aux("Site_ID_Aux")) & " AND NTLogin='" & request("NTLogin") & "'"
        Set rsNTLogin = Server.CreateObject("ADODB.Recordset")
        rsNTLogin.Open SQL_Login, conn, 3, 3
        
        if rsNTLogin.EOF then
  
          ' Build New Account Request:  Set to NEWFLAG = -2 to show Reciprocal Site Access Request

          sqlfRx = sqlfR & ",Site_ID,NewFlag,Reg_Request_Date,Comment) "
          sqlRx  = sqlR & "," & rsSite_Aux("Site_ID_Aux") & "," & CInt(-2) & ",'" & Now & "','This new account request was automatically created by the " & Site_Description & " Extranet Site for the purpose of reciprocal site access. Please complete the Group Affiliation and Reciprocal Site section of this account then update.')"
          sqlRx  = sqlfRx & sqlRx
          conn.execute (sqlRx)
          
        else
        
          ' Update Existing Account       

          sqluRx = "UPDATE UserData SET " & sqluR
          sqluRx = sqluRx & ",ChangeID=" & killquote(request("ChangeID")) & ",ChangeDate=" & "'" & killquote(request("ChangeDate")) & "'"
          sqluRx = sqluRx & " WHERE Site_ID=" & rsSite_Aux("Site_ID_Aux") & " AND NTLogin='" & request("NTLogin") & "'"
          conn.execute (sqluRx)

        end if
        
        rsNTLogin.close
        set rsNTLogin = nothing
        
      end if
  
      rsSite_Aux.MoveNext
  
    loop
  
  end if
    
  rsSite_Aux.Close
  set rsSite_Aux = nothing

end sub

' --------------------------------------------------------------------------------------

sub ErrorHandler

  if CInt(CMS_Region(User_Region)) = CInt(False) or CInt(Send_2_CMS_Debug) = CInt(True) then
    Screen_Title = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Users Account Error Report Screen",Alt_Language,conn)
    Bar_Title = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=NormalBoldGold>" & Translate("Users Account Error Report Screen",Login_Language,conn) & "</FONT>"
    Navigation = False
    Top_Navigation = False
    Content_Width = 95  ' Percent

    %>
    <!--#include virtual="/SW-Common/SW-Header.asp"-->
    <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
    <%
  
    response.write "<FONT CLASS=Normal><B>" & Translate("Use the [ Back ] button of your browser to return to the previous screen to correct the error(s) listed below",Login_Language,conn) & "<BR>" & Translate("or click on [ Main Menu ] to return to the Account Administration screen.",Login_Language,conn) & "</B><BR><BR>"
    response.write "<FONT CLASS=NormalRed><UL>" & ErrorMessage & "</UL></FONT><BR><BR>"
    response.write Translate("Email any application errors noted to be reported to the Extranet Webmaster or if you have questions regarding this site or the site administration tools, to:",Login_Language,conn) & " "
    '>>>>>>>>>
    'Replaced Kelly's mail id on 20-11-2007
    response.write "<A HREF=""mailto:extranetalerts@fluke.com;"">Extranet Group</A>, " & Translate("Fluke Extranet Webmaster",Login_Language,conn) & ".<BR><BR></FONT>"
    '>>>>>>>>>>>
    %>
    <INPUT TYPE="BUTTON" Value=" Main Menu " onclick="Redirect('default.asp?Site_ID=<%=Site_ID%>');" CLASS=NavLeftHighlight1 ID="Button1" NAME="Button1"><BR><BR>
  
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  
    <SCRIPT LANGUAGE="JavaScript">
    <!--
    function Redirect(MyURL){ 
      window.location = MyURL;
    } 
    // -->
    </SCRIPT>
    <%
    
  else

   Action = "Error"
   strPost_QueryString = "ErrorMessage='" & Replace(Replace(ErrorMessage,"<LI>",""),"</LI>","") & "'"
   Call Send_2_CMS     ' Send Data to CM Sytem

 end if  

end sub

' --------------------------------------------------------------------------------------

sub Send_EMail

  Mailer.QMessage = False
  Mailer.Subject  = MailSubject
  Mailer.BodyText = MailMessage

  if Mailer.SendMail then

  else
    ErrorMessage = ErrorMessage & vbCrLf & "<LI>" & Translate("Send Email Failure.",Login_Language,conn) & "<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Mailer.Response & ". " & Translate("Report this error to the Site Administrator.",Login_Language,conn) & "</LI>"
  end if   

end sub

Call Disconnect_SiteWide
%>

