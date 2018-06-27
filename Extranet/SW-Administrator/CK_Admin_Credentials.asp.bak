<%

Admin_Access      = 0 ' No Access
Multiple_Accounts = 0

Dim wtbadm
wtbadm = false

'response.write request.querystring & "<P>"

if not isblank(Session("LOGON_USER")) then
  Admin_Name = Session("LOGON_USER")
elseif not isblank(Request.QueryString("LOGON_USER")) then
  Admin_Name = Request.QueryString("LOGON_USER")
elseif not isblank(Request.form("LOGON_USER")) then
  Admin_Name = Request.form("LOGON_USER")
else
  Admin_Name = ""
end if

if not isblank(Session("Password")) then
  Admin_Password = Session("Password")
elseif not isblank(Request.QueryString("LOGON_Password")) then
  Admin_Password = Request.QueryString("LOGON_Password")
elseif not isblank(Request.QueryString("Password")) then
  Admin_Password = Request.QueryString("Password")
elseif not isblank(Request.form("LOGON_Password")) then
  Admin_Password = Request.form("LOGON_Password")
elseif not isblank(Request.form("Password")) then
  Admin_Password = Request.form("Password")
else
  Admin_Password = ""
end if

if not isblank(Request.QueryString("Site_ID")) and isnumeric(Request.QueryString("Site_ID")) then
  Site_ID = Request.QueryString("Site_ID")
elseif not isblank(Request.form("Site_ID")) and isnumeric(Request.form("Site_ID")) then
  Site_ID = Request.form("Site_ID")
elseif not isblank(Session("Site_ID")) and isnumeric(Session("Site_ID")) then
  Site_ID = Session("Site_ID")
elseif isblank(Site_ID) then
  Site_ID = 100  
end if

if Instr(Admin_Name, "\") > 0 then %><%
  Admin_Name = Mid(Admin_Name, InStrRev(Admin_Name, "\") + 1) '"
end if

'response.write Admin_Access & "<BR>"
'response.write Site_ID & "<BR>"
'response.write Admin_Name & "<BR>"
'response.write Admin_Password & "<BR>"
'response.flush
'response.end

' Check for Multiple Accounts if Site_ID is not known

if Site_ID = 100 or Site_ID = 101 then

  SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Admin_Name & "' AND UserData.Password='" & Admin_Password & "' AND UserData.NewFlag=" & CInt(False)
  Set rsLogin = Server.CreateObject("ADODB.Recordset")
  rsLogin.Open SQL, conn, 3, 3

  Multiple_Accounts = rsLogin.RecordCount

  do while not rsLogin.EOF

    if (instr(1,LCase(rsLogin("SubGroups")),LCase("domain"))         > 0 _
      or instr(1,LCase(rsLogin("SubGroups")),LCase("administrator")) > 0 _
      or instr(1,LCase(rsLogin("SubGroups")),LCase("account"))       > 0 _
      or instr(1,LCase(rsLogin("SubGroups")),LCase("content"))       > 0 _
      or instr(1,LCase(rsLogin("SubGroups")),LCase("submitter"))     > 0 _
      or instr(1,LCase(rsLogin("SubGroups")),LCase("metrics"))       > 0 _
      or instr(1,LCase(rsLogin("SubGroups")),LCase("literature"))    > 0 _            
      or instr(1,LCase(rsLogin("SubGroups")),LCase("forum"))         > 0) then
      
      if instr(1,LCase(rsLogin("SubGroups")),LCase("wtbadm")) > 0 then
        wtbadm = true
      end if
      
      Admin_Access = 1      ' At least one Admin account exists
      
      Admin_ID             = rsLogin("ID")
      Admin_FirstName      = rsLogin("FirstName")
      Admin_MiddleName     = rsLogin("MiddleName")
      Admin_LastName       = rsLogin("LastName")
      Admin_Company        = rsLogin("Company")
      Admin_EMail          = rsLogin("EMail")
      
      exit do 
    
    end if
    
    rsLogin.MoveNext
    
  loop  

  rsLogin.close
  set rsLogin = nothing

else
  
  SQL = "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Admin_Name & "' AND UserData.Password='" & Admin_Password & "' AND UserData.Site_ID=" & Site_ID & " AND UserData.NewFlag=" & CInt(False)
  Set rsLogin = Server.CreateObject("ADODB.Recordset")
  rsLogin.Open SQL, conn, 3, 3
 
  if not rsLogin.EOF then
  
    if instr(1,LCase(rsLogin("SubGroups")),"domain") > 0 then
      Admin_Access = 9
    elseif instr(LCase(rsLogin("SubGroups")),"administrator") > 0 then
      Admin_Access = 8
    elseif instr(LCase(rsLogin("SubGroups")),"account") > 0 then
      Admin_Access = 6
    elseif instr(LCase(rsLogin("SubGroups")),"content") > 0 then
      Admin_Access = 4              
    elseif instr(LCase(rsLogin("SubGroups")),"literature") > 0 then
      Admin_Access = 3              
    elseif instr(LCase(rsLogin("SubGroups")),"submitter") > 0 then
      Admin_Access = 2
    elseif instr(LCase(rsLogin("SubGroups")),"metrics") > 0 then
      Admin_Access = 1
    end if
     
    if Admin_Access > 0 then
      
      Admin_ID             = rsLogin("ID")
      Admin_Site_ID        = rsLogin("Site_ID")
      Admin_Password       = rsLogin("Password")
      Admin_FirstName      = rsLogin("FirstName")
      Admin_MiddleName     = rsLogin("MiddleName")
      Admin_LastName       = rsLogin("LastName")
      Admin_Company        = rsLogin("Company")
      Admin_Country        = rsLogin("Business_Country")
      Admin_EMail          = rsLogin("EMail")
      Admin_Region         = rsLogin("Region")
      Admin_Account_Region = rsLogin("Account_Region")

      if rsLogin("RTE_Enabled") = CInt(True) then
        Admin_RTE_Enabled  = true
      else  
        Admin_RTE_Enabled  = false
      end if
      
      if instr(1,LCase(rsLogin("SubGroups")),LCase("wtbadm")) > 0 then
        wtbadm = true
      end if
          
      if isblank(Session("Language")) and isblank(Request.QueryString("Language")) and isblank(Request.form("Language")) then
        Login_Language       = rsLogin("Language")
        Session("Language")  = Login_Language
      end if
           
    end if
         
  end if

  rsLogin.close
  set rsLogin = nothing
 
end if

'response.write Admin_Access & "<BR>"
'response.write Site_ID & "<BR>"
'response.write Admin_Name & "<BR>"
'response.write Admin_Password & "<BR>"
'response.flush
'response.end

if Admin_Access = 0 or isblank(Admin_Name) or isblank(Site_ID) then
  response.redirect "/register/login.asp?Site_ID=100"
end if

Session("LOGON_USER") = Admin_Name
Session("Password")   = Admin_Password
Session("Site_ID")    = Site_ID
%>
