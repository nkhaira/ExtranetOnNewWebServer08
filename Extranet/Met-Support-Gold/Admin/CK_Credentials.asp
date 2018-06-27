<%

SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Admin_Name & "' AND UserData.Password='" & Admin_Password & "' AND UserData.NewFlag=" & CInt(False)
Set rsLogin = Server.CreateObject("ADODB.Recordset")
rsLogin.Open SQL, conn, 3, 3

bIsMetcalAdmin = false

do while not rsLogin.EOF
    if (instr(1,LCase(rsLogin("Groups")),LCase("met-support-gold")) > 0) then
      bIsMetcalAdmin = true
      exit do 
    end if
    
    rsLogin.MoveNext
loop  

'response.write("admin_access: " & admin_access & "<BR>")
'response.write("bIsMetcalAdmin: " & bIsMetcalAdmin & "<BR>")

'response.end

if not (admin_access <> 1 or admin_access <> 8 or admin_access <> 9)  or not bIsMetcalAdmin then
   Session("ErrorString") = "<LI>" & Translate("You do not have sufficient permissions to view this page.",Login_Language,conn)
   response.redirect "/register/default.asp"
end if

%>