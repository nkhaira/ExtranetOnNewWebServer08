<!--#include virtual="/include/functions_string.asp"-->
<%

if isblank(request("ClientID")) or isblank(request("TrackID")) or isblank(request("CourseID")) or isblank(request("Band")) then

  response.write "Invalid Gateway parameters."

else  

  %>
  <!--#include virtual="/connections/connection_SiteWide.asp"-->
  <%

  Call Connect_Sitewide

  SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Session("Logon_User") & "'"
  Set rsUser = Server.CreateObject("ADODB.Recordset")
  rsUser.Open SQL, conn, 3, 3
  
  if not rsUser.EOF then		

    response.write "<HTML>" & vbCrLf
    response.write "<HEAD>" & vbCrLf
    response.write "<TITLE>Account Verified</TITLE>" & vbCrLf
    response.write "</HEAD>" & vbCrLf
    response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
    response.write "<FORM ACTION=""https://w3app.caliber.com/coursemanager/CaliberStartDefault.asp"" METHOD=""POST"" TARGET=""control"">"

    response.write "<INPUT TYPE=HIDDEN NAME=ClientId VALUE=""" & request("ClientID") & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=TrackId VALUE="""  & request("TrackID")  & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=CourseID VALUE=""" & request("CourseID") & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=Band VALUE="""     & request("Band")     & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=Fname VALUE="""    & rsUser("FirstName") & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=Lname VALUE="""    & rsUser("LastName")  & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=Email VALUE="""    & rsUser("Email")     & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=Login VALUE="""    & rsUser("NTLogin")   & """>"
    response.write "<INPUT TYPE=HIDDEN NAME=Password VALUE=""(secured)"">"

    response.write "</FORM>"
    response.write "</FORM>" & vbCrLf
    response.write "</BODY>" & vbCrLf
    response.write "</HTML>" & vbCrLf

  else
  
    response.write "Invalid User Gateway parameters."
    
  end if
  
  rsUser.close
  set rsUser = nothing
  Call Disconnect_Sitewide   
  
end if
%>

