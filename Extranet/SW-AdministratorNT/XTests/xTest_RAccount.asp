<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

session("Logon_User") = "whitlock-user"

Call Connect_SiteWide

    SQLUser = "SELECT NTLogin, ExpirationDate, Site_ID FROM UserData WHERE UserData.NTLogin='" & Session("Logon_User") & "' AND NewFlag=0"
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQLUser, conn, 3, 3
    
    do while not rsUser.EOF
    
      if CDate(rsUser("ExpirationDate")) < CDate("9/9/9999") then

        ' Get Auto Renew Info from Site
        SQLSite = "SELECT ID, Renew_Days FROM Site WHERE ID=" & rsUser("Site_ID")    
        Set rsSite = Server.CreateObject("ADODB.Recordset")
        rsSite.Open SQLSite, conn, 3, 3
   
        if not rsSite.EOF then
          'If Greater than current expiration Date, then Push Date out.
          response.write rsSite("Renew_Days") & "<BR>"
          response.write rsUser("Expirationdate") & "<BR>"
          
          if CDate(myDate) > CDate(rsUser("ExpirationDate") then
          if CDate(DateAdd("d",CInt(rsSite("Renew_Days")),Date())) > CDate(rsUser("ExpirationDate")) then
            SQLU = "UPDATE UserData SET ExpirationDate='" & CDate(DateAdd("d",CInt(rsSite("Renew_Days")),Date())) & "' " &_
                   "WHERE NTLogin='" & Session("Logon_User") & "' AND Site_ID=" & rsUser("Site_ID")          
            
            response.write SQLU & "<P>"
            'conn.Execute (SQLU)
          else  
          end if
        end if
        rsSite.close
        set rsSite = Nothing
      end if
      
      rsUser.MoveNext
    loop
    
    rsUser.close
    set rsUser = nothing
    
Call Disconnect_SiteWide
%>