<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<HTML>
<BODY>
<%
if request.form("userid") <> "" then
    ' is is valid login - if so redirect to welcome page
    Connect_SiteWide
    sql = "select count(*) as mycnt from PomonaCustomers" & vbcrlf &_
        "where login = " & trim(request.form("userid")) & vbcrlf &_
        "and password = " & trim(request.form("password"))
    set dbRS = conn.Execute(sql)
    
    if dbRS("mycnt") > 0 then
        set dbRS = nothing
        Disconnect_Sitewide
        Write_welcome_form
    else
        Response.write "Invalid Login - Password combo<BR>" & vbcrlf
        Write_login_form
    end if
else 
    Write_login_form
end if

%>
</form>
</body>
</html>
<%
' ------------------------------- end of main -------------------------------
sub Write_login_form
%>
<form name="launch" method="POST">
<Table>
    <TR><TD>Login:</td>
        <TD><input type="test" name="userid" size=10></td>
    </tr>
    <TR><TD>Password:</td>
        <TD><input type="test" name="password" size=10></td>
    </tr>
</table>
<input type="submit" value="Login">
<%
end sub

sub Write_welcome_form
%>
Pomona Electronics is updgrading their extranet.  Now would be a good time to pre-register.
<form name="launch" method="POST" action="http://support.fluke.com/pomona-sales/welcome.asp">
<input type="hidden" name="puserid" value="<%=trim(request.form("userid"))%>">
<input type="submit" value="Register Now">
<%
end sub
%>


