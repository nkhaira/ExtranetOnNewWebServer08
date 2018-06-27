<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
Call Connect_Sitewide

      SQL = "SELECT ID, SubGroups " &_
            "FROM UserData " &_
            "WHERE SubGroups like '%,,%'"
   
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sql, conn, 3, 3
      
      do while not rs.EOF
      
        SQLU = "UPDATE UserData SET SubGroups='" & replace(rs("SubGroups"),",,",",") & "' WHERE ID=" & rs("ID")
        conn.execute SQLU
      
        rs.MoveNext
      
      loop

      rs.Close
      set rs = nothing
      
      Call Disconnect_Sitewide
      
      response.write "Done"
%>
