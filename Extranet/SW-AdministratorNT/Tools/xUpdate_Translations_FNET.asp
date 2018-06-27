<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' DO NOT RUN w/o Modifications

' Author: Kelly Whitlock

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 2 * 60

Call Connect_SiteWide

SQL = "SELECT  Title, Description, OracleID " &_
      "FROM dbo.Calendar_FNET_Translations " &_
      "WHERE (OracleID IS NOT NULL)"

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  if not isnull(rsID("Title")) then
  
    SQLU = "UPDATE dbo.Calendar SET Title=N'" & replace(rsID("Title"),"'","''") & "'"
    
    if not isnull(rsID("Description")) then 
    
      SQLU = SQLU & ", Description=N'" & replace(rsID("Description"),"'","''") & "'"
    
    end if
    
    SQLU = SQLU & ",Status=1"
  
    SQLU = SQLU & " WHERE Item_Number='" & rsID("OracleID") & "' AND Site_ID=82 AND Language <> 'eng'"
    
    response.write SQLU & "<P>"
    
    conn.execute SQLU
  
  end if  

  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>