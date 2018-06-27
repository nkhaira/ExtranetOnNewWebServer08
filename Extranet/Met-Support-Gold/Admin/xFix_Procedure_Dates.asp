<%@Language="VBScript" Codepage=65001%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------
' Setup DB Connection
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

SQL = "Select Procedure_ID, Date FROM Metcal_Procedures"
Set rsDates = Server.CreateObject("ADODB.Recordset")
rsDates.Open SQL, conn, 3, 3
                
do while not rsDates.EOF

  MyDate = CDate(rsDates("Date"))
  response.write FormatDateTime(MyDate,"0") & "<BR>"
  
  SQLUpdate = "UPDATE Metcal_Procedures SET Date='" & MyDate & "' WHERE Procedure_ID=" & rsDates("Procedure_ID")
  conn.execute SQLUpdate
    
  rsDates.MoveNext

loop

rsDates.close
set rsDates = nothing

Call Disconnect_SiteWide

%>
