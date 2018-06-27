<%@ LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
server.scripttimeout = 120  '  2 Minutes for no uploads

Call Connect_SiteWide

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'Tek %' or Instrument like 'Tektronix %'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"Tek ","Tektronix "))
  newstr = Trim(replace(newstr,"TEK ","Tektronix "))
  newstr = Trim(replace(newstr,"TEK.","Tektronix "))
  newstr = Trim(replace(newstr,"TEKTRONIX ","Tektronix "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'AGILENT %'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"AGILENT ","Agilent "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'BALLANTINE %'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"BALLANTINE ","Ballantine "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing


'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'B&K PRECISION %' or instrument like 'BK Precision %'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"B&K PRECISION ","B&K "))
  newstr = Trim(replace(newstr,"BK Precision ","B&K "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'DANA %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"DANA ","Dana "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'DATA TECH %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"DATA TECH ","Data Tech "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'DATA-PRECISION %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"DATA-PRECISION ","Data-Precision "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'EXTECH %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"EXTECH ","Extech "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'FLUKE %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"FLUKE ","Fluke "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'KEITHLEY %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"KEITHLEY ","Keithley "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'METRAWATT %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"METRAWATT ","Metrawatt "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'PHILIPS %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"PHILIPS ","Philips "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'RACAL DANA %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"RACAL DANA ","Racal Dana "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'TRIPLETTA %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"TRIPLETT ","Triplett "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'VALHALLA %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"VALHALLA ","Valhalla "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'WESTON %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"WESTON ","Weston "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'HEWLETT PACKARD %'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"HEWLETT PACKARD ","HP "))
  newstr = Trim(replace(newstr,"Hewlett Packard ","HP "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like 'LeCroy %'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"Lecroy ","LeCroy "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing


'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like '%: :%'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),": : ",": "))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing


'---

sql = "SELECT * FROM Metcal_Procedures WHERE Instrument like '%ÿ%'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"ÿ","-"))
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing

'---

sql = "SELECT * FROM Metcal_Procedures WHERE Revision like 'V %' or Revision like ' %'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn, 3, 3

do while not rs.EOF

  newstr = Trim(replace(rs("Instrument"),"V ",""))
  newstr = Trim(newstr)
  sqlu = "UPDATE Metcal_Procedures SET Instrument='" & newstr & "' WHERE Procedure_ID=" & rs("Procedure_ID")
  conn.execute sqlu
  rs.MoveNext

loop

rs.close
set rs = nothing



Call Disconnect_SiteWide
%>



