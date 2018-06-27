<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 2 * 60

Call Connect_SiteWide

SQL = "SELECT Procedure_ID, Description " &_
      "FROM   dbo.Metcal_Procedures " &_
      "ORDER BY Procedure_ID"

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  'response.write rsID("Procedure_ID") & "<BR>"
  'response.flush

  Description = rsID("Description")
  
  if not isblank(Description) then
  
    Description = Replace(Description,"<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">","")
    Description = Replace(Description,"<<","<")    
    Description = Replace(Description,">>",">")
    Description = Replace(Description,"<html>","")
    Description = Replace(Description,"<HTML>","")
    Description = Replace(Description,"<head>","")
    Description = Replace(Description,"<HEAD>","")
    Description = Replace(Description,"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">","")
    Description = Replace(Description,"<meta name=""generator"" content=""NoteTab Pro"">","")
    Description = Replace(Description,"<title>*** Your Title Here ***</title>","")
    Description = Replace(Description,"<TITLE>*** Your Title Here ***</TITLE>","")
    Description = Replace(Description,"</head>","")
    Description = Replace(Description,"</HEAD>","")
    Description = Replace(Description,"<body>","")
    Description = Replace(Description,"<BODY>","")
    Description = Replace(Description,"</body>","")
    Description = Replace(Description,"</BODY>","")
    Description = Replace(Description,"</html>","")
    Description = Replace(Description,"</HTML>","")
    Description = Replace(Description," ","")
    Description = Replace(Description,Chr(10),"")
    Description = Replace(Description,Chr(13),"")
    Description = Replace(Description,"<p>","<P>")
    Description = Replace(Description,"</p>","</P>")
    Description = Replace(Description,"<br>","<BR>")
    Description = Replace(Description,"<BR><BR>","<P>")
    Description = Replace(Description,"<P>","<P>"   + Chr(13) + Chr(10))
    Description = Replace(Description,"</P>","</P>" + Chr(13) + Chr(10))
    Description = Replace(Description,"<BR>","<BR>" + Chr(13) + Chr(10))
    Description = Replace(Description,"'","''")
    Description = Replace(Description,"'","''")    
    Description = Replace(Description,"&nbsp;"," ")
    Description = Replace(Description,"&NBSP;"," ")
    Description = Replace(Description,Chr(9),"")
    
    NewDescription = ""
    for i = 1 to len(Description)
      char = asc(mid(Description,i,1))
      if char > 9 and char < 255 then
        NewDescription = NewDescription & chr(char)
      end if
    next
    
    SQL = "UPDATE dbo.Metcal_Procedures SET Description='" & Mid(NewDescription,1,6000) & "' WHERE Procedure_ID=" & rsID("Procedure_ID")
    conn.execute SQL
    
  end if  

  rsID.MoveNext

loop

rsID.close
set rsID = nothing

Call Disconnect_SiteWide
%>