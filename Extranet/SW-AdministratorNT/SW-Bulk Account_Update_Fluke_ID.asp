 <!-- Begin ASP Source Code -->
  <%@ LANGUAGE="VBSCRIPT" %>
  
  <%    
        Dim Sql, X, Sqlu
        Dim objRS
        Dim objCMD
        Dim Conn, Conn_prd
        Dim strConnectionString_SiteWide
        Dim strConnectionString_SiteWide_PRD
        
        'Use this connection to retrieve data from Raytek_Accounts table
        'which is present only on the DEV database
        strConnectionString_SiteWide = "DRIVER={SQL Server};SERVER=Evtibg18" &_
																				";UID=SITEWIDE_WEB" &_
																				";DATABASE=FLUKE_SITEWIDE" &_
																				";pwd=tuggy_boy"
																				
        'Use this connection to update records to different servers (DEV\TEST\PRD)
        strConnectionString_SiteWide_PRD = "DRIVER={SQL Server};SERVER=FLKTST18.data.ib.fluke.com" &_
																				";UID=SITEWIDE_WEB" &_
																				";DATABASE=FLUKE_SITEWIDE" &_
																				";pwd=tuggy_boy"
        
        Set Conn = Server.CreateObject("ADODB.Connection")
        conn.ConnectionTimeout = 30
				conn.Open strConnectionString_SiteWide
				
				Set Conn_prd = Server.CreateObject("ADODB.Connection")
        conn_prd.ConnectionTimeout = 30
				conn_prd.Open strConnectionString_SiteWide_PRD
        
        Set objRS = Server.CreateObject("ADODB.Recordset")
        objRS.ActiveConnection = Conn
        objRS.CursorType = 3                    'Static cursor.
        objRS.LockType = 2                      'Pessimistic Lock.
        objRS.Source = "Select * from Raytek_Accounts"
        objRS.Open
   %>
  
   <%
      Response.Write("Original Data")

      'Printing out original spreadsheet headings and values.

      'Note that the first recordset does not have a "value" property
      'just a "name" property.  This will spit out the column headings.

      Response.Write("<TABLE><TR>")
      For X = 0 To objRS.Fields.Count - 1
         Response.Write("<TD>" & objRS.Fields.Item(X).Name & "</TD>")
      Next
      Response.Write("</TR>")
      objRS.MoveFirst

      While Not objRS.EOF
         Response.Write("<TR>")
         For X = 0 To objRS.Fields.Count - 1
            Response.write("<TD>" & objRS.Fields.Item(X).Value)
         Next
         sqlu = "UPDATE Userdata SET Business_System='" & objRS("Business System").Value & "'" & vbcrlf
         sqlu = sqlu & ", Fluke_ID='" &  objRS("Fluke Customer ID").Value & "'" & vbcrlf
         sqlu = sqlu & " WHERE ID=" & objRS("Account ID").Value
				 response.write "updated"
				 conn_prd.execute sqlu
         objRS.MoveNext
         Response.Write("</TR>")
      Wend
      Response.Write("</TABLE>")

      

      'ADO Object clean up.

      objRS.Close
      Set objRS = Nothing

      conn.Close
      Set conn = Nothing
      
      conn_prd.Close
      Set conn_prd = Nothing
   %>
   <!-- End ASP Source Code -->