<%
  
  Sub connect()
  	strConnectionString = "SW-QUIZ"
  	set DataConn = Server.CreateObject("ADODB.Connection")
  	DataConn.open strConnectionString
  End Sub
  
  Sub disconnect()
  	DataConn.Close
  	set DataConn = nothing
  End Sub
 
  
%>

