<%
	on error resume next

	Dim oHTTPComm
	Dim strData
	Dim strRecData
	Dim lRecDataLen
	Dim bResponse
				
	Set oHTTPComm = Server.CreateObject("VBHTTPComm.cVBHTTPComm")
	with oHTTPComm
		.TransferMethod = 1
		.RemotePort = 80

' ---------------- For testing on Fluke servers ----------------
		.RemoteHostIP = "129.196.132.58"
		.LocalHostIP = "129.196.132.59"
		.HostName = "www.dev.fluke.com"
		.TargetFile = "/syncforce/response.asp"
		.RefererFile = "http://129.196.132.59/syncforce/default.asp"
		
' ---------------- Internal Testing ----------------------------
'		.RemoteHostIP = "216.9.4.31"
'		.LocalHostIP = "192.168.0.5"
'		.HostName = "216.9.4.31"
'		.TargetFile = "/response.asp"
'		.RefererFile = "http://192.168.0.5:81/default.asp"
'---------------------------------------------------------------

		response.write("OpenConnection: " & .OpenConnection & "<BR>")
		response.write("<BR>HTTP_Err: " & .GetErrorDescription & "<BR>")

		.AddData "foo_1", "foo_1:foo_1_test"
		.AddData "foo_2", "foo_2:foo_2_test"
		.AddData "foo_3", "foo_3:foo_3_test"
		.AddData "foo_4", "foo_4:foo_4_test"
		.AddData "foo_5", "foo_5:foo_5_test"    
		response.write("<BR>SendData: " & .SendData & "<BR><BR>")
		
		response.write("<BR>HTTP_Err: " & .GetErrorDescription & "<BR>")

		bResponse = true
		do while .IsDataAvailable and bResponse
			'strRecData = .ReceiveData(cStr(strRecData), cLng(iRecDataLen))
			bResponse = .ReceiveData(strRecData, iRecDataLen)
			response.write("<BR>bResponse: " & bResponse & "<BR>")
			response.write("<BR>strRecData: " & strRecData & "<BR><BR>")
			response.write("Length: " & iRecDataLen & "<BR>")
		loop
		response.write("<BR>HTTP_Err: " & .GetErrorDescription & "<BR>")
		response.write("<BR>Err: " & err.description & "<BR>")
	end with
%>
