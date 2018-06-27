<%

Response.Write("Test FTP") 

'set obj = server.CreateObject("FTPclient.FTP") 
'obj.Hostname="Fluke.ingest.cdn.level3.net"
'obj.Username="fluke"
'obj.Password="KaxETa6vsg"

'obj.UploadFile "D:\Extranet\SW-Administrator\Test_Asset2.PDF","FNet/Download/Asset/Test_Asset123.PDF"

set obj = server.CreateObject("FTPclient.FTP") 
obj.Hostname="ftp.fluke.com/"
obj.Username="flukecontent"
obj.Password="c0nt3nt4F1uke"

obj.UploadFile "D:\IIS\websites\Extranet\SW-Administrator\Test_Asset2.PDF","content/fluke/Test/Fluke/Download/Asset/Test_Asset123.PDF"

if err.number <> 0 then 
Response.Write err.Description
end if
Response.Write "uploaded"

'ftp_address          = "ftp://Fluke.ingest.cdn.level3.net"
'	ftp_username         = "fluke"
'	ftp_password         = "KaxETa6vsg"
%>