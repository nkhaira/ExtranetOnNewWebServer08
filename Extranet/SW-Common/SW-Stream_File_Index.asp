<%
dim fso, fldr, file, strDownloadDirectory

strDownloadDirectory = Server.mappath("/find-sales/Download/Asset/")
response.write strDownloadDirectory & "<P>"

response.write "<FORM METHOD=POST ACTION=""/SW-Common/SW-Stream_File.asp"">" & vbCrLf
response.write "<SELECT NAME=""FileName"">" & vbCrLf

set fso = server.createobject("scripting.filesystemobject")
set fldr = fso.getfolder(strDownloadDirectory)
for each file in fldr.files
	response.write "<OPTION VALUE=""" & strDownloadDirectory & "\" & file.name & """>" & file.name & "</OPTION>" & vbCrLf
next
set fso = nothing
set fldr = nothing
set file = nothing
response.write "</SELECT>" & vbCrLf
response.write "<INPUT TYPE=SUBMIT VALUE=""download file"">"
response.write "</FORM>" & vbCrLf
%>