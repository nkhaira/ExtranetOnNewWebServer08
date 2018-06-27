<%  

strFile = Request("filename")
strCategory = Request("category")

select case uCase(strCategory)
  case "METCAL"
	strPath = "/upload/metcal/procedures/"
  strPath = "/met-support-gold/download/metcal/procedures/"
  case else
	response.end
end select

downloadFile(strPath & strFile)

Function downloadFile(strFile)  
  Dim strFileName, iFilelength
  Dim oStream, fso, oFile

  ' get full path of specified file  
  strFilename = server.MapPath(strFile)  
 
  ' clear the buffer  
  Response.Buffer = True  
  Response.Clear  
 
  ' create stream  
  Set oStream = Server.CreateObject("ADODB.Stream")  
  oStream.Open  
 
  ' Set as binary  
  oStream.Type = 1  
 
  ' load in the file  
  on error resume next  
 
 
  ' check the file exists  
  Set fso = Server.CreateObject("Scripting.FileSystemObject")  
  if not fso.FileExists(strFilename) then  
	Response.Write("<h1>Error:</h1>" & strFilename & " does not exist<p>")  
	Response.End  
  end if  
 
 
  ' get length of file  
  Set oFile = fso.GetFile(strFilename)  
  iFilelength = oFile.size  
 
  oStream.LoadFromFile(strFilename)  
  if err then  
	Response.Write("<h1>Error: </h1>" & err.Description & "<p>")  
	response.End  
  end if  

  ' send the headers to the users browser  
  Response.AddHeader "Content-Disposition", "attachment; filename=" & oFile.name  
  Response.AddHeader "Content-Length", iFilelength  
  Response.CharSet = "UTF-8"  
  Response.ContentType = "application/octet-stream"  
 
  ' output the file to the browser  
  Response.BinaryWrite oStream.Read  
  Response.Flush  
 
 
  ' Clean up  
  oStream.Close  
  Set oStream = Nothing  
 
End Function  
 
%>  
 
