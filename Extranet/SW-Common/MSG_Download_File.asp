<%@Language="VBScript" Codepage=65001%>
<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       8/21/2005
' Name:       Met-Support-Gold File Stream (Secure Stream)
' --------------------------------------------------------------------------------------

strFile     = request("filename")
strCategory = request("category")

select case uCase(strCategory)
  case "METCAL"
  	strPath = "/upload/metcal/procedures/"
    strPath = "/met-support-gold/download/metcal/procedures/"
  case "PORTOCAL"
  	strPath = "/upload/portocal/procedures/"
    strPath = "/met-support-gold/download/portocal/procedures/"
  case else
  	response.end
end select

Call DownloadFile(strPath & strFile)

' --------------------------------------------------------------------------------------

function DownloadFile(strFile)
  
  Dim strFileName, iFileSize
  Dim objStream, fso, oFile

  ' Get full path of specified file  
  strFilename = server.MapPath(strFile)  
 
  with response
  
    response.Buffer = True  
    response.Clear
  
    ' Check file exitsts
      
    Set fso = Server.CreateObject("Scripting.FileSystemObject")  
    if not fso.FileExists(strFilename) then  
  	  .write("<H1>Error:</H1> " & strFilename & " cannot be found.<P>")  
      .flush
  	  .end
    else
      strFile = Mid(strFileName, InstrRev(strFileName, "\") + 1)
    end if  
 
    ' Open Stream
    
    adTypeBinary    = 1  
  
    Set objStream   = Server.CreateObject("ADODB.Stream")  
  
    objStream.Open  
    objStream.Type  = adTypeBinary
    objStream.LoadFromFile(strFilename)    
  
    intFileSize = objStream.size
  
    ' Send Headers to the users browser
    
    .AddHeader "Content-Disposition", "attachment; filename=" & strFile  
    .AddHeader "Content-Length", intFileSize  
    .CharSet = "UTF-8"  
    .ContentType = "application/octet-stream"
    .CacheControl = "no-store"
    .flush
 
    chunk = 2048

    ' Send File to User via Chunk Method
  
    for i = 1 to intFileSize \ chunk 
      if not .IsClientConnected then exit for 
      .BinaryWrite objStream.Read(chunk)
      .flush
    next 
 
    if intFileSize mod chunk > 0 Then 
      if .IsClientConnected then 
        .BinaryWrite objStream.Read(intFileSize Mod chunk)
        .flush
      end if 
    end if 
    
  end with  
 
  ' Clean up  
  objStream.Close  
  Set objStream = nothing  
 
end function  
%>  
 
