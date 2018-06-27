<%@ Language=VBScript %>
<%

Dim FileURL
  FileURL = "http://support.dev.fluke.com/find-sales/download/asset/190c_dm.pdf"
'  FileURL = "http://support.dev.fluke.com/find-sales/download/thumbnail/87Bhandfloor.jpg"
'  FileURL = "http://support.dev.fluke.com/find-sales/download/asset/00_09_erep.xls"
'  FileURL = "http://support.dev.fluke.com/find-sales/download/asset/sweepstakes_webbnn.gif"  
'  FileURL = "http://support.dev.fluke.com/find-sales/download/asset/technical_training_scm_families.ppt"    
'  FileURL = "http://support.dev.fluke.com/find-sales/download/asset/scopemeter_newsletter_58.zip"    
'  FileURL = "http://support.fluke.com/find-sales/download/asset/87-retrofit_catalog_copy.doc"
  

DownloadFlag = false
  
Success = StreamFile(FileURL, DownloadFlag)

function StreamFile(FileURL, DownloadFlag)

  response.buffer = true
  
  Dim FileName, FileType, myStream

  FileName = LCase(mid(FileURL,  instrrev(FileURL,  "/") + 1))
  FileType = LCase(mid(FileName, instrrev(FileName, ".") + 1))
  
  select case FileType
    case "pdf"
      response.ContentType = "application/pdf"
    case "jpeg", "jpg"
      response.ContentType = "image/jpeg"
    case "gif"
      response.ContentType = "image/gif"
    case "xls"
      response.ContentType = "application/vnd.ms-excel"
    case "doc", "dot"
      response.ContentType = "application/msword"
    case "ppt", "pps", "pot"
      response.ContentType = "application/vnd.ms-powerpoint"
    case "zip"
      response.ContentType = "application/x-zip-compressed"
      response.AddHeader "content-disposition","attachment;filename=" & FileName
      DownLoadFlag = false
    case "txt"
      response.ContentType = "text/plain"
    case "asp", "aspx", "html", "htm"
      response.ContentType = "text/html"
  end select
  
'  response.CacheControl = "no-store"
'  response.CacheControl = "no-cache"

  
  if CInt(DownloadFlag) = CInt(true) then
    Response.AddHeader "content-disposition","attachment;filename=" & FileName
'  else  
'    Response.AddHeader "content-disposition","inline;filename=" & FileName    
  end if
  
  set myStream = CreateObject("MSXML2.ServerXMLHTTP")
  
  myStream.open "GET", FileURL, false 
  myStream.setTimeouts 10000,10000,10000,10000 
  myStream.send ()
  
  response.BinaryWrite myStream.responseBody
  
  set myStream = nothing

  StreamFile = True
  
end function
%>
