<%@Language="VBScript" Codepage=65001%>

<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
response.Buffer = true

Dim strFileToDownload, strFileContents, bShowDialog

bShowDialog = true

FileName = "/find-sales/download/asset/9030111_ENG_A_W.PDF"

'strFileToDownload = Request.QueryString("FileName")   ' Mapped Path
strFileToDownload = Server.Mappath(FileName)   ' Mapped Path

strFileContents   = FileStream(strFileToDownload,bShowDialog)

function FileStream(strFileName, bShowDialog)

  adTypeBinary = 1
  
  Set objStream = Server.CreateObject("ADODB.Stream")
  objStream.Open
  objStream.Type = adTypeBinary
  objStream.LoadFromFile strFileName
  
  intFileSize = objStream.Size
  
  strFile = Mid(strFileName, InstrRev(strFileName, "\") + 1) %><%
  
  ' Output File Name
  
  strFile   = Replace(strFile," ","_")  ' Convert spaces to underscores
  FileOrig  = strFileName
  FileRoot  = LCase(Mid(strFileName, 1, InstrRev(strFileName, ".") - 1))
  FileExtn  = LCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))
  
  ' Get Content Type for this file
  
  Call Connect_SiteWide
  
  SQL = "SELECT ContentType FROM Asset_Type WHERE File_Extension='" & FileExtn & "'"
  Set rsType = Server.CreateObject("ADODB.Recordset")
  rsType.Open SQL, conn, 3, 3
  
  if not rsType.EOF then
    ContentType = rsType("ContentType")
  else
    ContentType = "application/octet-stream"
  end if
  
  rsType.close
  set rsType = nothing
  
  Call Disconnect_SiteWide
  
  ' Stream the File
  
  with response
  
    if CInt(bShowDialog) = CInt(True) then
      .AddHeader "Content-Disposition", "attachment; filename=" & strFile
    end if

    .AddHeader     "Content-Length", intFileSize  
    .Charset     = "UTF-8"
    .ContentType = ContentType
    .BinaryWrite   objStream.Read
    .Flush

  end with
  
  objStream.Close
  Set objStream = Nothing

end function
%>