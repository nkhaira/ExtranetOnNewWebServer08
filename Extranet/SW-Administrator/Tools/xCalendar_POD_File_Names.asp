<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' The purpose of this script is to rename both the physical file name and DB file name
' per the naming convention for Print on Demand Files.

Session.timeout      = 240      ' Set to 4 Hours
Server.ScriptTimeout = 60 * 10  ' 10-Minutes

Dim DoWork, DoWhat

if isblank(request("DoWhat")) then
  response.write "What do you want to do?  List or Rename"
  response.end
elseif UCase(request("DoWhat")) = "LIST" then
  DoWork = false
elseif UCase(request("DoWhat")) = "RENAME" then
  DoWork = true
end if

Call Connect_SiteWide

Dim FileNotFoundCount
Dim FileOKCount
Dim FileRenamedCount

FileNotFoundCount = 0
FileOKCount       = 0
FileRenamedCount  = 0

SQL = "SELECT ID, Item_Number, Cost_Center, [Language], Revision_Code, File_Name_POD " &_
      "FROM dbo.Calendar " &_
      "WHERE (Item_Number IS NOT NULL) AND (LEN(Item_Number) = 7) AND (File_Name_POD IS NOT NULL) " 


Set objFSO = CreateObject("Scripting.FileSystemObject")
MyServerPath = "/"
MyServerPath = Server.MapPath(MyServerPath)

Set rsID = Server.CreateObject("ADODB.Recordset")
rsID.Open SQL, conn, 3, 3

do while not rsID.EOF

  if isblank(rsID("Cost_Center")) then
    CC = "0000"
  elseif rsID("Cost_Center") = 0 then
    CC = "0000"
  elseif len(rsID("Cost_Center")) <> 4 then
    CC = "0000"  
  else
    CC = rsID("Cost_Center")
  end if

  POD_Name = "POD/" & UCase(rsID("Item_Number") & "_" & CC & "_" & rsID("Language") & "_" & rsID("Revision_Code") & "_P.PDF")

  if UCase(rsID("File_Name_POD")) <> UCase(POD_Name) then
  
    From_File = Server.MapPath("/" & rsID("File_Name_POD"))
    To_File   = Server.MapPath("/" & Pod_Name)
    
    response.write "FF: " & From_File & "<BR>"
    response.write "TF: " & To_File & "<P>"
    
    If objFSO.FileExists(From_File) Then

      if DoWork then
      
        if objFSO.FileExists(To_File) then
          objFSO.MoveFile To_File , To_File & ".BAK"
        end if
        objFSO.MoveFile From_File , To_File
    
        SQL = "UPDATE Calendar SET File_Name_POD='" & POD_NAME & "'" & " WHERE ID=" & rsID("ID")
        conn.execute SQL
      end if  

      response.write "<FONT COLOR=GREEN>R</FONT>: " & POD_Name & "<P>"
      FileRenamedCount  = FileRenamedCount + 1
    
    else
    
      response.write "<FONT COLOR=RED>X</FONT>: " & rsID("Item_Number") & " " & rsID("File_Name_POD") & "<P>"
      FileNotFoundCount = FileNotFoundCount + 1 
    end if
  
  else
  
    From_File = Server.MapPath("/" & rsID("File_Name_POD"))
        
    If objFSO.FileExists(From_File) Then
      'response.write "OK: " & rsID("File_Name_POD") & "<BR>"
      FileOKCount = FileOKCount + 1
    else
      response.write "<FONT COLOR=ORANGE>NF</FONT>: " & rsID("Item_Number") & " " & rsID("File_Name_POD") & "<P>"
      FileNotFoundCount = FileNotFoundCount + 1
    end if
      
  end if  

  response.flush
  rsID.MoveNext

loop

rsID.close
set rsID   = nothing
Set objFSO = nothing

response.write "<P>Done<P>"
response.write "FileNotFoundCount: " & FileNotFoundCount & "<BR>"
response.write "FileOKCount: " & FileOKCount & "<BR>"
response.write "FileRenamedCount: " & FileRenamedCount & "<BR>"

response.write "<P>Remember to run Lit_Status_Build"

Call Disconnect_SiteWide
%>