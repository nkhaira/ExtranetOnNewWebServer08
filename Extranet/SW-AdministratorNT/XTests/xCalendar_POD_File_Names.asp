<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Session.timeout = 240 ' Set to 4 Hours
Server.ScriptTimeout = 2 * 60

Call Connect_SiteWide

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

  if UCase(rsID("File_Name_POD")) <> POD_Name then
  
    From_File = MyServerPath & "\" & replace(rsID("File_Name_POD"),"/","\")
    To_File   = MyServerPath & "\" & replace(POD_Name,"/","\")
  
    If objFSO.FileExists(From_File) Then
      objFSO.MoveFile From_File , To_File
    end if
  
    SQL = "UPDATE Calendar SET File_Name_POD='" & POD_NAME & "'" & " WHERE ID=" & rsID("ID")
    conn.execute SQL

  end if  

  rsID.MoveNext

loop

rsID.close
set rsID   = nothing
Set objFSO = nothing

conn.execute("exec lit_status_build",,adCmdText)

Call Disconnect_SiteWide

response.write "Done"
%>