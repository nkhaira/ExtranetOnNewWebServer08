<%@ Language=VBScript %>
<% Option Explicit %>

<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->

<%
'-----------------------------------------------------------------------
' FileUpEE Generic FileServer Auto-Process ASP Script
'
' This file server script will auto process any request submitted
' to it by a FileUpEE web server script (/SW-Administrator/Calendar_Admin.asp)
'
' The file is uploaded to the web server from the client and then forwarded
' to the FileServer via SOAP to be saved
'
' fileserver.asp -- This script goes on the fileserver (Original Script Example)
' FileUpEE_FileServer.asp -- this script
'
' Author: Kelly Whitlock
' Date:   10/25/2005
'-----------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' Declarations
Dim Debug_Flag
Debug_Flag = false

Dim strfile

Dim FileUpEE_TempPath
FileUpEE_TempPath = "c:\temp"

' Instantiate the FileUp object
Dim oFileUpEE
Set oFileUpEE = Server.CreateObject("SoftArtisans.FileUpEE")

' Set the TransferStage to the appropriate value
oFileUpEE.TransferStage = saFileServer
oFileUpEE.DynamicAdjustScriptTimeout(saFileServer) = true
oFileUpEE.TempStorageLocation(saFileServer) = FileUpEE_TempPath

if Debug_Flag then
  oFileUpEE.AuditLogDestination(saFileServer) = saLogFile
  oFileUpEE.AuditLogFile(saFileServer) = "D:\Inetpub\extranet\SW-FileUp_Log\FS_Audit_Log.txt"

  oFileUpEE.DebugLevel = 3
  oFileUpEE.DebugLogFile = "D:\Inetpub\extranet\SW-FileUp_Log\FS-Debug_Log.txt"
end if  

' Handle errors gracefully for the remainder of the script

On Error Resume Next

'-----------------------------------------------------------------------	
' Call ProcessRequest to Read the Submitted Upload from WebServer
' Parameters for ProcessRequest:
'	1) Request - ASP Request object
'	2) SOAP Request? - True, the web server sends SOAP requests to the fileserver
'	3) Auto Process? - True, we want to auto-process the request from the web server
'    Note: "SOAP Request" will typically be True in the file server script
'-----------------------------------------------------------------------  

oFileUpEE.ProcessRequest request, true, true

if err.Number <> 0 Then
  response.write "<B>FileServer ProcessRequest Error</B><BR>" & err.Number & ": " & err.Description & " (" & err.Source & ")"
  repsonse.Status = 500
  response.end
end if

on error goto 0

' Instantiate the Archive and FileSystem Object

Dim oArch
set oArch = Server.CreateObject("SoftArtisans.Archive")
Dim FileObj
set FileObj = Server.CreateObject("Scripting.FileSystemObject")			

Dim Archive_Log_Path, ArchiveNameNew, ArchivePath, ArchiveSize, ArchiveFile, ArchiveSQL
Dim FileTemp

Dim oFile

Archive_Log_Path = "/SW-FileUp_Log/Create_Archive_File.Log"
Archive_Log_Path = server.mappath(Archive_Log_Path)
ArchiveSize      = 0

if Debug_Flag then
  Set FileTemp = FileObj.CreateTextFile(Archive_Log_Path, True)
  FileTemp.writeline("Begin Create Archive Log" & " " & Now())
end if

' Loop through Files to find File to Archive

for each oFile in oFileUpEE.Files

  if Debug_Flag then
    FileTemp.writeline(oFile.Name)
  end if
  
  ' Create Archive File of File_Name Only and skip those that are already ZIP or EXE types
  
  if oFile.size <> 0 and instr(LCase(oFile.Name),".zip") = 0 and instr(LCase(oFile.Name),".exe") = 0 then
    
    if instr(LCase(oFile.DestinationPath),"\asset\") > 0 then ' LCase(oFile.Name) = "file_name" then

		  ArchiveNameNew = Left(oFile.DestinationPath, InStrRev(oFile.DestinationPath, "."))
      ArchiveNameNew = UCase(replace(LCase(ArchiveNameNew),"\asset\","\archive\") & "ZIP")

      ' Delete old Archive if it exists since it is being replaced
 			if FileObj.FileExists(ArchiveNameNew) then
    		FileObj.DeleteFile ArchiveNameNew, true
 			end if
      
      ' Create Archive
  		oArch.ArchiveType = 1 ' Zip File
	  	oArch.CreateArchive(ArchiveNameNew)
      oArch.Addfile oFile.DestinationPath, false    
  		oArch.CloseArchive()
      
    	'-----------------------------------------------------------------------  
      ' We need to check the new archive file size here and store it since the Web Server does not
      ' have access to file information on the FileServer
      '-----------------------------------------------------------------------        

      if FileObj.FileExists(ArchiveNameNew) then
     		set ArchiveFile = FileObj.GetFile(ArchiveNameNew)
        ArchiveSize = ArchiveFile.size
        set ArchiveFile = nothing
      end if

      Call Connect_SiteWide
      
      ArchivePath = replace(ArchiveNameNew,"\","/")
      ArchivePath = mid(ArchivePath,instr(LCase(ArchivePath),"download/"))
      
      ArchiveSQL = "INSERT dbo.Calendar_Temp (Name, Value1, Value2) " &_
                   "VALUES ('Archive_Name','" & ArchivePath & "'," & ArchiveSize & ")"
            
      conn.execute ArchiveSQL
       
      Call Disconnect_SiteWide

      '-----------------------------------------------------------------------  
            
      if Debug_Flag then
        FileTemp.writeLine("File to Archive: " & oFile.DestinationPath)    
        FileTemp.writeLine("Archive Name: " & ArchiveNameNew)
        FileTemp.writeLine("File To Add : " & oFile.DestinationPath)      
      end if  
      
      exit for ' There is only one Archive file produced from the Asset Directory
      
    end if
    
  end if
  
next

if Debug_Flag then
  FileTemp.writeline("End Create Archive Log")
  FileTemp.close
  set FileTemp = nothing
end if  

set FileObj = nothing
set oArch   = nothing

' Send the SOAP Response Back to the WebServer

oFileUpEE.SendResponse Response

if err.Number <> 0 then
  response.Write "<B>FileServer SendResponse Error</B><BR>" & Err.Description & " (" & Err.Source & ")"
  repsonse.Status = 500
  response.end
end if

set oFileUpEE = nothing
%>
