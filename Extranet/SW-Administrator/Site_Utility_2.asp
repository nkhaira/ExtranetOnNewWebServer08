<% Option Explicit %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>Site Utility Domain Administration</TITLE>
</HEAD>
<BODY>

<%
' --------------------------------------------------------------------------------------
' Global Variables
' --------------------------------------------------------------------------------------

Dim TabStop
Dim NewLine

Dim TestFilePath
TestFilePath = server.mappath("/")
response.write testfilepath
Dim TestDrive
TestDrive = Mid(TestFilePath,1,1)

' --------------------------------------------------------------------------------------
' Constants returned by Drive.DriveType
' --------------------------------------------------------------------------------------
Const DriveTypeRemovable = 1
Const DriveTypeFixed = 2
Const DriveTypeNetwork = 3
Const DriveTypeCDROM = 4
Const DriveTypeRAMDisk = 5

' --------------------------------------------------------------------------------------
' Constants returned by File.Attributes
' --------------------------------------------------------------------------------------
Const FileAttrNormal   = 0
Const FileAttrReadOnly = 1
Const FileAttrHidden = 2
Const FileAttrSystem = 4
Const FileAttrVolume = 8
Const FileAttrDirectory = 16
Const FileAttrArchive = 32 
Const FileAttrAlias = 1024
Const FileAttrCompressed = 2048

' --------------------------------------------------------------------------------------
' Constants for opening files
' --------------------------------------------------------------------------------------
Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

Call Main

' --------------------------------------------------------------------------------------

sub Main

   Dim FSO

   TabStop = Chr(9)
   NewLine = Chr(10)
   
   Set FSO = CreateObject("Scripting.FileSystemObject")

   Print GenerateDriveInformation(FSO) & NewLine & NewLine
   Print GenerateTestInformation(FSO) & NewLine & NewLine
   
end sub   

' --------------------------------------------------------------------------------------
' ShowDriveType
' Purpose: 
'    Generates a string describing the drive type of a given Drive object.
' Demonstrates the following 
'  - Drive.DriveType
' --------------------------------------------------------------------------------------

Function ShowDriveType(Drive)

   Dim S
   
   Select Case Drive.DriveType
     Case DriveTypeRemovable
        S = "Removable"
     Case DriveTypeFixed
        S = "Fixed"
     Case DriveTypeNetwork
        S = "Network"
     Case DriveTypeCDROM
        S = "CD-ROM"
     Case DriveTypeRAMDisk
        S = "RAM Disk"
     Case Else
        S = "Unknown"
   End Select

   ShowDriveType = S

End Function

' --------------------------------------------------------------------------------------
' ShowFileAttr
' Purpose: 
'    Generates a string describing the attributes of a file or folder.
' Demonstrates the following 
'  - File.Attributes
'  - Folder.Attributes
' --------------------------------------------------------------------------------------

Function ShowFileAttr(File) ' File can be a file or folder

   Dim S
   Dim Attr
   
   Attr = File.Attributes

   If Attr = 0 Then
      ShowFileAttr = "Normal"
      Exit Function
   End If

   If Attr And FileAttrDirectory Then S = S & "Directory "
   If Attr And FileAttrReadOnly Then S = S & "Read-Only "
   If Attr And FileAttrHidden Then S = S & "Hidden "
   If Attr And FileAttrSystem Then S = S & "System "
   If Attr And FileAttrVolume Then S = S & "Volume "
   If Attr And FileAttrArchive Then S = S & "Archive "
   If Attr And FileAttrAlias Then S = S & "Alias "
   If Attr And FileAttrCompressed Then S = S & "Compressed "

   ShowFileAttr = S

End Function

' --------------------------------------------------------------------------------------
' GenerateDriveInformation
' Purpose: 
'    Generates a string describing the current state of the 
'    available drives.
' Demonstrates the following 
'  - FileSystemObject.Drives 
'  - Iterating the Drives collection
'  - Drives.Count
'  - Drive.AvailableSpace
'  - Drive.DriveLetter
'  - Drive.DriveType
'  - Drive.FileSystem
'  - Drive.FreeSpace
'  - Drive.IsReady
'  - Drive.Path
'  - Drive.SerialNumber
'  - Drive.ShareName
'  - Drive.TotalSize
'  - Drive.VolumeName
' --------------------------------------------------------------------------------------

Function GenerateDriveInformation(FSO)

   Dim Drives
   Dim Drive
   Dim S

   Set Drives = FSO.Drives
   
   S = "Number of drives:" & TabStop & Drives.Count & NewLine & NewLine
   
   ' Construct 1st line of report.
   S = S &_
       "<TABLE CELLSPACING=2 BORDER=0><TR>" &_
       "<TD>Drive</TD>" &_
       "<TD COLSPAN=5>&nbsp;</TD>" &_
       "<TD>Total</TD>" &_
       "<TD>Free</TD>" &_       
       "<TD>Available</TD>" &_       
       "<TD>Serial</TD>" &_
       "</TR>"       

   ' Construct 2nd line of report.
   S = S &_
       "<TR>" &_
       "<TD>Letter</TD>" &_
       "<TD>Path</TD>" &_
       "<TD>Type</TD>" &_
       "<TD>Ready?</TD>" &_       
       "<TD>Name</TD>" &_       
       "<TD>System</TD>" &_
       "<TD>Space</TD>" &_
       "<TD>Space</TD>" &_
       "<TD>Space</TD>" &_
       "<TD>Number</TD>" &_                            
       "</TR>"       
      
   ' Construct 3nd line of report.
   S = S &_
       "<TR>" &_
       "<TD>&nbsp;</TD>" &_
       "<TD>&nbsp;</TD>" &_
       "<TD>&nbsp;</TD>" &_
       "<TD>&nbsp;</TD>" &_       
       "<TD>&nbsp;</TD>" &_       
       "<TD>&nbsp;</TD>" &_
       "<TD>KByte</TD>" &_
       "<TD>KByte</TD>" &_
       "<TD>KByte</TD>" &_
       "<TD>&nbsp;</TD>" &_                            
       "</TR>"       
   
   For Each Drive In Drives
   
      S = S & "<TR>"
      S = S & "<TD>" & Drive.DriveLetter & "</TD>"
      S = S & "<TD>" & Drive.Path & "</TD>"
      S = S & "<TD>" & ShowDriveType(Drive) & "</TD>"
      S = S & "<TD>" & Drive.IsReady & "</TD>"

      If Drive.IsReady Then
         If DriveTypeNetwork = Drive.DriveType Then
            S = S & "<TD>" & Drive.ShareName & "</TD>" 
         Else
            S = S & "<TD>" & Drive.VolumeName & "</TD>" 
         End If
         S = S & "<TD>" & Drive.FileSystem & "</TD>"
         S = S & "<TD>" & FormatNumber(CDbl(Drive.TotalSize / 1024),0) & "</TD>"
         S = S & "<TD>" & FormatNumber(CDbl(Drive.FreeSpace / 1024),0) & "</TD>"
         S = S & "<TD>" & FormatNumber(CDbl(Drive.AvailableSpace / 1024),0) & "</TD>"                  
         S = S & "<TD>" & Hex(Drive.SerialNumber) & "</TD>"
      End If

      S = S & "</TR>"

   Next
   
   S = S & "</TABLE>"

   GenerateDriveInformation = S

End Function

' --------------------------------------------------------------------------------------
' GenerateFileInformation
' Purpose: 
'    Generates a string describing the current state of a file.
' Demonstrates the following 
'  - File.Path
'  - File.Name
'  - File.Type
'  - File.DateCreated
'  - File.DateLastAccessed
'  - File.DateLastModified
'  - File.Size
' --------------------------------------------------------------------------------------

Function GenerateFileInformation(File)

   Dim S

   S = NewLine & "Path:" & TabStop & File.Path
   S = S & NewLine & "Name:" & TabStop & File.Name
   S = S & NewLine & "Type:" & TabStop & File.Type
   S = S & NewLine & "Attribs:" & TabStop & ShowFileAttr(File)
   S = S & NewLine & "Created:" & TabStop & File.DateCreated
   S = S & NewLine & "Accessed:" & TabStop & File.DateLastAccessed
   S = S & NewLine & "Modified:" & TabStop & File.DateLastModified
   S = S & NewLine & "Size" & TabStop & File.Size & NewLine

   GenerateFileInformation = S

End Function

' --------------------------------------------------------------------------------------
' GenerateFolderInformation
' Purpose: 
'    Generates a string describing the current state of a folder.
' Demonstrates the following 
'  - Folder.Path
'  - Folder.Name
'  - Folder.DateCreated
'  - Folder.DateLastAccessed
'  - Folder.DateLastModified
'  - Folder.Size
' --------------------------------------------------------------------------------------

Function GenerateFolderInformation(Folder)

   Dim S

   S = "Path:" & TabStop & Folder.Path
   S = S & NewLine & "Name:" & TabStop & Folder.Name
   S = S & NewLine & "Attribs:" & TabStop & ShowFileAttr(Folder)
   S = S & NewLine & "Created:" & TabStop & Folder.DateCreated
   S = S & NewLine & "Accessed:" & TabStop & Folder.DateLastAccessed
   S = S & NewLine & "Modified:" & TabStop & Folder.DateLastModified
   S = S & NewLine & "Size:" & TabStop & Folder.Size
   S = S & NewLine & NewLine

   GenerateFolderInformation = S

End Function

' --------------------------------------------------------------------------------------
' GenerateAllFolderInformation
' Purpose: 
'    Generates a string describing the current state of a
'    folder and all files and subfolders.
' Demonstrates the following 
'  - Folder.Path
'  - Folder.SubFolders
'  - Folders.Count
' --------------------------------------------------------------------------------------

Function GenerateAllFolderInformation(Folder)

   Dim S
   Dim SubFolders
   Dim SubFolder
   Dim Files
   Dim File

   S = "<FONT COLOR=BLUE><B>Folder</B></FONT>:" & TabStop & Folder.Path & NewLine & NewLine
   Set Files = Folder.Files

   If 1 = Files.Count Then
      S = S & "<FONT COLOR=RED>There is 1 file</FONT>" & NewLine
   Else
      S = S & "<FONT COLOR=RED>There are " & Files.Count & " files</FONT>" & NewLine
   End If

   If Files.Count <> 0 Then
      For Each File In Files
         S = S & GenerateFileInformation(File)
      Next
   End If

   Set SubFolders = Folder.SubFolders

   If 1 = SubFolders.Count Then
      S = S & NewLine & "<FONT COLOR=RED>There is 1 sub folder</FONT>" & NewLine & NewLine
   Else
      S = S & NewLine & "<FONT COLOR=RED>There are " & SubFolders.Count & " sub folders</FONT>" &_
      NewLine & NewLine
   End If

   If SubFolders.Count <> 0 Then
      For Each SubFolder In SubFolders
         S = S & GenerateFolderInformation(SubFolder)
      Next
      S = S & NewLine & NewLine
      For Each SubFolder In SubFolders
         S = S & GenerateAllFolderInformation(SubFolder)
      Next
   End If

   GenerateAllFolderInformation = S

End Function

' --------------------------------------------------------------------------------------
' GenerateTestInformation
' Purpose: 
'    Generates a string describing the current state of the C:\Test
'    folder and all files and subfolders.
' Demonstrates the following 
'  - FileSystemObject.DriveExists
'  - FileSystemObject.FolderExists
'  - FileSystemObject.GetFolder
' --------------------------------------------------------------------------------------

Function GenerateTestInformation(FSO)

   Dim TestFolder
   Dim S

   If Not FSO.DriveExists(TestDrive) Then Exit Function
   If Not FSO.FolderExists(TestFilePath) Then Exit Function

   Set TestFolder = FSO.GetFolder(TestFilePath)

   GenerateTestInformation = GenerateAllFolderInformation(TestFolder) 

End Function

' --------------------------------------------------------------------------------------

Sub Print(x)
   Response.Write "<PRE><FONT FACE=""Arial"" SIZE=""2"">"
   Response.Write x
   Response.Write "</FONT></PRE>"
End Sub

%>
</BODY>
</HTML>
