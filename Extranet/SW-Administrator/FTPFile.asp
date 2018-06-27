<%
Dim sSourceFile, sFTPLoc

sSourceFile=Request.QueryString("sSourceFile")
sFTPLoc=Request.QueryString("sFTPLoc")
 
'Const ForWriting = 2 ' Input OutPut mode
'Const Create = True 

'Dim MyFile 
'Dim FSO ' FileSystemObject
'Dim TSO ' TextStreamObject

' Use MapPath function to get the Physical Path of file

'MyFile = Server.MapPath("textfile.txt")

'Set FSO = Server.CreateObject("Scripting.FileSystemObject")
'Set TSO = FSO.OpenTextFile(MyFile, ForWriting, Create)

'TSO.write "This is first line in this text File" & vbcrlf 
' Vbcrlf is next line character
'TSO.write "This is Second line in this text file" & vbcrlf
'TSO.write "Writen by devasp visitor at " & Now() 
'TSO.WriteLine ""
'TSO.WriteLine sSourceFile & vbcrlf
'TSO.WriteLine sFTPLoc & vbcrlf

'Response.Write " Three lines are writen to textfile.txt <br>"
'Response.Write " Local time at server is " & Now()

' close TextStreamObject and 
' destroy local variables to relase memory
'TSO.close
'Set TSO = Nothing
'Set FSO = Nothing

' FTP Code

    
    set obj = server.CreateObject("FTPclient.FTP") 
    
    'sFTPLocPrefix = "fnetimages/content/FNet_Dev/FNet/"
      obj.Hostname="ftp.flukenetworks.com"
       obj.Username="fnetimages"
       obj.Password="FN3tImag3s"
     
     'sFTPLocPrefix = "FNet/"
     '  obj.Hostname="Fluke.ingest.cdn.level3.net"
     '   obj.Username="fluke"
     '   obj.Password="KaxETa6vsg"  
       
       obj.UploadFile sSourceFile,sFTPLoc
    

 %>