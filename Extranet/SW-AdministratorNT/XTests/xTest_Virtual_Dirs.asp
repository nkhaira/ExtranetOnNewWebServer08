<%
   Dir2Ck = 5
   
   Set FileObj = Server.CreateObject("Scripting.FileSystemObject")			
   
   select case Dir2Ck
   
    case 1
       TempFile = FileUpEE_TempPath & "\FileUp_Test_File_1.txt"
       response.write TempFile & "<BR>"
       Set FileObjTemp = FileObj.CreateTextFile(TempFile, True, False)
       FileObjTemp.WriteLine(Now())
       FileObjTemp.Close
       Set FileObj = Server.CreateObject("Scripting.FileSystemObject")
       FileObj.CopyFile TempFile, Replace(TempFile,"_1.txt", "_2.txt"), true
     
     case 2  
       TempFile = Server.MapPath("/find-sales/download/asset") & "\FileUp_Test_File_1.txt"
       response.write TempFile & "<BR>"
       Set FileObjTemp = FileObj.CreateTextFile(TempFile, True, False)
       FileObjTemp.WriteLine(Now())
       FileObjTemp.Close
    
    case 3
       TempFile = Server.MapPath("/find-sales/download/archive") & "\FileUp_Test_File_1.txt"
       response.write TempFile & "<BR>"
       Set FileObjTemp = FileObj.CreateTextFile(TempFile, True, False)
       FileObjTemp.WriteLine(Now())
       FileObjTemp.Close
    
    case 4
       TempFile = Server.MapPath("/find-sales/download/thumbnail") & "\FileUp_Test_File_1.txt"
       response.write TempFile & "<BR>"
       Set FileObjTemp = FileObj.CreateTextFile(TempFile, True, False)
       FileObjTemp.WriteLine(Now())
       FileObjTemp.Close
       
    case 5
       TempFile = Server.MapPath("/upload/metcal/procedures") & "\FileUp_Test_File_1.txt"
       response.write TempFile & "<BR>"
       Set FileObjTemp = FileObj.CreateTextFile(TempFile, True, False)
       FileObjTemp.WriteLine(Now())
       FileObjTemp.Close

   end select
   
   set FileObj     = nothing
   set FileObjTemp = nothing
   response.end
%>
