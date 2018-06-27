<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' The purpose of this script is to rename both the physical file name and DB file name
' per the naming convention for Assets, Archives and Thumbnails.
'
' Note: Change Site_ID

Dim Site_ID
Site_ID = 82

Session.timeout      = 240      ' Set to 4 Hours
Server.ScriptTimeout = 60 * 10  ' 10 Minutes

Call Connect_SiteWide

Set objFSO = CreateObject("Scripting.FileSystemObject")

SQL = "SELECT ID, Site_Code FROM Site WHERE Enabled=-1 ORDER BY ID"

Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

do while not rsSite.EOF

  if CInt(rsSite("ID")) = Site_ID then

    SQL = "SELECT ID, Item_Number, Cost_Center, [Language], Revision_Code, File_Name, Archive_Name, Thumbnail, UDate " &_
          "FROM dbo.Calendar " &_
          "WHERE (Item_Number IS NOT NULL) AND (LEN(Item_Number) = 7) AND (File_Name IS NOT NULL) AND Site_ID=" & rsSite("ID")
    
    
    SitePath = "/" & rsSite("Site_Code") & "/"
    
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
    
      Asset_Extn = UCase(Mid(rsID("File_Name"), InstrRev(rsID("File_Name"), ".") + 1))
      Asset_Name = "Download/Asset/" & UCase(rsID("Item_Number"))
      if rsID("Item_Number") < 9000000 then
        Asset_Name = Asset_Name & "_" & CC
      end if
      Asset_Name = Asset_Name & "_" & UCase(rsID("Language") & "_" & rsID("Revision_Code"))
      
      if not isblank(rsID("Archive_Name")) then    
        Archive_Extn = UCase(Mid(rsID("Archive_Name"), InstrRev(rsID("Archive_Name"), ".") + 1))
        Archive_Name = "Download/Archive/" & UCase(rsID("Item_Number"))
        if rsID("Item_Number") < 9000000 then
          Archive_Name = Archive_Name & "_" & CC
        end if
        Archive_Name = Archive_Name & "_" & UCase(rsID("Language") & "_" & rsID("Revision_Code"))
      else
        Archive_Extn = ""
        Archive_Name = ""
      end if  
          
      if not isblank(rsID("Thumbnail")) then
        Thumb_Extn = UCase(Mid(rsID("Thumbnail"), InstrRev(rsID("Thumbnail"), ".") + 1))
        Thumb_Name = "Download/Thumbnail/" & UCase(rsID("Item_Number") & "_" & rsID("Revision_Code"))
      else
        Thumb_Extn = ""
        Thumb_Name = ""
      end if  
  
      select case Asset_Extn
        case "PDF"
          Asset_Name   = Asset_Name   & "_W." & Asset_Extn
          if not isblank(Archive_Name) then
            Archive_Name = Archive_Name & "_W." & Archive_Extn
          end if  
          if not isblank(Thumb_Name) then
            Thumb_Name = Thumb_Name & "_T." & Thumb_Extn
          end if  
        case else
          Asset_Name   = Asset_Name   & "_X." & Asset_Extn
          if not isblank(Archive_Name) then
            Archive_Name = Archive_Name & "_X." & Archive_Extn
          end if  
          if not isblank(Thumb_Name) then
            Thumb_Name = Thumb_Name & "_T." & Thumb_Extn
          end if  
      end select
        
      response.write "Asset Original  File Name A: " & UCase(rsID("File_Name")) & "<BR>"
      response.write "Asset Rename to File Name B: " & UCase(Asset_Name) & "<BR>"
      
      if UCase(rsID("File_Name")) <> UCase(Asset_Name) then
      
        From_File = Server.MapPath(SitePath & rsID("File_Name"))
        To_File   = Server.MapPath(SitePath & Asset_Name)
        response.write "Asset From File Path C: " & From_File & "<BR>"
        response.write "Asset To   File Path D: " & To_File
        if not objFSO.FileExists(From_File) then
          response.write "<FONT COLOR=RED> Source Not Found</FONT>"        
        elseif objFSO.FileExists(From_File) and not objFSO.FileExists(To_File) Then
          objFSO.MoveFile From_File , To_File
          response.write "<FONT COLOR=GREEN> Renamed</FONT>"
        else
          response.write "<FONT COLOR=BLUE> No Change</FONT>"        
        end if
        response.write "<BR>"
        
        if not isblank(Archive_Name) then
          From_File = Server.MapPath(SitePath & rsID("Archive_Name"))
          To_File   = Server.MapPath(SitePath & Asset_Name)
          response.write "Archive From E: " & From_File & "<BR>"
          response.write "Archive To   F: " & To_File
          if not objFSO.FileExists(From_File) then
            response.write "<FONT COLOR=RED> Source Not Found</FONT>"        
          elseif objFSO.FileExists(From_File) and not objFSO.FileExists(To_File) Then
            objFSO.MoveFile From_File , To_File
            response.write "<FONT COLOR=GREEN> Renamed</FONT>"
          else
            response.write "<FONT COLOR=BLUE> No Change</FONT>"        
          end if
          response.write "<BR>"        
        end if  
        
        if not isblank(Thumb_Name) then
          From_File = Server.MapPath(SitePath & rsID("Thumbnail"))
          To_File   = Server.MapPath(SitePath & Thumb_Name)
          response.write "Thumbnail From G: " & From_File & "<BR>"
          response.write "Thumbnail to   H: " & To_File
          if not objFSO.FileExists(From_File) then
            response.write "<FONT COLOR=RED> Source Not Found</FONT>"        
          elseif objFSO.FileExists(From_File) and not objFSO.FileExists(To_File) Then
            objFSO.MoveFile From_File , To_File
            response.write "<FONT COLOR=GREEN> Renamed</FONT>"
          else
            response.write "<FONT COLOR=BLUE> No Change</FONT>"        
          end if
          response.write "<BR>"
        end if
      
        SQL = "UPDATE Calendar SET " &_
              "File_Name='" & Asset_Name & "' "
              
        if not isblank(Archive_Name) then            
          SQL = SQL & ", Archive_Name='" & Archive_Name & "' "
        else
          SQL = SQL & ", Archive_Name=NULL "
          SQL = SQL & ", Archive_Size=0 "
        end if
        
        if not isblank(Thumb_Name) then
          SQL = SQL & ", Thumbnail='" & Thumb_Name & "' "
        else  
          SQL = SQL & ", Thumbnail=NULL "
          SQL = SQL & ", Thumbnail_Size=0 "        
        end if
        
        SQL = SQL & ", UDate='" & rsID("UDate") & "' "
                    
        SQL = SQL & "WHERE Item_Number='" & rsID("Item_Number") & "' AND Revision_Code='" & rsID("Revision_Code") & "'"
  
        conn.execute SQL
        
        response.write "<BR>"
        
      else
        response.write "<BR>"
      end if
        
      response.flush
        
      rsID.MoveNext
    
    loop
    
    rsID.close
    set rsID   = nothing

  end if
  
  rsSite.MoveNext
  
loop

rsSite.close
set rsSite = nothing

set objFSO = nothing

Call Disconnect_SiteWide

response.write "Done"
%>