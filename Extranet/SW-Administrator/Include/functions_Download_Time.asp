<%
'--------------------------------------------------------------------------------------
'
' Author  : K. D. Whitlock
' Date    : 11/01/2000
' Purpose : Determine typical file download time based on internet connection speed
'
' File_Size       : kbytes or Mbytes depending on nBytes setting
' nBytes          : kb = Kbytes, mb or else = MBytes
' User_Speed      : From user table
' Display_Method  : 0 = in-line string, -1 Table
'--------------------------------------------------------------------------------------

Function Download_Time(File_Size, nBytes, User_Speed, Display_Method)

  ' Uses table "Download_Time" from SiteWide
  
  if LCase(nBytes) = "mb" then
    Multiplier = 1024
  else  ' Default to Kbytes
    Multiplier = 1
  end if  

  SQL = "SELECT Download_Time.* FROM Download_Time WHERE Download_Time.Enabled=" & CInt(True)
  
  if User_Speed > 0 then
    SQL = SQL & " AND ID=" & User_Speed
  end if
  
  SQL = SQL & " ORDER BY DownLoad_Time.bps, DownLoad_Time.Description"
  Set rsDownload = Server.CreateObject("ADODB.Recordset")
  rsDownload.Open SQL, conn, 3, 3

  if Display_Method = True then
    response.write "<TABLE Border=1>"
  end if
  
  do while not rsDownload.EOF  

    speedDescription = rsDownload("Description")
    speed = rsDownload("bps")

    TotalTime = ((File_Size * Multiplier) / speed)
    TotalHours = Int((TotalTime / 3600))
    TotalHoursMod = (TotalTime mod 3600)
    TotalMin = Int(TotalHoursMod/60)
    TotalMinMod =  (TotalHoursMod mod 60)
    TotalSec = Int(TotalMinMod)
    sTotalHours = ""
    sTotalMin   = ""
    sTotalSec   = ""

    if TotalHours < 10 then sTotalHours = "0"
    sTotalHours = sTotalHours & TotalHours & "h "
    if TotalMin   < 10 then sTotalMin = "0"
    sTotalMin = sTotalMin & TotalMin & "m "
    if TotalSec   < 10 then sTotalSec = "0"
    sTotalSec = sTotalSec & TotalSec & "s "
    
    if Display_Method = True then

      response.write "<TR>"
      response.write "<TD CLASS=Small>" & SpeedDescription & "</TD>"
      response.write "<TD CLASS=Small>" & sTotalHours & "</TD>"
      response.write "<TD Class=Small>" & sTotalMin & "</TD>"
      response.write "<TD CLASS=Small>" & sTotalSec & "</TD>"
      response.write "</TR>"
    
    else
      
      response.write SpeedDescription & ": " & sTotalHours & sTotalMin & sTotalSec & "<BR>"
      
    end if
    
    rsDownload.MoveNext
    
  Loop
  
  if Display_Method = True then
    response.write "</TABLE>"
  end if

end function
%>