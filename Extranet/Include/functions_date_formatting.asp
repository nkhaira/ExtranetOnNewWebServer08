<%
'
' Date Formating Functions
' Author: K. David Whitlock
' Date:   03/17/2000

Function FormatDate(iDateFormat, TempDate)

  if Len(DatePart("d", TempDate)) = 1 then
  	strDay = "0" & DatePart("d", TempDate)
  else
  	strDay = DatePart("d", TempDate)
  end if
  
  if Len(DatePart("m", TempDate)) = 1 then
  	strMonth = "0" & DatePart("m", TempDate)
  else
  	strMonth = DatePart("m", TempDate)
  end if
  
  if iDateFormat = 1 then
  	FormatDate = strMonth & "/" & strDay & "/" & DatePart("yyyy", TempDate)
  elseif iDateFormat = 2 then
  	FormatDate = strDay & "/" & strMonth & "/" & DatePart("yyyy", TempDate)
  else
  	FormatDate = DatePart("yyyy", TempDate) & "/" & strMonth & "/" & strDay
  end if

end function

' --------------------------------------------------------------------------------------

Function FormatDateText(iDateFormat, TempDate)

  Dim strDay
  Dim strMonth
  Months="January,February,March,April,May,June,July,August,September,October,November,December"
  arrayMonths = Split(Months,",")
  
  if Len(DatePart("d", TempDate)) = 1 then
  	strDay = "0" & DatePart("d", TempDate)
  else
  	strDay = DatePart("d", TempDate)
  end if
  
  if DatePart("m", TempDate) < 1 or DatePart("m", TempDate) > 12 then
  	strMonth = "Invalid Month"
  else
  	strMonth = arrayMonths(DatePart("m",TempDate)-1)
  end if
  
  if iDateFormat = 1 then
  	FormatDateText = strMonth & ", " & strDay & "," & DatePart("yyyy", TempDate)
  elseif iDateFormat = 2 then
  	FormatDateText = strDay & " " & strMonth & " " & DatePart("yyyy", TempDate)
  else
  	FormatDateText = DatePart("yyyy", TempDate) & " " & strMonth & " " & strDay
  end if

end function

' --------------------------------------------------------------------------------------

Function FormatTime(yourTime)

    Dim TotalTime
    TotalTime = yourTime

    TotalHours = Int((TotalTime / 3600000))
    TotalHoursMod = (TotalTime mod 3600000)
    TotalMin = Int(TotalHoursMod/60000)
    TotalMinMod =  (TotalHoursMod mod 60000)
    TotalSec = Int(TotalMinMod / 1000)
    sTotalHours = ""
    sTotalMin   = ""
    sTotalSec   = ""

    if TotalHours < 10 then sTotalHours = "0"
    sTotalHours = sTotalHours & TotalHours & ":"
    if TotalMin   < 10 then sTotalMin = "0"
    sTotalMin = sTotalMin & TotalMin & ":"
    if TotalSec   < 10 then sTotalSec = "0"
    sTotalSec = sTotalSec & TotalSec
    FormatTime = sTotalHours & sTotalMin & sTotalSec

 end function


%>