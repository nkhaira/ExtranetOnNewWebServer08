<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' --------------------------------------------------------------------------------------

Dim dDate     ' Date we're displaying calendar for
Dim tDate     ' Today's Date
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table
Dim View      ' 0 = Days have A HREF Links / 1 = No Links

View = 1      

' --------------------------------------------------------------------------------------
' Get selected date.  There are two ways to do this.
' First check if we were passed a full date in RQS("date").
' If so use it, if not look for seperate variables, putting them togeter into a date.
' Lastly check if the date is valid...if not use today
' --------------------------------------------------------------------------------------

tDate = Date()  

If IsDate(request("date")) Then
	dDate = CDate(request("date"))
Else
	If IsDate(request("month") & "-" & request("day") & "-" & request("year")) Then
		dDate = CDate(request("month") & "-" & request("day") & "-" & request("year"))
	Else
		dDate = Date()
	End If
End If

' Now we've got the date.  Now get Days in Month and the Day of the Week Calendar starts on

iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(Month(dDate), Year(dDate))

' --------------------------------------------------------------------------------------
' Get Calendar Items by day that fall Within this Month / Year
' --------------------------------------------------------------------------------------

Dim dID
Dim dID_Max
Dim dCategory_ID
Dim dCategory
Dim dLocation
Dim dBDate
Dim dEDate
Dim dPDate
Dim dUDate
Dim dTitle
Dim dStatus

SQL = "SELECT * FROM Calendar_Category WHERE Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Calendar_View=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
Set rsCategory = Server.CreateObject("ADODB.Recordset")
rsCategory.Open SQL, conn, 3, 3
                  
Do while not rsCategory.EOF            
    
  SQL = "SELECT Calendar.* FROM Calendar WHERE Calendar.Site_ID=" & CInt(Site_ID) & " AND Calendar.Category_ID=" & rsCategory("ID")
  
  ' Build LIKE Clause for Sub-Group Memberships
  
  if (Access_Level <=2 or Access_Level=6) then
    SQL = SQL & " AND (Calendar.SubGroups LIKE '%all%'"
    for i = 0 to UserSubGroups_Max
        SQL = SQL & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
    next
    SQL = SQL & ")"        
  end if   
  
  ' Determine if Active or Archive
  
  if (Access_Level <=2 or Access_Level=6) then
    SQL = SQL & " AND Calendar.Status=1 AND "
  else ' Show all for Admin
    SQL = SQL & " AND (Calendar.Status=1 OR Calendar.Status=0) AND "
  end if
  
  if (Access_Level <=2 or Access_Level=6) then SQL = SQL & "Calendar.LDate<='" & Now() & "' AND "       ' Do not show prior to pre-announce date
  SQL = SQL & "((Calendar.BDate>='" & Month(dDate) & "/1/" & Year(dDate) & "' AND Calendar.BDate<='" & Month(dDate) & "/" & iDIM & "/" & Year(dDate) & "') OR "
  SQL = SQL & " (Calendar.EDate>='" & Month(dDate) & "/1/" & Year(dDate) & "' AND Calendar.EDate<='" & Month(dDate) & "/" & iDIM & "/" & Year(dDate) & "')) "    

  ' Restricted Countries
  
  if (Access_Level <= 2 or Access_Level = 6) then
    SQL = SQL & " AND (Calendar.Country='none' OR Calendar.Country LIKE '%" & Login_Country & "%') "
  end if 

  SQL = SQL & "ORDER BY Calendar.BDate, Calendar.Category_ID"          
  Set rsCalendar = Server.CreateObject("ADODB.Recordset")
  rsCalendar.Open SQL, conn, 3, 3
  
  'response.write SQL & "<BR><BR>"
  
  do while not rsCalendar.EOF       ' Use tilda ~ as delimeter
    dID            = dID            & "~" & rsCalendar("ID")
    dCategory_ID   = dCategory_ID   & "~" & rsCalendar("Category_ID")
    temp = rsCategory("Title")
    if lcase(mid(temp,len(temp))) = "s" then
      dCategory    = dCategory      & "~" & mid(temp,1,len(temp)-1)
    else
      dCategory    = dCategory      & "~" & Temp
    end if  
    dCategory_Code = dCategory_Code & "~" & rsCategory("Code")
    dLocation      = dLocation      & "~" & rsCalendar("Location")
    dBDate         = dBDate         & "~" & rsCalendar("BDate")
    dEDate         = dEDate         & "~" & rsCalendar("EDate")
    dTitle         = dTitle         & "~" & rsCalendar("Title")
    dPDate         = dPDate         & "~" & rsCalendar("PDate")
    dUDate         = dUDate         & "~" & rsCalendar("UDate")    
    dStatus        = dStatus        & "~" & rsCalendar("Status")
    rsCalendar.MoveNext
  loop
  
  rsCalendar.close
  set rsCalendar = nothing

 rsCategory.MoveNext 
loop
   
rsCategory.close
set rsCategory=nothing

dID             = Split(dID,"~")
dCategory_ID    = Split(dCategory_ID,"~")
dCategory       = Split(dCategory,"~")
dCategory_Code  = Split(dCategory_Code,"~")
dLocation       = Split(dLocation,"~")
dBDate          = Split(dBDate,"~")
dEDate          = Split(dEDate,"~")
dPDate          = Split(dPDate,"~")
dUDate          = Split(dUDate,"~")
dTitle          = Split(dTitle,"~")
dStatus         = Split(dStatus,"~")

dID_Max = Ubound(dID)

' --------------------------------------------------------------------------------------
' Outer Table
' --------------------------------------------------------------------------------------

response.write "<DIV ALIGN=CENTER>"

Call Table_Begin_Calendar

%>      

<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0 WIDTH="100%">
  <TR>
    <TD ALIGN="center" COLSPAN=7>
    	<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
    		<TR>
    			<TD WIDTH= "2%" ALIGN="center" CLASS=NormalBoldGold>
            <%
            response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&date=" & SubtractOneMonth(dDate) & """><IMG SRC=""/images/Left_Arrow.gif"" BORDER=0></A>" & vbCrLf
            %>                      
          </TD>
    			<TD WIDTH="96%" ALIGN="center" CLASS=NormalBoldGold>
            <%= Translate(MonthName(Month(dDate)),Login_Language,conn) & "  " & Year(dDate) %>
          </TD>
    			<TD WIDTH="2%" ALIGN="center" CLASS=NormalBoldGold>
            <%
            response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&date=" & AddOneMonth(dDate) & """><IMG SRC=""/images/Right_Arrow.gif"" BORDER=0></A>" & vbCrLf
            %>                      
          </TD>
    		</TR>
    	</TABLE>
    </TD>
  </TR>
	<TR>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="13%" CLASS=NORMALBOLD><%=Translate("Sunday",Login_Language,conn)%></TD>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="15%" CLASS=NORMALBOLD><%=Translate("Monday",Login_Language,conn)%></TD>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="15%" CLASS=NORMALBOLD><%=Translate("Tuesday",Login_Language,conn)%></TD>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="14%" CLASS=NORMALBOLD><%=Translate("Wednesday",Login_Language,conn)%></TD>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="15%" CLASS=NORMALBOLD><%=Translate("Thursday",Login_Language,conn)%></TD>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="15%" CLASS=NORMALBOLD><%=Translate("Friday",Login_Language,conn)%></TD>
		<TD ALIGN="center" BGCOLOR="White" WIDTH="13%" CLASS=NORMALBOLD><%=Translate("Saturday",Login_Language,conn)%></TD>
	</TR>
<%

' Write spacer cells at beginning of first row if month doesn't start on a Sunday.

If iDOW <> 1 Then
	Response.Write vbTab & "<TR>" & vbCrLf
	iPosition = 1
	Do While iPosition < iDOW
		Response.Write vbTab & vbTab & "<TD BGCOLOR=""#BBBBBB"">&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
End If

' Write days of month in proper day slots

iCurrent = 1
iPosition = iDOW
Do While iCurrent <= iDIM

	' If we're at the begginning of a row then write TR

	If iPosition = 1 Then
		Response.Write vbTab & "<TR>" & vbCrLf
	End If
	
  Current_Day = DateSerial(Year(dDate),Month(dDate),iCurrent)

	' If the day we're writing is the selected day then highlight it.

	if Year(tDate) = Year(dDate) and Month(tDate) = Month(dDate) and iCurrent = Day(tDate) Then

		response.Write vbTab & vbTab & "<TD HEIGHT=""70"" VALIGN=TOP BGCOLOR=""#FFFFFF"" CLASS=SMALLBOLD>"
    response.write "<SPAN CLASS=NORMALBOLDRED>" & iCurrent & "</SPAN><BR>"
    
    for event_day = 1 to dID_Max

      if DateSerial(Year(dBDate(event_day)), Month(dBDate(event_day)), Day(dBDate(event_day))) <= current_day and DateSerial(Year(dEDate(event_day)), Month(dEDate(event_day)), Day(dEDate(event_day))) >= current_day then
        response.write "<SPAN TITLE=""See Detail Below"">" & Translate(Trim(dCategory(event_day)),Login_Language,conn) & "</SPAN><BR>"
      end if
    next    
    
    response.Write "</TD>" & vbCrLf

	else
  
    if view = 1 then
  		response.Write vbTab & vbTab & "<TD HEIGHT=""70"" VALIGN=TOP BGCOLOR=""#EEEEEE"" CLASS=SMALLBOLD>"
      response.write "<SPAN CLASS=NORMALBOLD>" & iCurrent & "</SPAN><BR>"  
    else  
  		response.Write vbTab & vbTab & "<TD HEIGHT=""70"" VALIGN=TOP BGCOLOR=""#EEEEEE"" CLASS=SMALLBOLD><A HREF=""calendar_grid.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & Year(dDate) & """>"
      response.write "<SPAN CLASS=NORMALBOLD>" & iCurrent & "</SPAN></A><BR>"
    end if
    
    for event_day = 1 to dID_Max

      if DateSerial(Year(dBDate(event_day)), Month(dBDate(event_day)), Day(dBDate(event_day))) <= current_day and DateSerial(Year(dEDate(event_day)), Month(dEDate(event_day)), Day(dEDate(event_day))) >= current_day then
        response.write "<SPAN TITLE=""See Detail Below"">" & Translate(Trim(dCategory(event_day)),Login_Language,conn) & "</SPAN><BR>"
      end if
        
    next
        
    Response.Write "</TD>" & vbCrLf

	end if
	
	' End of Row

	If iPosition = 7 Then
		Response.Write vbTab & "</TR>" & vbCrLf
		iPosition = 0
	End If
	
	' Increment Counters

	iCurrent = iCurrent + 1
	iPosition = iPosition + 1
Loop

' Write spacer cells at end of last row if month doesn't end on a Saturday.

If iPosition <> 1 Then
	Do While iPosition <= 7
		Response.Write vbTab & vbTab & "<TD BGCOLOR=""#BBBBBB"">&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
	Response.Write vbTab & "</TR>" & vbCrLf
End If

response.write "</TABLE>"

Call Table_End_Calendar

' Schedule of Calendar Events

if dID_Max > 0 then

  response.write "<BR>"

  Call Table_Begin_Calendar

  response.write "<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0 WIDTH=""100%"">" & vbCrLf
  response.write "<TR>"
  response.write "  <TD COLSPAN=7 HEIGHT=16 BGCOLOR=""#666666"" ALIGN=CENTER CLASS=NormalBoldGold>" & Translate("Schedule of Events for",Login_Language,conn) & " " & Translate(MonthName(Month(dDate)),Login_Language,conn) & " " & Year(dDate) & "</TD>"
  response.write "</TR>"
  response.write "<TR>"
  response.write "  <TD COLSPAN=7 HEIGHT=4 BGCOLOR=""White"" ALIGN=CENTER></TD>"
  response.write "</TR>"
  response.write "<TR>"
  response.write "  <TD BGCOLOR=""#666666""              CLASS=SmallBoldGold>" & Translate("Event",Login_Language,conn) & "</TD>"
  response.write "  <TD BGCOLOR=""#666666"" COLSPAN=3    CLASS=SmallBoldGold>" & Translate("Title or Subject",Login_Language,conn) & "</TD>"
  response.write "  <TD BGCOLOR=""#666666"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("Begin Date",Login_Language,conn) & "</TD>"
  response.write "  <TD BGCOLOR=""#666666"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("End Date",Login_Language,conn) & "</TD>"
  response.write "  <TD BGCOLOR=""#666666"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("Detail",Login_Language,conn) & "</TD>"
  response.write "</TR>"
  
  for i = 0 to dID_Max
  
    if not isblank(dID(i)) then
      response.write "<TR>"      
      response.write "  <TD BGCOLOR=""White"" CLASS=SMALLBOLD>" & Translate(dCategory(i),Login_Language,conn) & "</TD>"
      response.write "  <TD BGCOLOR=""White"" CLASS=SMALL COLSPAN=3><B>" & dTitle(i) & "</B>"
      if not isblank(dLocation(i)) then response.write "<BR>" & dLocation(i)
      response.write "  </TD>"

      response.write "  <TD BGCOLOR="
      select case dStatus(i)
        case 1      ' Live
          if Access_Level = 0 then  
            response.write """White"""
          else
            response.write """#99FF99"""
          end if              
        case 2      ' Archive
          response.write """#AAAAFF"""            
        case else   ' Pending
          response.write """Yellow"""
      end select
      response.write " CLASS=SMALL ALIGN=CENTER>" & Day(dBDate(i)) & " " & Translate(MonthName(Month(dBDate(i))),Login_Language,conn) & " " & Year(dBDate(i)) & "</TD>"

      response.write "  <TD BGCOLOR="
      select case dStatus(i)
        case 1      ' Live
          if Access_Level = 0 then  
            response.write """White"""
          else
            response.write """#99FF99"""
          end if              
        case 2      ' Archive
          response.write """#AAAAFF"""            
        case else   ' Pending
          response.write """Yellow"""
      end select      
      response.write " CLASS=SMALL ALIGN=CENTER>" & Day(dEDate(i)) & " " & Translate(MonthName(Month(dEDate(i))),Login_Language,conn) & " " & Year(dEDate(i)) & "</TD>"      

      response.write "  <TD BGCOLOR=""White"" CLASS=SMALLBOLD ALIGN=CENTER>"
      
      response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID
      response.write "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & dCategory_Code(i) & "&CINN=" & dCategory_ID(i) & "#" & dID(i) & """>"
      response.write "<IMG SRC=""/images/calendar_button.gif"" WIDTH=16 BORDER=0 VSPACE=0 ALIGN=ABSMIDDLE></A>"     
      response.write    "</TD>"
      response.write "</TR>"

    end if

  next

  response.write "</TABLE>" & vbCrLf

  Call Table_End_Calendar

end if                          
  
response.write "</DIV>"

%><!-- Calendar Begin --><%
' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------

sub Table_Begin_Calendar()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder WIDTH=""100%"">" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/SideNav_TL_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/SideNav_TR_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"" WIDTH=""100%"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End_Calendar()
    response.write "        </TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/SideNav_BL_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/SideNav_BR_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "    </TABLE>" & vbCrLf
end sub  

'--------------------------------------------------------------------------------------

Function GetDaysInMonth(iMonth, iYear)
	Select Case iMonth
		Case 1, 3, 5, 7, 8, 10, 12
			GetDaysInMonth = 31
		Case 4, 6, 9, 11
			GetDaysInMonth = 30
		Case 2
			If IsDate("February 29, " & iYear) Then
				GetDaysInMonth = 29
			Else
				GetDaysInMonth = 28
			End If
	End Select
End Function

' --------------------------------------------------------------------------------------

Function GetWeekdayMonthStartsOn(iMonth, iYear)
	GetWeekdayMonthStartsOn = WeekDay(CDate(iMonth & "/1/" & iYear))
End Function

' --------------------------------------------------------------------------------------

Function SubtractOneMonth(dDate)
  Dim iDay, iMonth, iYear	
  	iDay = Day(dDate)
  	iMonth = Month(dDate)
  	iYear = Year(dDate)
  
  	If iMonth = 1 Then
  		iMonth = 12
  		iYear = iYear - 1
  	Else
  		iMonth = iMonth - 1
  	End If
  	
  	If iDay > GetDaysInMonth(iMonth, iYear) Then iDay = GetDaysInMonth(iMonth, iYear)
  
  	SubtractOneMonth = CDate(iMonth & "-" & iDay & "-" & iYear)
End Function

' --------------------------------------------------------------------------------------

Function AddOneMonth(dDate)
  Dim iDay, iMonth, iYear	
  	iDay = Day(dDate)
  	iMonth = Month(dDate)
  	iYear = Year(dDate)
  
  	If iMonth = 12 Then
  		iMonth = 1
  		iYear = iYear + 1
  	Else
  		iMonth = iMonth + 1
  	End If
  	
  	If iDay > GetDaysInMonth(iMonth, iYear) Then iDay = GetDaysInMonth(iMonth, iYear)
  
  	AddOneMonth = CDate(iMonth & "-" & iDay & "-" & iYear)
End Function

' --------------------------------------------------------------------------------------
%>
