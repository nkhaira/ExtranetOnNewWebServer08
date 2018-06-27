<%
' Author: David Whitlock
%>

<TITLE>Fluke Events Calendar</TITLE>
<%

' ***Begin Function Declaration***

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

Function GetWeekdayMonthStartsOn(iMonth, iYear)
	GetWeekdayMonthStartsOn = WeekDay(CDate(iMonth & "/1/" & iYear))
End Function

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

' ***End Function Declaration***

response.write  "<BODY BGCOLOR=""White"" ALINK=""Black"" LINK=""Black"" VLINK=""Black"">"

Dim dDate     ' Date we're displaying calendar for
Dim tDate     ' Today's Date
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table


' Get selected date.  There are two ways to do this.
' First check if we were passed a full date in RQS("date").
' If so use it, if not look for seperate variables, putting them togeter into a date.
' Lastly check if the date is valid...if not use today

tDate = Date()  

If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
		' The elegant solution for those of you running IIS4
		'If Request.QueryString.Count <> 0 Then Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
	End If
End If

'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(Month(dDate), Year(dDate))

%>

<!-- Outer Table is simply to get the pretty border-->

<TABLE WIDTH="460" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
<TR>
<TD>      
<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0 WIDTH="100%">
 <TR>
		<TD ALIGN="center" COLSPAN=7>
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH= "2%" ALIGN="center">
            <FONT FACE="Verdana" SIZE=2 COLOR="#FFCC00">
            <A HREF="calendar_grid.asp?view=<%=Request.QueryString("view")%>&date=<%= SubtractOneMonth(dDate) %>"><IMG SRC="/images/Left_Arrow.gif" BORDER=0></A>
            </FONT>
          </TD>
					<TD WIDTH="96%" ALIGN="center">
            <FONT FACE="Verdana" SIZE=2 COLOR="#FFCC00">
            <B><%= MonthName(Month(dDate)) & "  " & Year(dDate) %></B>
            </FONT>
          </TD>
					<TD WIDTH="2%" ALIGN="center">
          <FONT FACE="Verdana" SIZE=2 COLOR="#FFCC00">
          <A HREF="calendar_grid.asp?view=<%=Request.QueryString("view")%>&date=<%= AddOneMonth(dDate) %>"><IMG SRC="/images/Right_Arrow.gif" BORDER=0></A>
          </FONT>
          </TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Sun</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Mon</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Tue</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Wed</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Thu</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Fri</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR="White"><FONT COLOR="Black" FACE="Verdana"><B>Sat</B></FONT><BR><IMG SRC="/images/spacer.gif" WIDTH=60 HEIGHT=1 BORDER=0></TD>
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
	
	' If the day we're writing is the selected day then highlight it somehow.
	If Year(tDate) = Year(dDate) and Month(tDate) = Month(dDate) and iCurrent = Day(tDate) Then
		Response.Write vbTab & vbTab & "<TD VALIGN=TOP BGCOLOR=""#FFFFFF""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""Red""><B>" & iCurrent & "</B></FONT><BR>"
    Response.Write "<FONT SIZE=1>"
    Response.Write "<BR>"
    Response.Write "<BR>"
    Response.Write "<BR>"        
    Response.Write "</FONT>"        
    Response.Write "</TD>" & vbCrLf

	ElseIf iCurrent = Day(dDate) Then
		Response.Write vbTab & vbTab & "<TD VALIGN=TOP BGCOLOR=""#FFFFFF""><FONT FACE=""Arial"" SIZE=""2""><B>" & iCurrent & "</B></FONT>"
    Response.Write "<FONT SIZE=1>"
    Response.Write "<BR>"
    Response.Write "<BR>"
    Response.Write "<BR>"
    Response.Write "</FONT>"        
    Response.Write "</TD>" & vbCrLf
	Elseif Request.QueryString("view") = "1" then
		Response.Write vbTab & vbTab & "<TD VALIGN=TOP BGCOLOR=""#EEEEEE""><FONT FACE=""Arial"" SIZE=""2""><B>" & iCurrent & "</FONT></B></BR>"
    Response.Write "<FONT SIZE=1>"
    Response.Write "<BR>"
    Response.Write "<BR>"
    Response.Write "<BR>"    
    Response.Write "</FONT>"        
    Response.Write "</TD>" & vbCrLf
	Else
		Response.Write vbTab & vbTab & "<TD VALIGN=TOP BGCOLOR=""#EEEEEE""><B><A HREF=""calendar_grid.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & Year(dDate) & """><FONT FACE=""Arial"" SIZE=""2"">" & iCurrent & "</FONT></A></B></BR>"
    Response.Write "<FONT SIZE=1>"
    Response.Write "<BR>"
    Response.Write "<BR>"
    Response.Write "<BR>"    
    Response.Write "</FONT>"        
    Response.Write "</TD>" & vbCrLf
	End If
	
	' If we're at the endof a row then write /TR
	If iPosition = 7 Then
		Response.Write vbTab & "</TR>" & vbCrLf
		iPosition = 0
	End If
	
	' Increment variables
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
%>
</TABLE>
</TD>
</TR>
</TABLE>
