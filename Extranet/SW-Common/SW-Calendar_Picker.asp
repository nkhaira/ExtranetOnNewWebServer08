<HTML>
<HEAD>
  <TITLE>Date selector</TITLE>

<SCRIPT LANGUAGE="JavaScript">

function setDate() {

    this.dateField   = opener.dateField;
    this.inDate      = dateField.value;

    // SET DAY MONTH AND YEAR TO TODAY'S DATE
    var now   = new Date();
    var day   = now.getDate();
    var month = now.getMonth();
    var year  = now.getYear();
    if (year < 1000) {
        year += 1900;
    }

    // IF A DATE WAS PASSED IN THEN PARSE THAT DATE
    if (inDate.indexOf('/')) {
      
        var inMonth = inDate.substring(0,inDate.indexOf("/"));
        //var inMonth = inDate.substring(inDate.indexOf("/") + 1, inDate.lastIndexOf("/"));
            if (inMonth.substring(0,1) == "0" && inMonth.length > 1)
                inMonth = inMonth.substring(1,inMonth.length);
            inMonth = parseInt(inMonth);
        
        var inDay   = inDate.substring(inDate.indexOf("/") + 1, inDate.lastIndexOf("/"));
        //var inDay   = inDate.substring(0,inDate.indexOf("/"));
            if (inDay.substring(0,1) == "0" && inDay.length > 1)
                inDay = inDay.substring(1,inDay.length);
            inDay = parseInt(inDay);

        var inYear  = parseInt(inDate.substring(inDate.lastIndexOf("/") + 1, inDate.length));

        if (inDay) {
            day = inDay;
        }
        if (inMonth) {
            month = inMonth-1;
        }
        if (inYear) {
            year = inYear;
        }
    }
    this.focusDay                           = day;
    document.calControl.month.selectedIndex = month;
    document.calControl.year.value          = year;
    displayCalendar(day, month, year);
}


function setToday() {
    // SET DAY MONTH AND YEAR TO TODAY'S DATE
    var now   = new Date();
    var day   = now.getDate();
    var month = now.getMonth();
    var year  = now.getYear();
    if (year < 1000) {
        year += 1900;
    }
    this.focusDay                           = day;
    document.calControl.month.selectedIndex = month;
    document.calControl.year.value          = year;
    displayCalendar(day, month, year);
}


function isFourDigitYear(year) {
    if (year.length != 4) {
        alert ("Sorry, the year must be four-digits in length.");
        document.calControl.year.select();
        document.calControl.year.focus();
    }
    else {
        return true;
    }
}


function selectDate() {
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)) {
        var day   = 0;
        var month = document.calControl.month.selectedIndex;
        displayCalendar(day, month, year);
    }
}


function setPreviousYear() {
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)) {
        var day   = 0;
        var month = document.calControl.month.selectedIndex;
        year--;
        document.calControl.year.value = year;
        displayCalendar(day, month, year);
    }
}


function setPreviousMonth() {
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)) {
        var day   = 0;
        var month = document.calControl.month.selectedIndex;
        if (month == 0) {
            month = 11;
            if (year > 1000) {
                year--;
                document.calControl.year.value = year;
            }
        }
        else {
            month--;
        }
        document.calControl.month.selectedIndex = month;
        displayCalendar(day, month, year);
    }
}


function setNextMonth() {
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)) {
        var day   = 0;
        var month = document.calControl.month.selectedIndex;
        if (month == 11) {
            month = 0;
            year++;
            document.calControl.year.value = year;
        }
        else {
            month++;
        }
        document.calControl.month.selectedIndex = month;
        displayCalendar(day, month, year);
    }
}


function setNextYear() {
    var year  = document.calControl.year.value;
    if (isFourDigitYear(year)) {
        var day   = 0;
        var month = document.calControl.month.selectedIndex;
        year++;
        document.calControl.year.value = year;
        displayCalendar(day, month, year);
    }
}


function displayCalendar(day, month, year) {       

    day     = parseInt(day);
    month   = parseInt(month);
    year    = parseInt(year);
    var i   = 0;
    var now = new Date();

    if (day == 0) {
        var nowDay = now.getDate();
    }
    else {
        var nowDay = day;
    }
    var days         = getDaysInMonth(month+1,year);
    var firstOfMonth = new Date (year, month, 1);
    var startingPos  = firstOfMonth.getDay();
    days += startingPos;

    // MAKE BEGINNING NON-DATE BUTTONS BLANK
    for (i = 0; i < startingPos; i++) {
        document.calButtons.elements[i].value = "   ";
        
        //Aanpassing om lege buttons weg te laten  (nog 2x wijziging in onderstaande stuk)
        document.calButtons.elements[i].height = 0;
        document.calButtons.elements[i].width = 0;
    }

    // SET VALUES FOR DAYS OF THE MONTH
    for (i = startingPos; i < days; i++)  
    {
        document.calButtons.elements[i].value = i-startingPos+1;
        document.calButtons.elements[i].onClick = "returnDate"
        
        //Aanpassing om lege buttons weg te laten  (nog 1x wijziging in onderstaande stuk)
        document.calButtons.elements[i].height = 25;
        document.calButtons.elements[i].width = 25;
    }

    // MAKE REMAINING NON-DATE BUTTONS BLANK
    for (i=days; i<42; i++)  {
        document.calButtons.elements[i].value = "   ";
        
        //Aanpassing om lege buttons weg te laten  (laatste wijziging)
        document.calButtons.elements[i].height = 0;
        document.calButtons.elements[i].width = 0;
    }

    // GIVE FOCUS TO CORRECT DAY
    document.calButtons.elements[focusDay+startingPos-1].focus();
}


// GET NUMBER OF DAYS IN MONTH
function getDaysInMonth(month,year)  {
    var days;
    if (month==1 || month==3 || month==5 || month==7 || month==8 ||
        month==10 || month==12)  days=31;
    else if (month==4 || month==6 || month==9 || month==11) days=30;
    else if (month==2)  {
        if (isLeapYear(year)) {
            days=29;
        }
        else {
            days=28;
        }
    }
    return (days);
}


// CHECK TO SEE IF YEAR IS A LEAP YEAR
function isLeapYear (Year) {
    if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0)) {
        return (true);
    }
    else {
        return (false);
    }
}


// SET FORM FIELD VALUE TO THE DATE SELECTED
function returnDate(inDay)
{
    var day   = inDay;
    var month = (document.calControl.month.selectedIndex)+1;
    var year  = document.calControl.year.value;

    if ((""+month).length == 1)
    {
        month="0"+month;
    }
    if ((""+day).length == 1)
    {
        day="0"+day;
    }
    if (day != "   ") {
        // dateField.value = day + "-" + month + "-" + year;
        dateField.value = month + "/" + day + "/" + year;        
        window.close()
    }
}


// -->
</SCRIPT>
</HEAD>

<BODY BGCOLOR="" ONLOAD="setDate()">

<CENTER>
<FORM NAME="calControl" onSubmit="return false;">
<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR><TD COLSPAN=7>
<CENTER>
<SELECT NAME="month" onChange='selectDate()'>
   <OPTION>January</OPTION>
   <OPTION>February</OPTION>
   <OPTION>March</OPTION>
   <OPTION>April</OPTION>
   <OPTION>May</OPTION>
   <OPTION>June</OPTION>
   <OPTION>July</OPTION>
   <OPTION>August</OPTION>
   <OPTION>September</OPTION>
   <OPTION>Oktober</OPTION>
   <OPTION>November</OPTION>
   <OPTION>December</OPTION>
</SELECT>
<INPUT NAME="year" TYPE=TEXT SIZE=4 MAXLENGTH=4 onChange="selectDate()">
</CENTER>
</TD>
</TR>

<TR>
<TD COLSPAN=7>
<CENTER><INPUT 
TYPE=BUTTON NAME="previousYear" VALUE="<<"    onClick="setPreviousYear()"><INPUT 
TYPE=BUTTON NAME="previousYear" VALUE=" < "   onClick="setPreviousMonth()"><INPUT 
TYPE=BUTTON NAME="previousYear" VALUE="Today" onClick="setToday()"><INPUT 
TYPE=BUTTON NAME="previousYear" VALUE=" > "   onClick="setNextMonth()"><INPUT 
TYPE=BUTTON NAME="previousYear" VALUE=">>"    onClick="setNextYear()">
</CENTER>
</TD>
</TR>
</FORM>

<FORM NAME="calButtons">

<TR HEIGHT=10><TD></TD></TR>

<TR><TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>Su</B></FONT></CENTER></TD>
    <TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>Mo</B></FONT></CENTER></TD>
    <TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>Tu</B></FONT></CENTER></TD>
    <TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>We</B></FONT></CENTER></TD>
    <TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>Th</B></FONT></CENTER></TD>
    <TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>Fr</B></FONT></CENTER></TD>
    <TD><CENTER><FONT SIZE=-1 FACE="Arial,Helv,Helvetica" COLOR="black"><B>Sa</B></FONT></CENTER></TD></TR>

<TR><TD><INPUT TYPE="button" NAME="but00"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but01"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but02"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but03"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but04"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but05"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but06"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD></TR>

<TR><TD><INPUT TYPE="button" NAME="but07"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but08"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but09"  value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but10" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but11" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but12" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but13" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD></TR>

<TR><TD><INPUT TYPE="button" NAME="but14" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but15" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but16" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but17" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but18" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but19" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but20" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD></TR>

<TR><TD><INPUT TYPE="button" NAME="but21" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but22" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but23" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but24" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but25" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but26" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but27" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD></TR>

<TR><TD><INPUT TYPE="button" NAME="but28" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but29" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but30" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but31" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but32" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but33" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but34" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD></TR>

<TR><TD><INPUT TYPE="button" NAME="but35" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but36" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but37" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but38" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but39" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but40" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD>
    <TD><INPUT TYPE="button" NAME="but41" value="    " onClick="returnDate(this.value)" style="WIDTH: 25px"></TD></TR>

</TABLE>
</FORM>
</BODY>
</HTML>
