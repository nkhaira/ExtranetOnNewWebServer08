<%
' K. D. Whitlock
' 05/01/2001
' Based in-part on JavaScript code snipits by R. Mortez

Dim Translations_On
Translations_On = true

dim Site_ID
Site_ID = 0
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_Sitewide     ' Used Only for Translation Purposes

response.write "<HTML>" & vbCrLf
response.write "<HEAD>" & vbCrLf
response.write "<TITLE>"

if Translations_On then
  response.write Translate("Calendar",Login_Language,conn)
else
  response.write "Calendar"
end if
response.write "</TITLE>"

Dim Meta_Charset
Dim Meta_Language
Dim ISO_Code

if IsObject(conn) then
  SQL = "SELECT Language.* FROM Language WHERE Language.Code='" & Login_Language & "'"
  Set rsMeta = Server.CreateObject("ADODB.Recordset")
  rsMeta.Open SQL, conn, 3, 3
  Meta_Charset  = rsMeta("Meta_Charset")
  Meta_Language = rsMeta("Description")
  Meta_ISO_Code = rsMeta("Code")  
  rsMeta.close
  set rsMeta = nothing
end if    

response.write Meta_Charset & vbCrLf
response.write "<META HTTP-EQUIV=""Content-language"" CONTENT=""" & Meta_ISO_Code & """>" & vbCrLf
response.write "<META NAME=""LANGUAGE"" CONTENT=""" & Meta_Language & """>" & vbCrLf
response.write "<META NAME=""AUTHOR"" CONTENT=""K. David Whitlock - David.Whitlock@fluke.com"">" & vbCrLf
%>

<STYLE>
<!--
.hdr  {background:Black;color:#FFCC00;border-width:1px;border-color:black;border-style:solid}
.hdrA {background:Black;color:#FFCC00;border-width:1px;border-color:black;border-style:solid;cursor:hand;}
.ndt  {position:absolute;width:19;height:19;}
.bdt  {position:absolute;width:19;height:19;}
.dt   {position:absolute;width:19;height:19;cursor:hand;}
.sdt  {position:absolute;width:19;height:19;}
 //-->
</STYLE>
</HEAD>

<BODY BGCOLOR=#FFFFFF TEXT=#000000 LINK=#000000 VLINK=#000000 ALINK=#000000 onLoad="DoLoad()">
<BASEFONT FACE="Arial,Helvetica,Geneva,Swiss,Sans Serif">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=1 style="border-width:2px;border-color:black;border-style:solid;font:8pt arial;background:#F1F1F1;">
<TR><TD HEIGHT=20>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=1 HEIGHT=20 style="font: 8pt arial;"><TR>
<TD WIDTH=16 ID=PrevDiv CLASS="hdrA"><IMG ID=Prev SRC="/images/calendar/prev.gif" onClick="PC();" ALT="Show Previous Month"></TD>
<TD WIDTH=101 ALIGN=MIDDLE CLASS="hdr"><SPAN ID=MonthTitle></SPAN>&nbsp;&nbsp;<SPAN ID=YearTitle></SPAN></TD>
<TD WIDTH=16 ID=NextDiv CLASS="hdrA"><IMG ID=Next SRC="/images/calendar/next.gif" onClick="NC();" ALT="Show Next Month"></TD>
</TR></TABLE>
</TD></TR>
<TR><TD>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR><TD><IMG ID=WeekImg SRC="/images/calendar/week1.gif"></TD></TR>
<TR><TD HEIGHT=1 BGCOLOR=#000000></TD></TR>
<TR><TD ALIGN=MIDDLE STYLE="position:relative;">
<IMG ID="SelDate" CLASS="sdt" STYLE="display:none;" SRC="/images/calendar/seldate.gif">
<IMG ID=MonthImg SRC="/images/calendar/Calendar_Blank.gif" STYLE="position:relative;left:0;top:0;">
<DIV ID=BKIMG1>
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
<IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif"><IMG SRC="/images/calendar/date.gif">
</DIV>
<IMG ID="Today" CLASS="ndt" STYLE="display:none;" SRC="/images/calendar/ring.gif" onClick="TC()">
</TD></TR>
<TR><TD HEIGHT=1 BGCOLOR=#000000></TD></TR>
</TABLE>
</TD></TR>
<TR><TD HEIGHT=20 ALIGN=MIDDLE><A STYLE="color: #336699" HREF="javascript:CC()"><%if Translations_On then response.write Translate("Close Calendar",Login_Language,conn) else response.write "Close Calendar"%></A></TD></TR>
</TABLE>

<SCRIPT>
<!--
var g_fCalLoaded=false;
var da=document.all;
var wp=window.parent;
var cf=wp.document.all.CalFrame;
var bdc=da.BKIMG1.children;
var dMin;var dMax;
var XOff=2;var YOff=1;
var XSize=20;var YSize=20;
var g_dC=-1;var g_mC=-1;var g_yC=-1;
var g_dI=-1;var g_mI=-1;var g_yI=-1;

function DoLoad()
{
for(i=0;i<7;i++)
{
	for(j=0;j<6;j++)
	{
		var t=j*7+i;
		bdc[t].day=t+1;
		bdc[t].onclick=BC;
		bdc[t].className="dt";
		bdc[t].style.left=da.MonthImg.offsetLeft+XOff+XSize*i-1;
		bdc[t].style.top=da.MonthImg.offsetTop+YOff+YSize*j;
	}
}
}

function TC()
{
if(event.srcElement.className=="dt")
{
	var dt=new Date();
	wp.SetDate(dt.getDate(),dt.getMonth()+1,dt.getFullYear());
	cf.style.display="none";
}
event.cancelBubble = true;
}

function BC()
{
if(event.srcElement.className=="dt")
{
	var iDay = event.srcElement.day;
	iDay-=GetDOW(1,g_mC,g_yC);
	wp.SetDate(iDay,g_mC,g_yC);
	cf.style.display="none";
}
event.cancelBubble=true;
}

function CC() {cf.style.display="none";}

function NC()
{
if(g_mC==12) SetDate(g_dC,1,g_yC+1);
else SetDate(g_dC,g_mC+1,g_yC);
}

function PC()
{
if(g_mC==1) SetDate(g_dC,12,g_yC-1);
else SetDate(g_dC,g_mC-1,g_yC);
}

function SetInputDate(day,month,year) {g_dI = day;g_mI = month;g_yI = year;}

function FmtTitle(str)
{
var r=str.charAt(0);
for(i=1;i<str.length;i++) r=r+""+str.charAt(i);
//for(i=1;i<str.length;i++) r=r+"&nbsp;"+str.charAt(i);
return "<B>" + r + "</B>";		
}

function SetMinMax(min,max) {dMin=min;dMax=max;}

function SetDate(day, month, year)
{
da.WeekImg.src="/images/calendar/week"+wp.GetDowStart()+".gif";
da.MonthImg.src="/images/calendar/w"+GetDOW(1,month,year)+"d"+GetMonthCount(month,year)+".gif";
da.MonthTitle.innerHTML=FmtTitle(rgMN[month-1]);
da.YearTitle.innerHTML=FmtTitle(year.toString());
var dt=new Date();
var s,n,v,d;

d="none";
if(month==dt.getMonth()+1&&year==dt.getFullYear())
{
	iBox=dt.getDate()+GetDOW(1,dt.getMonth()+1,dt.getFullYear())-1;
	if(ValidDate(dt.getDate(),dt.getMonth()+1,dt.getFullYear())) n="dt";
	else n="bdt";
	da.Today.className=n;

	da.Today.style.left=bdc[iBox].style.left;
	da.Today.style.top=bdc[iBox].style.top;
	d="block";
}
da.Today.style.display=d;

d="none";
if(-1!=g_dI&&month==g_mI&&year==g_yI)
{
	iBox=g_dI+GetDOW(1,g_mI,g_yI)-1;
	da.SelDate.style.left=bdc[iBox].style.left;
	da.SelDate.style.top=bdc[iBox].style.top;
	d="block";
}
da.SelDate.style.display=d;

if( year<dMin.getFullYear() || (year==dMin.getFullYear()&&month<=(dMin.getMonth()+1)) ) {n="hdr";v="hidden";}
else {n="hdrA";v="visible";}
da.PrevDiv.className=n;
da.Prev.style.visibility=v;

if( year>dMax.getFullYear() || (year==dMax.getFullYear()&&month>=(dMax.getMonth()+1)) ) {n="hdr";v="hidden";}
else {n="hdrA";v="visible";}
da.NextDiv.className=n;
da.Next.style.visibility=v;

var i=0;
var iMin=GetDOW(1,month,year);
var iMax=GetMonthCount(month,year)+GetDOW(1,month,year);

for(;i<iMin;i++) {bdc[i].src="/images/calendar/nodate.gif";bdc[i].className="ndt";}
if( year<dMin.getFullYear() || (year==dMin.getFullYear()&&month<(dMin.getMonth()+1)) || year>dMax.getFullYear() || (year==dMax.getFullYear()&&month>(dMax.getMonth()+1)) )
{
	for(;i<iMax;i++) {bdc[i].src="/images/calendar/baddate.gif";bdc[i].className="bdt";}
}
else if(month==(dMin.getMonth()+1))
{
	iBox=dMin.getDate()+GetDOW(1,dMin.getMonth()+1,dMin.getFullYear())-1;
	for(;i<iMax;i++)
	{
		if(i<iBox) {s="/images/calendar/baddate.gif";n="bdt";}
		else {s="/images/calendar/date.gif";n="dt";}
		bdc[i].src=s;bdc[i].className=n;
	}
}
else if(month==(dMax.getMonth()+1))
{
	iBox=dMax.getDate()+GetDOW(1,dMax.getMonth()+1,dMax.getFullYear())-1;
	for(;i<iMax;i++)
	{
		if(i<iBox+1) {s="/images/calendar/date.gif";n="dt";}
		else {s="/images/calendar/baddate.gif";n="bdt";}
		bdc[i].src=s;bdc[i].className=n;
	}
}
else
{
	for(;i<iMax;i++) {bdc[i].src="/images/calendar/date.gif";bdc[i].className="dt";}
}
for(;i<42;i++) {bdc[i].src="/images/calendar/nodate.gif";bdc[i].className="ndt";}

g_dC=day;
g_mC=month;
g_yC=year;
}

function ValidDate(day,month,year)
{
if( year<dMin.getFullYear() || (year==dMin.getFullYear()&&month<(dMin.getMonth()+1)) || (year==dMin.getFullYear()&&month==(dMin.getMonth()+1)&&day<dMin.getDate()) ) return false;
else if( year>dMax.getFullYear() || (year==dMax.getFullYear()&&month>(dMax.getMonth()+1)) || (year==dMax.getFullYear()&&month==(dMax.getMonth()+1)&&day>dMax.getDate()) ) return false;
else return true;
}

function GetMonthCount(month,year)
{
var c=rgMC[month-1];
if((2==month)&&IsLeapYear(year)) c++;
return c;
}

function IsLeapYear(year) {return( 0==year%4 && ((year%100!=0)||(year%400==0)) );}

function GetDOW(day,month,year)
{
var dt=new Date(year,month-1,day);
return (dt.getDay()+(7-wp.GetDowStart()))%7;
}

var rgMN=new Array(12);

rgMN[0] ="<%if Translations_On then response.write Translate("January",Login_Language,conn)   else response.write "January"%>";
rgMN[1] ="<%if Translations_On then response.write Translate("February",Login_Language,conn)  else response.write "February"%>";
rgMN[2] ="<%if Translations_On then response.write Translate("March",Login_Language,conn)     else response.write "March"%>";
rgMN[3] ="<%if Translations_On then response.write Translate("April",Login_Language,conn)     else response.write "April"%>";
rgMN[4] ="<%if Translations_On then response.write Translate("May",Login_Language,conn)       else response.write "May"%>";
rgMN[5] ="<%if Translations_On then response.write Translate("June",Login_Language,conn)      else response.write "June"%>";
rgMN[6] ="<%if Translations_On then response.write Translate("July",Login_Language,conn)      else response.write "July"%>";
rgMN[7] ="<%if Translations_On then response.write Translate("August",Login_Language,conn)    else response.write "August"%>";
rgMN[8] ="<%if Translations_On then response.write Translate("September",Login_Language,conn) else response.write "September"%>";
rgMN[9] ="<%if Translations_On then response.write Translate("October",Login_Language,conn)   else response.write "October"%>";
rgMN[10]="<%if Translations_On then response.write Translate("November",Login_Language,conn)  else response.write "November"%>";
rgMN[11]="<%if Translations_On then response.write Translate("December",Login_Language,conn)  else response.write "December"%>";

var rgMC=new Array(12);
rgMC[0]=31;
rgMC[1]=28;
rgMC[2]=31;
rgMC[3]=30;
rgMC[4]=31;
rgMC[5]=30;
rgMC[6]=31;
rgMC[7]=31;
rgMC[8]=30;
rgMC[9]=31;
rgMC[10]=30;
rgMC[11]=31;

g_fCalLoaded=true;
//-->
</SCRIPT>

</BODY>
</HTML>

<%
 Call Disconnect_SiteWide
%>

