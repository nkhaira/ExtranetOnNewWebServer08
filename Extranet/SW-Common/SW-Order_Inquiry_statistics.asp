<%@ Language="VBScript" CODEPAGE="65001" %>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<%
' --------------------------------------------------------------------------------------
' Author:     P. Barbee
' Date:       3/18/2002
'
' Shows statistics for Order Inquiry usage.  This page is viewable by anyone logged onto
' the extranet by only navigable if you have the Site Statistics button - it is not translated
' --------------------------------------------------------------------------------------

response.buffer = true

Dim ErrMessage
ErrMessage = ""

Call Connect_SiteWide

Dim bShowAll
if Request.QueryString("all") = "yes" then
	bShowAll = True
	ShowText = "all sites"
else
	bShowAll = False
	ShowText = "this site"
end if

if Session("Logon_user") <> "" then
	%>
	<!-- #include virtual="/SW-Common/SW-Security_Module.asp" -->
	<%
else
  response.redirect "/register/default.asp"
	site_id = 3
end if
%>
<!-- #include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

' --------------------------------------------------------------------------------------
' Start building the page
' --------------------------------------------------------------------------------------
Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Content_width	  ' Percent

Screen_Title    = Site_Description & " - " & Translate("Order Inquiry Statisics",Login_Language,conn)
Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & Translate("Order Inquiry Statisics",Login_Language,conn) & "</SPAN>"
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95

BackURL = Session("BackURL")
%>

<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->

<%

response.write "<SPAN CLASS=Heading3>Order Inquiry Statisics for " & ShowText
response.write "</SPAN><P>" & vbcrlf

' Substitute Navigation (because we'r using ...No-Navigation)
with Response
	.write "<SPAN CLASS=SmallBold>"
	.write "<A HREF=""" & Request("BackURL") & """>"
	.write Translate("Home",Login_Language,conn) & "</a> | " & vbcrlf
end with

if bShowAll then
	response.write "<A HREF=""SW-Order_Inquiry_statistics.asp"">Statistics for this site</A>" & vbcrlf
else
	response.write "<A HREF=""SW-Order_Inquiry_statistics.asp?all=yes"">Statistics for all sites</A>" & vbcrlf
end if

response.write "</span><P>" & vbcrlf

' now generate the report they want

sSql = "select year(view_time) as myyear" & vbcrlf &_
	", datepart(wk,view_time) as mywk" & vbcrlf &_
	",count(*) as mycnt" & vbcrlf &_
	"from activity" & vbcrlf &_
	"where cid = 9007" & vbcrlf &_
	"and scid = 1" & vbcrlf

if not bShowAll then
	sSql = sSql & "and site_id = " & site_id & vbcrlf
end if
	
sSql = sSql & "group by year(view_time),datepart(wk,view_time)" & vbcrlf &_
	"order by myyear desc, mywk desc"

set dbRS = conn.Execute(sSql)

'Response.write "SQL " & replace(sSql,vbcrlf,"<BR>"&vbcrlf) & vbcrlf

if dbRS.EOF then
	response.write "Damn, no results!"
else
	' put the data into an array
	i = 0
	MaxCnt = 0
	do until dbRS.EOF
		ReDim Preserve Data(i)
		cnt = dbRS("mycnt") 
		Data(i) = dbRS("myyear") & ":" & dbRS("mywk") & ":" & cnt
		if cnt > MaxCnt then
			MaxCnt = cnt
		end if
		i = i + 1
		dbRS.MoveNext
	loop
	dbRS.close
	
	' display the array
	
	' what's the first day of this week?
	dSunday = DateValue(DateAdd("d",(1-WeekDay(Date)),Date))
	' what is our scaling value? 300 pixels is max
	MyScale = Cdbl(300/MaxCnt)
	
	response.write "Number of Order Inquiry results per week starting<BR>" & vbcrlf
	response.write "<TABLE>" & vbcrlf
	for each var in Data
		arrt = split(var,":")
		cnt = arrt(2)
		Response.write "  <TR><TD>" & dSunday & "</td><TD ALIGN=""RIGHT"">" & cnt & "</td>"
		Response.write "<TD ALIGN=""LEFT""><HR SIZE=5 COLOR=""#FFCC00"" WIDTH=" & Cint(cnt*MyScale)
		Response.write "></td></tr>" & vbcrlf
		
		' decrement the date NOTE - this depends on weeks being sequential...
		dSunday = DateAdd("d",-7,dSunday)
	Next
	response.write "</table>" & vbcrlf
end if
set dbRS = nothing
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

Call Disconnect_SiteWide
%>

