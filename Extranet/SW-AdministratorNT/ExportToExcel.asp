<%@ Language="VBScript" CODEPAGE="65001" %>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
' --------------------------------------------------------------------------------------
' Author:     P.Deshpande.
' Date:       06/1/2007
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------
on error resume next
Response.Buffer = false
Session.timeout = 60            ' Set to 1 Hour
Server.ScriptTimeout = 5 * 60   ' Set to 5 Minutes
'Response.end 
Call Connect_SiteWide

dim objcmd,rs
set objcmd=server.createobject("ADODB.Command")
set objcmd.ActiveConnection=conn
objcmd.CommandText="getassets"

Set objPara1 = objcmd.CreateParameter("@siteid", 3, 1)
Set objPara2 = objcmd.CreateParameter("@language",129, 1,3)
Set objPara3 = objcmd.CreateParameter("@categoryid", 3, 1)
Set objPara4 = objcmd.CreateParameter("@GroupId", 3, 1)
Set objPara5 = objcmd.CreateParameter("@Country", 3, 1)
Set objPara6 = objcmd.CreateParameter("@Submitted_By", 3, 1)
Set objPara7 = objcmd.CreateParameter("@Campaign", 3, 1)
Set objPara8 = objcmd.CreateParameter("@SortBy", 3, 1)

objcmd.Parameters.append objPara1
objcmd.Parameters.append objPara2
objcmd.Parameters.append objPara3
objcmd.Parameters.append objPara4
objcmd.Parameters.append objPara5
objcmd.Parameters.append objPara6
objcmd.Parameters.append objPara7
objcmd.Parameters.append objPara8

objPara1.value  = Request.QueryString("Site_ID")
if Request.QueryString("language") <> "" then
    objPara2.value = Request.QueryString("language")
end if
            
if Request.QueryString("categoryid") > 0 then
    objPara3.value = Request.QueryString("categoryid") 
end if

if Request.QueryString("GroupId") > 0 then
    objPara4.value = Request.QueryString("GroupId") 
end if

if Request.QueryString("Country") <> "" then
     objPara5.value = Request.QueryString("Country") 
end if
 
if Request.QueryString("Submitted_By") > 0 then
     objPara6.value = Request.QueryString("Submitted_By") 
end if

if Request.QueryString("Campaign") > 0 then
     objPara7.value = Request.QueryString("Campaign") 
end if

if Request.QueryString("SortBy") > 0 then
     objPara8.value = Request.QueryString("SortBy") 
end if
  
'Response.AddHeader "Content-Disposition", "attachment; filename=Asset_List.xls" 
'Response.Charset     = "utf-16"
Set Rs = objcmd.Execute()
'dim rsUserClone
dim iUserfieldCount
'set rsUserClone=Server.CreateObject("adodb.recordset")
'rsUserClone.CursorLocation=3
'for iUserfieldCount=0 to Rs.fields.count-1
'	rsUserClone.Fields.append Rs.fields(iUserfieldCount).name,200,500,64
'next
'rsUserClone.Open()
Asset_ID_Old = 0
Asset_Record_Count = 0
Dim sqlHtml
sqlHtml = ""
sqlHtml = sqlHtml & "<table border=1>"
sqlHtml = sqlHtml &  "<tr>"
for iCol=0 to 20
				sqlHtml = sqlHtml & "<td ><b>" & Rs.fields(iCol).Name & "</b></td>"
next
sqlHtml = sqlHtml &  "</tr>"
do while not Rs.EOF
	if Rs.fields("ID").value <> Asset_ID_Old then
		'Asset_Record_Count = Asset_Record_Count + 1
		Asset_ID_Old = Rs("ID")
		'rsUserClone.AddNew
		'for iUserfieldCount=0 to Rs.fields.count-1
		'	rsUserClone.Fields(iUserfieldCount).Value = Rs.fields(iUserfieldCount).value & ""
		'next
		'rsUserClone.Update
  '		while not Rs.eof
			sqlHtml = sqlHtml &  "<tr>"
			for iCol=0 to 20
				sqlHtml = sqlHtml &  "<td>" & Server.HTMLEncode(Rs.fields(iCol).Value & "") & "</td>"
			next
			sqlHtml = sqlHtml &  "</tr>"
			'Rs.movenext
'		wend
	end if
	Rs.MoveNext
loop
sqlHtml = sqlHtml &  "</table>"
Rs.close
'Response.Write sqlHtml
Response.write "Done"
'set Rs = rsUserClone
'Response.Write rs.recordcount
'response.End
'Rs.movefirst
'if Rs.eof <> true then
	
'end if

Set objPara1 = nothing
Set objPara2 = nothing
Set objPara3 = nothing
Set objPara4 = nothing
Set objPara5 = nothing
Set objPara6 = nothing
Set objPara7 = nothing
Set objPara8 = nothing

set rs=nothing
set objcmd = nothing
if err.number <> 0 then
	Response.Write err.Description
end if

%>
