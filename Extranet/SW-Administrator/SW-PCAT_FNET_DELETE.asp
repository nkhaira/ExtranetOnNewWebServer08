<!--#include virtual="/sw-administrator/SW-PCAT_FNET_IISSERVER.asp"-->
<%
on error resume next
if clng(oFileUpEE.form("pidDelete"))> 0 then
	'Modified for Pcat-Asset Relationship on 26-01-2006
	'PcatValidateSql="select * from calendar where clone=" & oFileUpEE.form("ID")  & _
	'" and site_id= " & CInt(Site_ID)
	''''Validation for clones.
	'set rsValidate=conn.execute(PcatValidateSql)
	'if not(rsValidate.eof) then
	'	showdeleteerror("Child records already exists for this record.Unable to delete the record.<br>")
	'	rsValidate.close
	'	set rsvalidate=nothing
	'end if
	'rsValidate.close
	'set rsvalidate=nothing

	PcatValidateSql="select * from calendar where PID=(select distinct PID from calendar where id=" & oFileUpEE.form("ID") & ")"  & _
	" and language ='" & oFileUpEE.form("oldLanguage") & "'" & " and id <> " & oFileUpEE.form("ID") & " and site_id= " & CInt(Site_ID) & _
	" and clone = " & oFileUpEE.form("ID")
	'Response.Write PcatValidateSql
	'response.End
	'''Validation for Locale and catalog.
	set rsValidate=conn.execute(PcatValidateSql)

	if not(rsValidate.eof) then
		showdeleteerror("Change the language of this record first and then try to delete.Unable to delete the record.<br>")
		rsValidate.close
		set rsvalidate=nothing
	end if
	rsValidate.close


	set objAssetDelete=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
	'Pass the actual url here
	call objAssetDelete.open("POST",striisserverpath,0,0,0)
	call objAssetDelete.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	
	if clng(oFileUpEE.form("Clone"))=clng(oFileUpEE.form("ID")) then
		strclone = false
		DeleteAsset=true
	else
		strclone=true
		DeleteAsset=false
	end if 

	set rsLanguage=conn.execute("select iso2 from language where code='" & oFileUpEE.form("oldLanguage") & "'")
	if not(rsLanguage.eof) then
		strLanguage=rsLanguage.fields(0).value
	end if
	rsLanguage.close
	set rsLanguage=nothing
	               
	strparameters="operation=D" & "&isclone=" & strclone & "&assetpid=" & oFileUpEE.form("pidDelete") & _
	"&language=" & strLanguage & "&DeleteAll=" & DeleteAsset & "&setRelationship=" & DeleteAsset & _
	"&itemNumber=" & oFileUpEE.form("oldItemNumber") & "&SiteID=" & Site_ID
	call objAssetDelete.send(strparameters)

	if err.number<>0 then
		showdeleteerror("Unable to delete the records<br>" & err.Description & "<br>" )   
	end if
	on error goto 0
end if
sub showdeleteerror(strmessage)
	BackURL= "/sw-administrator/default.asp?Site_ID=" & Site_ID & "&ID=edit_record&Category_ID=" & oFileUpEE.form("Category_ID")
	response.write "<HTML>" & vbCrLf
	response.write "<HEAD>" & vbCrLf
	response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
	response.write "<TITLE>Error</TITLE>" & vbCrLf
	response.write "</HEAD>"
	response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
	response.write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """>" & vbCrLf
	response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """>" & vbCrLf
	response.write "<DIV ALIGN=CENTER>"
	Call Nav_Border_Begin
	response.write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
	'response.write "If your browser does not automatically return to the edit screen<BR>within 5 seconds, click on the [Continue] link below.<P>"
	Response.Write strmessage & "<br>"
'    Response.Write err.Description & "<br>"
	response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
	response.write "</TD></TR></TABLE>" & vbCrLf
	Call Nav_Border_End
	response.write "</FORM>" & vbCrLf
	response.write "</DIV>"
	response.write "</BODY>"
	response.write "</HTML>"
	'response.flush
	on error goto 0
	Response.End
end sub
'>>>>>>>>>>>
%>