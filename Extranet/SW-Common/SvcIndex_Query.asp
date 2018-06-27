<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

%>

<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->

<%

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

dim partpages
dim extrapage

dim iLimit
Dim iMaxRows
Dim strDoc_Type
Dim iDoc_Num_Min
Dim iDoc_Num_Max
Dim strModel
Dim rsResults
Dim strSQL
Dim dDate_Month
Dim dDate_Day
Dim dDate_Year
Dim iSort
Dim strHighlight
Dim dDate

' --------------------------------------------------------------------------------------
' Get passed in values from the request object
' --------------------------------------------------------------------------------------
iMaxRows = request("Rows")
iLimit = request("limit")
strDoc_Type = request("Doc_Type")
iDoc_Num_Min = request("Doc_Num_Min")
iDoc_Num_Max = request("Doc_Num_Max")
strModel = request("model")
dDate_Month = request("Date_Month")
dDate_Day = request("Date_Day")
dDate_Year = request("Date_Year")
dDate = request("date")
iSort = request("sort")
strHighlight = request("highlight")

' --------------------------------------------------------------------------------------
' Initialize remaining variables
' --------------------------------------------------------------------------------------

if iMaxRows <> "" then
	iMaxRows = Cint(iMaxRows)
else
	iMaxRows = 10
end if
	
if dDate = "" or not IsDate(dDate) then
	if dDate_Month <> "" and dDate_Day <> "" and dDate_Year <> "" then
		if IsNumeric(dDate_Month) and IsNumeric(dDate_Day) and IsNumeric(dDate_Year) then
			dDate = cDate(dDate_Month & "/" & dDate_Day & "/" & dDate_Year)
		end if
	end if
end if

Set rsResults = Server.CreateObject("ADODB.Recordset")


' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

set rsResults = GetResults(strDoc_Type, iDoc_Num_Min, iDoc_Num_Max, strModel, dDate, iSort, strHighlight)

if iLimit = "" or not IsNumeric(iLimit) then
	iLimit = 100
end if

if iLimit > rsResults.recordcount then
	partpages=1
else
	partpages = rsResults.recordcount \ iLimit
	partpages = cint(partpages)
	extrapage = rsResults.recordcount MOD iLimit
	if extrapage > 0 then
		partpages = partpages + 1
	end if
end if
'	if partpages > 25 then
'		partpages = 25
'	end if
	
Session("Perpage") = iLimit
Session("partpages") = partpages

response.redirect ("/SW-Common/SvcIndex_Results.asp?whatpage=1&Doc_Type=" & strDoc_Type & "&Doc_Num_Min=" & iDoc_Num_Min & "&Doc_Num_Max=" & iDoc_Num_Max & "&model=" & strModel & "&date=" & dDate & "&sort=" & iSort & "&highlight=" & strHighlight)

' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------

Function GetResults(strDoc_Type, iDoc_Num_Min, iDoc_Num_Max, strModel, dDate, iSort, strHighlight)
	Dim strSQL
	Dim strFilter

	strFilter = "dummy All Q UD SSU SD MSU MET SPF IM ESU DTE CIS USU GENERAL OBSOLETE"
	strSQL = "SELECT * FROM SVC_Master"

	if strDoc_Type <> "All" and strDoc_Type <> "" then
		Select Case strDoc_Type
			case "OBSOLETE"
				strSQL = strSQL& " WHERE Model='OBSOLETE'"
			qedit = qedit + 1
		case "GENERAL"
			strSQL = strSQL & " WHERE Model='GENERAL-INFORMATION'"
			qedit = qedit + 1
		case else
			strSQL = strSQL & " WHERE svc_master.Doc_Type='" & strDoc_Type & "'"
			qedit = qedit + 1
		end select
	end if
	
	if iDoc_Num_Min <> "" and iDoc_Num_Max = "" and IsNumeric(iDoc_Num_Min) then
		if InStr(strFilter,strDoc_Type) = 0 then
			if qedit > 0 then
				strSQL = strSQL & " AND svc_master.Doc_Num= '" & iDoc_Num_Min & "'"
			else
				strSQL = strSQL & " WHERE svc_master.Doc_Num= '" & iDoc_Num_Min & "'"
			end if
			qedit = qedit + 1
		end if
	end if
	
	if iDoc_Num_Min = "" and iDoc_Num_Max <> "" and IsNumeric(iDoc_Num_Max) then
		if InStr(strFilter,strDoc_Type) = 0 then
			if qedit > 0 then
				strSQL = strSQL & " AND svc_master.Doc_Num= '" & iDoc_Num_Max & "'"
			else
				strSQL = strSQL & " WHERE svc_master.Doc_Num= '" & iDoc_Num_Max & "'"
			end if
			qedit = qedit + 1
		end if
	end if
	
	if iDoc_Num_Min <> "" and iDoc_Num_Max <> "" and IsNumeric(iDoc_Num_Min) and IsNumeric(iDoc_Num_Max) then
		if InStr(strFilter,strDocType) = 0 then
			if qedit > 0 then
				strSQL = strSQL & " AND val(svc_master.Doc_Num)> " & cint(iDoc_Num_Min) & " and val(svc_master.Doc_Num)< " & cint(iDoc_Num_Max)
			else
				strSQL = strSQL & " WHERE val(svc_master.Doc_Num)> " & cint(iDoc_Num_Min) & " and val(svc_master.Doc_Num)< " & cint(iDoc_Num_Max)
			end if
			qedit = qedit + 1
		end if
	end if
		
	if strModel <> "" then
		strModel = changecard(strModel)
		if InStr(strModel,"*") then
			if qedit > 0 then
				strSQL = strSQL & " AND Model LIKE '" & strModel & "'"
			else
				strSQL = strSQL & " WHERE Model LIKE '" & strModel & "'"
			end if
		else
			if qedit > 0 then
				strSQL = strSQL & " AND Model = '" & strModel & "'"
			else
				strSQL = strSQL & " WHERE Model = '" & strModel & "'"
			end if
		end if
		qedit = qedit + 1
	end if
		
	if IsDate(dDate) then
		if qedit > 0 then
			strSQL = strSQL & " AND (Doc_Date > " & dDate & " OR Doc_Rev > " & dDate & ")"
		else
			strSQL = strSQL & " WHERE (Doc_Date > " & dDate & " OR Doc_Rev > " & dDate & ")"
		end if
	end if

	if iSort = "" or iSort = "1" then
	  strSQL = strSQL & " ORDER BY svc_master.Model, svc_master.Doc_Sequence, svc_master.Doc_Num"
	elseif iSort = "2" then
	  strSQL = strSQL & " ORDER BY svc_master.Model, svc_master.Assembly, svc_master.Doc_Sequence, svc_master.Doc_Num"
	elseif iSort = "3" then
	  strSQL = strSQL & " ORDER BY svc_master.Model, svc_master.Doc_Sequence, svc_master.Class, SVC_Master.Assembly"
	end if  
  
'response.write strSQL

	rsResults.open strSQL,conn,3,1,1

	set GetResults = rsResults
	
end function


Function changecard(str)
  dim i
  dim tempstr
  for i = 1 to len(str)
	  if Mid(str,i,1) = "*" then
  		tempstr=tempstr & "%"
  		if i <> 1 then
  			exit for	
  		end if
  		if i = len(str)-1 then
  			exit for
  		end if
    else
      if Mid(str,i,1) = "'" then
    		tempstr=tempstr
     	else
     		tempstr=tempstr & mid(str,i,1)
     	end if
    end if
  next
  changecard = tempstr
end function

' --------------------------------------------------------------------------------------

Call Disconnect_SiteWide
%>