<%@ Language="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->

<% 
Call Connect_SiteWide

sql = "select procs.procedure_id [Procedure ID], Instrument, AdjThreshold, authors.description [Author], " &_
    	"company.description [Company], procs.Date, primcal.Description [Primary Calibrator], Revision, " &_
    	"type.description [Type], [5500CAL_Ready], Restricted, In_Revision, Buy, source.description [Source], " &_
    	"pricepoint.description [Price Point], ZipFileName, CreateDate, UpdateBy, UpdateDate " &_
      "from metcal_procedures procs " &_
    	"inner join metcal_categories authors " &_
    	"on procs.author_id = authors.metcal_category_id " &_
    	"inner join metcal_categories company " &_
  		"on procs.company_id = company.metcal_category_id " &_
    	"inner join metcal_categories primcal " &_
  		"on procs.primcalibrator_id = primcal.metcal_category_id " &_
	    "inner join metcal_categories type " &_
  		"on procs.type_id = type.metcal_category_id " &_
	    "inner join metcal_categories source " &_
  		"on procs.source_id = source.metcal_category_id " &_
	    "inner join metcal_categories pricepoint " &_
  		"on procs.price_point_id = pricepoint.metcal_category_id"

'response.write "sql: " & sql & "<BR>"
'response.end

set rsDB = conn.execute(sql)

with response

  if rsDB.EOF then
    set rsDB= nothing
    %><HTML><BODY>No results found<P><%=Replace(sql,vbcrlf,"<BR>"&vbcrlf)%></BODY></HTML><%

  else

    .Buffer = true
    .Clear  
    .ContentType = "application/vnd.ms-excel;charset=" & charset
    
  	.write "<TABLE BORDER=1>"
    .write "<TR>"

    for each oField in rsDB.Fields
      .write "<TD VALIGN=""TOP"" BGCOLOR=""#99CCFF""><B>" & oField.name & "</B></TD>"
    next

    .write "</TR>"

    Dim icounter
    icounter = 0

  	do while not rsDB.eof 'and iCounter < 558
	  	.write "<TR>"
  		for each oField in rsDB.Fields
	  		if uCase(oField.name) <> "DESCRIPTION" then
		  		.write "<TD VALIGN=""TOP"">" & rsDB(oField.name) & "" & "</TD>"
  			end if
  		next
    
	  	.write "</TR>"
		
      rsDB.MoveNext
      icounter= icounter + 1
  	loop

	  .write "</TABLE>"
    
  end if

  .flush

end with

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------

function CleanupDescription(strDescription)
	if InStr(1, uCase(strDescription), "<BODY>") then
		strDescription = mid(strDescription, InStr(1, uCase(strDescription), "<BODY>") + 6)
	end if

	if InStr(1, uCase(strDescription), "</BODY>") then
		strDescription = mid(strDescription, 1, InStr(1, uCase(strDescription), "</BODY>") - 1)
	end if

	cleanupDescription = strDescription
end function
%>