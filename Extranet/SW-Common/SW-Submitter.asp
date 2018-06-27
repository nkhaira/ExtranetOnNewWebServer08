  <%
  
  ' Content Submitter Administrator -- Add Content or Event

  response.write "<FORM NAME=""SW-Submitter"">"
  response.write "<TABLE WIDTH=""100%"" BORDER=1 BORDERCOLOR=""GRAY"" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>" & vbCrLf
  response.write "  <TR>" & vbCrLf
  response.write "    <TD WIDTH=""100%"" BGCOLOR=""#FFCC00"">" & vbCrLf
  response.write "      <TABLE WIDTH=""100%"" CELLPADDING=4 BORDER=0>" & vbCrLf
  response.write "        <TR>" & vbCrLf
  response.write "          <TD BGCOLOR=""#FFCC00"" WIDTH=""50%"" CLASS=Normal>" & vbCrLf
  response.write "            <B>Add</B> - Content or Event into Category:" & vbCrLf
  response.write "          </TD>" & vbCrLf
  response.write "          <TD BGCOLOR=""#FFCC00"" WIDTH=""50%"" CLASS=Normal>" & vbCrLf
  response.write "            <SELECT CLASS=Normal LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='/sw-administrator/Calendar_edit.asp?ID=add&Site_ID=" & Site_ID & "&Category_ID='+this.options[this.selectedIndex].value"" NAME=""Category_ID"">" & vbCrLf
  
  SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=-1" & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title" & vbCrLf
  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.Open SQL, conn, 3, 3
      
  response.write "              <OPTION VALUE="""">Select from list</OPTION>" & vbCrLf
                                
  Do while not rsCategory.EOF
    response.write "            <OPTION VALUE=""" & rsCategory("ID") & """>" & RestoreQuote(rsCategory("Title")) & "</OPTION>" & vbCrLf
    rsCategory.MoveNext 
  loop
              
  rsCategory.close
  set rsCategory=nothing

  response.write "            </SELECT>" & vbCrLf
  response.write "          </TD>" & vbCrLf
  response.write "        </TR>" & vbCrLf

  ' Content Submitter Administrator -- Edit Content or Event

  response.write "        <TR>" & vbCrLf

  if request("ID") = "edit_record" then
    response.write "        <TD BGCOLOR=""Black"" WIDTH=""50%"" CLASS=NormalGold>" & vbCrLf
  else
    response.write "        <TD BGCOLOR=""#FFCC00"" WIDTH=""50%"" CLASS=Normal>" & vbCrLf
  end if

  response.write "            <B>Edit</B> - Content or Event from Category:" & vbCrLf
  response.write "          </TD>" & vbCrLf

  if request("ID") = "edit_record" then
    response.write "        <TD BGCOLOR=""Black"" WIDTH=""50%"" CLASS=NormalGold>" & vbCrLf
  else

    'Check Category Items/Events Pending Approval
    
    SQL =       "SELECT Calendar.* "
    SQL = SQL & "FROM Calendar "
    SQL = SQL & "WHERE Calendar.Site_ID=" & CInt(Site_ID)
    
    select case Access_Level
      Case 4,8,9
        SQL = SQL & " AND Calendar.Status=0" & " AND Calendar.Review_By=" & Login_ID
      Case 2
        SQL = SQL & " AND Calendar.Status=0" & " AND Calendar.Submitted_By=" & Login_ID
    end select    
  
    Set rsApproval = Server.CreateObject("ADODB.Recordset")
    rsApproval.Open SQL, conn, 3, 3  
  
    if not rsApproval.EOF then
      response.write "      <TD BGCOLOR=""#FF0000"" WIDTH=""50%"" CLASS=Normal>" & vbCrLf
    else
      response.write "      <TD BGCOLOR=""#FFCC00"" WIDTH=""50%"" CLASS=Normal>" & vbCrLf
    end if
    
    rsApproval.close
    set rsApproval = nothing  

  end if

  response.write "            <SELECT CLASS=Normal LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='default.asp?ID=edit_record&Site_ID=" & Site_ID & "&Category_ID='+this.options[this.selectedIndex].value"" NAME=""Category_ID"">" & vbCrLf

  if request("ID") <> "edit_record" then    
    response.write "            <OPTION VALUE="""">Select from list</OPTION>" & vbCrLf
  end if
              
  select case Access_Level
    case 4,6,8,9              

    SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=-1" & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
    Set rsCategory = Server.CreateObject("ADODB.Recordset")
    rsCategory.Open SQL, conn, 3, 3
                                    
    Do while not rsCategory.EOF
      if request("ID") = "edit_record" and CInt(request("Category_ID")) = rsCategory("ID") then
        response.write "        <OPTION SELECTED VALUE=""" & rsCategory("ID") & """>" & RestoreQuote(rsCategory("Title")) & "</OPTION>" & vbCrLf
      else
        response.write "        <OPTION VALUE=""" & rsCategory("ID") & """>" & RestoreQuote(rsCategory("Title")) & "</OPTION>" & vbCrLf
      end if                
      rsCategory.MoveNext 
    loop
                  
    rsCategory.close
    set rsCategory=nothing            
                                    
  end select
              
  if request("ID") = "edit_record" and CInt(request("Category_ID")) = 9998 then
    if Access_Level = 2 then
      response.write "          <OPTION VALUE=""9998"" SELECTED>View Submit Queue</OPTION>" & vbCrLf
    else
      response.write "          <OPTION VALUE=""9998"" SELECTED>View Approval Queue</OPTION>" & vbCrLf
    end if                  
  else
    if Access_Level = 2 then              
      response.write "          <OPTION VALUE=""9998"" CLASS=NavLeftHighlight1>View Submit Queue</OPTION>" & vbCrLf
    else
      response.write "          <OPTION VALUE=""9998"" CLASS=NavLeftHighlight1>View Approval Queue</OPTION>" & vbCrLf
    end if  
  end if

  if Access_Level >= 8 and request("ID") = "edit_record" and CInt(request("Category_ID")) = 9999 then
    response.write "            <OPTION VALUE=""9999"" SELECTED>View Approval Queue - All</OPTION>" & vbCrLf
  elseif Access_Level >= 8 and request("ID") = "edit_record" then
    response.write "            <OPTION VALUE=""9999"" CLASS=NavLeftHighlight1>View Approval Queue - All</OPTION>" & vbCrLf
  end if

  response.write "            </SELECT>" & vbCrLf
  response.write "          </TD>" & vbCrLf
  response.write "        </TR>" & vbCrLf
  response.write "      </TABLE>" & vbCrLf
  response.write "    </TD>" & vbCrLf
  response.write "  </TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf
  response.write "</FORM>" & vbCrLf
  
%>