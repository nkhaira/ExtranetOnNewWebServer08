<%

' PID_System variable, Integer, 0=FNET, 1=FIND

if not IsNumeric(Calendar_ID) then
  ' Insert Code to get New PID here
  PID_Value = "4444"
  
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""PCat"" VALUE=""" & PID_Value & """>" & vbCrLf

else
  response.write "<INPUT TYPE=""HIDDEN"" NAME=""PCat"" VALUE=""" & rs("PID") & """>" & vbCrLf  
end if

' Build PCAT Relationship Fields
with response

  ' Category
  ' Field Title
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
  .write Translate("PCat Category",Login_Language,conn) & ":"
  .write "</TD>" & vbCrLf

  ' Required Icon or Space
  .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
  .write      "&nbsp;" ' Or required Icon
  .write "</TD>"
  
  ' Field
  .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"

  if IsNumeric(Calendar_ID) then    ' Edit Mode
    .write "Edit Code Goes Here"
  else
    .write "Add Code Goes Here"
  end if

  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf

  ' Product Relationship
  ' Field Title
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
  .write Translate("PCat Products",Login_Language,conn) & ":"
  .write "</TD>" & vbCrLf

  ' Required Icon or Space
  .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
  .write      "&nbsp;" ' Or required Icon
  .write "</TD>"
  
  ' Field
  .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"

  if IsNumeric(Calendar_ID) then    ' Edit Mode
    .write "Edit Code Goes Here"
  else
    .write "Add Code Goes Here"
  end if

  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf

  ' Locale
  ' Field Title
  .write "<TR>" & vbCrLf
  .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
  .write Translate("PCat Locale",Login_Language,conn) & ":"
  .write "</TD>" & vbCrLf

  ' Required Icon or Space
  .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
  .write      "&nbsp;" ' Or required Icon
  .write "</TD>"
  
  ' Field
  .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"
  
  if IsNumeric(Calendar_ID) then    ' Edit Mode
    .write "Edit Code Goes Here"
  else
    .write "Add Code Goes Here"
  end if
  
  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  
end with
%>

