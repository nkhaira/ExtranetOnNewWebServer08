<%
  stuff = Array(56,100,23.333,87,22,77,10)
  labelstuff = array("label1","label2","label3","label4","label5","label6","label7")
  Response.Write makechart("Site Activity",stuff,labelstuff,null, "#F6F6F6",0, 200,10,false)

function makechart(title, numarray, labelarray, color, bgcolor, bordersize, maxheight, maxwidth, addvalues)
     
 dim tablestring 
 'max value is maximum table value
 dim max 
 'maxlength maximum length of labels
 dim maxlength
 dim tempnumarray
 dim templabelarray
 dim heightarray
 Dim colorarray
 'value to multiplie chart values by to get relative size 
 Dim multiplier
 'if data valid
 if maxheight > 0 and maxwidth > 0 and ubound(labelarray) = ubound(numarray) then
  'colorarray: color of each bars if more bars then colors loop through
  colorarray = array("red","blue","yellow","navy","orange","purple","green")
  templabelarray = labelarray
  tempnumarray = numarray
  heightarray = array()
  max = 0
  maxlength = 0
  tablestring = "<TABLE bgcolor='" & bgcolor & "' border='" & bordersize & "'>" & _
    "<tr><td><TABLE border='0' cellspacing='1' cellpadding='0'>" & vbCrLf
  'get maximum value
  for each stuff in tempnumarray
   if stuff > max then max = stuff end if 
  next
  'calculate multiplier
  multiplier = maxheight/max
  'populate array
  for counter = 0 to ubound(tempnumarray)
   if tempnumarray(counter) = max then 
    redim preserve heightarray(counter)
    heightarray(counter) = maxheight
   else
    redim preserve heightarray(counter) 
    heightarray(counter) = tempnumarray(counter) * multiplier 
   end if 
  next 

  'set title 
  tablestring = tablestring & "<TR><TH colspan='" & ubound(tempnumarray)+1 & "'>" & _
     "<FONT FACE='Verdana, Arial, Helvetica' SIZE='1'><U>" & title & "</TH></TR>" & _
      vbCrLf & "<TR>" & vbCrLf
    'loop through values
  for counter = 0 to ubound(tempnumarray) 
    tablestring = tablestring & vbTab & "<TD valign='bottom' align='center' >" & _
    "<FONT FACE='Verdana, Arial, Helvetica' SIZE='1'>" & _
    "<table border='0' cellpadding='0' width='" & maxwidth & "'><tr>" & _
    "<tr><td valign='bottom' bgcolor='" 
    if not isNUll(color) then 
     'if color present use that color for bars
     tablestring = tablestring & color
    else
     'if not loop through colorarray
     tablestring = tablestring & colorarray(counter mod (ubound(colorarray)+1))
    end if
    tablestring = tablestring & "' height='" & _
     round(heightarray(counter),2) & "'><img src='chart.gif' width='1' height='1'>" & _
     "</td></tr></table>"
    if addvalues then
     'print actual values
     tablestring = tablestring & "<BR>" & tempnumarray(counter)
    end if 
    tablestring = tablestring & "</TD>" & vbCrLf
  next
 
  tablestring = tablestring & "</TR>" & vbCrLf
  'calculate max lenght of labels
  for each stuff in labelarray
   if len(stuff) >= maxlength then maxlength = len(stuff)
  next
  'print labels and set each to maxlength
  for each stuff in labelarray
   tablestring = tablestring & vbTab & "<TD align='center'><" & _
    "FONT FACE='Verdana, Arial, Helvetica' SIZE='1'><B> " 
   for count = 0 to round((maxlength - len(stuff))/2)
    tablestring = tablestring & " "
   next
   if maxlength mod 2 <> 0 then tablestring = tablestring & " "
   tablestring = tablestring & stuff 
   for count = 0 to round((maxlength - len(stuff))/2)
    tablestring = tablestring & " "
   next
   tablestring = tablestring & " </TD>" & vbCrLf
  next
   
  tablestring = tablestring & "</TABLE></td></tr></table>" & vbCrLf
  makechart = tablestring
 else
  Response.Write "Error Function Makechart: maxwidth and maxlength have to be greater " & _
  " then 0 or number of labels not equal to number of values"
 end if 
end function
%>
