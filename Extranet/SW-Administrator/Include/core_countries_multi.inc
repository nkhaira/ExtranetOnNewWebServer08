<!--#include virtual="/include/adovbs.inc"-->
<%
' --------------------------------------------------------------------------------------

Dim cmdCountries
Dim rsCountries
Dim strOutput
Dim NameOfFormField
Dim Users_Country
Dim StyleAttribute
Dim SelectPrompt

' --------------------------------------------------------------------------------------

Sub DisplayCountryList(NameOfFormField, Users_Country, SelectPrompt, StyleAttribute)

	'set cmdCountries = Server.CreateObject("ADODB.Command")
	'set cmdCountries.ActiveConnection = dbConnFormData
	'cmdCountries.CommandText = "Forms_GetCountries"
	'cmdCountries.CommandType = adCmdStoredProc

	'set rsCountries = cmdCountries.Execute

  SQL = "SELECT * FROM Country WHERE Enable=" & CInt(True) & " ORDER BY Name"
  Set rsCountries = Server.CreateObject("ADODB.Recordset")
  rsCountries.Open SQL, conn, 3, 3

' Begin select list with desired name and display item on select list which is initially displayed

	Response.write "<SELECT SIZE=10 NAME=""" & NameOfFormField & """ CLASS=""" & StyleAttribute & """ MULTIPLE>" & vbcrlf
  
  if Trim(Users_Country) = "" or IsNull(Users_Country) then
    if SelectPrompt <> "" then
    	response.write "<OPTION VALUE="""" SELECTED>" & SelectPrompt & "</OPTION>" & vbcrlf
      response.write "<OPTION VALUE=""""></OPTION>"
    end if      
    response.write "<OPTION VALUE=""US"">United States</OPTION>"
    response.write "<OPTION VALUE=""UM"">United States Minor Outlying Islands</OPTION>"
    response.write "<OPTION VALUE=""""></OPTION>"
	end if
  
	Do while not rsCountries.EOF

    if Users_Country = "" then
    
      if instr(1,rsCountries.Fields("Name"),"United States") = 0 then
      
        if rsCountries.Fields("Abbrev") <> "UK" then
      		response.write "<OPTION VALUE=""" & rsCountries.Fields("Abbrev") & """"

  	    	if Instr(1,User_Country,rsCountries.Fields("Abbrev")) > 0 then
      			response.write " SELECTED>"
          else
            response.write ">"  
    		  end if

          response.write rsCountries.Fields("Name") & "</OPTION>" & vbCrLf
        end if  
      end if
    else
      if rsCountries.Fields("Abbrev") <> "UK" then
    		response.write "<OPTION VALUE=""" & rsCountries.Fields("Abbrev") & """"

  	    if instr(1,Users_Country,rsCountries.Fields("Abbrev")) > 0 then
      		response.write " SELECTED>"
        else
          response.write ">"  
    		end if

        response.write rsCountries.Fields("Name") & "</OPTION>" & vbCrLf      
      end if  
    end if  

		rsCountries.MoveNext
	Loop

  Response.write("</SELECT>")

  rsCountries.close
  set rsCountries = nothing
  'set cmdCountries = nothing

End Sub

' --------------------------------------------------------------------------------------

Sub displayCountryTable(StyleAttribute)

	'set cmdCountries = Server.CreateObject("ADODB.Command")
	'set cmdCountries.ActiveConnection = dbConnFormData
	'cmdCountries.CommandText = "Forms_GetCountries"
	'cmdCountries.CommandType = adCmdStoredProc

	'set rsCountries = cmdCountries.Execute

  SQL = "SELECT * FROM Country WHERE Enable=" & CInt(True) & " ORDER BY Name"
  Set rsCountries = Server.CreateObject("ADODB.Recordset")
  rsCountries.Open SQL, conn, 3, 3

	Response.write "<TABLE>"
	
	Do while Not rsCountries.EOF
    response.write "<TR>"
		response.write "<TD CLASS=""" & StyleAttribute & """>" & rsCountries.Fields("Abbrev") & "</TD>"
		response.write "<TD CLASS=""" & StyleAttribute & """>" & rsCountries.Fields("Name") & "</TD>"
    response.write "</TR>"
		rsCountries.MoveNext
	Loop

  Response.write "</TABLE><BR><BR>"

  rsCountries.close
  set rsCountries = nothing
  'set cmdCountries = nothing

End Sub

%>