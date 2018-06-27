<%
if AFOVBS_LOADED <> True then
  %><!--#include virtual="/include/adovbs.inc"--><%
end if

' --------------------------------------------------------------------------------------

Dim cmdCountries
Dim rsCountries
Dim strOutput
Dim NameOfFormField
Dim Users_Country
Dim StyleAttribute
Dim SelectPrompt

' --------------------------------------------------------------------------------------

Sub DisplayCountrySelect(Users_Country, StyleAttribute)

  ' This subroutine displays only the OPTION values, you must provide <SELECT> and </SELECT>
  
	'set cmdCountries = Server.CreateObject("ADODB.Command")
	'set cmdCountries.ActiveConnection = dbConnFormData
	'cmdCountries.CommandText = "Forms_GetCountries"
	'cmdCountries.CommandType = adCmdStoredProc

	'set rsCountries = cmdCountries.Execute

  SQLx = "SELECT * FROM Country WHERE Enable=" & CInt(True) & " ORDER BY Name"
  Set rsCountries = Server.CreateObject("ADODB.Recordset")
  rsCountries.Open SQLx, conn, 3, 3
  
' Begin select list with desired name and display item on select list which is initially displayed

	Do while not rsCountries.EOF

    if Users_Country = "" then
    
      if instr(1,rsCountries.Fields("Name"),"United States") = 0 then
      
        if rsCountries.Fields("Abbrev") <> "UK" then
      		response.write "<OPTION VALUE=""" & rsCountries.Fields("Abbrev") & """"

  	    	if rsCountries.Fields("Abbrev") = Users_Country then
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

  	    if rsCountries.Fields("Abbrev") = Users_Country then
      		response.write " SELECTED>"
        else
          response.write ">"  
    		end if

        response.write rsCountries.Fields("Name") & "</OPTION>" & vbCrLf      
      end if  
    end if  

		rsCountries.MoveNext
	Loop

  rsCountries.close
  set rsCountries = nothing
  'set cmdCountries = nothing

End Sub

%>