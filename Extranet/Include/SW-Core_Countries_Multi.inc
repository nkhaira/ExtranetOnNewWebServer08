<%

' --------------------------------------------------------------------------------------

Dim NameOfFormField
Dim Users_Country
Dim StyleAttribute
Dim SelectPrompt
Dim ShowEmbargo

' --------------------------------------------------------------------------------------

Sub DisplayCountryList(NameOfFormField, Users_Country, SelectPrompt, StyleAttribute, ShowEmbargo)

  SQL = "SELECT Country.* FROM Country ORDER BY Country.Abbrev"
  Set rsCountries = Server.CreateObject("ADODB.Recordset")
  rsCountries.Open SQL, conn, 3, 3

' Begin select list with desired name and display item on select list which is initially displayed

	Response.write "<SELECT NAME=""" & NameOfFormField & """ CLASS=""" & StyleAttribute & """ MULTIPLE>" & vbcrlf
  
  if (Trim(Users_Country) = "" or IsNull(Users_Country)) then
    if SelectPrompt <> "" then
    	response.write "<OPTION VALUE="""" SELECTED>" & SelectPrompt & "</OPTION>" & vbcrlf
      response.write "<OPTION VALUE=""""></OPTION>"
    end if      
	end if
  
	Do while not rsCountries.EOF
  
    if ShowEmbargo = false and rsCountries.Fields("Enable") = false then
      
    else
  		response.write "<OPTION VALUE=""" & rsCountries.Fields("Abbrev") & """"
  
	    if instr(1,Users_Country,rsCountries.Fields("Abbrev")) > 0 then
    		response.write " SELECTED>"
      else
        response.write ">"  
  		end if

       response.write rsCountries.Fields("Name") & "</OPTION>" & vbCrLf      

    end if    

		rsCountries.MoveNext
    
	Loop

  Response.write("</SELECT>")

  rsCountries.close
  set rsCountries = nothing
  set SQL = nothing

End Sub


%>