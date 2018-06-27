<%

with response
  .write "<!--" & vbCrLf
  .write "'===========================================================================" & VbCrLF
  .write "' Site ID                 = " & request("Site_ID") & vbCrLf
  .write "' Site Code               = " & request("Site_Code") & vbCrLf
  .write "' Site Description        = " & request("Site_Description") & vbCrLf
  .write "' Name                    = " & request("Prefix") & " " & request("FirstName") & " " & request("LastName") & vbCrLf
  .write "' Job Title               = " & request("Job_Title") & vbCrLf  
  .write "' Company                 = " & request("Company") & vbCrLf
  .write "' MailStop                = " & request("Business_MailStop") & vbCrLf
  .write "' Address 1               = " & request("Business_Address") & vbCrLf
  .write "' Address 2               = " & request("Business_Address_2") & vbCrLf
  .write "' City                    = " & request("Business_City") & vbCrLf
  .write "' State                   = " & request("Business_State") & vbCrLf
  .write "' State Other             = " & request("Business_State_Other") & vbCrLf
  .write "' Postal Code             = " & request("Business_Postal_Code") & vbCrLf
  .write "' Country                 = " & request("Business_Country") & vbCrLf
  .write "' EMail                   = " & request("EMail") & vbCrLf
  .write "' Language                = " & request("Language") & vbCrLf
  .write "' Registration Date       = " & request("Reg_Request_Date") & VbCrLf
  .write "' Promotion Complete URL  = " & request("Promotion_Complete_URL") & VbCrLf  
  .write "'===========================================================================" & VbCrLF
  .write "-->" & vbCrLf & vbCrLf
end with
  
response.write "<HTML><TITLE></TITLE><BODY BGCOLOR=""white""></BODY>" & VbCrLF
response.write "<FONT FACE=""Arial"" SIZE=2>"
response.write "This page is still under construction."
response.write "</FONT>"
response.write "</BODY></HTML>"
%>