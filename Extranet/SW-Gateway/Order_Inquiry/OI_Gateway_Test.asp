<%
' Script to Call Partner Portal Gateway to Auto Logon User for Order Inquiry Views.
' K. D. Whitlock
' 04/08/2003

NTLogin  = "cj.kokee"                 ' Logon User Name (Test Account with MFG/PRO Account Access)
Password = "Partner Portal"           ' Password

Site_ID  = 3                          ' Site ID Number for Find-Sales
BackURL  = "http://www.SyncForce.nl"  ' Where to go when user clicks on [Home] Button

with response

  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf
  .write "<TITLE>Gateway</TITLE>" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
  .write "<FORM NAME=""Gateway"" ACTION=""http://support.fluke.com/sw-gateway/order_inquiry/default.asp"" METHOD=""POST"">"
  .write "<INPUT TYPE=""HIDDEN"" NAME=""NTLogin"" VALUE="""  & NTLogin  & """>" & vbCrLf
  .write "<INPUT TYPE=""HIDDEN"" NAME=""PASSWORD"" VALUE=""" & Password & """>" & vbCrLf        
  .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE="""  & Site_ID  & """>" & vbCrLf
	.write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE="""  & BackURL  & """>" & vbCrLf
  .write "</FORM>" & vbCrLf
  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf
  
end with  
  
%>
