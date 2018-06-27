<%
varText = ""
'Response.Write " Form variables"& "<BR />"
'For Each item In Request.Form
'   Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
'Next
'Response.Write " QueryString values :"& "<BR />"
'For Each item In Request.QueryString 
'  Response.Write "Key: " & item & " - Value: " & Request.QueryString(item) & "<BR />"
'Next
if  (not IsNull( Request.Cookies("CDNAuthDatapaq")) and   Request.Cookies("CDNAuthDatapaq")<>""  )  then
  varText = varText & " CDNAuthDatapaq Cookies value: " &URLDecode(LCase(Request.Cookies("CDNAuthDatapaq")))& vbCrLf
else
  varText = varText & " CDNAuthDatapaq Cookies value is blank or not available "& vbCrLf
end if  
if (not IsNull( Request.QueryString("URL")) and   Request.QueryString("URL")<>"" )  then 
   varText = varText & " URL value: "&URLDecode(LCase(Request.QueryString("URL")))& vbCrLf 
else
   varText = varText & " URL value is blank or not available "& vbCrLf 
end if   
varText = varText & " ServerVariables :"& vbCrLf 
For Each item In Request.ServerVariables  
 varText = varText & "Key: " & item & " - Value: " & Request.ServerVariables(item) & vbCrLf 
Next
''======================================================
'Start send mail 
''======================================================
    
            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")    
            'adding new email method
            %>
            <!--#include virtual="/connections/connection_email_new.asp"-->
            <%
            'Mailer.QMessage       = False
            'Mailer.ReturnReceipt  = False
            'Mailer.Priority       = 1
            'Mailer.RemoteHost  = "mail.evt.danahertm.com:25"
            'Mailer.TimeOut     = 120
            'Mailer.WordWrap    = False
            'Mailer.WordWrapLen = 150           
            ''Mailer.FromName       = "girish deshpande"
            ''Mailer.FromAddress    = "girish.deshpande@fluke.com"     
            
            'Mailer.FromName       = "Santosh Tembhare"
            'Mailer.FromAddress    = "Santosh.tembhare@fluke.com"     

            msg.From = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"
             
                      
            'Mailer.AddRecipient "Sreejith Nair", "Sreejith.Nair@fluke.com"
            'Mailer.Addcc "girish deshpande", "girish.deshpande@zensar.in"  
            
            
            'Mailer.AddRecipient "Santosh Tembhare", "Santosh.tembhare@fluke.com"
            'Mailer.Addcc "Santosh Tembhare", "Santosh.tembhare@fluke.com"            
            'Mailer.Subject  = "Datapaq CDN -- test"
            'Mailer.BodyText = varText
            'Mailer.SendMail 

            msg.To = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"
            msg.Cc = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"
            msg.Subject = "Datapaq CDN -- test"
            msg.TextBody = varText
            msg.Send
			Set conf = Nothing
			Set msg = Nothing

 ''======================================================
'End send mail 
''======================================================     
         
				
             intStatusCode = 301 
             			 
				Function URLDecode(str) 
                   str = Replace(str, "+", " ")        
                    For i = 1 To Len(str) 
                        sT = Mid(str, i, 1) 
                        If sT = "%" Then 
                            If i+2 < Len(str) Then 
                                sR = sR & _ 
                                    Chr(CLng("&H" & Mid(str, i+1, 2))) 
                                i = i+2 
                            End If 
                        Else 
                            sR = sR & sT 
                        End If 
                    Next 
                    URLDecode = sR 
                End Function 
                if (not IsNull( Request.QueryString("URL")) and   Request.QueryString("URL")<>"" )  then 
				 	'Get the CDNAuthDatapaq cookie value
					if  (not IsNull( Request.Cookies("CDNAuthDatapaq")) and   Request.Cookies("CDNAuthDatapaq")<>""  )  then
					 	if (URLDecode(LCase(Request.QueryString("URL"))) = URLDecode(LCase(Request.Cookies("CDNAuthDatapaq")))) then
							
							intStatusCode = 200	'For Valid
							
							varText = " Status Code :"& intStatusCode &" CDN Pass "&  vbCrLf & varText														   
							
						else
							intStatusCode = 301	'For Not Valid
							varText = " Status Code :"& intStatusCode &" CDN Fail"&  vbCrLf & varText									
									   
					    end if
					end if
				  end if
				                  
                 
                 
                 'Response.Write(varText)
                 'Response.End ()
				
				'Add header to responce				
				Response.Status = intStatusCode
				if (intStatusCode = 301) then				
					Response.Redirect ("http://support.fluke.com/SW-common/cdnfail.asp")	
'			    else
'			    	Response.Redirect ("http://support.fluke.com/SW-common/cdnPass.asp")					
			    end if
			   
			   
			  
		 
			 
%>