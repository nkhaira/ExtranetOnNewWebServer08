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
varText = varText & " ServerVariables :"& vbCrLf 
For Each item In Request.ServerVariables  
  varText = varText & "Key: " & item & " - Value: " & Request.ServerVariables(item) & vbCrLf 
Next
''======================================================
'Start send mail 
''======================================================
    
            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")    
            Mailer.QMessage       = False
            Mailer.ReturnReceipt  = False
            Mailer.Priority       = 1
            Mailer.RemoteHost  = "mail.fluke.com"
            Mailer.TimeOut     = 120
            Mailer.WordWrap    = False
            Mailer.WordWrapLen = 150           
             Mailer.FromName       = "fluke.com "
             Mailer.FromAddress    = "info@fluke.com"          
             
            Mailer.AddRecipient "jigar joshi", "jigar.joshi@zensar.in"            
            Mailer.AddCC "girish deshpande", "girish.deshpande@zensar.in"
            Mailer.Subject  = "CDN price list test"
            Mailer.BodyText = varText
            Mailer.SendMail 
 ''======================================================
'End send mail 
''======================================================     
         
				
             intStatusCode = 401 
             			 
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
				 	'Get the CDNAuth cookie value
					if  (not IsNull( Request.Cookies("CDNAuth")) and   Request.Cookies("CDNAuth")<>""  )  then
					 	if (URLDecode(LCase(Request.QueryString("URL"))) = URLDecode(LCase(Request.Cookies("CDNAuth")))) then
							intStatusCode = 202	'For Valid
							'   Response.Write(" CDN Pass"& "<BR>")							   
							
						else
										intStatusCode = 401	'For Not Valid
									'	Response.Write(" CDN Fail"& "<BR>")
									   
					    end if
					end if
				  end if

				'Add header to responce
				'Response.StatusCode = intStatusCode;
				Response.Status = intStatusCode
				if (intStatusCode = 401) then				
					Response.Redirect ("http://support.fluke.com/SW-common/cdnfail.asp")	
			    else
			    	Response.Redirect ("http://support.fluke.com/SW-common/cdnPass.asp")					
			    end if
			   
			   
			  
		 
			 
%>