<%
           ' if (not IsNull( Request.QueryString("URL"))) then 
		' 	Response.Write(" Querystring = " & Request.QueryString("URL") & "<BR>")
		'		Response.Write(" Decoded Querystring = " & URLDecode(LCase(Request.QueryString("URL"))) & "<BR>")
		'	else
		'		Response.Write(" No  Querystring <BR>")
		'	end if 
		'	if  (not IsNull( Request.Cookies("CDNAuth")))  then
		'			Response.Write("Cookie value = " & Request.Cookies("CDNAuth") & "<BR>")
		'			Response.Write("Decoded Cookie value = " & URLDecode(LCase(Request.Cookies("CDNAuth"))) & "<BR>")
		'	else
		'			Response.Write("cookie value isnull"& "<BR>")
		'	end if
								
           '  dim  intStatusCode           
				
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
			    'else
			    '	Response.Redirect ("http://support.fluke.com/SW-common/cdnPass.asp")					
			    end if
			   
			   
			  
		 
			 
%>