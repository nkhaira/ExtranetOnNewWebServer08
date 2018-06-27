<%
 option explicit
 Response.Buffer = True
 
 'declare a variable for the reference page, 
 'the XMLHTTP Object, and the regular expressions used  

 Dim RefPage, objXMLHTTP, RegEx

 'set the RefPage variable to the "ref" querystring
 'the JavaScript function above passes the current page URL
 'You can use the Request.ServerVariables("HTTP_REFERER") to 
 'get the page as a last option if needed
   
 RefPage = Request.QueryString("ref")
 if RefPage = "" then
   response.write "<h3>Invalid reference page</h3>"
   response.end
 end if
 
 'set the objXMLHTTP object to the XMLHTTP object from Microsoft
 Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
 
  
 'perform the HTTP "GET" method via the XMLHTTP object to retrieve 
 'the called page
 objXMLHTTP.Open "GET", RefPage, False
 objXMLHTTP.Send
 
 'give RefPage the text(HTML) from the call above 
 RefPage = objXMLHTTP.responseText
 
 'Create built In Regular Expression object that 
 'is now included with VBScript version5
 Set RegEx = New RegExp
 RegEx.Global = True
 
 
 'Set the pattern To look For <!-- START PPOMIT -->  tags
 RegEx.Pattern =  "<!-- START PPOMIT -->"
 
 'replace the comment pattern with a rare ASCII character
 'i've choosed No 253 in this case, I suppose if you're speaking
 'a language other than English this may not be the case but 
 'the reasoning here is so it will be unique and not interfere 
 'this the main page content 
 RefPage = RegEx.Replace(refpage,( chr(253) ))
 
 'Set the pattern To look For <!-- END PPOMIT --> tags
 RegEx.Pattern = "<!-- END PPOMIT -->"
 
 'This time make it replace the comment with another rare 
 'ASCII character, NOT the one used above
 RefPage = RegEx.Replace(refpage,( chr(254) ))
 
 'Use this regular expression to "cut out" HTML between the 
 'start and end comments now the new ASCII characters
 RegEx.Pattern = chr(253) & "[^" & chr(254) & "]*" & chr(254)
 
 'This will perform the actual striping
 RefPage = RegEx.Replace(refpage, " " )
 
 'Don't forget to tidy up :-)
 Set RegEx = Nothing
 Set objXMLHTTP = Nothing
 
 'Output your Printer Friendly Page!
 Response.Write RefPage
%>

