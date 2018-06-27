<%

' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       6/4/2003
'             Sandbox
'
' Post from Print Page button on parent page, sends current URL to this page.
' Start / End Markers are used to deliniate are to be printed.
' Window closes when done.
' --------------------------------------------------------------------------------------

 option explicit
 Response.Buffer = True
 
 ' Declare a variable for the reference page, 
 ' the XMLHTTP Object, and the regular expressions used  

 Dim RefPage, objXMLHTTP, RegEx

 ' Set the RefPage variable to the "ref" querystring the JavaScript function above passes the current page URL
 ' You can use the Request.ServerVariables("HTTP_REFERER") to get the page as a last option if needed
   
 RefPage = Request.QueryString("ref")
 
 if RefPage = "" then
   response.write "<h3>Invalid Print Area Reference Page</h3>"
   response.end
 end if
 
 Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
   
 ' Perform the HTTP "GET" method via the XMLHTTP object to retrieve the called page

 objXMLHTTP.Open "GET", RefPage, False
 objXMLHTTP.Send
 
 ' Give RefPage the text(HTML) from the call above 

 RefPage = objXMLHTTP.responseText
 
 ' Create built In Regular Expression object that is now included with VBScript version 5

 Set RegEx = New RegExp
 RegEx.Global = True
 
 ' Set the pattern to look for Start tag
 
 RegEx.Pattern =  "<!-- Start Print Area -->"
 
 ' Replace the comment pattern with a rarely used ASCII character #253 so as not to interfere with body text
 ' Need to check what happens with DBCS such as Simplified Chineese
  
 RefPage = RegEx.Replace(refpage,( chr(253) ))
 
 ' Set the pattern to look for End tag

 RegEx.Pattern = "<!-- End Print Area -->"
 
 ' Replace the comment pattern with a rarely used ASCII character #254 so as not to interfere with body text
 ' Need to check what happens with DBCS such as Simplified Chineese

 RefPage = RegEx.Replace(refpage,( chr(254) ))
 
 ' Use this regular expression to "cut out" HTML between the start and end comments now the new ASCII characters

 RegEx.Pattern = chr(253) & "[^" & chr(254) & "]*" & chr(254)
 
 'This will perform the actual striping

 RefPage = RegEx.Replace(refpage, " " )
 
 ' Clean up

 Set RegEx = Nothing
 Set objXMLHTTP = Nothing
 
 'Output print area to printer

 Response.Write RefPage
%>

<SCRIPT Language="JavaScript">

