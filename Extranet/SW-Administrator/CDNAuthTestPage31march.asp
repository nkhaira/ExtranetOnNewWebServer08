<!--#include virtual="/include/functions_string.asp"-->
<script type='text/javascript' language='javascript'> 
 function SaveClicked(ctrl)
    {
     document.getElementById('hidSave').value = ctrl.value;
    }
</script>
<%
Dim txtPath
txtPath = ""
					
'Response.Write("Test FTP") 
%>
<form id='frmCDNAuthTest' name='frmCDNAuthTest' method='POST' action=''>
 <input type='hidden' id='hidSave' name='hidSave'/>
<p style="height: 32px">
        
<%
               
                 'txtPath = request.Form("txtPath") 			     
                 'Response.Cookies("CDNAuth") = txtPath '"http://download.fluke.com/pricelist/Download/Asset/9032065_ENG_A_X.XLS"
                 'Response.Cookies("CDNAuth").Domain = ".fluke.com"
                 'Response.Cookies("CDNAuth").Expires = Date() + 1			
			     'Response.Write(" CDNAuth cookie value from response object : "&Request.Cookies("CDNAuth"))
                 
                 'Response.Write("<a href="& txtPath &">Your file is available for download</a>." ) 
                 'Response.Write(txtPath)
                 'Response.Write(Request.Form("hidSave"))
                 'Response.End 
                if (Request.Form("hidSave") <> "") then
                    if (Request.Form("hidSave") = "9032071_ENG_A_X") then                    
                      txtPath = "9032071"
                    end if 
                    if  (Request.Form("hidSave") = "9032072_ENG_A_X") then					 
					  txtPath = "9032072"
					end if
					if  (Request.Form("hidSave") = "90320701_ENG_A_X") then					 
					  txtPath = "90320701"
					end if  
					if  (Request.Form("hidSave") = "9032067_ENG_A_X") then					  
					   txtPath = "9032067"
					end if  
					if  (Request.Form("hidSave") = "9032063_ENG_A_X") then
					 txtPath = "9032063"					
					end if  
					if  (Request.Form("hidSave") = "9032066_ENG_A_X") then
					  txtPath = "9032066"
					end if  
					if  (Request.Form("hidSave") = "9032065_ENG_A_X") then
					  txtPath = "9032065"
					end if  
                    if  (Request.Form("hidSave") = "Download") then
                      txtPath =  request.Form("txtPath") 
                    end if
                  Document = txtPath;			      
                  Response.Cookies("CDNAuth") = txtPath
                  Response.Cookies("CDNAuth").Domain = ".fluke.com"
                  Response.Cookies("CDNAuth").Expires = Date() + 1	
                  Response.Redirect(txtPath)		
			      'Response.Write(Request.Cookies("CDNAuth"))	
			      'Response.Write("<a href="& txtPath &">Your file is available for download</a>." ) 			     
			        response.write "<HTML>" & vbCrLf
					response.write "<HEAD>" & vbCrLf
					response.write "<TITLE>Find_It</TITLE>" & vbCrLf
					response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso-8859-1"">" & vbCrLf
					response.write "</HEAD>" & vbCrLf
					response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
					response.write "<FORM NAME=""FORM1"" ACTION=""/SW-Common/SW-Find_It_New.asp"" METHOD=""POST"">" & vbCrLf
					  
					if request("Locator") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""Locator"" VALUE=""" & request("Locator") & """>" & vbCrLf
					end if
					  
					if request("SW-Locator") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""SW-Locator"" VALUE=""" & request("SW-Locator") & """>" & vbCrLf
					end if
					  
					if Document <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""Document"" VALUE=""" & Document & """>" & vbCrLf
					end if
					  
					if request("Style") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""Style"" VALUE=""" & request("Style") & """>" & vbCrLf
					end if
					  
					if request("Verify") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""Verify"" VALUE=""" & request("Verify") & """>" & vbCrLf
					end if
					  
					if request("CMS_Site") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""CMS_Site"" VALUE=""" & request("CMS_Site") & """>" & vbCrLf
					end if
					  
					if request("CMS_Path") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""CMS_Path"" VALUE=""" & request("CMS_Path") & """>" & vbCrLf
					end if
					  
					if request("SRC") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""SRC"" VALUE=""" & request("SRC") & """>" & vbCrLf
					end if
					  
					if request("Debug") <> "" then
						response.write "<INPUT TYPE=""HIDDEN"" NAME=""Debug"" VALUE=""" & request("Debug") & """>" & vbCrLf
					end if
					  
					LastSite = request.ServerVariables("HTTP_REFERER")
					  
					if isblank(LastSite) then
						LastSite = "[Unknown HTTP_Referer]"
					end if
					  
					if not isblank(Document) then
						LastSite = "URL Link from: " & Server.URLEncode(LastSite & "?" & request.QueryString)
					elseif not isblank(request("Locator")) then
						LastSite = "Partner Portal Subscription URL Link: " & Server.URLEncode("http://Support.Fluke.com/Find_It.asp?Locator=" & request("Locator"))
					elseif not isblank(request("SW-Locator")) then
						LastSite = "Partner Portal Asset URL Link: " & Server.URLEncode("http://Support.Fluke.com/Find_It.asp?SW-Locator=" & request("SW-Locator"))
					else
						LastSite = "Unknown Method 1 by using URL Link: " & Server.URLEncode(LastSite & "?" & request.QueryString)
					end if
					  
					response.write "<INPUT TYPE=""HIDDEN"" NAME=""Referer"" VALUE=""" & LastSite & """>" & vbCrLf
					  
					response.write "</FORM>" & vbCrLf
					response.write "</BODY>" & vbCrLf
					response.write "</HTML>" & vbCrLf
  		    
			    end if

%> 
<span style='font-size:9.0pt;font-family:Arial;
mso-fareast-font-family:"Times New Roman";color:black;mso-ansi-language:EN-US;
mso-fareast-language:EN-US;mso-bidi-language:AR-SA'>Have created below pricelist
asset on production location from dev environment ,You can download available Asset with below links :</span>
 <br>
 <input type='submit' id="Submit1" name='btnSave'  onclick='return SaveClicked(this);' value='9032071_ENG_A_X'/>
 <br>
 <input type='submit' id="Submit2" name='btnSave'  onclick='return SaveClicked(this);' value='9032072_ENG_A_X'/>
 <br>
 <input type='submit' id="Submit3" name='btnSave'  onclick='return SaveClicked(this);' value='90320701_ENG_A_X'/>
 <br>
 <input type='submit' id="Submit4" name='btnSave'  onclick='return SaveClicked(this);' value='9032067_ENG_A_X'/>
 <br>
 <input type='submit' id="Submit5" name='btnSave'  onclick='return SaveClicked(this);' value='9032063_ENG_A_X'/>
 <br>
 <input type='submit' id="Submit6" name='btnSave'  onclick='return SaveClicked(this);' value='9032066_ENG_A_X'/>
 <br>
 <input type='submit' id="Submit7" name='btnSave'  onclick='return SaveClicked(this);' value='9032065_ENG_A_X'/>
<br>

<p class=MsoNormal style='margin-Center:-.25in'><b style='mso-bidi-font-weight:
normal'><span style='font-size:9.0pt;font-family:Arial;color:black'> In addition, you can upload new pricelist asset </span></b><span
style='font-size:9.0pt;font-family:Arial;color:black'>, <b style='mso-bidi-font-weight:
normal'>and verify it exists in prod FTP location. <o:p></o:p></b></span></p>

<p class=MsoNormal style='margin-Center:-.25in'><span style='font-size:9.0pt;
font-family:Arial;color:black'><span style='mso-spacerun:yes'>  
</span>Insert uploaded file name in “</span>Pricelist Asset File Name to download:” <span
style='font-size:9.0pt;font-family:Arial;color:black'>Textbox and click on Download button to download pricelist asset from CDN location <o:p></o:p></span></p>
<br>
Enter your "Pricelist Asset" File Name to download:
<input id="txtPath" type="text" value="" NAME="txtPath"/> 
<!--input id="Submit1" type="submit" value="Submit" onclick="return Submit1_onclick()" NAME="Submit1"/-->
 
<input type='submit' id='btnSave' name='btnSave'  onclick='return SaveClicked(this);' value='Download'/>
   

 </form>



