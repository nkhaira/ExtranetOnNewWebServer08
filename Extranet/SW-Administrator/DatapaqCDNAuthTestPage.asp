<script type='text/javascript' language='javascript'> 
 function SaveClicked()
    {
     document.getElementById('hidSave').value = 'Save';
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
Enter your CDN FilePath to download:
<input id="txtPath" type="text" value="" NAME="txtPath"/> 
<!--input id="Submit1" type="submit" value="Submit" onclick="return Submit1_onclick()" NAME="Submit1"/-->
 <input type='submit' id='btnSave' name='btnSave'  onclick='return SaveClicked();' value='Save'/>
           
<%
                  
                if (Request.Form("hidSave") <> "") then
                  txtPath = request.Form("txtPath") 			      
                  Response.Cookies("CDNAuthDatapaq") = txtPath
                  Response.Cookies("CDNAuthDatapaq").Domain = ".fluke.com"
                  Response.Cookies("CDNAuthDatapaq").Expires = Date() + 1			
			      'Response.Write(Request.Cookies("CDNAuthDatapaq"))	
			      Response.Write("<a href="& txtPath &">Your file is available for download..</a>." ) 	
			     
			    end if
			    Response.Write("<br><br> CDNAuthDatapaq cookie value from response object : "&Request.Cookies("CDNAuthDatapaq"))    
                

%> 
</form>





