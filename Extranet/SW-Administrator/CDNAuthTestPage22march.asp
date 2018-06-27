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
               
                 'txtPath = request.Form("txtPath") 			     
                 'Response.Cookies("CDNAuth") = txtPath '"http://downloads.fluke.com/pricelist/Download/Asset/9032065_ENG_A_X.XLS"
                 'Response.Cookies("CDNAuth").Domain = ".fluke.com"
                 'Response.Cookies("CDNAuth").Expires = Date() + 1			
			     'Response.Write(" CDNAuth cookie value from response object : "&Request.Cookies("CDNAuth"))
                 
                 'Response.Write("<a href="& txtPath &">Your file is available for download</a>." ) 
                 'Response.Write(txtPath)
                 'Response.Write(Request.Form("hidSave"))
                 'Response.End 
                if (Request.Form("hidSave") <> "") then
                  txtPath = request.Form("txtPath") 			      
                  Response.Cookies("CDNAuth") = txtPath
                  Response.Cookies("CDNAuth").Domain = ".fluke.com"
                  Response.Cookies("CDNAuth").Expires = Date() + 1			
			      'Response.Write(Request.Cookies("CDNAuth"))	
			      Response.Write("<a href="& txtPath &">Your file is available for download</a>." ) 		    
			    end if

%> 
</form>

<%
                 'Response.Cookies("CDNAuth1") = txtPath '"http://downloads.fluke.com/pricelist/Download/Asset/9032065_ENG_A_X.XLS"
                 'Response.Cookies("CDNAuth").Domain = ".fluke.com"
                 'Response.Cookies("CDNAuth1").Expires = Date() + 1			
			     'Response.Write(" CDNAuth1 cookie value from response object : "&Request.Cookies("CDNAuth1"))
                 
%> 



