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
<HTML>
<HEAD>
<TITLE>Find_It</TITLE>
<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>
</HEAD> 
<BODY BGCOLOR='White' onLoad='document.forms[0].submit()'>
<FORM NAME='FORM1' ACTION='/SW-Common/SW-Find_It-test.asp' METHOD='POST' ID='Form1'>
					  
<!--<form id='frmCDNAuthTest' name='frmCDNAuthTest'  method='POST' action=''>-->
 <input type='hidden' id='hidSave' name='hidSave'/>
<p style="height: 32px">
        

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
</span>Insert uploaded file name in "</span>Pricelist Asset File Name to download:" <span
style='font-size:9.0pt;font-family:Arial;color:black'>Textbox and click on Download button to download pricelist asset from CDN location <o:p></o:p></span></p>
<br>
Enter your "Pricelist Asset" File Name to download:
<input id="txtPath" type="text" value="" NAME="txtPath"/> 
<!--input id="Submit1" type="submit" value="Submit" onclick="return Submit1_onclick()" NAME="Submit1"/-->
 
<input type='submit' id='btnSave' name='btnSave'  onclick='return SaveClicked(this);' value='Download'/>
   
</FORM>
</BODY>
</HTML>



