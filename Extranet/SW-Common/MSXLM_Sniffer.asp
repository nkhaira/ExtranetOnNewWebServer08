<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>MSXML Sniffer</title>
<script language="JavaScript"> 
function sniff(){
	var xml = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><cjb></cjb>";
	var xsl = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><x:stylesheet version=\"1.0\" xmlns:x=\"http://www.w3.org/1999/XSL/Transform\" xmlns:m=\"urn:schemas-microsoft-com:xslt\"><x:template match=\"/\"><x:value-of select=\"system-property('m:version')\" /></x:template></x:stylesheet>";
	//var xsl = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><x:stylesheet version=\"1.0\" xmlns:x=\"http://www.w3.org/TR/WD-xsl\"></x:stylesheet>";

	var x = null;
	    
	try{ 
	    x = new ActiveXObject("Msxml2.DOMDocument"); 
	    x.async = false;
	    if (x.loadXML(xml)){
	   		sniffer.msxml2.checked = true;
	   		document.getElementById("advice1").innerText = "";
	   	}
	}catch(e){
		document.getElementById("msxml2reason").innerText = e.description;
		document.getElementById("advice2").innerText = "";
	}
	 
	try{ 
	    x = new ActiveXObject("Msxml2.DOMDocument.2.6"); 
	    x.async = false;
	    if (x.loadXML(xml)) 
	    	  sniffer.msxml2v26.checked = true;
	}catch(e){document.getElementById("msxml2v26reason").innerText = e.description} 

	try{ 
	    x = new ActiveXObject("Msxml2.DOMDocument.3.0"); 
	    x.async = false;
	    if (x.loadXML(xml)) 
	    	  sniffer.msxml2v30.checked = true;
	}catch(e){document.getElementById("msxml2v30reason").innerText = e.description}

	try{ 
	    x = new ActiveXObject("Msxml2.DOMDocument.4.0"); 
	    x.async = false;
	    if (x.loadXML(xml)) 
	    	  sniffer.msxml2v40.checked = true;
	}catch(e){document.getElementById("msxml2v40reason").innerText = e.description}

	try{ 
	    x = new ActiveXObject("Microsoft.XMLDOM");  
	    x.async = false;
	    if (x.loadXML(xml))
	    	  sniffer.msxml.checked = true;
	}catch(e){document.getElementById("msxmlreason").innerText = e.description} 

	try{
		var s = new ActiveXObject("Microsoft.XMLDOM"); 
		s.async = false;
		if (s.loadXML(xsl)){
			try{
				var op = x.transformNode(s);
				if (op.indexOf("stylesheet") == -1){
					sniffer.replace.checked = true;
					document.getElementById("replacereason").innerText = "Replace V" + op.substr(op.lastIndexOf(">")+1);
					document.getElementById("advice2").innerText = "";
				}else
					if (sniffer.msxml2.checked)
						document.getElementById("replacereason").innerText = "Side-By-Side";
			}catch(e){
				if (sniffer.msxml2.checked)
						document.getElementById("replacereason").innerText = "Side-By-Side";
			}
		}
	}catch(e){}
}
</script> 
 
</head>

<body onload="sniff()" bgcolor="#2288ff">

<h1 align="center"><font color="#000000">MSXML Sniffer</font></h1>
<form name="sniffer">
  <div align="center">
    <center>
    <table border="1" width="1%">
      <tr>
        <td width="1%" nowrap style="color:#000000;">MSXML</td>
        <td width="1%"><input type="checkbox" name="msxml" value="ON" disabled></td>
        <td id="msxmlreason" nowrap style="color:#000000;">Installed</td>
      </tr>
      <tr>
        <td width="1%" nowrap style="color:#000000;">MSXML2</td>
        <td width="1%"><input type="checkbox" name="msxml2" value="ON" disabled></td>
        <td id="msxml2reason" nowrap style="color:#000000;">Installed</td>
      </tr>
      <tr>
        <td width="1%" nowrap style="color:#000000;">MSXML2 v2.6</td>
        <td width="1%"><input type="checkbox" name="msxml2v26" value="ON" disabled></td>
        <td id="msxml2v26reason" nowrap style="color:#000000;">Installed</td>
      </tr>
      <tr>
        <td width="1%" nowrap style="color:#000000;">MSXML2 v3.0</td>
        <td width="1%"><input type="checkbox" name="msxml2v30" value="ON" disabled></td>
        <td id="msxml2v30reason" nowrap style="color:#000000;">Installed</td>
      </tr>
      <tr>
        <td width="1%" nowrap style="color:#000000;">MSXML2 v4.0</td>
        <td width="1%"><input type="checkbox" name="msxml2v40" value="ON" disabled></td>
        <td id="msxml2v40reason" nowrap style="color:#000000;">Installed</td>
      </tr>
      <tr>
        <td width="1%" nowrap style="color:#000000;">Mode</td>
        <td width="1%"><input type="checkbox" name="replace" value="ON" disabled></td>
        <td id="replacereason" nowrap style="color:#000000;">&nbsp;</td>
      </tr>
    </table>
    </center>
  </div>
</form>
<p id="advice1" style="color:#000000;">You are using an old version of MSXML. It is recomended that you download
an upgrade from <a href="http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/596/msdncompositedoc.xml" target="xxx">Microsoft.</a>
Or you can use the automatic <a href="JavaScript:parent.changeContent1('/xml/utils/instalmsxml.xml');">Install MSXML</a> utility in this section.</p>
<p id="advice2" style="color:#000000;">Although you have the new version of MSXML you are running it in "side-by-side" mode. 
To get the full benifits of this new version you should run xmlinst (from Microsoft to switch to
"replace" mode.</p>
</body>

</html>
