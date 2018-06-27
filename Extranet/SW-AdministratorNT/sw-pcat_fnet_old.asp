<%@ Language=VBScript %>
<HTML>
	<HEAD>
		<script language="javascript">
<!--
	function AddRemoveOptions(strFrom,strTo)
		{	
		//alert(document.frmpcat.lstAProducts.options.length);
		var i=0;
		var objfrom;
		var objto;
		objfrom=eval('document.frmpcat.' + strFrom);
		objto=eval('document.frmpcat.' + strTo);
		for(i=(objfrom.options.length-1);i>=0;i--)
			{	
				if (objfrom.options[i].selected==true) 
					{
					var optnew;
					optnew = document.createElement("OPTION") 
					optnew.text=objfrom.options[i].text;
					optnew.value=objfrom.options[i].value;
					objto.options.add(optnew);
					objfrom.options.remove(i);
					sortList(strTo);
					}
			}
	}
	
	function sortList(objfrom)
	{	
		var objfrom;
		var objto;
		var i;
		var pos;
		var optnew;
		//new 
		strFrom=eval('document.frmpcat.' + objfrom);
		//objto=eval('document.frmpcat.' + strTo);
		//alert(strFrom.options.length);
		var listarray=new Array(strFrom.options.length-1);
		//var listValue=new Array(objfrom.options.length-1);
		
		for(i=(strFrom.options.length-1);i>=0;i--)
			{
				listarray[i]=strFrom.options[i].text + '$_' + strFrom.options[i].value;
				//listValue[i]=objfrom.options[i].value;
			}
			listarray.sort();
		/*for(i=0;i<=listarray.length-1;i++)
		{
			alert(listarray[i]);
		}*/
			
		for(i=(strFrom.options.length-1);i>=0;i--)
			{
				strFrom.options.remove(i);
			}
		for(i=0;i<=listarray.length-1;i++)
			{
				optnew = document.createElement("OPTION") ;
				//alert(listarray[i]);
				
				pos=listarray[i].indexOf('$_');
				//alert(listarray[i].substring(0,pos));
				//alert(listarray[i].substring(pos+3,listarray[i].length-pos+3));
				//alert(pos);
				optnew.text=listarray[i].substring(0,pos);
				//alert(optnew.text);
				optnew.value=listarray[i].substring(pos+3,listarray[i].length-pos+3);
				//alert(optnew.value);
				//alert(optnew.text);
				strFrom.options.add(optnew);
			}
	}
	
//-->
		</script>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
	</HEAD>
	<body onload="sortList('lstAProducts');sortList('lstALocales');">
		<P>&nbsp;</P>
		<P>&nbsp;</P>
		<P>&nbsp;</P>
		<form name="frmpcat" method="post">
			<TABLE cellSpacing="1" cellPadding="1" width="504" border="1" height="234">
				<TR>
					<TD width="99">PCat Category:</TD>
					<TD width="173"><SELECT name="cboCategory">
							<OPTION selected>Select from List</OPTION>
							<OPTION value="1">Network Tools</OPTION>
						</SELECT></TD>
					<TD width="26"></TD>
					<TD></TD>
				</TR>
				<TR>
					<TD width="99">PCat Products</TD>
					<TD width="173">
						<TABLE cellSpacing="1" cellPadding="1" width="160" border="1" height="70">
							<TR>
								<TD>Available Products</TD>
							</TR>
							<TR>
								<TD><SELECT multiple size="3" name="lstAProducts">
										<OPTION value="1">Network Inspector</OPTION>
										<OPTION value="2">Optview</OPTION>
									</SELECT></TD>
							</TR>
						</TABLE>
					</TD>
					<TD width="26">
						<TABLE height="61" cellSpacing="1" cellPadding="1" width="24" border="1">
							<TR>
								<TD><INPUT type="button" value=">" name="btnAproducts" onclick="AddRemoveOptions('lstAProducts','lstSProducts')"></TD>
							</TR>
							<TR>
								<TD><INPUT type="button" value="<" name="btnRproducts" onclick="AddRemoveOptions('lstSProducts','lstAProducts')"></TD>
							</TR>
						</TABLE>
					</TD>
					<TD>
						<TABLE cellSpacing="1" cellPadding="1" width="176" border="1" height="82">
							<TR>
								<TD>Selected&nbsp;Products</TD>
							</TR>
							<TR>
								<TD><SELECT multiple size="3" name="lstSProducts">
									</SELECT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD width="99">PCat Locale</TD>
					<TD width="173">
						<TABLE cellSpacing="1" cellPadding="1" width="160" border="1" height="100">
							<TR>
								<TD>Available Locales</TD>
							</TR>
							<TR>
								<TD><SELECT multiple size="4" name="lstALocales">
										<OPTION value="1">Usen</OPTION>
										<OPTION value="2">Caen</OPTION>
										<OPTION value="3">Cafr</OPTION>
										<OPTION value="4">Brpt</OPTION>
									</SELECT></TD>
							</TR>
						</TABLE>
					</TD>
					<TD width="26">
						<TABLE height="61" cellSpacing="1" cellPadding="1" width="16" border="1">
							<TR>
								<TD><INPUT type="button" value=">" name="btnAlocales" onclick="AddRemoveOptions('lstALocales','lstSLocales')"></TD>
							</TR>
							<TR>
								<TD><INPUT type="button" value="<" name="btnRlocales" onclick="AddRemoveOptions('lstSLocales','lstALocales')"></TD>
							</TR>
						</TABLE>
					</TD>
					<TD>
						<TABLE cellSpacing="1" cellPadding="1" width="176" border="1" height="70">
							<TR>
								<TD>Selected Locales</TD>
							</TR>
							<TR>
								<TD><SELECT multiple size="4" name="lstSLocales">
									</SELECT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
