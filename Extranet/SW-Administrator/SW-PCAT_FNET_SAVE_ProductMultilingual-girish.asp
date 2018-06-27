<!--#include virtual="/sw-administrator/SW-PCAT_FNET_IISSERVER.asp"-->
<!--#include virtual="/include/functions_string.asp"-->

<script type='text/javascript' language = 'javascript'>
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=1900;
var maxYear=9999;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strMonth=dtStr.substring(0,pos1)
	var strDay=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		//alert("The date format should be : mm/dd/yyyy")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		//alert("Please enter a valid month")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		//alert("Please enter a valid day")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		//alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		//alert("Please enter a valid date")
		return false
	}
return true
}

function ValidateForm(){
	
	var dtfrom=document.getElementById("txtStartDate");
	var dtend=document.getElementById("txtEndDate");
	if (isDate(dtfrom.value)==false || isDate(dtend.value)==false)
		return false;
    else
        return true;
 }
function ValidateStartEndDates(){
	
	var dtfrom=document.getElementById("txtStartDate");
	var dtend=document.getElementById("txtEndDate");
	
	if (new Date(dtfrom.value) > new Date(dtend.value))
		return false;
    else
        return true;
 }
</script>

<script type='text/javascript' language='javascript'> 

        function AddRemoveOptions(strFrom,strTo) {	
        var i=0;
        var objfrom = document.getElementById(strFrom);
        var objto   = document.getElementById(strTo);

        /* //alert(objfrom.options.length + " : " + objto.options.length); */
        for(i=(objfrom.options.length-1);i>=0;i--) {	
        if (objfrom.options[i].selected==true) {
            var optnew;
            optnew = document.createElement("OPTION") 
            optnew.text=objfrom.options[i].text;
            optnew.value=objfrom.options[i].value;
            if(checkifexists(strTo, objfrom.options[i].value)==false)
            {
                alert("Value ''" + objfrom.options[i].text + "'' is already present!");
//                //return false;
            }
            else
            {
                objto.options.add(optnew);
            }    
//            //objfrom.options.remove(i);
//            //sortList(strTo);
        }
        }
        }

        function checkifexists(strTo, productid)
        {   var iproductcount;
            var objto   = document.getElementById(strTo);
            for (iproductcount=0;iproductcount<objto.options.length;iproductcount++)
            {   
                if (objto.options[iproductcount].value==productid)
                {
                    return false;
                }
            } 
        }

        function RemoveOption(strFrom) {
        var i;
        var objfrom = document.getElementById(strFrom);
        for(i=(objfrom.options.length-1);i>=0;i--) {	
        if (objfrom.options[i].selected==true) {
        objfrom.options.remove(i);
                }
            }
        }
        
    function SaveClicked()
    {
        var oLocales, strTemp, oLocalesVal, products, len, selLocales;
        strTemp = "";
        strTemp = document.getElementById("Locales").value + ",";
        var desc = document.getElementById("txtDesc").value;
	  var desc1 = document.getElementById("txtDesc1").value;	
        selLocales = document.getElementById("SelLocales");
	  var productURL=document.getElementById("txtProductUrl").value;

        if (desc == '')
        {
            alert('Please enter valid title.');
            return false;
        }
	  if (desc1 == '')
        {
            alert('Please enter valid description.');
            return false;
        }

        var temp = ValidateForm();
        if(!temp)
        {
	
            alert('Please enter valid start and end dates.');
		document.getElementById("txtStartDate").focus();
            return false;
        }
	 var tempdate=ValidateStartEndDates();
	 if(!tempdate)
	 {
		alert('End date must be greater or equal to Start date.');
            return false;
 	 }
	//var tempyear=ValidateStartEndYears();
	 

        else
        {

		// Validate url

		if(productURL!="")
		{
		var tomatch= /https?:\/\/[A-Za-z0-9\.-]{3,}\.[A-Za-z]{3}/
	    	 if (tomatch.test(productURL)){	         
      	   
	     }
		else{
		alert('Please enter valid product URL. Sample URL format: http://us.fluke.com');
		return false;
		}

		}
		// end validate
            document.getElementById('hidSave').value = 'Save';
            
            oLocales = document.getElementById('Locales');
            oLocalesVal = document.getElementById('Locales').value;
            oLocalesVal = oLocalesVal.substring(0, 2);
            if(oLocales != null && selLocales != null)
            {
                if (oLocales.length)
                {
                    for (products = 0; products < selLocales.length; products++)
                    {
                        if(selLocales[products])
                            strTemp = strTemp + selLocales.options[products].value + ",";
                    }
                }
////                    for (products = 0; products < oLocales.length; products++)
////                    {
////                        if (oLocalesVal == oLocales.options[products].value.substring(0, 2))
////                            strTemp = strTemp + oLocales.options[products].value + ",";
////                    }
////                }
                len = strTemp.length;
                strTemp = strTemp.substring(0, len-1);
            }
            document.getElementById('hidCheckbox').value = document.getElementById('chkAllLocales').checked;
            document.getElementById('hidAllLocales').value = strTemp;
            document.frmAssetMultilingual.submit();
            return true;
        }
    }
    function ChangeLocale()
    {
        var cnt, Locales, arr;
        Locales = document.getElementById("Locales");
        /*//alert(Locales.options.length);
        //alert(myarray.length);*/
        document.getElementById("txtDesc").value = '';
        document.getElementById("txtDesc1").value = '';
        document.getElementById("txtStartDate").value = '';
        document.getElementById("txtEndDate").value = '';
        document.getElementById("txtProductUrl").value = '';
        ////for(cnt=0; cnt < Locales.options.length; cnt++)
        ////{
        var localeSel = Locales.value;
            for(arr=0; arr< myarray.length; arr++)
            {
               if(myarray[arr][0].toLowerCase() == localeSel.toLowerCase())
               {
                    document.getElementById("txtDesc").value = myarray[arr][1];
                    document.getElementById("txtDesc1").value = myarray[arr][2];
                    document.getElementById("txtStartDate").value = myarray[arr][3];
                    document.getElementById("txtEndDate").value = myarray[arr][4];
                    document.getElementById("txtProductUrl").value = myarray[arr][5];
               }
            }
        ////}
    }
</script>

<%
Dim SaveRecord, strPID, Site_ID, objcol, Products, strparameters, txtDesc1, Locales1, objinfo2, objinfo1, objinfo, desc, objinfo3, objinfo4, objinfo5, objinfo6
Dim txtStartDate, txtEndDate, txtDesc2, txtProductUrl, chkAllLocale, strLocales, hidAllLocales, hidCheckbox

SaveRecord=true
strPID = Request.QueryString("ID")
Site_ID= Request.QueryString("SiteId")


 if (Request.Form("hidSave") <> "") then
        txtDesc1 = Request.Form("txtDesc")
        Locales1 = Request.Form("Locales")
        hidAllLocales = Request.Form("hidAllLocales")
        hidCheckbox = Request.Form("hidCheckbox")

        if (hidCheckbox = "true") then
          strLocales = hidAllLocales
        else
          strLocales = Locales1
        end if

''Response.Write strLocales
''Response.End
        txtStartDate = CDate(Request.Form("txtStartDate"))
        txtEndDate = CDate(Request.Form("txtEndDate"))
        if (Site_ID = 3 Or Site_ID = 46) then
			txtDesc2 = Server.URLEncode(server.HTMLEncode(Request.Form("txtDesc1")))
		else
			txtDesc2 = Request.Form("txtDesc1")
		end if
        'RI-771
        txtProductUrl = Server.URLEncode(Request.Form("txtProductUrl"))
           
        chkAllLocale = Request.Form("chkAllLocales")

        'if (err.number <> 0) Then
          'Response.Write err.Description
          'Response.Write "Please provide correct dates."
        'end if
        ''Response.Write txtDesc1
        ''Response.Write Locales1
        ''Response.End ************StartDate,EndDate,Desc
        
        strparameters =  "operation=UL&assetpid=" & strPID & "&ProdDesc=" & txtDesc1 & "&Locale=" & strLocales & "&SiteID=" & Site_ID & "&StartDate=" & txtStartDate & "&EndDate=" & txtEndDate & "&Desc=" & txtDesc2 & "&ProductUrl=" & txtProductUrl
        
        set Products = server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
		Call Products.open("POST", striisserverpath, 0, 0, 0)
		Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
		Call Products.send(strparameters)

		strProducts = Products.responseXML.XML
		set objxml = Server.CreateObject("msxml2.domdocument")
		Call objxml.loadxml(strProducts)

		if not(objxml is nothing ) then
		    set objcol = objxml.selectsingleNode("ProductId")
		end if
		''if (objcol <> null And objcol.text <> null And objcol.text <> "true") then
		if (err.Description <> "") then
            Showerror(err.Description)
			Response.End
        else
            ''Response.Write "<script type='text/javascript' language='javascript'> document.getElementById('message').value = alert('Data Saved successfully!');<script>"
            ''Response.Write "<script type='text/javascript' language='javascript'> function msg() {document.getElementById('message').value = 'Data Saved Successfully!';} body.onload = setTimeout('msg()',1000);<script>"
            Response.Write "Data Saved Successfully."
        end if
  end if


''on error resume next
with Response
    set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
    '' ***Pass the actual url here
    call Products.open("POST", striisserverpath,0,0,0)
    call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    
    ''Response.Write strPID + "  "
    ''Response.Write Request.QueryString("ID")
    ''Response.End
    call Products.send("operation=RL&assetpid=" & strPID & "&SiteID=" & Site_ID)
    strProducts=Products.responseXML.XML

    set objxml=Server.CreateObject("msxml2.domdocument")
    call objxml.loadxml(strProducts)
    icol = 0
    set objcol=objxml.selectsingleNode("Root")
    ''Response.Write strPID
    '''Response.Write(objcol.childnodes.length)
    ''response.End
    if not(objcol is nothing) then
			  .Write "<script type='text/javascript' language='javascript'> var myarray = new Array(" & objcol.childnodes.length & ");"
			  for icol=0 to objcol.childnodes.length-1
			      .Write " myarray[" & icol & "]=new Array(5);"
                  set objinfo=objcol.childnodes(icol)
                  set objinfo1=objinfo.childnodes(0)
                  set objinfo2=objinfo.childnodes(1)
                  set objinfo3=objinfo.childnodes(2)
                  set objinfo4=objinfo.childnodes(3)
                  set objinfo5=objinfo.childnodes(4)
                  set objinfo6=objinfo.childnodes(5)
                  .Write "myarray[" & icol & "][0] = '" & objinfo1.text & "';"
                  .Write "myarray[" & icol & "][1] = """ & objinfo2.text & """;"
                  .Write "myarray[" & icol & "][2] = """ & objinfo3.text & """;"
                  .Write "myarray[" & icol & "][3] = '" & DateValue(objinfo4.text) & "';"
                  .Write "myarray[" & icol & "][4] = '" & DateValue(objinfo5.text) & "';"
                  .Write "myarray[" & icol & "][5] = """ & objinfo6.text & """;"
			  next
      .Write "</script>"
    end if
%>

<%
end with
 %>

<html>
<head>
<link rel='STYLESHEET' href="/SW-Common/SW-Style.css">
</head>
<body>
    <form id='frmAssetMultilingual' name='frmAssetMultilingual' method='POST' action=''>
    <table width="100%">
        <tr>
            <td colspan='2'><h1>Assets Multilingual Management</h1>
		<font size="1">Fields marked with (<font color="red">*</font>) are compulsory.</font><br>
            These entries define what translated language titles and descriptions will appear on the selected (language/locales) and when it will appear.<br /><br />
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left style="width:25%;">
               Language-Locale:</td>
            <td CLASS=Medium ALIGN=left>
                <select name="Locales" id="Locales" onchange="ChangeLocale();">
                    <option value="cs-cz">cs-cz</option>
			  <option value="de-at">de-at</option>
                    <option value="de-de">de-de</option>
                    <option value="de-ch">de-ch</option>	
                    <option value="da-dk">da-dk</option>				
                    <option value="en-au">en-au</option>
                    <option value="en-ca">en-ca</option>
                    <option value="en-gb">en-gb</option>
                    <option value="en-ie">en-ie</option>
			  <option value="en-in">en-in</option>
                    <option value="en-sg">en-sg</option>
			  <option value="en-tt">en-tt</option>
                    <option value="en-tw">en-tw</option>
                    <option value="en-us">en-us</option>
                    <option value="es-es">es-es</option>
		        <option value="es-us">es-us</option>
		        <option value="es-mx">es-mx</option>	
		        <option value="es-cl">es-cl</option>			        
		        <option value="es-ar">es-ar</option>				
		        <option value="es-co">es-co</option>
		        <option value="es-pe">es-pe</option>
			  <option value="es-ec">es-ec</option>
			  <option value="es-cr">es-cr</option>
			  <option value="es-ve">es-ve</option>
			  <option value="es-gt">es-gt</option>
			  <option value="es-bo">es-bo</option>
			  <option value="es-do">es-do</option>
			  <option value="es-sv">es-sv</option>
			  <option value="es-uy">es-uy</option>
			  <option value="fr-be">fr-be</option>
                    <option value="fr-ca">fr-ca</option>
                    <option value="fr-ch">fr-ch</option>
                    <option value="fr-fr">fr-fr</option>
                    <option value="fi-fi">fi-fi</option>
                    <option value="it-ch">it-ch</option>
                    <option value="it-it">it-it</option>
			  <option value="nl-be">nl-be</option>
                    <option value="nl-nl">nl-nl</option>		
                    <option value="no-no">no-no</option>
			  <option value="pl-pl">pl-pl</option>		
                    <option value="pt-br">pt-br</option>
                    <option value="pt-pt">pt-pt</option>
			  <option value="ro-ro">ro-ro</option>
			  <option value="ru-ru">ru-ru</option>
                    <option value="sv-se">sv-se</option>
			  <option value="tr-tr">tr-tr</option>
                    <option value="zh-cn">zh-cn</option>
			</select>
                &nbsp;&nbsp;&nbsp;&nbsp; <input type="checkbox" checked id="chkAllLocales" name="chkAllLocales" style='display:none;' value="Apply to all locales" />
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
                Title<font color="red">*</font>:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='512' style="width: 300px;" id='txtDesc' name='txtDesc' />
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
                Description<font color="red">*</font>:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='512' style="width: 300px;" id='txtDesc1' name='txtDesc1' accept="text/html" />
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
                Start Date<font color="red">*</font>:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='10' style="width: 100px;" id='txtStartDate' name='txtStartDate' />
		  (mm/dd/yyyy)
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
                End Date<font color="red">*</font>:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='10' style="width: 100px;" id='txtEndDate' name='txtEndDate' />
	  (mm/dd/yyyy)
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
                Product Url:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='512' style="width: 300px;" id='txtProductUrl' name='txtProductUrl' />
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td colspan='2' align='left'> <b>Copy to:</b> <br />Select the additional locales you wish to update <br />(Note: Hold the shift key to select multiple language/locales)
            </td>
        </tr>        
        <tr>
            <td>
            </td>
            <td>
                <table border='0' cellpadding='0' cellspacing='0'>
                    <tr>
                        <td>
                            Available Locales <br />
                            <select size='8' multiple name="AllLocales" id="AllLocales">
                    <option value="cs-cz">cs-cz</option>
			  <option value="de-at">de-at</option>
                    <option value="de-de">de-de</option>
                    <option value="de-ch">de-ch</option>	
                    <option value="da-dk">da-dk</option>				
                    <option value="en-au">en-au</option>
                    <option value="en-ca">en-ca</option>
                    <option value="en-gb">en-gb</option>
                    <option value="en-ie">en-ie</option>
			  <option value="en-in">en-in</option>
                    <option value="en-sg">en-sg</option>
			  <option value="en-tt">en-tt</option>
                    <option value="en-tw">en-tw</option>
                    <option value="en-us">en-us</option>
                    <option value="es-es">es-es</option>
		        <option value="es-us">es-us</option>
		        <option value="es-mx">es-mx</option>	
		        <option value="es-cl">es-cl</option>	
		        <option value="es-ar">es-ar</option>				
		        <option value="es-co">es-co</option>
		        <option value="es-pe">es-pe</option>
			  <option value="es-ec">es-ec</option>
			  <option value="es-cr">es-cr</option>
			  <option value="es-ve">es-ve</option>
			  <option value="es-gt">es-gt</option>
			  <option value="es-bo">es-bo</option>
			  <option value="es-do">es-do</option>
			  <option value="es-sv">es-sv</option>
			  <option value="es-uy">es-uy</option>
			  <option value="fr-be">fr-be</option>
                    <option value="fr-ca">fr-ca</option>
                    <option value="fr-ch">fr-ch</option>
                    <option value="fr-fr">fr-fr</option>
                    <option value="fi-fi">fi-fi</option>
                    <option value="it-ch">it-ch</option>
                    <option value="it-it">it-it</option>
			  <option value="nl-be">nl-be</option>
                    <option value="nl-nl">nl-nl</option>		
                    <option value="no-no">no-no</option>
			  <option value="pl-pl">pl-pl</option>		
                    <option value="pt-br">pt-br</option>
                    <option value="pt-pt">pt-pt</option>
			  <option value="ro-ro">ro-ro</option>
			  <option value="ru-ru">ru-ru</option>
                    <option value="sv-se">sv-se</option>
			  <option value="tr-tr">tr-tr</option>
                    <option value="zh-cn">zh-cn</option>
                            </select>
                        </td>
                        <td>
                            &nbsp;&nbsp;<input type="button" value=">" class="NavLeftHighlight1" name="btnAproducts" onclick="AddRemoveOptions('AllLocales','SelLocales')" />&nbsp;&nbsp;
                            <br />
                            <br />
                            &nbsp;&nbsp;<input type="button" value="<" class="NavLeftHighlight1" name="btnRproducts" onclick="RemoveOption('SelLocales')" />&nbsp;&nbsp;
                        </td>
                        <td>
                            Selected Locales <br />
                            <select size='8' multiple name="SelLocales" id="SelLocales">
                            </select>
                        </td>
                    </tr>
                </table>
            </td>
            
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
        </tr>                        
        <tr>
            <td CLASS=Medium ALIGN=left>
               <input type='button' id='btnSave' name='btnSave' CLASS='NavLeftHighlight1' onclick='return SaveClicked();' value='Save'/>
            </td>
            <td CLASS=Medium ALIGN=left>
               <input type='button' id='btnClose' name='btnClose' CLASS='NavLeftHighlight1' onclick='window.close();' value='Close' />
               <input type='hidden' id='hidSave' name='hidSave'/>
               <input type='hidden' id='hidAllLocales' name='hidAllLocales'/>
               <input type='hidden' id='hidCheckbox' name='hidCheckbox'/>
            </td>
        </tr>
    </table>
    <div id="message">
        <script language='javascript' type="text/javascript"> ChangeLocale(); </script> 
    </div>
    </form>
</body>
</html>

