<!--#include virtual="/sw-administratorNT/SW-PCAT_FNET_IISSERVER.asp"-->
<!--#include virtual="/include/functions_string.asp"-->

<script type='text/javascript' language='javascript'> 
    function SaveClicked()
    {
        var desc = document.getElementById("txtDesc").value;
        if (desc == '')
        {
            alert('Please enter valid title.');
            return false;
        }
        else
        {
            document.getElementById('hidSave').value = 'Save';
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
               }
            }
        ////}
    }
    
    function ChangeCheckLocale()
    {
        var Locales, chkbox;
        chkbox = document.getElementById("ChangeCheckLocales");
        Locales = document.getElementById("PCat_Locales");
        for (var i=0; i<Locales.length; i++) 
        {
            if (chkbox.checked)
              Locales[i].selected = true;
            else
              Locales[i].selected = false;
        }
    }
</script>

<%
Dim SaveRecord, strPID, Site_ID, objcol, Products, strparameters, txtDesc1, Locales1, objinfo2, objinfo1, objinfo, desc, objinfo3, objinfo4, objinfo5
Dim txtStartDate, txtEndDate, txtDesc2

SaveRecord=true
strPID = Request.QueryString("ID")
Site_ID= Request.QueryString("SiteId")


 if (Request.Form("hidSave") <> "") then
        txtDesc1 = Request.Form("txtDesc")
        Locales1 = Request.Form("Locales")
        txtStartDate = Request.Form("txtStartDate")
        txtEndDate = Request.Form("txtEndDate")
        txtDesc2 = Request.Form("txtDesc1")
        ''Response.Write txtDesc1
        ''Response.Write Locales1
        ''Response.End ************StartDate,EndDate,Desc
        
        strparameters =  "operation=UL&assetpid=" & strPID & "&ProdDesc=" & txtDesc1 & "&Locale=" & Locales1 & "&SiteID=" & Site_ID & "&StartDate=" & txtStartDate & "&EndDate=" & txtEndDate & "&Desc" & txtDesc2
        
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
                  .Write "myarray[" & icol & "][0] = '" & objinfo1.text & "';"
                  .Write "myarray[" & icol & "][1] = '" & objinfo2.text & "';"
                  .Write "myarray[" & icol & "][2] = '" & objinfo3.text & "';"
                  .Write "myarray[" & icol & "][3] = '" & CDate(objinfo4.text) & "';"
                  .Write "myarray[" & icol & "][4] = '" & CDate(objinfo5.text) & "';"
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
    <table>
        <tr>
            <td colspan='2'><h1>Assets Multilingual Management</h1>
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
               Locales:</td>
            <td CLASS=Medium ALIGN=left>
                <input type='checkbox' id="ChangeCheckLocales" name="ChangeCheckLocales" onchange="ChangeCheckLocale();" CLASS="Medium" value="Select All Locales" />
                <br />
                <br />
                <select LANGUAGE="JavaScript" multiple size="5" NAME="PCat_Locales" CLASS="Medium">
                <%
                  if not(objcol is nothing) then
			         for icol=0 to objcol.childnodes.length-1
                        set objinfo=objcol.childnodes(icol)
                        set objinfo1=objinfo.childnodes(0)
                        .Write objinfo1.text
			         next
                  end if
                %>
                </select>
            </td>
        </tr>
        <tr>
            <td CLASS=Medium ALIGN=left>
                Title:
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
                Description:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='512' style="width: 300px;" id='txtDesc1' name='txtDesc1' />
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
                Start Date:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='15' style="width: 100px;" id='txtStartDate' name='txtStartDate' />
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
                End Date:
            </td>
            <td CLASS=Medium ALIGN=left>
                <input type='text' value='' maxlength='15' style="width: 100px;" id='txtEndDate' name='txtEndDate' />
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
               <input type='button' id='btnSave' name='btnSave' CLASS=NavLeftHighlight1 onclick='return SaveClicked();' value='Save'/>
            </td>
            <td CLASS=Medium ALIGN=left>
               <input type='button' id='btnClose' name='btnClose' CLASS=NavLeftHighlight1 onclick='window.close();' value='Close' />
               <input type='hidden' id='hidSave' name='hidSave'/>
            </td>
        </tr>
    </table>
    <div id="message"></div>
    </form>
</body>
</html>

