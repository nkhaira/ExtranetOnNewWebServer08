<%response.Buffer=false%>
<html>
<body>
<form method="post" >
<% Dim strSelectedCat, strRequestID
strRequestID = Request.QueryString("ID") 
%>
<input type="hidden" name="hdFTP" value="1" />
<div id="prepage" style="position:absolute; font-family:arial; font-size:16; left:0px; top:20px; background-color:white; height:100%; width:100%;"> 
<TABLE width="100%"><TR><TD align="center"><B>Please wait while file loads to CDN...</B></TD></TR></TABLE>
</div>
<table  width="700"  cellpadding="0">
        <tr>
            <td><b>File FTP Utility</b>
            </td>
        </tr>
    <%if (strRequestID = "" or strRequestID = null) then %>
        <tr >
            <td><b>Please select categories and click on below button for bulk FTP:</b><br />
                (<font color="red"> Note: Only categories with CDN Enabled from AMS Edit-Category Options page will be displayed.</font>)
            </td>
        </tr>
        <tr>
            <td >
                <%
                   response.write (DisplayCategories())  
                %>
            </td>
        </tr>
        <tr >
            <td><input type="submit" value="Submit"/> (<font color="red">Only last 12 months Files will be FTPed.</font>)
            </td>
        </tr>
    <%end if%>
</table>

</form>

<% 
Dim isSubmit, strErrorMesages, SQL, tempCount, strFTPLocation, strSelectedCatTemp, isAlreadyFTPed
isSubmit = Request.Form("hdFTP")
isAlreadyFTPed = false
if ((isSubmit = 1 and len(strSelectedCat) > 0) OR (strRequestID <> "" OR strRequestID <> null)) then
    server.scripttimeout = 28800 '8 hrs, earlier it was 1200
    
    ''0. Initialize error variables so that all error can log at a time
    strErrorMesages = ""    
    tempCount=0
    ''1. Connect to fluke_sidewide database to get list of Product Softwares
    ''1.1 Connection
    Dim connFlukeSitewide, connFnetPE
	set connFlukeSitewide = Server.CreateObject("ADODB.Connection")
	connFlukeSitewide.ConnectionTimeout = 1200 '20 Minutes
	connFlukeSitewide.Open GetFlukeSitewideConn()
	connFlukeSitewide.CommandTimeout = 28800 '8 hrs, earlier it was 1200
	
	''1.2 ProductEngine2 connection
	set connFnetPE = Server.CreateObject("ADODB.Connection")
	connFnetPE.ConnectionTimeout = 1200
	connFnetPE.Open GetPEConn()
	connFnetPE.CommandTimeout = 28800 '8 hrs, earlier it was 1200
	
    if (isSubmit = 1 and len(strSelectedCat) > 0) then
        strSelectedCatTemp = Replace(strSelectedCat,"[","")
        strSelectedCatTemp = Replace(strSelectedCatTemp, "]", "")
        strSelectedCatTemp = Left(strSelectedCatTemp, Len(strSelectedCatTemp) - 2)
        
        SQL = "SELECT Calendar.ID, Category_Id,Item_Number,PID, File_Name, Language.ISO2 as Lang,CDN_Required, CDNFilePath " 
        SQL = SQL & " FROM Calendar LEFT OUTER JOIN Language  "
        SQL = SQL & " ON Language.Code = Calendar.Language "
        SQL = SQL & " WHERE Site_ID = 82 and not (pid = -1 or pid = 0) AND DATEDIFF(month,Udate, getdate()) <= 12 "
        SQL = SQL & " AND (CDN_Required =0 OR CDN_Required is null )  "
        SQL = SQL & " AND Category_Id IN ("  & strSelectedCatTemp & ") " 
    elseif (strRequestID <> "" OR strRequestID <> null) then
        SQL = "SELECT Calendar.ID, Category_Id,Item_Number,PID, File_Name, Language.ISO2 as Lang,CDN_Required, CDNFilePath " 
        SQL = SQL & " FROM Calendar LEFT OUTER JOIN Language  "
        SQL = SQL & " ON Language.Code = Calendar.Language "
        SQL = SQL & " WHERE Site_ID = 82 and not (pid = -1 or pid = 0) "
        SQL = SQL & " AND (CDN_Required =0 OR CDN_Required is null )  "
        SQL = SQL & " AND Calendar.ID = "  & strRequestID
    end if
    
    Set rsAssets = Server.CreateObject("ADODB.Recordset")
    rsAssets.Open SQL, connFlukeSitewide, 3, 3  
    
    ''Open 
    ''2. Loop the list
    
    if not rsAssets.EOF Then 
        rsAssets.MoveFirst()
        if (strRequestID <> "" OR strRequestID <> null) then
            ''Response.write ("The File is FTPing. Please wait...<br>")
        else
            Response.Write "Start FTPing:<br>"
        end if
        
    else
        if (strRequestID <> "" OR strRequestID <> null) then
            isAlreadyFTPed = true
            Response.Write "<b><font color='red'>Error:</font></b> The file is already FTPed OR the Asset Id does not exist. Please check.<br>" 
        else
            Response.Write "<b><font color='red'>Error:</font></b> No file found for FTP.<br>"
        end if
    end if
    
    Do While not rsAssets.EOF 
      
      strFTPLocation = ""
      tempCount = tempCount +1
            
      ''3. FTPed the asset file if not already FTPed
      if isnull(rsAssets("CDNFilePath")) OR rsAssets("CDNFilePath") = "" then
           ''3.1 check if the file is not already FTPed
           ''3.2 FTP the file
           ''if (ValidateFTPFileName("FNet/" & rsAssets("File_Name"),rsAssets("ID"))) Then
                strFTPFilePath = FTPFile(Server.MapPath("/PortWeb/" & rsAssets("File_Name")), rsAssets("File_Name")) ''FTP the file
                if (strFTPFilePath <> "") then
                    ''4. Update the record by setting CDN_Required = -1 and CDNFilePath = <FPTed path>
                    UpdateCalendarFTPDetail rsAssets("ID"), -1, rsAssets("File_Name")
                    
                    ''5. Update The CDN information in ProductEngine2 ProductLocalizedProperties table
                    UpdatePEFTPDetail rsAssets("PID"), -1, rsAssets("File_Name"), rsAssets("Lang") 
                    
                end if
           ''end if
      end if
      
      rsAssets.MoveNext()
    Loop
            
    ''6. Clear objects
    rsAssets.close
    set rsAssets = nothing
    
    connFlukeSitewide.close
	set connFlukeSitewide = nothing
	
	connFnetPE.Close
    set connFnetPE = nothing
    
    ''7. Write all errors at a time
    if strErrorMesages <> "" then
        Response.Write "<br><b>Errors are:</b><br>"
        Response.Write strErrorMesages
        Response.Write "<br>"
    end if
 
    if (strRequestID <> "" OR strRequestID <> null)  then
        if isAlreadyFTPed = false then Response.Write "The file FTPed successfully and related Database has been updated.<br>" 
    else
            Response.Write "<br>END File FTPs"
    end if
    
end if 


' --------------------------------------------------------------------------------------
' Added on 10th Aug 2010, for CDN File FTP
' To FTP given file
' --------------------------------------------------------------------------------------
function FTPFile(sSourceFile, sFTPLocation)
    Dim sFTPLoc, sFTPLocPrefix 
    sFTPLocPrefix = ""
    
    set obj = server.CreateObject("FTPclient.FTP") 
    ''Seperate FTP server name and their credentials depending on Current server
    if instr(UCase(Request.ServerVariables("SERVER_NAME")),"DTMEVTVSDV15") > 0 OR instr(UCase(Request.ServerVariables("SERVER_NAME")),"DEV") > 0 then
        ''Comment/Uncomment to locate Dev FTP server
        'sFTPLocPrefix = "FNet/"
        'obj.Hostname="Fluke.ingest.cdn.level3.net"
        'obj.Username="fluke"
        'obj.Password="KaxETa6vsg"
        
        sFTPLocPrefix = "fnetimages/content/FNet_Dev/FNet/"
        obj.Hostname="ftp.flukenetworks.com"
        obj.Username="fnetimages"
        obj.Password="FN3tImag3s"
        
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"TEST") > 0 then
         ''Comment/Uncomment to locate Test FTP server
         'sFTPLocPrefix = "FNet/"
         'obj.Hostname="Fluke.ingest.cdn.level3.net"
         'obj.Username="fluke"
         'obj.Password="KaxETa6vsg"
        
          sFTPLocPrefix = "fnetimages/content/FNet_Test/FNet/"
        obj.Hostname="ftp.flukenetworks.com"
        obj.Username="fnetimages"
        obj.Password="FN3tImag3s"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"PRD") > 0 then
        sFTPLocPrefix = "FNet/"
        obj.Hostname="Fluke.ingest.cdn.level3.net"
        obj.Username="fluke"
        obj.Password="KaxETa6vsg"
    else
	    'Default to PRODUCTION
	    sFTPLocPrefix = "FNet/"
	    obj.Hostname="Fluke.ingest.cdn.level3.net"
        obj.Username="fluke"
        obj.Password="KaxETa6vsg"
	end if 
	''End 
    
    sFTPLoc = sFTPLocPrefix & sFTPLocation
    
    On error resume next
    obj.UploadFile sSourceFile,sFTPLoc

    if err.number <> 0 then 
        ''Response.Write err.Description
        strErrorMesages = strErrorMesages & "<LI>File not found: The file """ & sFTPLoc & """ is not present on File Server to FTP.</LI>"
        sFTPLoc = ""
    end if
    On error goto 0
    
    FTPFile = sFTPLoc
end function

' --------------------------------------------------------------------------------------
'Added on 18th Aug 2010, for CDN Implementation
'Validation for CDN FTP File name duplicacy
' --------------------------------------------------------------------------------------
function ValidateFTPFileName(sFileNameToFTP, sID)
    Dim sSQL, sIsValid 
    sIsValid = True
    
    sSQL = "SELECT CDNFilePath FROM Calendar WHERE Site_ID=" & CInt(Site_ID) & " AND CDNFilePath = '" & sFileNameToFTP & "'"
    if sID <> "add" then
        sSQL = sSQL & " AND ID <> " & sID
    end if
    
    Set rsFileName = Server.CreateObject("ADODB.Recordset")
    rsFileName.Open sSQL, connFlukeSitewide, 3, 3  

    if not rsFileName.EOF then
        if not (isnull(rsFileName("CDNFilePath")) OR rsFileName("CDNFilePath") = "") then
            sIsValid = False
        end if
    end if

    rsFileName.close
    set rsFileName = nothing   
    
    ValidateFTPFileName = sIsValid
    
    if (sIsValid = False) Then
        strErrorMesages = strErrorMesages & "<LI>File already exists: The file name """ & sFileNameToFTP & """ already exists on FTP Server. Please rename the file and upload again.</LI>"
    end if
end function
'End

' --------------------------------------------------------------------------------------
' Added on 20th Aug 2010
' To Save FTPed file details in 
' --------------------------------------------------------------------------------------
function UpdateCalendarFTPDetail(sID, sCDNRequried, sFTPLocation)
    Dim sSQL
    sSQL = "UPDATE Calendar SET CDN_Required = -1,CDNFilePath='" & sFTPLocation & "' WHERE Site_ID=82 AND Id = " & sID 
    connFlukeSitewide.Execute(sSQL)
end function
'End

' --------------------------------------------------------------------------------------
' Added on 20th Aug 2010
' To Save FTPed file details in ProductEngine2
' --------------------------------------------------------------------------------------
function UpdatePEFTPDetail(iPID, sCDNRequried, sFTPLocation, sLang)
    Dim sSQL
     
    sSQL = "EXEC AMS_AssetDetails_SAVE " & iPID & ", " & sCDNRequried & ", '" & sFTPLocation & "', '" & sLang & "'"
    'Response.Write sSQL
    connFnetPE.Execute(sSQL)
     
end function
'End

' --------------------------------------------------------------------------------------
' Added on 20th Aug 2010
' GEt connection stringfor Fluke_Sitewide
' --------------------------------------------------------------------------------------
function GetFlukeSitewideConn()
    Dim str
    if instr(UCase(Request.ServerVariables("SERVER_NAME")),"DTMEVTVSDV15") > 0 OR instr(UCase(Request.ServerVariables("SERVER_NAME")),"DEV") > 0 then
        str = "DRIVER={SQL Server};SERVER=EVTIBG18.TC.FLUKE.COM;UID=SITEWIDE_WEB;DATABASE=FLUKE_SITEWIDE;pwd=tuggy_boy"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"TEST") > 0 then
        str = "DRIVER={SQL Server};SERVER=FLKTST18.DATA.IB.FLUKE.COM;UID=SITEWIDE_WEB;DATABASE=FLUKE_SITEWIDE;pwd=tuggy_boy"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"PRD") > 0 then
        str = "DRIVER={SQL Server};SERVER=FLKPRD18.DATA.IB.FLUKE.COM;UID=SITEWIDE_WEB;DATABASE=FLUKE_SITEWIDE;pwd=tuggy_boy"
    else
	    'Default to PRODUCTION
	    str = "DRIVER={SQL Server};SERVER=FLKPRD18.DATA.IB.FLUKE.COM;UID=SITEWIDE_WEB;DATABASE=FLUKE_SITEWIDE;pwd=tuggy_boy"
	end if 
	''End 
   
    GetFlukeSitewideConn = str
end function
'End

' --------------------------------------------------------------------------------------
' Added on 20th Aug 2010
' Get connection stringfor ProductEngine2
' --------------------------------------------------------------------------------------
function GetPEConn()
    Dim str
    if instr(UCase(Request.ServerVariables("SERVER_NAME")),"DTMEVTVSDV15") > 0 OR instr(UCase(Request.ServerVariables("SERVER_NAME")),"DEV") > 0 then
        str = "DRIVER={SQL Server};SERVER=dtmevtsvdb02.danahertm.com;UID=Fnet_Web_SQL;DATABASE=ProductEngine2;pwd=?Twink123"
        'str = "DRIVER={SQL Server};SERVER=dtmevtsvdb02.danahertm.com;UID=FnetRead;DATABASE=ProductEngine2;pwd=?ReadFnet"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"TEST") > 0 then
        str = "Provider=SQLOLEDB;SERVER=dtmflksvdb01.data.ib.fluke.com;DATABASE=ProductEngine2;Persist Security Info=True;Integrated Security=SSPI;"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"PRD") > 0 then
        str = "Provider=SQLOLEDB;SERVER=dtmflkmsql04.data.ib.fluke.com;DATABASE=ProductEngine2;Persist Security Info=True;Integrated Security=SSPI;"
    else
	    'Default to PRODUCTION
	    str = "Provider=SQLOLEDB;SERVER=dtmflkmsql04.data.ib.fluke.com;DATABASE=ProductEngine2;Persist Security Info=True;Integrated Security=SSPI;"
	end if 
	''End 
   
    GetPEConn = str
end function
'End

' --------------------------------------------------------------------------------------
' Added on 13th Sept 2010
' To display category ceckboxes
' --------------------------------------------------------------------------------------
function DisplayCategories() 
    ''Get previouslly checked checkboxes
   
    strSelectedCat = ""
    for each x in Request.Form 
        if (instr(x,"Cat_") > 0) then
            'Response.Write("<br>" & x & " = " & Request.Form(x)) 
            strSelectedCat = strSelectedCat & "[" & Request.Form(x) & "], "
        end if
    next 
    
 
    Dim connCategory, sqlCategory,strCatList
	set connCategory = Server.CreateObject("ADODB.Connection")
	connCategory.ConnectionTimeout = 1200
	connCategory.Open GetFlukeSitewideConn()
	
	sqlCategory = "SELECT ID, Title  from Calendar_Category Where Site_Id = 82 AND CDN_Implementation = -1 " 
	sqlCategory = sqlCategory & " ORDER BY Title "
        
    Set rsCat = Server.CreateObject("ADODB.Recordset")
    rsCat.Open sqlCategory, connCategory, 3, 3  
    
    strCatList = ""
    if not rsCat.EOF Then rsCat.MoveFirst()
    Do While not rsCat.EOF 
       strCatList = strCatList & " <INPUT TYPE=""Checkbox"" NAME=""Cat_" & rsCat("ID") & """ VALUE=""" & rsCat("ID") & """" 
       if (instr(strSelectedCat, "[" & rsCat("ID") & "]")) then
            strCatList = strCatList & " CHECKED "
       end if
       strCatList = strCatList & " >&nbsp;&nbsp;" & rsCat("TITLE") & "<BR>"
	   rsCat.MoveNext()
    Loop
	
	connCategory.close
	set connCategory = nothing
	
	DisplayCategories = strCatList
end function

%>
<script language="javascript" type="text/javascript">
    waitPreloadPage();
    function waitPreloadPage() 
    {
        //alert("test");
         //DOM
        if (document.getElementById)
        {
            document.getElementById('prepage').style.visibility='hidden';
        }
        else
        {
            if (document.layers)
            { //NS4
                document.prepage.visibility = 'hidden';
            }
            else 
            { //IE4
                document.all.prepage.style.visibility = 'hidden';
             }
        }
    }

</script>
</body>
</html>