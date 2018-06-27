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
            <td><b>File FTP Utility (DataPaq) </b>
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
            <td><input type="submit" value="Submit"/> <font color="red"></font>
            </td>
        </tr>
    <%end if%>
</table>

</form>

<% 
Dim isSubmit, strErrorMesages, SQL, tempCount, strFTPLocation, strSelectedCatTemp, isAlreadyFTPed
isSubmit = Request.Form("hdFTP")
isAlreadyFTPed = false
'Response.Write("isSubmit" &isSubmit)

if ((isSubmit = 1 and len(strSelectedCat) > 0) OR (strRequestID <> "" OR strRequestID <> null)) then
    server.scripttimeout = 28800 '8 hrs, earlier it was 1200
    
    ''0. Initialize error variables so that all error can log at a time
    strErrorMesages = ""    
    tempCount=0
    ''1. Connect to fluke_sidewide database to get list of Product Softwares
    ''1.1 Connection
    Dim connFlukeSitewide
	set connFlukeSitewide = Server.CreateObject("ADODB.Connection")
	connFlukeSitewide.ConnectionTimeout = 1200 '20 Minutes
	connFlukeSitewide.Open GetFlukeSitewideConn()
	connFlukeSitewide.CommandTimeout = 28800 '8 hrs, earlier it was 1200
		
    if (isSubmit = 1 and len(strSelectedCat) > 0) then
        strSelectedCatTemp = Replace(strSelectedCat,"[","")
        strSelectedCatTemp = Replace(strSelectedCatTemp, "]", "")
        strSelectedCatTemp = Left(strSelectedCatTemp, Len(strSelectedCatTemp) - 2)
        
        SQL = "SELECT TOP 50 Calendar.ID, Category_Id,Item_Number,PID, File_Name, Language.ISO2 as Lang,CDN_Required, CDNFilePath " 
        SQL = SQL & " FROM Calendar LEFT OUTER JOIN Language  "
        SQL = SQL & " ON Language.Code = Calendar.Language "
        'SQL = SQL & " WHERE Site_ID = 29  AND DATEDIFF(day,Udate, getdate()) <= 2 "
        SQL = SQL & " WHERE Site_ID = 29  "
        ' SQL = SQL & " and not (pid = -1 or pid = 0)"
        SQL = SQL & " AND (CDN_Required =0 OR CDN_Required is null )  "
        SQL = SQL & " AND Category_Id IN ("  & strSelectedCatTemp & ")  ORDER BY Calendar.ID DESC " 
    elseif (strRequestID <> "" OR strRequestID <> null) then
        SQL = "SELECT Calendar.ID, Category_Id,Item_Number,PID, File_Name, Language.ISO2 as Lang,CDN_Required, CDNFilePath " 
        SQL = SQL & " FROM Calendar LEFT OUTER JOIN Language  "
        SQL = SQL & " ON Language.Code = Calendar.Language "
        SQL = SQL & " WHERE Site_ID = 29  "
        ' SQL = SQL & " and not (pid = -1 or pid = 0) " 
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
            Response.write ("The File is FTPing. Please wait...<br>")
        else
            Response.Write "Start FTPing:<br>"
        end if
        
    else
        if (strRequestID <> "" OR strRequestID <> null) then
            isAlreadyFTPed = true
            Response.Write "<b><font color='red'>Error:</font></b> The file is already FTPed OR the Asset Id does not exist. Please check.<br>" 
        else
            'Response.Write "<b><font color='red'>Error:</font></b> No file found for FTP.<br>"
            strErrorMesages = "<b><font>No file found for FTP.</font></b> <br>"
        end if
    end if
    
    Do While not rsAssets.EOF 
      
      strFTPLocation = ""
      tempCount = tempCount +1  
      if not isnull(rsAssets("File_Name")) and rsAssets("File_Name") <> "" then  
      strFTPFilePath = ""
       strFTPFilePath = FTPFile(Server.MapPath("/Datapaq/" & rsAssets("File_Name")), rsAssets("File_Name")) ''FTP the file
        if (strFTPFilePath <> "") then
          ''4. Update the record by setting CDN_Required = -1 and CDNFilePath = <FPTed path>
          UpdateCalendarFTPDetail rsAssets("ID"), -1, rsAssets("File_Name")
        end if 
      end if      
      ''3. FTPed the asset file if not already FTPed
      'if isnull(rsAssets("CDNFilePath")) OR rsAssets("CDNFilePath") = "" then
           ''3.1 check if the file is not already FTPed
           ''3.2 FTP the file
           ''if (ValidateFTPFileName("FNet/" & rsAssets("File_Name"),rsAssets("ID"))) Then
           'if (ValidateFTPFileName(rsAssets("File_Name"),rsAssets("ID"))) Then
              '  strFTPFilePath = FTPFile(Server.MapPath("/Datapaq/" & rsAssets("File_Name")), rsAssets("File_Name")) ''FTP the file
               ' Response.Write (rsAssets("File_Name")) 
           'end if   
           
             ' if (strFTPFilePath <> "") then
                    ''4. Update the record by setting CDN_Required = -1 and CDNFilePath = <FPTed path>
                 '   UpdateCalendarFTPDetail rsAssets("ID"), -1, rsAssets("File_Name")
                                       
             ' end if
           ''end if
      'end if
      
      rsAssets.MoveNext()
    Loop
            
    ''6. Clear objects
    rsAssets.close
    set rsAssets = nothing
    
    connFlukeSitewide.close
	set connFlukeSitewide = nothing
	
	'connFnetPE.Close
    'set connFnetPE = nothing
    
    ''7. Write all errors at a time
    if strErrorMesages <> "" then
        Response.Write "<br><b>Errors are:</b><br>"
        Response.Write strErrorMesages
        Response.Write "<br>"
    else
        Response.Write "The file FTPed successfully and related Database has been updated.<br>" 
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
       
            sFTPLocPrefix = "content/Datapaq/Dev/datapaq/"
            obj.Hostname="ftp.fluke.com/"
            obj.Username="flukecontent"
            obj.Password="c0nt3nt4F1uke"
        
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"TEST") > 0 then
          ''test location   
            sFTPLocPrefix = "content/Datapaq/Test/datapaq/"
            obj.Hostname="ftp.fluke.com/"
            obj.Username="flukecontent"
            obj.Password="c0nt3nt4F1uke"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"PRD") > 0 then
             sFTPLocPrefix = "Fluke/datapaq/"
             obj.Hostname="Fluke.ingest.cdn.level3.net"
             obj.Username="fluke"
             obj.Password="KaxETa6vsg"
    else
	        'sFTPLocPrefix = "content/Datapaq/Dev/datapaq/"
            'obj.Hostname="ftp.fluke.com/"
            'obj.Username="flukecontent"
            'obj.Password="c0nt3nt4F1uke"
            
            sFTPLocPrefix = "Fluke/datapaq/"
             obj.Hostname="Fluke.ingest.cdn.level3.net"
             obj.Username="fluke"
             obj.Password="KaxETa6vsg"
	end if 
	''End 
    
    sFTPLoc = sFTPLocPrefix & sFTPLocation
    
    On error resume next
    obj.UploadFile sSourceFile,sFTPLoc
    
    '--------------------
    Dim fsoRenameDataPaq_FileServerFile
    Set fsoRenameDataPaq_FileServerFile = CreateObject("Scripting.FileSystemObject")
    
    Dim DataPaqFleNameBeforeModification : DataPaqFleNameBeforeModification  =  fsoRenameDataPaq_FileServerFile.getfilename(sSourceFile)
    Dim DataPaqFileNameAfterModification : DataPaqFileNameAfterModification = "AmsToCDNUploaded" & DataPaqFleNameBeforeModification
    
    Dim DataPaqCdnUploadedFilePath : DataPaqCdnUploadedFilePath = Replace(sSourceFile,DataPaqFleNameBeforeModification, DataPaqFileNameAfterModification)
    
    if fsoRenameDataPaq_FileServerFile.FileExists(sSourceFile) Then
        if fsoRenameDataPaq_FileServerFile.FileExists(DataPaqCdnUploadedFilePath) then
            on error resume next
            fsoRenameDataPaq_FileServerFile.DeleteFile DataPaqCdnUploadedFilePath, True ' Delete Record, True=Read Only and non-Read Only
            on error goto 0
        end if 
                              
        fsoRenameDataPaq_FileServerFile.MoveFile sSourceFile,DataPaqCdnUploadedFilePath
     end if 
    '------------------------

    if err.number <> 0 then       
        'strErrorMesages = strErrorMesages & "<LI> File not found: The file """ & sFTPLoc & """ is not present on File Server to FTP.</LI>"
        strErrorMesages = strErrorMesages & "<LI> File not FTPed: The file """ & sFTPLoc & """ is not present on File Server or not   FTPed to CDN location.</LI>"
        sFTPLoc = ""
    end if
    On error goto 0
    
    FTPFile = sFTPLoc
end function

' --------------------------------------------------------------------------------------

'Validation for CDN FTP File name duplicacy
' --------------------------------------------------------------------------------------
function ValidateFTPFileName(sFileNameToFTP, sID)
    'Response.Write("Filename" & sFileNameToFTP) 
    Dim sSQL, sIsValid 
    sIsValid = True
    
   ' sSQL = "SELECT CDNFilePath FROM Calendar WHERE Site_ID=" & CInt(Site_ID) & " AND CDNFilePath = '" & sFileNameToFTP & "'"
    sSQL = "SELECT CDNFilePath FROM Calendar WHERE Site_ID=29 AND CDNFilePath = '" & sFileNameToFTP & "'"
     
    if sID <> "add" then
        sSQL = sSQL & " AND ID <> " & sID
    end if
    'Response.Write(sSQL)
    
    Set rsFileName = Server.CreateObject("ADODB.Recordset")
    rsFileName.Open sSQL, connFlukeSitewide, 3, 3  

    if not rsFileName.EOF then
        'Response.Write ("path=" & rsFileName("CDNFilePath"))
        if not (isnull(rsFileName("CDNFilePath")) OR rsFileName("CDNFilePath") = "") then
            sIsValid = False
            
        end if
    end if

    rsFileName.close
    set rsFileName = nothing   
    
    ValidateFTPFileName = sIsValid
    
    if (sIsValid = False) Then
    '    strErrorMesages = strErrorMesages & "<LI>File already exists: The file name """ & sFileNameToFTP & """ already exists on FTP Server. Please rename the file and upload again.</LI>"
    end if
end function
'End

' --------------------------------------------------------------------------------------
' Added on 20th Aug 2010
' To Save FTPed file details in 
' --------------------------------------------------------------------------------------
function UpdateCalendarFTPDetail(sID, sCDNRequried, sFTPLocation)
    Dim sSQL
    sSQL = "UPDATE Calendar SET CDN_Required = -1,CDNFilePath='" & sFTPLocation & "' WHERE Site_ID=29 AND Id = " & sID 
    connFlukeSitewide.Execute(sSQL)
end function
'End



' --------------------------------------------------------------------------------------
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
	    'str = "DRIVER={SQL Server};SERVER=EVTIBG18.TC.FLUKE.COM;UID=SITEWIDE_WEB;DATABASE=FLUKE_SITEWIDE;pwd=tuggy_boy"
	    str = "DRIVER={SQL Server};SERVER=FLKPRD18.DATA.IB.FLUKE.COM;UID=SITEWIDE_WEB;DATABASE=FLUKE_SITEWIDE;pwd=tuggy_boy"
	end if 
	''End 
   
    GetFlukeSitewideConn = str
end function


' --------------------------------------------------------------------------------------
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
	
	sqlCategory = "SELECT ID, Title  from Calendar_Category Where Site_Id = 29 AND CDN_Implementation = -1 " 
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