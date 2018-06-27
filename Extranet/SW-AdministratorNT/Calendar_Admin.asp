<%@Language="VBScript" Codepage=65001%>
<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->
<%
' --------------------------------------------------------------------------------------
'
' Author: Kelly Whitlock
'
' --------------------------------------------------------------------------------------
'response.buffer = true
' --------------------------------------------------------------------------------------
' Declarations for FileUpEE and Archive
' --------------------------------------------------------------------------------------

Dim Script_Debug
Script_Debug = false

''Nitin Code Changes Start
Dim FileUpEE_Flag, FileUpEE_Remote_Flag, FileUpEE_TempPath, ServerName, AssetCategorySQL
''Nitin Code Changes End

ServerName = UCase(request.ServerVariables("SERVER_NAME"))

FileUpEE_TempPath = Server.MapPath("/SW-FileUp_Temp")
' --------------------------------------------------------------------------------------



Dim oFileUpEE
Dim oFile
Dim intSAResult
Dim oFormItem
Dim oSubItem

Dim arch

if CInt(request("FileUpEE_Flag")) = CInt(true)        then  ' Set in Calendar_Edit.asp
  FileUpEE_Flag       = true
else
  FileUpEE_Flag       = false
end if

if CInt(request("FileUpEE_Remote_Flag")) = CInt(true) then  ' Set in Calendar_Edit.asp
  FileUpEE_Remote_Flag = true
else
  FileUpEE_Remote_Flag = false
end if

Set Arch      = Server.CreateObject("SoftArtisans.Archive")
Arch.TempPath = FileUpEE_TempPath

Set oFileUpEE = Server.CreateObject("SoftArtisans.FileUpEE")
oFileUpEE.TransferStage = saWebServer             ' Must be set for WebServer as opposed to FileServer
oFileUpEE.DynamicAdjustScriptTimeout(saWebServer) = true
oFileUpEE.TempStorageLocation(saWebServer) = FileUpEE_TempPath

if Script_Debug then
  oFileUpEE.DebugLevel = 3
  oFileUpEE.DebugLogFile = Server.MapPath("/SW-FileUp_Log/WS-Debug_Log.txt")

  oFileUpEE.AuditLogDestination
  oFileUpEE.AuditLogFile(saWebServer) = Server.MapPath("/SW-FileUp_Log/WS_Audit_Log.txt")
end if  

if isblank(request.querystring("cProgressID")) and isblank(request.querystring("wProgressID")) then
  oFileUpEE.ProgressIndicator(saClient)    = false
  oFileUpEE.ProgressIndicator(saWebServer) = false
else  
  oFileUpEE.ProgressIndicator(saClient)    = true
  oFileUpEE.ProgressIndicator(saWebServer) = true
  oFileUpEE.ProgressID(saClient)           = CInt(request.querystring("cProgressID"))
  oFileUpEE.ProgressID(saWebServer)        = CInt(request.querystring("wProgressID"))
end if  

on error resume next
oFileUpEE.ProcessRequest Request, False, False
if err.Number <> 0 then
	response.write "<B>WebServer ProcessRequest Error</B><BR>" & Err.Description & " (" & Err.Source & ")"
	response.end
end if
on error goto 0

oFileUpEE.TargetURL = "http://" & request.ServerVariables("SERVER_NAME") & "/SW-FileUp_Transfer/FileUpEE_FileServer.asp"
oFileUpEE.OverwriteFiles = true

Dim File_Name_Flag, File_Name_POD_Flag, Thumbnail_Flag

File_Name_Flag     = false
File_Name_POD_Flag = false
Thumbnail_Flag     = false

Dim ScriptTimeOut
ScriptTimeOut = 120       '  2 Minutes for no uploads

for each oFormItem in oFileUpEE.Files
  if oFormItem.size <> 0 then
    select case LCase(oFormItem.Name)
      case "file_name"
        File_Name_Flag     = true
        ScriptTimeOut = 600   ' 10 Minutes for large file uploads
      case "file_name_pod"
        File_Name_POD_Flag = true
        ScriptTimeOut = 600   ' 10 Minutes for large file uploads
      case "thumbnail"
          Thumbnail_Flag   = true
    end select    
  end if
next

server.scripttimeout = ScriptTimeOut

' --------------------------------------------------------------------------------------
' Posting Form Debug
' --------------------------------------------------------------------------------------
Script_Debug = false

if Script_Debug = true then

  with response
	for each oFormItem in oFileUpEE.Form
		.write oFormItem.Name & ": "
		if IsObject(oFormItem.Value) then
			for each oSubItem in oFormItem.Value
				.write oSubItem.Value & " "
			next
		else
		  .write oFormItem.Value
		end if
		.write "<P>"
	next
  
  for each oFormItem in oFileUpEE.Files
		.write oFormItem.Name & ": " & oFileUpEE.Files(oFormItem.Name).ClientFileName & "<P>"
  next

  .flush
  .end
  
  end with

end if  

' --------------------------------------------------------------------------------------
' Declarations for Main
' --------------------------------------------------------------------------------------

Dim Screen_Title
Dim Bar_Title
Dim Top_Navigation
Dim Navigation
Dim Content_Width

Dim Site_ID
Dim Site_Code
Dim Site_Description
Dim Show_View
Dim Record_ID

Dim BTime   ' Start of File Upload
Dim ETime   ' End   of File Upload
Dim Subscription_Early
Dim Bytes   ' Transfer Bytes
Dim Upload_Status
Dim Path_Source
Dim Path_Destination

Dim Stream_Only
Stream_Only = false

Dim error_msg
error_msg = ""

Dim Show_PID, PID_System, PID_Enabled

Show_PID = false

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual= "/include/functions_String.asp"-->
<!--#include virtual= "/include/functions_File.asp"-->
<!--#include virtual= "/include/functions_DB.asp"-->
<!--#include virtual= "/include/functions_Translate.asp"-->
<!--#include virtual= "/sw-administratorNT/Calendar_Show_Values.asp"-->
<!--#include virtual= "/connections/connection_SiteWide.asp"-->
<%

Dim Login_Language, Alt_Language
Login_Language = "eng"
Alt_Language   = "eng"
Call Connect_SiteWide



%>
<!--#include virtual="/sw-administratorNT/CK_Admin_Credentials_Calendar_Admin.asp"-->
<%

' -------------------------------------------------------------------------------------- 
' Declarations for Upload
' --------------------------------------------------------------------------------------

Site_Description = ""
Site_Code        = oFileUpEE.form("Path_Site")
Site_ID          = oFileUpEE.form("Site_ID")
Record_ID        = LCase(oFileUpEE.form("ID"))

SQL = "SELECT PID_Enabled, PID_System FROM Site WHERE ID=" & CInt(Site_ID)
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3  

if not rsSite.EOF then
  PID_Enabled = rsSite("PID_Enabled")
  PID_System  = rsSite("PID_System")

  if CInt(PID_Enabled) = CInt(false) then
    Show_PID = false
  end if
end if

rsSite.close
set rsSite = nothing

' --------------------------------------------------------------------------------------
' Required for PCat Interface
' --------------------------------------------------------------------------------------

Category_ID = oFileUpEE.form("Category_ID")
Call Get_Show_Values
' --------------------------------------------------------------------------------------
' No Action
' --------------------------------------------------------------------------------------

if not isblank(oFileUpEE.form("Nav_Main_Menu")) and isblank(err_message) then

  HomeURL = "default.asp?Site_ID=" & Site_ID
  if LCase(Record_ID) <> "add" then
    HomeURL = HomeURL & "&ID=" & "edit_record&Category_ID=" & oFileUpEE.form("Category_ID")
  end if
  Set oFileUpEE = Nothing
  Call Disconnect_SiteWide
  response.redirect HomeURL   

end if

' --------------------------------------------------------------------------------------
' Delete Record
' --------------------------------------------------------------------------------------

if not isblank(oFileUpEE.form("Nav_Delete")) and isblank(err_message) then
   if Show_PID = true then
         if  PID_System = 0 then  
              SQL = "SELECT Calendar.ID, Calendar.File_Name, Calendar.Archive_Name, Calendar.Thumbnail, Calendar.Include, File_Name_POD, Archive_Name_POD FROM Calendar WHERE Calendar.ID = " & Record_ID & _
              " union SELECT Calendar.ID, Calendar.File_Name, Calendar.Archive_Name, Calendar.Thumbnail, Calendar.Include, File_Name_POD, Archive_Name_POD FROM Calendar WHERE Calendar.clone = " & Record_ID
         elseif PID_System = 1 then  
              SQL = "SELECT Calendar.ID, Calendar.File_Name, Calendar.Archive_Name, Calendar.Thumbnail, Calendar.Include, File_Name_POD, Archive_Name_POD FROM Calendar WHERE Calendar.ID = " & Record_ID
         end if
   else
          SQL = "SELECT Calendar.ID, Calendar.File_Name, Calendar.Archive_Name, Calendar.Thumbnail, Calendar.Include, File_Name_POD, Archive_Name_POD FROM Calendar WHERE Calendar.ID = " & Record_ID
   end if
   
  Set rsDelete = Server.CreateObject("ADODB.Recordset")
  rsDelete.Open SQL, conn, 3, 3
  
  if not rsDelete.EOF then

    Dim DeleteFile(1,5)

    DeleteFile(0,0) = "File_Name"
    DeleteFile(0,1) = "Archive_Name"
    DeleteFile(0,2) = "Thumbnail"
    DeleteFile(0,3) = "Include"
    DeleteFile(0,4) = "File_Name_POD"
    DeleteFile(0,5) = "Archive_Name_POD"
    DeleteFile(1,0) = rsDelete(DeleteFile(0,0))
    DeleteFile(1,1) = rsDelete(DeleteFile(0,1))
    DeleteFile(1,2) = rsDelete(DeleteFile(0,2))
    DeleteFile(1,3) = rsDelete(DeleteFile(0,3))
    DeleteFile(1,4) = rsDelete(DeleteFile(0,4))
    DeleteFile(1,5) = rsDelete(DeleteFile(0,5))
   
    rsDelete.close
    set rsDelete = nothing
    
    'When deleting a record, delete all the assets if their occurances for this category are 1.  If > 1 the
    'asset is used by another asset record and should not be deleted.
    
    for ii = 0 to 5
      if not isblank(DeleteFile(1,ii)) then
          SQL = "SELECT Calendar." & DeleteFile(0,ii) & " FROM Calendar WHERE Calendar.Site_ID=" & Site_ID & " AND Calendar." & DeleteFile(0,ii) & "='" & DeleteFile(1,ii) & "'"
        Set rsCheck = Server.CreateObject("ADODB.Recordset")
        rsCheck.Open SQL, conn, 3, 3
              
        if rsCheck.RecordCount = 1 then                     ' Used by this record only.

          set FileObj = Server.CreateObject("Scripting.FileSystemObject")

          select case ii
            case 0,1,2,3
              MyFilePath = Server.MapPath("/" & Site_Code & "/" & DeleteFile(1,ii))
            case 4,5
              MyFilePath = Server.MapPath("/" & DeleteFile(1,ii))
          end select
          
          if FileObj.FileExists(MyFilePath) then
            on error resume next
            FileObj.DeleteFile MyFilePath, True ' Delete Record, True=Read Only and non-Read Only
            on error goto 0
          end if  
  
          set FileObj = Nothing
        end if
   
        rsCheck.close
        set rsCheck = nothing
   
      end if  
    next          
    
    ' PCAT Interface
    if CInt(Show_PID) = true then
        if     PID_System = 0 then
            %><!--#include virtual="/sw-administratorNT/SW-PCAT_FNET_DELETE.asp"--><%
        elseif PID_System = 1 then
            %><!--#include virtual="/sw-administratorNT/SW-PCAT_FIND_DELETE.asp"--><%
        end if
    end if        

    ' Now Delete the record from the Table
  	 SQL = "DELETE FROM Calendar WHERE Calendar.ID = " & Record_ID
    conn.execute (SQL)
    if Show_PID = true then
         if  PID_System = 0 then  
              SQL = "DELETE FROM Calendar WHERE Calendar.clone = " & Record_ID
              conn.execute (SQL)
         elseif PID_System = 1 then  
              'Place holder.
         end if
    else
        'Place holder.  
    End if   
  end if  

  BackURL = cstr(oFileUpEE.form("HomeURL")) & "?Site_ID=" & Site_ID & "&Category_ID=" & oFileUpEE.form("Category_ID") & "&ID=edit_record"

  Call Disconnect_SiteWide
  response.redirect BackURL  

end if   

' --------------------------------------------------------------------------------------
' Update Record
' --------------------------------------------------------------------------------------

if (not isblank(oFileUpEE.form("Nav_Update"))     or _
    not isblank(oFileUpEE.form("Nav_Clone"))      or _
    not isblank(oFileUpEE.form("Nav_Duplicate"))) and _
    isblank(err_message) then

  if not isblank(oFileUpEE.form("Nav_Clone"))     or _
     not isblank(oFileUpEE.form("Nav_Duplicate")) then
     Record_ID = "add"
  end if
  
  Dim Path_Site
  Dim Path_Site_POD
  Dim Path_Category

  Path_Site     = oFileUpEE.form("Path_Site")
  Path_Site_POD = oFileUpEE.form("Path_Site_POD")
  Path_Category = oFileUpEE.form("Path_Category")
 
  Dim Subgroups
  Subgroups = ""
  if IsObject(oFileUpEE.Form("SubGroups")) then
		for each oSubItem in oFileUpEE.form("SubGroups")
  		Subgroups = Subgroups & oSubItem.Value & ", "
	  next
		if instr(Subgroups,",") > 0 then
   		Subgroups = Mid(Subgroups,1,len(Subgroups)-2)
  	    end if
  else
    Subgroups = oFileUpEE.form("SubGroups")
  end if  

   'Code added/Updated for (Gold silo aggregation)options Entitlements (Fnet AMS) -- by zensar(11/17/08)
   if Site_ID = 82 then
    if(len(Subgroups) > 0) then
        Subgroups = Subgroups &", "& oFileUpEE.Form("EntSubGroups") 
       else
         Subgroups = oFileUpEE.Form("EntSubGroups") 
      end if    
   end if	
  'response.Write Subgroups
  sqlu = " UPDATE Calendar SET "

  ' Site ID
  if not isblank(Site_ID) then
    sqlu = sqlu & "Site_ID=" & Site_ID
  else
    error_msg = error_msg & "<LI>" & Translate("Missing Internal Site ID, Fatal - Report error to Extranet Admin Group(extranetalerts@fluke.com)",Login_Language,conn) & "</LI>"
  end if                 

  if lcase(Record_ID) = "add" then

    if not isblank(oFileUpEE.form("Nav_Clone")) or not isblank(oFileUpEE.form("Nav_Duplicate")) then    ' Reset to Current Admin
      sqlu = sqlu & ",Submitted_By=" & oFileUpEE.form("Admin_ID")
    elseif not isblank(oFileUpEE.form("Submitted_By_New")) then
      sqlu = sqlu & ",Submitted_By=" & oFileUpEE.form("Submitted_By_New")

    elseif not isblank(oFileUpEE.form("Submitted_By")) then                                       ' Original Creator of the Asset
      sqlu = sqlu & ",Submitted_By=" & oFileUpEE.form("Submitted_By")
    else
      sqlu = sqlu & ",Submitted_By=" & oFileUpEE.form("Admin_ID")
    end if

  else
    if not isblank(oFileUpEE.form("Submitted_By_New")) then
      sqlu = sqlu & ",Submitted_By=" & oFileUpEE.form("Submitted_By_New")
    end if
  end if   
      
  ' Approval Routing   
  if lcase(Record_ID) = "add" then

    if CInt(oFileUpEE.form("Admin_Access")) <= 2 or CInt(oFileUpEE.form("Admin_Access")) = 4 then
     
      if not isblank(oFileUpEE.form("Review_By_Group")) then              
        sqlu = sqlu & ",Review_By_Group=" & oFileUpEE.form("Review_By_Group") & ""         
        SQLApprovers = "Select Approvers.* FROM Approvers WHERE ID=" & oFileUpEE.form("Review_By_Group")

        Set rsApprovers = Server.CreateObject("ADODB.Recordset")
        rsApprovers.Open SQLApprovers, conn, 3, 3
               
        if not rsApprovers.EOF then
          if not isblank(rsApprovers("Approver_ID")) and rsApprovers("Approver_ID") > 0 then
            sqlu = sqlu & ",Review_By=" & rsApprovers("Approver_ID")
          end if
        end if  
         
        rsApprovers.close
        set rsApprovers = nothing
         
      else
        error_msg = error_msg & "<LI>" & Translate("Missing Notify Group to Approve this Submission",Login_Language,conn) & "</LI>"
      end if
       
    end if
   
  elseif isnumeric(Record_ID) then

    if CInt(oFileUpEE.form("Admin_Access")) >= 2 then

      if not isblank(oFileUpEE.form("Review_By_Group")) then
        sqlu = sqlu & ",Review_By_Group=" & oFileUpEE.form("Review_By_Group")
         
        SQLApprovers = "Select Approvers.* FROM Approvers WHERE ID=" & oFileUpEE.form("Review_By_Group")

        Set rsApprovers = Server.CreateObject("ADODB.Recordset")
        rsApprovers.Open SQLApprovers, conn, 3, 3
                
        if not rsApprovers.EOF then
          if not isblank(rsApprovers("Approver_ID")) and rsApprovers("Approver_ID") > 0 then         
            sqlu = sqlu & ",Review_By=" & rsApprovers("Approver_ID")
          end if    
        end if  
         
        rsApprovers.close
        set rsApprovers = nothing
         
      end if
     
    else
      sqlu = sqlu & ",Review_By=" & oFileUpEE.form("Admin_ID")        
    end if   

  end if          

  ' Category ID
  if not isblank(oFileUpEE.form("Category_ID")) then     
    sqlu = sqlu & ",Category_ID="& "" & killquote(oFileUpEE.form("Category_ID")) & ""       
  else
    error_msg = error_msg & "<LI>"& Translate("Missing Internal Category, Fatal - Report error to Extranet Admin Group (extranetalerts@fluke.com)",Login_Language,conn) & "</LI>"
  end if                 

  ' Mapped to Category ID (internal)
  if not isblank(oFileUpEE.form("Code")) then     
    sqlu = sqlu & ",Code="& "" & killquote(oFileUpEE.form("Code")) & ""
    Code = CInt(killquote(oFileUpEE.form("Code")))
    Show_Calendar = CInt(killquote(oFileUpEE.form("Show_Calendar")))
  else
    error_msg = error_msg & "<LI>" & Translate("Missing Internal Category (Code), Fatal - Report error to Extranet Admin Group (extranetalerts@fluke.com)",Login_Language,conn) & "</LI>"
  end if                 

  if not isblank(oFileUpEE.form("Status")) then
    if not isblank(oFileUpEE.form("Nav_Clone")) or not isblank(oFileUpEE.form("Nav_Duplicate")) then
      sqlu = sqlu & ",Status=0"
      Status = "0"
     
      sqlu = sqlu & ",Approved_By=0"
    else    
      sqlu = sqlu & ",Status="& "" & killquote(oFileUpEE.form("Status")) & ""
      Status = killquote(oFileUpEE.form("Status"))    ' MAC
      
      if CInt(oFileUpEE.form("Status")) > 0 then
        sqlu = sqlu & ",Approved_By=" & oFileUpEE.form("Admin_ID")
      else
        sqlu = sqlu & ",Approved_By=0"       
      end if
      
      sqlu = sqlu & ",Status_Comment=NULL"
      
    end if
  else
    sqlu = sqlu & ",Status=0"
  end if                 

  if not isblank(oFileUpEE.form("Sub_Category_New")) then 
    sqlu = sqlu & ",Sub_Category="& "'" & ReplaceQuote(oFileUpEE.form("Sub_Category_New")) & "'"
  elseif not isblank(oFileUpEE.form("Sub_Category")) then            
    sqlu = sqlu & ",Sub_Category="& "'" & ReplaceQuote(oFileUpEE.form("Sub_Category")) & "'"     
  else
    sqlu = sqlu & ",Sub_Category=NULL"
  end if
  
  ' Content_Group
  if not isblank(oFileUpEE.form("Content_Group")) then     
    sqlu = sqlu & ",Content_Group=" & oFileUpEE.form("Content_Group")
  end if

  ' Campaign
  if CInt(oFileUpEE.form("Content_Group")) > 0 and not isblank(oFileUpEE.form("Campaign")) then
    sqlu = sqlu & ",Campaign="& "" & killquote(oFileUpEE.form("Campaign")) & ""
  elseif CInt(oFileUpEE.form("Content_Group")) > 0 and isblank(oFileUpEE.form("Campaign")) then
    error_msg = error_msg & "<LI>" & Translate("Missing Production Introduction or Campaign Name",Login_Language,conn) & "</LI>"
  else
    sqlu = sqlu & ",Campaign=0"
  end if

  ' Product or Product Family
  if not isblank(oFileUpEE.form("Product_New")) then 
    sqlu = sqlu & ",Product="& "N'" & ReplaceQuote(oFileUpEE.form("Product_New")) & "'"
  elseif not isblank(oFileUpEE.form("Product")) then            
    sqlu = sqlu & ",Product="& "N'" & ReplaceQuote(oFileUpEE.form("Product")) & "'"     
  else
    error_msg = error_msg & "<LI>" & Translate("Missing Product or Product Family Description",Login_Language,conn) & "</LI>"         
  end if
   
  ''Nitin Code Changes Start
  ' Asset Categories
  if not isblank(oFileUpEE.form("PCat_SPortalCats")) then 
    '' ***** No Code *****
  else
    error_msg = error_msg & "<LI>" & Translate("Missing Asset Categories",Login_Language,conn) & "</LI>"         
  end if   
  ''Nitin Code Changes End
   
  'Title
  if not isblank(oFileUpEE.form("Title_B64")) then
  
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Title_B64")))

    DuplicateTitle     = false
    DuplicateTitleText = ""

    if CInt(Show_PID) = CInt(true) then
      if CInt(PID_System) = 0 then
        if isnumeric(oFileUpEE.form("ID")) then
          StrSql=" select id from calendar where ltrim(rtrim(title)) like '" & ReplaceQuote(myString) & "'" & _
          " and (id <> " & oFileUpEE.form("ID") & " and Category_id = " & oFileUpEE.form("Category_ID") & _
          " and site_id = " & CInt(Site_ID) & ")"
        else
          StrSql=" select id from calendar where ltrim(rtrim(title)) like '" & ReplaceQuote(myString) & "'" & _
          " and Category_id = " & oFileUpEE.form("Category_ID") & " and site_id = " & CInt(Site_ID)
        end if
        Set rsResult = Conn.execute(StrSql) 
        if not(rsResult.eof) then
          if clng(oFileUpEE.form("PCat"))>= 0 then
            DuplicateTitle = true  
            DuplicateTitleText = myString                 
          else
            sqlu = sqlu & ",Title=" & "N'" & ReplaceQuote(myString) & "'"
          end if  	                
        else
          sqlu = sqlu & ",Title=" & "N'" & ReplaceQuote(myString) & "'"
        end if
        rsResult.close
        set rsResult=nothing
      elseif CInt(PID_System) = 1 then
        sqlu = sqlu & ",Title="& "N'" & ReplaceQuote(myString) & "'"
      end if
    else
      sqlu = sqlu & ",Title="& "N'" & ReplaceQuote(myString) & "'"
    end if
  else
    error_msg = error_msg & "<LI>" & Translate("Missing Title",Login_Language,conn) & "</LI>"
  end if
   
  ' Description
  if not isblank(oFileUpEE.form("Description_B64")) then 
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Description_B64")))
    sqlu = sqlu & ",Description="& "N'" & ReplaceQuote(myString) & "'"
  else
    sqlu = sqlu & ",Description=NULL"
  end if

  ' Instructions
  if not isblank(oFileUpEE.form("Instructions_B64")) then
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Instructions_B64")))
    sqlu = sqlu & ",Instructions="& "N'" & ReplaceQuote(myString) & "'"
  else
    sqlu = sqlu & ",Instructions=NULL"
  end if

  ' Splash Header
  if not isblank(oFileUpEE.form("Splash_Header_B64")) then
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Splash_Header_B64")))
    sqlu = sqlu & ",Splash_Header="& "N'" & ReplaceQuote(mystring) & "'"
  elseif isblank(oFileUpEE.form("Code")) then
    sqlu = sqlu & ",Splash_Header=NULL"
  elseif oFileUpEE.form("Code") = 8000 or oFileUpEE.form("Code") = 8001 then
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Title_B64"))) & "</SPAN>"
    if not isblank(oFileUpEE.form("Description_B64")) then
      mystring = mystring & vbCrLf & vbCrLf & utf8Decode(decodeBase64(oFileUpEE.form("Description_B64")))
    end if  
    sqlu = sqlu & ",Splash_Header="& "N'<SPAN CLASS=MEDIUMBOLD>" & ReplaceQuote(mystring) & "'"
  else
    sqlu = sqlu & ",Splash_Header=NULL"
  end if

  ' Splash Footer
  if not isblank(oFileUpEE.form("Splash_Footer")) then
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Splash_Footer_B64")))
    sqlu = sqlu & ",Splash_Footer="& "N'" & ReplaceQuote(mystring) & "'"
  else
    sqlu = sqlu & ",Splash_Footer=NULL"
  end if

  ' Clone - If clone, move Item Number to Item Number 2 since clone is only to copy text fields not file fields.
  if not isblank(oFileUpEE.form("Nav_Clone")) and not isblank(oFileUpEE.form("Item_Number")) and not isblank(oFileUpEE.form("Revision_Code")) then

    sqlu = sqlu & ",Item_Number_2="& "'" & Trim(UCase(ReplaceQuote(oFileUpEE.form("Item_Number")))) & " " & oFileUpEE.form("Revision_Code") & " " & oFileUpEE.form("Content_Language") & "'"
    sqlu = sqlu & ",Item_Number=NULL"
    sqlu = sqlu & ",Item_Number_Show=" & CInt(False)
    
    sqlu = sqlu & ",Revision_Code=NULL"
    
    sqlu = sqlu & ",Cost_Center=0"

  else

    ' Reference Number (Primary)
    if not isblank(oFileUpEE.form("Item_Number")) and isnumeric (oFileUpEE.form("Item_Number")) and Len(oFileUpEE.form("Item_Number")) = 7 then
      sqlu = sqlu & ",Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'"
    elseif not isblank(oFileUpEE.form("Item_Number")) and Len(oFileUpEE.form("Item_Number")) = 9 then    
      sqlu = sqlu & ",Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'"
    elseif not isblank(oFileUpEE.form("Item_Number")) then
      error_msg = error_msg & "<LI>" & Translate("Format Error: Item Reference Number 1 must be numeric and a minimum 7-digits in length for Oracle Item Numbers or 5-digits + "-" + 3-alpha characters for European Literature Numbers.  Use Item Reference Number 2 for Legacy or Other References to this Asset.",Login_Language,conn) & "</LI>"
    elseif instr(1,LCase(Subgroups),"view") > 0 or instr(1,LCase(Subgroups),"shpcrt") > 0 or instr(1,LCase(Subgroups),"fedl") > 0 then
      error_msg = error_msg & "<LI>" & Translate("Missing Item Reference Number 1 for: 'eFulfillment Oracle', 'eFulfillment Digital Library, 'POD' or 'Literature Order Shopping Cart' reference.",Login_Language,conn) & "</LI>"  
    else  
      sqlu = sqlu & ",Item_Number=NULL"
    end if
  
    ' Cost Center
    if not isblank(oFileUpEE.form("Item_Number")) and not isblank(oFileUpEE.form("Cost_Center")) then 
      sqlu = sqlu & ",Cost_Center="& oFileUpEE.form("Cost_Center")
    else  
      sqlu = sqlu & ",Cost_Center=0"
    end if
    
    ' Reference Number Revision Code
    if not isblank(oFileUpEE.form("Revision_Code")) and not isblank(oFileUpEE.form("Item_Number")) then 
      sqlu = sqlu & ",Revision_Code="& "'" & UCase(ReplaceQuote(oFileUpEE.form("Revision_Code"))) & "'"
    else  
      sqlu = sqlu & ",Revision_Code=NULL"
    end if
    
    ' Reference Number Show
    if IsItChecked(oFileUpEE,"Item_Number_Show") = "on" and not isblank(oFileUpEE.form("Item_Number")) then
      sqlu = sqlu & ",Item_Number_Show=" & CInt(True)
    else
      sqlu = sqlu & ",Item_Number_Show=" & CInt(False)  
    end if  
  
    ' Reference Number (Secondary / Legacy)
    if not isblank(oFileUpEE.form("Item_Number_2")) then 
      sqlu = sqlu & ",Item_Number_2="& "'" & Trim(UCase(ReplaceQuote(oFileUpEE.form("Item_Number_2")))) & "'"
    else  
      sqlu = sqlu & ",Item_Number_2=NULL"
    end if

  end if
  
  ' ImageStore Image Locator ID
  if not isblank(oFileUpEE.form("Image_Locator")) then
    sqlu = sqlu & ",Image_Locator="& "'" & Image_Locator_Type & ReplaceQuote(oFileUpEE.form("Image_Locator")) & "'"
    SQLSite = "SELECT * FROM Site WHERE Site_Code='Image-Store'"
    Set rsSite = Server.CreateObject("ADODB.Recordset")
    rsSite.Open SQLSite, conn, 3, 3
    
    if not rsSite.EOF then
      Link_Name    = Replace(rsSite("URL"),"https://support.fluke.com","https://" & request.ServerVariables("SERVER_NAME"))
      Link_Name    = Replace(Link_Name,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))                           

      sqlu = sqlu & ",Link="& "'" & Link_Name & "/Default.asp?Locator=" & ReplaceQuote(oFileUpEE.form("Image_Locator")) & "'"

      sqlu = sqlu & ",Link_PopUp_Disabled=" & CInt(True)
    end if
    
    rsSite.close
    set rsSite = nothing  

  else
    sqlu = sqlu & ",Image_Locator=NULL"
  end if

  ' Forum ID 
  
  if not isblank(oFileUpEE.form("Show_Forum")) then
  
    if CInt(oFileUpEE.form("Show_Forum")) = CInt(True) then
  
      if not isblank(oFileUpEE.form("Forum_ID")) and isnumeric(oFileUpEE.form("Forum_ID")) then
        sqlu = sqlu & ",Forum_ID=" & oFileUpEE.form("Forum_ID")
      else
        error_msg = error_msg & "<LI>" & Translate("Missing Forum ID Number (Suggest: Using Content ID Number)",Login_Language,conn) & "</LI>"    
      end if
      
    end if  

  end if
  
  ' Forum Moderated and Moderator Name
  if IsItChecked(oFileUpEE,"Forum_Moderated") = "on" then
    if not isblank(oFileUpEE.form("Forum_Moderator_ID")) then
      sqlu = sqlu & ",Forum_Moderated=" & CInt(True)
      sqlu = sqlu & ",Forum_Moderator_ID=" & oFileUpEE.form("Forum_Moderator_ID")        
    else
      error_msg = error_msg & "<LI>" & Translate("Missing Forum Moderator Name for Moderated Forum",Login_Language,conn) & "</LI>"
    end if
  else  
    sqlu = sqlu & ",Forum_Moderated=" & CInt(False)
    sqlu = sqlu & ",Forum_Moderator_ID=NULL"
  end if  

  ' Mark as Confidential
  if IsItChecked(oFileUpEE,"Confidential") = "on" then
    sqlu = sqlu & ",Confidential=" & CInt(True)
  else
    sqlu = sqlu & ",Confidential=" & CInt(False)
  end if 

  ' Location (City, State, Country)
  if not isblank(oFileUpEE.form("Location_B64")) then
    mystring = utf8Decode(decodeBase64(oFileUpEE.form("Location_B64")))
    sqlu = sqlu & ",Location="& "N'" & ReplaceQuote(mystring) & "'"
  else
    sqlu = sqlu & ",Location=NULL"
  end if
   
  ' Beginning Date
  if isdate(oFileUpEE.form("BDate")) then   
    GoLiveDate = oFileUpEE.form("BDate")
    sqlu = sqlu & ",BDate="& "'" & killquote(oFileUpEE.form("BDate")) & "'"
    BDate = "'" & killquote(oFileUpEE.form("BDate")) & "'"  ' PIK / Campaign
  else
    GoLivedate = ""
    error_msg = error_msg & "<LI>" & Translate("Missing Beginning Date or Invalid Date Format",Login_Language,conn) & "</LI>"         
  end if

  ' Pre-Announce

  if isnumeric(oFileUpEE.form("LDays")) and not isdate(oFileUpEE.form("LDays")) then
  
    sqlu = sqlu & ",LDays="& killquote(oFileUpEE.form("LDays"))

    LDays = killquote(oFileUpEE.form("LDAYS"))
    
    if isdate(GoLiveDate) then
      GoLiveDate = DateAdd("d", (CInt(oFileUpEE.form("LDays")) * -1),GoLiveDate)
    end if
    
    sqlu = sqlu & ",LDate="& "'" & GoLiveDate & "'"

    LDays = oFileUpEE.form("LDAYS")  ' PIK / Campaign
    if isblank(LDays) then LDays = 0
    LDate = "'" & GoLiveDate & "'"        ' PIK / Campaign
  else
    sqlu = sqlu & ",LDays=0"
    sqlu = sqlu & ",LDate="& "'" & killquote(oFileUpEE.form("BDate")) & "'"

    LDays = "0"                                        ' PIK / Campaign
    LDate = "'" & killquote(oFileUpEE.form("BDate")) & "'"   ' PIK / Campaign   
  end if

  ' Ending Date
  if isdate(oFileUpEE.form("EDate")) then               
    sqlu = sqlu & ",EDate="& "'" & killquote(oFileUpEE.form("EDate")) & "'"
    EDate = "'" & killquote(oFileUpEE.form("EDate")) & "'"
  elseif isdate(oFileUpEE.form("BDate")) then
    sqlu = sqlu & ",EDate="& "'" & killquote(oFileUpEE.form("BDate")) & "'"
    EDate = "'" & killquote(oFileUpEE.form("BDate")) & "'"
  else              
    error_msg = error_msg & "<LI>" & Translate("Invalid Beginning or Ending Date",Login_Language,conn) & "</LI>"         
  end if 

  ' Expiration Date - Move to Archive
  if isnumeric(oFileUpEE.form("XDays")) and not isdate(oFileUpEE.form("XDays")) then                      
    sqlu = sqlu & ",XDays=" & killquote(oFileUpEE.form("XDays"))
    XDays = killquote(oFileUpEE.form("XDays"))
    XDate = killquote(oFileUpEE.form("XDays"))                  
    if isdate(oFileUpEE.form("EDate")) then               
      sqlu = sqlu & ",XDate="& "'" & DateAdd("d", CInt(oFileUpEE.form("XDays")),killquote(oFileUpEE.form("EDate"))) & "'"
      XDate = "'" & DateAdd("d", CInt(oFileUpEE.form("XDays")),killquote(oFileUpEE.form("EDate"))) & "'"      
    elseif isdate(oFileUpEE.form("BDate")) then
      sqlu = sqlu & ",XDate="& "'" & DateAdd("d", CInt(oFileUpEE.form("XDays")),killquote(oFileUpEE.form("BDate"))) & "'"
      XDate = "'" & DateAdd("d", CInt(oFileUpEE.form("XDays")),killquote(oFileUpEE.form("EDate"))) & "'"
    end if           
  end if

  ' Creation / Update Date/Time
  if isdate(oFileUpEE.form("PDate")) then               
    sqlu = sqlu & ",PDate='" & killquote(oFileUpEE.form("PDate")) & "'"
  else
    sqlu = sqlu & ",PDate=NULL"
  end if
   
  ' EEF Release Date
  if isdate(oFileUpEE.form("VDate")) then               
    sqlu = sqlu & ",VDate='" & killquote(oFileUpEE.form("VDate")) & "'"
    VDate = "'" & killquote(oFileUpEE.form("VDate")) & "'"
  else
    sqlu = sqlu & ",VDate=NULL"
    VDate = "NULL"
  end if

  ' Public Embargo Date
  if isdate(oFileUpEE.form("PEDate")) then               
    sqlu = sqlu & ",PEDate='" & killquote(oFileUpEE.form("PEDate")) & "'"
    PEDate = "'" & killquote(oFileUpEE.form("PEDate")) & "'"
  else
    sqlu = sqlu & ",PEDate=NULL"
    PEDate = "NULL"
  end if

  ' Last Update
  if isdate(oFileUpEE.form("UDate")) then               
    sqlu = sqlu & ",UDate="& "'" & killquote(oFileUpEE.form("UDate")) & "'"
  end if

  ''Nitin Code Changes Start
  ' URL Link
  if isblank(oFileUpEE.form("Image_Locator")) then
    if not isblank(oFileUpEE.form("URLLink")) and (instr(1,LCase(oFileUpEE.form("URLLink")),"http://") = 1 or instr(1,LCase(oFileUpEE.form("URLLink")),"https://") = 1 or instr(1,LCase(oFileUpEE.form("URLLink")),"ftp://") = 1 or instr(1,LCase(oFileUpEE.form("URLLink")),"/") >= 1 or instr(1,LCase(oFileUpEE.form("URLLink")),"javascript") = 1) then
      if instr(1,oFileUpEE.form("URLLink")," ") > 0 then
        error_msg = error_msg & "<LI>" & Translate("Format Error: URL Link cannot contain spaces.",Login_Language,conn) & "</LI>"
      else  
        sqlu = sqlu & ",Link="& "'" & ReplaceQuote(oFileUpEE.form("URLLink")) & "'"
      end if
    elseif not isblank(oFileUpEE.form("URLLink")) then
      error_msg = error_msg & "<LI>" & Translate("Format Error: URL Link to other Internet page, must begin with &quot;http://, https://, ftp://, or / (virtual path on this server)&quot;",Login_Language,conn) & "</LI>"
    else
      sqlu = sqlu & ",Link=NULL"  
    end if
  end if  
  ''Nitin Code Changes End
  
  ' Determines whether or not a Pop-Up window is launched when asset is clicked
  if isblank(oFileUpEE.form("Image_Locator")) then
    if IsItChecked(oFileUpEE,"Link_PopUp_Disabled") = "on" then
      sqlu = sqlu & ",Link_PopUp_Disabled=" & CInt(True)
    else  
      sqlu = sqlu & ",Link_PopUp_Disabled=" & CInt(False)
    end if
  end if
    
  ' Site Administrator Lock on Asset to disable change.
  if IsItChecked(oFileUpEE,"Locked") = "on" then
    sqlu = sqlu & ",Locked=" & CInt(True)
  else
    sqlu = sqlu & ",Locked=" & CInt(False)    
  end if
  
  ' Determines what group(s) 1=US, 2=Europe, 3=Intercon can see asset

  if not isblank(SubGroups) or _
         IsItChecked(oFileUpEE,"SubGroups_1") = "on" or _
         IsItChecked(oFileUpEE,"SubGroups_2") = "on" or _
         IsItChecked(oFileUpEE,"SubGroups_3") = "on" then
      
    ' Reset for a clone
    if not isblank(oFileUpEE.form("Nav_Clone")) then
      SubGroups = replace(SubGroups,"view, ","")
      SubGroups = replace(SubGroups,"fedl, ","")
      SubGroups = replace(SubGroups,"shpcrt, ","")
      SubGroups = replace(SubGroups,"nomac, ","")
    end if

    if isblank(oFileUpEE.form("Item_Number")) then
      SubGroups = replace(SubGroups,"view, ","")
      SubGroups = replace(SubGroups,"fedl, ","")
      SubGroups = replace(SubGroups,"shpcrt, ","")
    end if
  
    for ii = 1 to 3
      if IsItChecked(oFileUpEE,"SubGroups_" & Trim(CStr(ii))) = "on" then
        SQLsg = "SELECT SubGroups.*, SubGroups.Order_Num "
        SQLsg = SQLsg & "FROM SubGroups "
        SQLsg = SQLsg & "WHERE SubGroups.Site_ID=" & Site_ID & " AND SubGroups.Enabled=" & CInt(True) & " AND SubGroups.Region=" & Trim(CStr(ii)) & " "
        SQLsg = SQLsg & "ORDER BY SubGroups.Order_Num"
  
        Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
        rsSubGroups.Open SQLsg, conn, 3, 3
        
        Do while not rsSubGroups.EOF
          if instr(1,SubGroups,rsSubGroups("Code")) = 0 then
            if isblank(SubGroups) then
              SubGroups = SubGroups & rsSubGroups("Code")
            else
              SubGroups = SubGroups & ", " & rsSubGroups("Code")
            end if
          end if
          rsSubGroups.MoveNext
        loop
        
        rsSubGroups.Close
        set rsSubGroups = nothing

      end if
    next
    
    if instr(1,lcase(SubGroups),"all") > 0 then             ' Override above if selected
    
      if isblank(oFileUpEE.form("Nav_Clone")) then
        if instr(1,lcase(SubGroups),"view") > 0 and instr(1,lcase(SubGroups),"fedl") = 0 then              ' Check for User Viewable and add to ALL
          sqlu = sqlu & ",SubGroups="& "'all, view'"
        elseif instr(1,lcase(SubGroups),"view") = 0 and instr(1,lcase(SubGroups),"fedl") > 0 then          ' Check for User Viewable and add to ALL
          sqlu = sqlu & ",SubGroups="& "'all, fedl'"
        elseif instr(1,lcase(SubGroups),"view") > 0 and instr(1,lcase(SubGroups),"fedl") > 0 then          ' Check for User Viewable and add to ALL
          sqlu = sqlu & ",SubGroups="& "'all, view, fedl'"
        else  
          sqlu = sqlu & ",SubGroups="& "'all'"
        end if
      else
        sqlu = sqlu & ",SubGroups="& "'all'"
      end if
    else
      sqlu = sqlu & ",SubGroups="& "'" & killquote(SubGroups) & "'"
    end if     
  else
    error_msg = error_msg & "<LI>" & Translate("Missing Group(s) that are allowed to view this information",Login_Language,conn) & "</LI>"
  end if

  ' Asset Language
  
  if not isblank(oFileUpEE.form("Nav_Clone")) then
    if CInt(Show_PID) = CInt(true) then
      if CInt(PID_System) = 0 then
        sqlu = sqlu & ",Language=NULL"
      else  
        if not isblank(oFileUpEE.form("Content_Language")) then 
          sqlu = sqlu & ",Language="& "'" & killquote(oFileUpEE.form("Content_Language")) & "'"         
        else
          sqlu = sqlu & ",Language="& "'eng'"         
        end if
      end if
    else  
      if not isblank(oFileUpEE.form("Content_Language")) then 
        sqlu = sqlu & ",Language="& "'" & killquote(oFileUpEE.form("Content_Language")) & "'"         
      else
        sqlu = sqlu & ",Language="& "'eng'"         
      end if
    end if  
  else  
    if not isblank(oFileUpEE.form("Content_Language")) then 
      sqlu = sqlu & ",Language="& "'" & killquote(oFileUpEE.form("Content_Language")) & "'"         
    else
      sqlu = sqlu & ",Language="& "'eng'"         
    end if
  end if  
  
  ''Added on 29th Oct 2009
    'if not isblank(oFileUpEE.form("Show_Marketing_Automation")) then
        'if oFileUpEE.form("Show_Marketing_Automation") = True then
        if CInt(oFileUpEE.form("Show_Marketing_Automation")) = CInt(True) then
            ''To Insert/Update Business Unit
            if not isblank(oFileUpEE.form("Business_Unit")) then 
              sqlu = sqlu & ",BusinessUnitCode="& "'" & killquote(oFileUpEE.form("Business_Unit")) & "'"         
            else
              sqlu = sqlu & ",BusinessUnitCode="& "''"         
            end if
            
            ''To Insert/Update Ad Pixel URL
            'sqlu = sqlu & ",AdPixel= N'" & oFileUpEE.form("txtAdPixel")  & "'"
            if not isblank(oFileUpEE.form("txtAdPixel_B64")) then 
                'response.Write "test " & utf8Decode(decodeBase64(oFileUpEE.form("txtAdPixel_B64")))
                mystring = HTMLDecode(oFileUpEE.form("txtAdPixel_B64"))
                sqlu = sqlu & ",AdPixel="& "N'" & mystring & "'"
            else
                sqlu = sqlu & ",AdPixel=NULL"
            end if
            
            
            ''To Insert/Update Form Processing URL
            ''Updated on 11th Nov 2009 to add querystring document=XXXX
            if not isblank(oFileUpEE.form("txtFormProcessingURL")) then 
                mystring = ""
                
                ''To add document querystring paramenter
                ''1. if the URL already have the parameter with right Item_Number then do nothing
                ''2. if the URL already have the parameter but not with correct Item_Number, then replace the old Item number with new one
                ''3. if document parameter is not present but "?" is in the url then add "&document=XXXX"
                ''4. if document parameter is not present and also no query string parameter then add "?document=XXXX"
                if InStr(oFileUpEE.form("txtFormProcessingURL"),"document=") > 0 then
                    if InStr(oFileUpEE.form("txtFormProcessingURL"),"document=" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number")))) = 0 then
                        dim strToReplace, arrToReplace, i
                        strToReplace = Replace(oFileUpEE.form("txtFormProcessingURL"),"?","&")
                        arrToReplace = split(strToReplace,"&")
                        
                        for i= 0 to uBound(arrToReplace)
                            if instr(arrToReplace(i),"document") > 0 then
                                mystring = Replace(oFileUpEE.form("txtFormProcessingURL"),arrToReplace(i),"document=" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))))
                                exit for
                            end if
                        Next 
                    else
                        mystring = oFileUpEE.form("txtFormProcessingURL")
                    end if
                else
                    if (inStr(oFileUpEE.form("txtFormProcessingURL"),"?") > 0) then
                        mystring = oFileUpEE.form("txtFormProcessingURL") & "&document=" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number")))
                    else
                        mystring = oFileUpEE.form("txtFormProcessingURL") & "?document=" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number")))
                    end if
                end if
                
                sqlu = sqlu & ",FormProcessingURL= N'" & mystring & "'"
            else
                sqlu = sqlu & ",FormProcessingURL=NULL"
            End if
        end if
    'end if
  ''End

  Subscription_Early = oFileUpEE.form("Subscription_Early")
    
  if isblank(oFileUpEE.form("Nav_Clone")) then
  
    ' Notify with Subscription Service
    if IsItChecked(oFileUpEE,"Subscription") = "on" then
      sqlu = sqlu & ",Subscription=" & CInt(True)
    else  
      sqlu = sqlu & ",Subscription=" & CInt(False)
    end if

    ' Notify with Subscription Service Early (Controls the AM or PM broadcast times of the XNet_Subscription Service Nightly Process
    if Subscription_Early = "-1" then
      sqlu = sqlu & ",Subscription_Early=" & CInt(True)
      Subscription_Early = CInt(True)
    else  
      sqlu = sqlu & ",Subscription_Early=" & CInt(False)
      Subscription_Early = CInt(False)
    end if

 else

    sqlu = sqlu & ",Subscription=" & CInt(False)
    sqlu = sqlu & ",Subscription_Early=" & CInt(False)
 end if

 ' Country Restrict/Exclude/No Exclusions
 ' a "0" (zero) as first element in the array indicates Include/Exclude toggle.

  if not isblank(oFileUpEE.form("Country")) and instr(1,oFileUpEE.form("Country_Reset"),"none") = 0 then
    
    Dim Country
    Country = ""
    
  	if IsObject(oFileUpEE.Form("Country")) then
      for each oSubItem in oFileUpEE.form("Country")
     		Country = Country & oSubItem.Value & ", "
    	next
     	if instr(Country,",") > 0 then
    		Country = Mid(Country,1,len(Country)-2)
    	end if
    else
      Country = oFileUpEE.form("Country") 
    end if  
 
    if instr(1,oFileUpEE.form("Country_Reset"),"0") > 0 then                              ' Exclude these countries
      sqlu = sqlu & ",Country=" & "'0, " & Country & "'"
    else                                                                            ' Include these countries
      sqlu = sqlu & ",Country=" & "'" & Country & "'"
    end if  
  else                                                                              ' No Restrictions 
    sqlu = sqlu & ",Country="& "'none'"
  end if

  ' Records the Parent ID
  if (isnumeric(oFileUpEE.form("Clone")) and oFileUpEE.form("Clone") <> Record_ID and isblank(oFileUpEE.form("Nav_Duplicate"))) then
    sqlu = sqlu & ",Clone="& killquote(oFileUpEE.form("Clone"))
  else     
    sqlu = sqlu & ",Clone=0"
  end if

  ' --------------------------------------------------------------------------------------  
  ' Begin Thumbnail Upload
  ' --------------------------------------------------------------------------------------  
  
  Thumbnail_Request = CInt(True)  
  
  if IsItChecked(oFileUpEE,"Delete_Thumbnail") = "on" and isblank(error_msg) then                                  
  
    sqlu = sqlu & ",Thumbnail=NULL"
    sqlu = sqlu & ",Thumbnail_Size=0"

    Thumbnail_Request = CInt(True)

  elseif not isblank(oFileUpEE.form("Thumbnail_Existing")) then

    sqlu = sqlu & ",Thumbnail='"& killquote(oFileUpEE.form("Thumbnail_Existing")) & "'"

    Thumbnail_Size = oFileUpEE.form("Thumbnail_Size")

    if isblank(Thumbnail_Size) then Thumbnail_Size = 0

    sqlu = sqlu & ",Thumbnail_Size=" & Thumbnail_Size
    
    Thumbnail_Request = CInt(False)
   
  elseif Thumbnail_Flag = true and isblank(error_msg) then
  
    MyPath = oFileUpEE.form("Path_Site")
    MyServerPath = "/" & MyPath & "/Download"
    MyServerPath = Server.MapPath(MyServerPath)
    MySubPath = oFileUpEE.form("Path_Thumbnail")
    MyFile = Mid(oFileUpEE.files("Thumbnail").ClientFilename, InstrRev(oFileUpEE.files("Thumbnail").ClientFilename, "\") + 1) %><%
      
    strFileName = Mid(MyFile, InstrRev(MyFile, "\") + 1)

    ' Filter Invalid Characters in File Name
    strFileName = Replace(strFileName," ","_")  ' Convert spaces to underscores

    if instr(1,strFileName,".") = 0 then
        error_msg = error_msg & "<LI><B>" & Translate("Upload Thumnail Image File Name Error",Login_Language,conn) & Translate(", Missing Filename Extension",Login_Language,conn) & "</B><BR>"      
    end if

    FileRoot = UCase(Mid(strFileName, 1, InstrRev(strFileName, ".") - 1))
    FileExtn = UCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))
    
    Item_Number_Ck = UCase(oFileUpEE.form("Item_Number"))
    Revi_Number_Ck = UCase(oFileUpEE.form("Revision_Code"))

    if isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 7 and not isblank(Revi_Number_Ck) then
      if  (Asc(Revi_Number_CK) >= asc("A") and Asc(Revi_Number_CK) <= asc("Z")) then
        FileRoot = Item_Number_Ck & "_" & Revi_Number_Ck & "_t"
        strFileName = FileRoot & "." & FileExtn
      end if
    elseif isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 7 Then   
        FileRoot = Item_Number_Ck & "_t"
        strFileName = FileRoot & "." & FileExtn
    end if
    
    strFileName = UCase(strFileName)
    
    select case LCase(FileExtn)
      case "jpg", "gif", "jpeg", "bmp"
      case else
        error_msg = error_msg & "<LI><B>" & Translate("Upload Thumbnail Image File Name Error",Login_Language,conn) & "</B><BR>"
        error_msg = error_msg & "<INDENT>" & Translate("File Name",Login_Language,conn) & ": """ & UCase(strFileName) & """ " & Translate("Invalid File name extension. File name extension must be one of the following",Login_Language,conn) & ": ""GIF"", ""JPG""</INDENT></LI>"
    end select
      
    if isblank(error_msg) then
    
      BTime = Now()
      ETime = BTime
      Path_Source      = oFileUpEE.Files("Thumbnail").ClientPath
      Path_Destination = MyServerPath & "\Thumbnail\" & strFileName
      Bytes = oFileUpEE.Files("Thumbnail").size
      Upload_Status = 0

      ' Transfer Settings for Asset File
      oFileUpEE.Files("Thumbnail").SaveAs Path_Destination
      oFileUpEE.Files.Remove("Thumbnail")
      Thumbnail_Flag = false

      ETime = Now()
      Upload_Status = CInt(true)

      sqlu = sqlu & ",Thumbnail='" & Replace((MySubPath & "/" & strFileName),"\","/") & "'"  
        
      sqlu = sqlu & ",Thumbnail_Size=" & Bytes

      ' Post Upload Status to DB
      sqlUpload = "INSERT INTO Calendar_Upload_Status "
      sqlUpload = sqlUpload & "(Site_ID,Account_ID,BTime,ETime,Path_Source,Path_Destination,Bytes,Status) "
      sqlUpload = sqlUpload & "VALUES (" & Site_ID & "," & Admin_ID & "," & "'" & BTime & "'," & "'" & ETime & "'," & "'" & Path_Source & "'," & "'" & Path_Destination & "'," & Bytes & "," & Upload_Status & ")"                    

      on error resume next
      conn.execute (sqlUpload)

      if err.Number <> 0 then
        error_msg = error_msg & "<LI><B>" & Translate("Unable to log file upload status to database.",Login_Language,conn) & " " & Translate("Source",Login_Language,conn) & ": " & Path_Source & "</B><BR>" &_
                                "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & "</INDENT></LI>"
      end if
    
      on error goto 0  
       
      set sqlUpload = nothing
      
      Thumbnail_Request = CInt(False)

    end if
    
  end if

  ' --------------------------------------------------------------------------------------  
  ' Begin File Upload (Low-Resolution Asset)
  ' --------------------------------------------------------------------------------------  
 
  if isblank(oFileUpEE.form("Nav_Clone")) then  

    if IsItChecked(oFileUpEE,"Delete_File") = "on" and isblank(error_msg) then
  
      sqlu = sqlu & ",File_Name=NULL"
      sqlu = sqlu & ",File_Size=0"
  
      sqlu = sqlu & ",Archive_Name=NULL"
      sqlu = sqlu & ",Archive_Size=0"
      
      sqlu = sqlu & ",Secure_Stream=0"

      Thumbnail_Request = False
      
    elseif isnumeric(oFileUpEE.form("File_Existing")) and isblank(error_msg) then
  
      SQLExisting = "SELECT Calendar.File_Name, Calendar.File_Size, Calendar.Archive_Name, Calendar.Archive_Size, Calendar.Thumbnail, Calendar.Thumbnail_Size, Calendar.ID " & vbCrLf &_
                    "FROM Calendar " & vbCrLf &_
                    "WHERE Calendar.ID=" & Clng(oFileUpEE.form("File_Existing"))
                    
      Set rsFile = Server.CreateObject("ADODB.Recordset")
      rsFile.Open SQLExisting, conn, 3, 3
      
      if not rsFile.EOF then
      
        if not isblank(rsFile("File_Name")) then 
          sqlu = sqlu & ",File_Name = '" & rsFile("File_Name") & "'"
        else        
          sqlu = sqlu & ",File_Name = NULL"
        end if   
  
        if not isblank(rsFile("File_Size")) then
          sqlu = sqlu & ",File_Size=" & rsFile("File_Size")
        else
          sqlu = sqlu & ",File_Size=0"
        end if  
          
        if not isblank(rsFile("Archive_Name")) then 
          sqlu = sqlu & ",Archive_Name='" & rsFile("Archive_Name") & "'"
        else        
          sqlu = sqlu & ",Archive_Name=NULL"
        end if   
  
        if not isblank(rsFile("Archive_Size")) then
          sqlu = sqlu & ",Archive_Size=" & rsFile("Archive_Size")
        else
          sqlu = sqlu & ",Archive_Size=0"
        end if  
  
        if isblank(oFileUpEE.form("Thumbnail")) then    ' Thumbnail already established
          if not isblank(rsFile("Thumbnail")) then 
            sqlu = sqlu & ",Thumbnail='" & rsFile("Thumbnail") & "'"
            Thumbnail_Request = CInt(True)        
          else        
            sqlu = sqlu & ",Thumbnail=NULL"
            Thumbnail_Request = CInt(False)        
          end if   
    
          if not isblank(rsFile("Thumbnail_Size")) then
            sqlu = sqlu & ",Thumbnail_Size=" & rsFile("Thumbnail_Size")
          else
            sqlu = sqlu & ",Thumbnail_Size=0"
          end if
        end if  
  
      end if
      
      rsFile.close
      set rsFile = nothing
      
    elseif not isblank(oFileUpEE.form("File_Existing")) and isblank(error_msg) then
       sqlu = sqlu & ",File_Name='" & oFileUpEE.form("File_Existing") & "'"
  
       if not isblank(oFileUpEE.form("File_Size")) then
         sqlu = sqlu & ",File_Size=" & oFileUpEE.form("File_Size")
       else
         sqlu = sqlu & ",File_Size=0"
       end if  
       '****Ri - 514 Commited By Zensar
       'if IsItChecked(oFileUpEE,"Secure_Stream") = "on" then
       '  sqlu = sqlu & ",Secure_Stream=-1"
       'else
         sqlu = sqlu & ",Secure_Stream=0"
       'end if
         
       if not isblank(oFileUpEE.form("Archive_Existing")) then
         sqlu = sqlu & ",Archive_Name='" & oFileUpEE.form("Archive_Existing") & "'"
       else  
         sqlu = sqlu & ",Archive_Name=NULL"
       end if    
       
       if not isblank(oFileUpEE.form("Archive_Size")) then     
         sqlu = sqlu & ",Archive_Size=" & oFileUpEE.form("Archive_Size")
       else  
         sqlu = sqlu & ",Archive_Size=0"
       end if
  
    elseif File_Name_Flag = true and isblank(error_msg) then
    
      MyPath = oFileUpEE.form("Path_Site")
      MyServerPath = "/" & MyPath & "/Download"
      MyServerPath = Server.MapPath(MyServerPath)
      MySubPath = oFileUpEE.form("Path_File")
      MyCategoryPath = oFileUpEE.form("Path_Category")        
    	MyFile = Mid(oFileUpEE.Files("File_Name").ClientFileName, InstrRev(oFileUpEE.Files("File_Name").ClientFileName, "\") + 1) %><%
    
      strFileName = Mid(MyFile, InstrRev(MyFile, "\") + 1) %><%
      
      ' Filter Invalid Characters in File Name
      strFileName = Replace(strFileName," ","_")  ' Convert spaces to underscores

      if instr(1,strFileName,".") = 0 then
          error_msg = error_msg & "<LI><B>" & Translate("Upload File Name Error",Login_Language,conn) & Translate(", Missing Filename Extension",Login_Language,conn) & "</B><BR>"      
      end if

      FileOrig  = strFileName
     	FileOExtn = UCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))
      
      FileRoot  = UCase(Mid(strFileName, 1, InstrRev(strFileName, ".") - 1))
    	FileExtn  = UCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))

      ' Check if valid Item Number and Revision, if so, use this for file name
      
      Item_Number_Ck = UCase(oFileUpEE.form("Item_Number"))
      Revi_Number_Ck = UCase(oFileUpEE.form("Revision_Code"))
      CC_Number_Ck   = oFileUpEE.form("Cost_Center")

	
        if isnumeric(Item_Number_CK) and Len(CStr(Item_Number_CK)) = 7 and not isblank(Revi_Number_Ck) and FileOExtn = "PDF" then
          if  (Asc(Revi_Number_Ck) >= asc("A") and Asc(Revi_Number_Ck) <= asc("Z")) then
            FileRoot = Item_Number_Ck
            if CLng(Item_Number_Ck) < 9000000 then            
              if Len(Trim(CStr(CC_Number_Ck))) = 4 then
                FileRoot = FileRoot & "_" & CC_Number_Ck
              else  
                FileRoot = FileRoot & "_0000"
              end if
            end if  
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "_" & Revi_Number_Ck
            FileRoot = FileRoot & "_w"
            strFileName = FileRoot & "." & FileExtn
          end if
        elseif isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 7 and not isblank(Revi_Number_Ck) and FileOExtn <> "PDF" then
          if  (Asc(Revi_Number_Ck) >= asc("A") and Asc(Revi_Number_Ck) <= asc("Z")) then
            FileRoot = Item_Number_Ck
            if CLng(Item_Number_Ck) < 9000000 then
              if Len(Trim(CStr(CC_Number_Ck))) = 4 then
                FileRoot = FileRoot & "_" & CC_Number_Ck
              else  
                FileRoot = FileRoot & "_0000"
              end if
            end if  
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "_" & Revi_Number_Ck            
            FileRoot = FileRoot & "_x"
            strFileName = FileRoot & "." & FileExtn
          end if
        elseif isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 7 and FileOExtn = "PDF" then   
            FileRoot = Item_Number_Ck
            if CLng(Item_Number_Ck) < 9000000 then            
              if Len(Trim(CStr(CC_Number_Ck))) = 4 then
                FileRoot = FileRoot & "_" & CC_Number_Ck
              else  
                FileRoot = FileRoot & "_0000"
              end if
            end if  
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "__"
            FileRoot = FileRoot & "_w"
            strFileName = FileRoot & "." & FileExtn
        elseif isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 7 and FileOExtn <> "PDF" then   
            FileRoot = Item_Number_Ck
            if CLng(Item_Number_Ck) < 9000000 then            
              if Len(Trim(CStr(CC_Number_Ck))) = 4 then
                FileRoot = FileRoot & "_" & CC_Number_Ck
              else  
                FileRoot = FileRoot & "_0000"
              end if
            end if  
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "__"
            FileRoot = FileRoot & "_x"
            strFileName = FileRoot & "." & FileExtn
        elseif not isnumeric(Item_Number_CK) and Len(CStr(Item_Number_CK)) = 9 and not isblank(Revi_Number_Ck) and FileOExtn = "PDF" then
          if  (Asc(Revi_Number_Ck) >= asc("A") and Asc(Revi_Number_Ck) <= asc("Z")) then
            FileRoot = Item_Number_Ck
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "_" & Revi_Number_Ck            
            FileRoot = FileRoot & "_w"
            strFileName = FileRoot & "." & FileExtn
          end if
        elseif not isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 9 and not isblank(Revi_Number_Ck) and FileOExtn <> "PDF" then
          if  (Asc(Revi_Number_Ck) >= asc("A") and Asc(Revi_Number_Ck) <= asc("Z")) then
            FileRoot = Item_Number_Ck
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "_" & Revi_Number_Ck            
            FileRoot = FileRoot & "_x"
            strFileName = FileRoot & "." & FileExtn
          end if
        elseif not isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 9 and FileOExtn = "PDF" then   
            FileRoot = Item_Number_Ck
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "__"
            FileRoot = FileRoot & "_w"
            strFileName = FileRoot & "." & FileExtn
        elseif not isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 9 and FileOExtn <> "PDF" then   
            FileRoot = Item_Number_Ck
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "__"
            FileRoot = FileRoot & "_x"
            strFileName = FileRoot & "." & FileExtn
        end if  
		'Code added to retain the file name as it is originaly,
		'This is the fix in case of work around to preserve the file name---by Zensar
		 if (isblank(Item_Number_Ck) and isblank(Revi_Number_Ck)) then
			strFileName = strFileName
		 else
	  		strFileName = UCase(strFileName)
	  	end if	
		'strFileName = UCase(strFileName)
	
	
      if script_debug then
        response.write "<BR>MyServerPath = '" & MyServerPath & "'"
        response.write "<BR>MyPath = '" & MyPath & "'"
        response.write "<BR>MySubPath = '" & MySubPath & "'"
        response.write "<BR>MyCategoryPath = '" & MyCategoryPath & "'"          
        response.write "<BR>MyFile = '" & MyFile & "'"
        response.write "<BR>FileName.Ext = '" & StrFileName & "'"
        response.write "<BR>FileRoot = '" & FileRoot & "'"
        response.write "<BR>FileExtn = '" & FileExtn & "'"
        response.write "<BR>"
        response.write "Server Save Path: '" & MyServerPath & "\" & MySubPath & "\" & MyCategoryPath & "\" & strFileName & "'<BR><BR>"       
      end if
      
      SQLAsset = "SELECT Asset_Type.* FROM Asset_Type WHERE Asset_Type.Enabled=" & CInt(True) & " AND Asset_Type.Upload_Authority<=" & CInt(Admin_Access) & " ORDER BY Asset_Type.File_Extension"
      Set rsAsset = Server.CreateObject("ADODB.Recordset")
'      rsAsset.Open SQLAsset, conn, 3, 3
      set rsAsset=conn.execute(SQLAsset)

      Upload_OK = False        

      Do while not rsAsset.EOF
        if LCase(FileExtn) = LCase(rsAsset("File_Extension")) then
          Upload_OK = True
          select case LCase(FileExtn)
            case "iso", "exe", "zip"
            Stream_Only = true
          end select
          exit do
        end if
        rsAsset.MoveNext
      loop

      select case Upload_OK
        case True
        case False
          rsAsset.MoveFirst
          error_msg = error_msg & "<LI><B>" & Translate("Upload File Name Error",Login_Language,conn) & "</B><BR>"
          error_msg = error_msg & "<INDENT>" & Translate("File Name",Login_Language,conn) & ": """ & UCase(strFileName) & """ " & Translate("Invalid File name extension.  Your account administration level only allows you to upload files with the following extension(s)",Login_Language,conn) & ":<BR>"
          do while not rsAsset.EOF
            error_msg = error_msg & "&quot;" & UCase(rsAsset("File_Extension")) & "&quot;, "
      			rsAsset.MoveNext
          loop
          error_msg = Mid(error_msg,1,len(error_msg)-1) & "</INDENT></LI>"
      end select
    
      rsAsset.Close
      set rsAsset = nothing
      SQLAsset    = ""
        
      ' Upload Asset File
      if isblank(error_msg) then

        BTime = Now()
        ETime = BTime
        Path_Source      = oFileUpEE.Files("File_Name").ClientPath
        Path_Destination = MyServerPath & "\" & MyCategoryPath & "\" & strFileName
        Bytes = oFileUpEE.Files("File_Name").size
        Upload_Status = 0
        
        ' Transfer Settings for Asset File Remote File Server
   	    ' oFileUpEE.Files("File_Name").DestinationPath = MyServerPath & "\" & MySubPath & "\" & MyCategoryPath & "\" & strFileName

        ' Transfer Settings for Asset File Local File Server
        oFileUpEE.Files("File_Name").SaveAs Path_Destination
        oFileUpEE.Files.Remove("File_Name")
        File_Name_Flag = false

        ETime = Now()
        Upload_Status = CInt(true)

        ' Only filename needs to be stored in the ProductEngine database. 
			     PcatSaveFilePath = strFileName  
					   
        sqlu = sqlu & ",File_Name='" & Replace((MySubPath & "/" & MyCategoryPath & "/" & strFileName),"\","/") & "'"
        sqlu = sqlu & ",File_Size=" & Bytes
        '****Ri - 514 Commited By Zensar
        'if IsItChecked(oFileUpEE,"Secure_Stream") = "on" or Stream_Only = true then
        '  sqlu = sqlu & ",Secure_Stream=-1"
        'else
          sqlu = sqlu & ",Secure_Stream=0"
        'end if

        ' Post Upload Status to DB
        sqlUpload = "INSERT INTO Calendar_Upload_Status "
        sqlUpload = sqlUpload & "(Site_ID,Account_ID,BTime,ETime,Path_Source,Path_Destination,Bytes,Status) "
        sqlUpload = sqlUpload & "VALUES (" & Site_ID & "," & Admin_ID & "," & "'" & BTime & "'," & "'" & ETime & "'," & "'" & Path_Source & "'," & "'" & Path_Destination & "'," & Bytes & "," & Upload_Status & ")"                    

        on error resume next
        conn.execute (sqlUpload)

        if err.Number <> 0 then
          error_msg = error_msg & "<LI><B>" & Translate("Unable to log file upload status to database.",Login_Language,conn) & " " & Translate("Source",Login_Language,conn) & ": " & Path_Source & "</B><BR>" &_
                                  "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & "</INDENT></LI>"
        end if
        
        on error goto 0  

        set sqlUpload = nothing
 
      end if

      ' --------------------------------------------------------------------------------------          
      ' Create Archive ZIP File - This is done on the FileServer via SW-FileUp_Upload
      ' --------------------------------------------------------------------------------------  
        
      if isblank(error_msg) then

        if UCase(FileExtn) <> "ZIP" and UCase(FileExtn) <> "EXE" then

          BTime = Now()
          ETime = BTime
          Path_Destination = MyServerPath & "\" & MyCategoryPath & "\" & strFileName
          Upload_Status    = 0

          ArchivePath      = "Download/Archive/" & UCase(FileRoot) & "." & "ZIP"
          ArchiveName      = MyServerPath & "\" & "Archive"   & "\" & UCase(FileRoot) & "." & "ZIP"
          ArchiveFile      = Path_Destination
          Arch.archivetype = 1                	'--- Set ArchiveType to ZIP=1, CAB=2 Format          

        	on error resume next
          Arch.CreateArchive ArchiveName        '--- Create Archive File
          if Err.Number <> 0 then
            error_msg = error_msg & "<LI><B>" & Translate("Archive File Creation Error",Login_Language,conn) & ": " & ArchiveName & "</B><BR>" &_ %><%
                                    "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & "</INDENT></LI>"
          end if

    			Arch.Addfile ArchiveFile, false       '--- True=Recursive Directory, False=None
        	Arch.CloseArchive
          
          if Err.Number = 0 then
            sqlf = sqlf & ",Archive_Name"              
            sql  =  sql & ",'" & ArchivePath & "'"
            sqlu = sqlu & ",Archive_Name='" & ArchivePath & "'"
  
           	Set fso = CreateObject("Scripting.FileSystemObject")
          	If fso.FileExists(ArchiveName) Then %><%               
           		Set f = fso.GetFile(ArchiveName) %><%
  
              sqlf = sqlf & ",Archive_Size"              
              sql  =  sql & "," & f.size
              sqlu = sqlu & ",Archive_Size=" & f.size
              Set f = nothing
            end if
  
            Set fso = nothing
          end if
        
          on error goto 0

        end if
                  
      end if

    elseif not isblank(error_msg) then 
        File_Name_Flag     = false
        File_Name_POD_Flag = false
        Thumbnail_Flag     = false    
    end if
  else
       'Added by zensar on 09-07-2006 for preserve file changes.
       if IsItChecked(oFileUpEE,"Preserve_Path") = "on" and isblank(error_msg) then
           if not isblank(oFileUpEE.form("File_Existing")) then
               sqlu = sqlu & ",File_Name='" & oFileUpEE.form("File_Existing") & "'"
           else
               sqlu = sqlu & ",File_Name=NULL"
           end if
           if not isblank(oFileUpEE.form("File_Size")) then
             sqlu = sqlu & ",File_Size=" & oFileUpEE.form("File_Size")
           else
             sqlu = sqlu & ",File_Size=0"
           end if  
           
           if not isblank(oFileUpEE.form("File_Page_Cnt")) then     
             sqlu = sqlu & ",File_Page_Count=" & oFileUpEE.form("File_Page_Cnt")
           else  
             sqlu = sqlu & ",File_Page_Count=0"
           end if  
             
           if not isblank(oFileUpEE.form("Archive_Existing")) then
             sqlu = sqlu & ",Archive_Name='" & oFileUpEE.form("Archive_Existing") & "'"
           else  
             sqlu = sqlu & ",Archive_Name=NULL"
           end if    
           
           if not isblank(oFileUpEE.form("Archive_Size")) then     
             sqlu = sqlu & ",Archive_Size=" & oFileUpEE.form("Archive_Size")
           else  
             sqlu = sqlu & ",Archive_Size=0"
           end if  
       end if    
  end if

  ' End File Upload & File Archive

  ' --------------------------------------------------------------------------------------  
  ' Begin POD File Upload
  ' --------------------------------------------------------------------------------------  
  
  if isblank(oFileUpEE.form("Nav_Clone")) and isblank(oFileUpEE.form("Nav_Duplicate")) then

    if IsItChecked(oFileUpEE,"Delete_File_POD") = "on" and isblank(error_msg) then

      sqlu = sqlu & ",File_Name_POD=NULL"
      sqlu = sqlu & ",File_Size_POD=0"
  
      sqlu = sqlu & ",Archive_Name_POD=NULL"
      sqlu = sqlu & ",Archive_Size_POD=0"
      
    elseif isnumeric(oFileUpEE.form("File_Existing_POD")) and isblank(error_msg) then

      SQLExisting = "SELECT Calendar.File_Name_POD, Calendar.File_Size_POD, Calendar.Archive_Name_POD, Calendar.Archive_Size_POD, Calendar.ID " & vbCrLf &_
                    "FROM Calendar " & vbCrLf &_
                    "WHERE Calendar.ID=" & Clng(oFileUpEE.form("File_Existing_POD"))
                    
      Set rsFile = Server.CreateObject("ADODB.Recordset")
      rsFile.Open SQLExisting, conn, 3, 3
      
      if not rsFile.EOF then
      
        if not isblank(rsFile("File_Name_POD")) then 
          sqlu = sqlu & ",File_Name_POD='" & rsFile("File_Name_POD") & "'"
        else        
          sqlu = sqlu & ",File_Name_POD=NULL"
        end if   
  
        if not isblank(rsFile("File_Size_POD")) then
          sqlu = sqlu & ",File_Size_POD=" & rsFile("File_Size_POD")
        else
          sqlu = sqlu & ",File_Size_POD=0"
        end if  
          
        if not isblank(rsFile("Archive_Name_POD")) then 
          sqlu = sqlu & ",Archive_Name_POD='" & rsFile("Archive_Name_POD") & "'"
        else        
          sqlu = sqlu & ",Archive_Name_POD=NULL"
        end if   
  
        if not isblank(rsFile("Archive_Size_POD")) then
          sqlu = sqlu & ",1Archive_Size_POD=" & rsFile("Archive_Size_POD")
        else
          sqlu = sqlu & ",1Archive_Size_POD=0"
        end if  
  
      end if
      
      rsFile.close
      set rsFile = nothing
      
    elseif not isblank(oFileUpEE.form("File_Existing_POD")) and isblank(error_msg) then

     sqlu = sqlu & ",File_Name_POD='" & lcase(oFileUpEE.form("File_Existing_POD")) & "'"

     if not isblank(oFileUpEE.form("File_Size_POD")) then
       sqlu = sqlu & ",File_Size_POD=" & oFileUpEE.form("File_Size_POD")
     else
       sqlu = sqlu & ",File_Size_POD=0"
     end if  
       
     if not isblank(oFileUpEE.form("Archive_Existing_POD")) then
       sqlu = sqlu & ",Archive_Name_POD='" & oFileUpEE.form("Archive_Existing_POD") & "'"
     else  
       sqlu = sqlu & ",Archive_Name_POD=NULL"
     end if    
     
     if not isblank(oFileUpEE.form("Archive_Size_POD")) then     
       sqlu = sqlu & ",Archive_Size_POD=" & oFileUpEE.form("Archive_Size_POD")
     else  
       sqlu = sqlu & ",Archive_Size_POD=0"
     end if  
  
    elseif File_Name_POD_Flag = true and isblank(error_msg) then
    
      MyPath       = oFileUpEE.form("Path_Site_POD")
      MyServerPath = "/Pod"
      MyServerPath = Server.MapPath(MyServerPath)
      MySubPath    = oFileUpEE.form("Path_File_POD")
    	 MyFile       = Mid(oFileUpEE.Files("File_Name_POD").ClientFileName, InstrRev(oFileUpEE.Files("File_Name_POD").ClientFileName, "\") + 1) %><%
    
      strFileName  = Mid(MyFile, InstrRev(MyFile, "\") + 1) %><%
  
      ' Filter Invalid Characters in File Name
      strFileName  = Replace(strFileName," ","_")  ' Convert spaces to underscores
  
      if instr(1,strFileName,".") = 0 then
          error_msg = error_msg & "<LI><B>" & Translate("Upload POD File Name Error",Login_Language,conn) & Translate(", Missing Filename Extension",Login_Language,conn) & "</B><BR>"      
      end if
  
      FileOrig  = strFileName
     	FileOExtn = UCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))
  
      FileRoot  = UCase(Mid(strFileName, 1, InstrRev(strFileName, ".") - 1))
    	FileExtn  = UCase(Mid(strFileName, InstrRev(strFileName, ".") + 1))
      
      ' Check if valid Item Number and Revision, if so, use this for file name
      Item_Number_Ck = UCase(oFileUpEE.form("Item_Number"))
      Revi_Number_Ck = UCase(oFileUpEE.form("Revision_Code"))
      CC_Number_Ck   = oFileUpEE.form("Cost_Center")
      
        if isnumeric(Item_Number_CK) and Len(CStr(Item_Number_CK)) = 7 and not isblank(Revi_Number_Ck) and FileOExtn = "PDF" then
          if  (Asc(Revi_Number_Ck) >= asc("A") and Asc(Revi_Number_Ck) <= asc("Z")) then
            FileRoot = Item_Number_Ck
            if Len(Trim(CStr(CC_Number_Ck))) = 4 then
              FileRoot = FileRoot & "_" & CC_Number_Ck
            else  
              FileRoot = FileRoot & "_0000"
            end if
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "_" & Revi_Number_Ck            
            FileRoot = FileRoot & "_p"
            strFileName = FileRoot & "." & FileExtn
          end if
        elseif isnumeric(Item_Number_Ck) and Len(CStr(Item_Number_Ck)) = 7 and FileOExtn = "PDF" then   
            FileRoot = Item_Number_Ck
            if Len(Trim(CStr(CC_Number_Ck))) = 4 then
              FileRoot = FileRoot & "_" & CC_Number_Ck
            else  
              FileRoot = FileRoot & "_0000"
            end if
            FileRoot = FileRoot & "_" & oFileUpEE.form("Content_Language")
            FileRoot = FileRoot & "__"
            FileRoot = FileRoot & "_p"
            strFileName = FileRoot & "." & FileExtn
        else
            FileRoot = FileRoot & "_p"
            strFileName = FileRoot & "." & FileExtn
        end if
      
      strFileName = UCase(strFileName)
  
      if script_debug then
        response.write "<BR>MyServerPath = '" & MyServerPath & "'"
        response.write "<BR>MyPath = '" & MyPath & "'"
        response.write "<BR>MySubPath = '" & MySubPath & "'"
        response.write "<BR>MyCategoryPath = '" & MyCategoryPath & "'"          
        response.write "<BR>MyFile = '" & MyFile & "'"
        response.write "<BR>FileName.Ext = '" & StrFileName & "'"
        response.write "<BR>FileRoot = '" & FileRoot & "'"
        response.write "<BR>FileExtn = '" & FileExtn & "'"
        response.write "<BR>"
        response.write "Server Save Path: '" & MyServerPath & "\" & MySubPath & "\" & strFileName & "'<BR><BR>"       
      end if
      
      Upload_OK = False
  
      if LCase(FileExtn) = "pdf" then
        Upload_OK = True
      end if
  
      select case Upload_OK
        case True
        case False
          error_msg = error_msg & "<LI><B>" & Translate("Upload POD File Name Error",Login_Language,conn) & "</B><BR>"
          error_msg = error_msg & "<INDENT>" & Translate("POD File Name",Login_Language,conn) & ": """ & UCase(strFileName) & """ " & Translate("Invalid File name extension.  Your account administration level only allows you to upload files with the following extension(s)",Login_Language,conn) & ": .PDF<BR>"
          error_msg = Mid(error_msg,1,len(error_msg)-1) & "</INDENT></LI>"
      end select
    
      ' Upload File
      if isblank(error_msg) then
      
        BTime = Now()
        ETime = BTime
        Path_Source      = oFileUpEE.Files("File_Name_POD").ClientPath
        Path_Destination = MyServerPath & "\" & strFileName
        Bytes = oFileUpEE.Files("File_Name_POD").size
        Upload_Status = 0
  
        ' Transfer Settings for Asset File Remote File Server
   	    ' oFileUpEE.Files("File_Name_POD").DestinationPath = MyServerPath & "\" & MySubPath & "\" & MyCategoryPath & "\" & strFileName

        ' Transfer Settings for Asset File Local File Server
        oFileUpEE.Files("File_Name_POD").SaveAs Path_Destination
        oFileUpEE.Files.Remove("File_Name_POD")
        File_Name_POD_Flag = false

        ETime = Now()
        Upload_Status = CInt(true)
        
        sqlu = sqlu & ",File_Name_POD='Pod/" & strFileName & "'"
        sqlu = sqlu & ",File_Size_POD=" & Bytes
        
        ' Post Upload Status to DB
        sqlUpload = "INSERT INTO Calendar_Upload_Status "
        sqlUpload = sqlUpload & "(Site_ID,Account_ID,BTime,ETime,Path_Source,Path_Destination,Bytes,Status) "
        sqlUpload = sqlUpload & "VALUES (" & Site_ID & "," & Admin_ID & "," & "'" & BTime & "'," & "'" & ETime & "'," & "'" & Path_Source & "'," & "'" & Path_Destination & "'," & Bytes & "," & Upload_Status & ")"                    
  
        on error resume next
        conn.execute (sqlUpload)
  
        if err.Number <> 0 then
          error_msg = error_msg & "<LI><B>" & Translate("Unable to log file upload status to database.",Login_Language,conn) & " " & Translate("Source",Login_Language,conn) & ": " & Path_Source & "</B><BR>" &_
                                  "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & "</INDENT></LI>"
        end if
        
        on error goto 0  
           
        set sqlUpload = nothing
        
      end if
    end if
  end if
  
  ' --------------------------------------------------------------------------------------  
  ' End POD File Upload
  ' --------------------------------------------------------------------------------------  
  
  if IsItChecked(oFileUpEE,"Thumbnail_Request") = "on" and Thumbnail_Request = CInt(True) then
    sqlu = sqlu & ",Thumbnail_Request=" & CInt(True)
  elseif  Thumbnail_Request = CInt(True) then
    sqlu = sqlu & ",Thumbnail_Request=" & CInt(True)
  else
    sqlu = sqlu & ",Thumbnail_Request=" & CInt(False)
  end if
  
  ' End Thumbnail Upload
  
  ' --------------------------------------------------------------------------------------  
  ' Send the Files from WebServer to the FileServer
  ' --------------------------------------------------------------------------------------  
  
  ' Do the Transfer of Files from WebServer to FileServer
  
	if File_Name_Flag = true or File_Name_POD_Flag = true or Thumbnail_Flag = true then
  
    on error resume next
  		intSAResult = oFileUpEE.SendRequest()
  		if Err.Number <> 0 then
  			error_msg = error_msg & "<SPAN CLASS=SmallRed>WebServer SendRequest Error</SPAN>:&nbsp;&nbsp;" & Err.Description & " (" & Err.Source & ")" & "<BR>"
  			error_msg = error_msg & "<SPAN CLASS=SmallRed>FileServer Returned:&nbsp;&nbsp;" & oFileUpEE.HttpResponse.BodyText
  		end If
  	on error goto 0
   
  	for each oFile in oFileUpEE.Files
  		if oFile.Size <> 0 Then
  			select case oFile.Processed
  				case saSaved            ' OK Display Nothing
  				case saExists
            error_msg = error_msg & "<SPAN CLASS=SmallRed>The file, " & oFile.Name & ", was not saved because OverwriteFile " & _
										                "is set to False and a file with the same name exists " & _
 									                  "in the destination directory of the FileServer.<BR>"
  				case saError
					  error_msg = error_msg & "<SPAN CLASS=SmallRed>An error has occurred&nbsp;" & oFile.Error & "<BR>"
  			end select
		  end If
  	next

  end if
  
  if File_Name_Flag = true then

  ' Since the Archive File was created on the FileServer, the Calendar_Temp table, contains the Archive File Size
  
    Calendar_Temp_ID = 0
    SQLArchive = "SELECT ID, Name, Value1, Value2 FROM Calendar_Temp WHERE Name='Archive_File' AND Value1='" & Archive_Name & "'"
    Set rsArchive = Server.CreateObject("ADODB.Recordset")
    rsArchive.Open SQLArchive, conn, 3, 3
    
    if not rsArchive.EOF then
      Calendar_Temp_ID = rsArchive("ID")
      sqlu = sqlu & ",Archive_Size=" & rsArchive("Value2")
    end if
    
    rsArchive.close
    set rsArchive  = nothing
    set SQLArchive = nothing
    
    if Calendar_Temp_ID <> 0 then
      'conn.execute "DELETE FROM Calendar_Temp WHERE ID=" & Calendar_Temp_ID
    end if
    
  end if
  
  ' --------------------------------------------------------------------------------------    
  ' Check for Error Message before action
  ' --------------------------------------------------------------------------------------           

  if isblank(error_msg) then
  
    ' Insert or Update Record      
    if lcase(Record_ID) = "add" then
      '>>>>>>>>>>>>>>>>>>>>>>>Modified by zensar to avoid Duplicate item number 12-09-2006>>>>>>>>>
      'ReadCommited Transaction.
      conn.IsolationLevel = 4096
      conn.BeginTrans
      Record_ID = Get_New_Record_ID("Calendar", "Content_Group", 0, conn)
      'Dummy statement below which locks the calendar table.Does not update any row in actual.
      conn.execute "UPDATE CALENDAR with (TABLOCK) SET Title = Title where 1 = 2  " 
      if CInt(Show_PID) = CInt(true) then
        if CInt(PID_System) = 0 then
          if DuplicateTitle = true then      
'           sqlu = sqlu & ",Title=" & "N'" & mid("[" & Record_ID & "] " & ReplaceQuote(DuplicateTitleText) & "'",1,128)
            sqlu = sqlu & ",Title=" & "N'" & mid(ReplaceQuote(DuplicateTitleText) & "'",1,128)            
          end if
        elseif CInt(PID_System) = 1 then
           'Placeholder
        end if
      else
          'Placeholder
      end if
      
      sqlu = sqlu & " WHERE ID=" & Record_ID
      
      sqlu = sqlu & " OPTION (ROBUST PLAN)"
      
''      response.write Replace(sqlu,",",",<BR>")
''      response.end

      'Check if it is Generic item number
      newItemNumber = Trim(ReplaceQuote(oFileUpEE.form("Item_Number")))
      if not isblank(oFileUpEE.form("Item_Number")) then
              set rsGIN = server.CreateObject("ADODB.Recordset")
              'Checks if item_number is present in Calendar table.
              sqlItemNumberCheck = "select id from calendar where item_number='" & _
              Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "' and site_id = " & cint(Site_ID)
              rsGIN.Open sqlItemNumberCheck,conn,adOpenStatic,adLockReadOnly
              if not(rsGIN.EOF) then
                   'If present generates new item number
                   if clng(oFileUpEE.form("Item_Number")) >= 9000000 then
                       newItemNumber = GetNextGenericNumber(cint(Site_ID))
                       sqlu = replace(sqlu,"Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'","Item_Number="& "'" & Trim(ReplaceQuote(newItemNumber)) & "'")
                   else
                       if CInt(Show_PID) = CInt(true) then
                            if CInt(PID_System) = 0 then
                               'error_msg = error_msg & "<LI><B>" & Translate("Duplicate oracle item number found.",Login_Language,conn) & " " & Translate("Record",Login_Language,conn) & ": " & Record_ID & "</B>"
                               if isblank(oFileUpEE.form("Nav_Duplicate")) then
                                   newItemNumber = GetNextGenericNumber(cint(Site_ID))
                                   sqlu = replace(sqlu,"Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'","Item_Number="& "'" & Trim(ReplaceQuote(newItemNumber)) & "'")
                               else
                                   sqlu = replace(sqlu,"Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'","Item_Number=NULL")
                               end if                                    
                            elseif CInt(PID_System) = 1 then
                               
                            end if
                       else
                            
                       end if
                   end if
                   'Replace old item number with new item number
              end if
              set rsGIN = nothing
      end if
       
      on error resume next
      if trim(error_msg)="" then
        conn.Execute(sqlu)  
      end if  
      
      if err.Number <> 0 then
      'Roll back transaction if error is present
        conn.RollbackTrans 
        error_msg = error_msg & "<LI><B>" & Translate("Unable to add record to database.",Login_Language,conn) & " " & Translate("Record",Login_Language,conn) & ": " & Record_ID & "</B><BR>" &_
                                "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & "</INDENT></LI><P>"

        if Admin_Access = 9 then
          error_msg = error_msg & "<P>" & sqlu & "<P>"
        end if
      else
          'Commit the transaction if no error is encountered.
          if trim(error_msg)="" then
            conn.CommitTrans 
          else
            conn.RollbackTrans 
          end if  
      end if
      
      
''Nitin Code Changes Start
AssetCategorySQL = "DELETE FROM dbo.Asset_Category WHERE (AssetId = " & Record_ID & ")"
conn.execute (AssetCategorySQL)

for each oFormItem in oFileUpEE.Form
	if oFormItem.Name ="PCat_SPortalCats" then
					if IsObject(oFormItem.Value) then
						for each oSubItem in oFormItem.Value
    					  AssetCategorySQL = "INSERT INTO [Asset_Category] ([AssetId] ,[CategoryId] ,[CreateDate] ,[CreatedBy]) VALUES (" & Record_ID & ", " & oSubItem.Value & ", getdate(), '')"
                          conn.execute (AssetCategorySQL)
						next
					else
						AssetCategorySQL = "INSERT INTO [Asset_Category] ([AssetId] ,[CategoryId] ,[CreateDate] ,[CreatedBy]) VALUES (" & Record_ID & ", " & oFormItem.Value & ", getdate(), '')"
                        conn.execute (AssetCategorySQL)
					end if
	end if
next
''Nitin Code Changes End            
      
      
      '>>>>>>>>>>>>>>>>>>>>>>>Modified by zensar to avoid Duplicate item number 12-09-2006>>>>>>>>> 
      on error goto 0
            
    else
      if CInt(Show_PID) = CInt(true) then
        if CInt(PID_System) = 0 then
          if DuplicateTitle = true then      
            sqlu = sqlu & ",Title=" & "N'" & mid(ReplaceQuote(DuplicateTitleText) & "'",1,128)            
          end if
        elseif CInt(PID_System) = 1 then
         'Placeholder
        end if
      else
        'Placeholder
      end if
      
      sqlu = sqlu & " WHERE ID=" & Record_ID
      sqlu = sqlu & " OPTION (ROBUST PLAN)"

						    ' Release Task  :   668
                            ' Updated by    :   Amol Jagtap
                            ' Description   :   To convert oFileUpEE.form("Clone") to Long from Int
						if not (isnumeric(oFileUpEE.form("Clone")) and CLng(trim(oFileUpEE.form("Clone"))) <> 0) then
  						PcatSaveFilePath = strFileName 
  	   end if
    
      on error resume next
      '>>>>>>>>>>>>>>>>>>>>>>>Modified by zensar to avoid Duplicate item number 12-09-2006>>>>>>>>>
      conn.IsolationLevel = 4096
      conn.BeginTrans
      'The below statement immediately serializes the other transactions if present for concurrent users.
      'Dummy update
      newItemNumber = Trim(ReplaceQuote(oFileUpEE.form("Item_Number")))
      conn.execute "UPDATE CALENDAR with (TABLOCK) SET Title = Title where ID = " & Record_ID
          
      if not isblank(oFileUpEE.form("Item_Number")) then
          set rsGIN = server.CreateObject("ADODB.Recordset")
          'Checks if item_number is present in Calendar table.
          sqlItemNumberCheck = "select id from calendar where item_number='" & _
                               Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "' and ID <> " & Record_ID & _
                               " and site_id= " & cint(Site_ID)
          rsGIN.Open sqlItemNumberCheck,conn,adOpenStatic,adLockReadOnly
          if not(rsGIN.EOF) then
               if clng(oFileUpEE.form("Item_Number")) >= 9000000 then
                   'If present generates new item number
                   newItemNumber = GetNextGenericNumber(cint(Site_ID))
                   'Replace old item number with new item number
                   sqlu = replace(sqlu,"Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'","Item_Number="& "'" & Trim(ReplaceQuote(newItemNumber)) & "'")
               else
                   if CInt(Show_PID) = CInt(true) then
																							if CInt(PID_System) = 0 then
																										'error_msg = error_msg & "<LI><B>" & Translate("Duplicate oracle item number found.",Login_Language,conn) & " " & Translate("Record",Login_Language,conn) & ": " & Record_ID & "</B>"
																										newItemNumber = GetNextGenericNumber(cint(Site_ID))
																										sqlu = replace(sqlu,"Item_Number="& "'" & Trim(ReplaceQuote(oFileUpEE.form("Item_Number"))) & "'","Item_Number="& "'" & Trim(ReplaceQuote(newItemNumber)) & "'")
																							elseif CInt(PID_System) = 1 then

																							end if
                  else

                  end if
               end if               
          end if
          set rsGIN = nothing
      end if
      
      if trim(error_msg)="" then
        conn.Execute(sqlu)  
      end if  
      
      if err.Number <> 0 then
        conn.RollBackTrans
        error_msg = error_msg & "<LI><B>" & Translate("Unable to update record in database.",Login_Language,conn) & " " & Translate("Record",Login_Language,conn) & ": " & Record_ID & "</B><BR>" &_
                                "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & Translate("Line Number",Login_Language,conn) & err.line & "</INDENT></LI>"

        if Admin_Access = 9 then
          error_msg = error_msg & "<P>" & sqlu & "<P>"
        end if
      else
        if trim(error_msg)="" then
            conn.CommitTrans
        else
            conn.RollBackTrans
        end if
      end if
      
      
''Nitin Code Changes Start
AssetCategorySQL = "DELETE FROM dbo.Asset_Category WHERE (AssetId = " & Record_ID & ")"
conn.execute (AssetCategorySQL)

for each oFormItem in oFileUpEE.Form
	if oFormItem.Name ="PCat_SPortalCats" then
					if IsObject(oFormItem.Value) then
						for each oSubItem in oFormItem.Value
    					  AssetCategorySQL = "INSERT INTO [Asset_Category] ([AssetId] ,[CategoryId] ,[CreateDate] ,[CreatedBy]) VALUES (" & Record_ID & ", " & oSubItem.Value & ", getdate(), '')"
                          conn.execute (AssetCategorySQL)
						next
					else
						AssetCategorySQL = "INSERT INTO [Asset_Category] ([AssetId] ,[CategoryId] ,[CreateDate] ,[CreatedBy]) VALUES (" & Record_ID & ", " & oFormItem.Value & ", getdate(), '')"
                        conn.execute (AssetCategorySQL)
					end if
	end if
next
''Nitin Code Changes End            
      
      
      Conn.execute "update calendar set status=" & killquote(oFileUpEE.form("Status")) & _
      " where clone = " & Record_ID & " and file_name ='" & trim(oFileUpEE.form("File_Existing")) & "'" & _
      " and site_id = " & cint(Site_ID)
      
      if IsItChecked(oFileUpEE,"Delete_File") = "on" and isblank(error_msg) then
            Conn.execute "update calendar set status=0 where id= " & Record_ID & " and site_id = " & cint(Site_ID)
            cloneSql = "select id,file_name from calendar where clone = " &  Record_ID & " and site_id= " & cint(Site_ID)    
            set rsClones=Conn.execute(cloneSql)
            Do while not(rsClones.eof)
                   if trim(rsClones.fields("file_name").value) = trim(oFileUpEE.form("File_Existing")) then
                       updateSql=" update calendar set " 
                       updateSql = updateSql & "File_Name=NULL,File_Size=NULL,Archive_Name=NULL,Archive_Size=NULL,File_Page_Count=NULL,Revision_Code=NULL,Status=0"
                       ' Last Update
                       if isdate(oFileUpEE.form("UDate")) then               
                        updateSql = updateSql & ",UDate="& "'" & killquote(oFileUpEE.form("UDate")) & "'"
                       end if
                       updateSql =updateSql & " where id= " & rsClones.fields("id").value
                       Conn.execute updateSql
                   end if    
                   rsClones.movenext
            loop
            set rsClones = nothing
      end if
      
      'Added by zensar on 09-07-2006 for preserve file changes.
      if IsItChecked(oFileUpEE,"Preserve_Path") = "on" and isblank(error_msg) then
            cloneSql = "select id,file_name from calendar where clone = " &  Record_ID & " and site_id= " & cint(Site_ID)    
            set rsClones=Conn.execute(cloneSql)
            Do while not(rsClones.eof)
                   if (isnull(rsClones.fields("file_name").value)= true or trim(rsClones.fields("file_name").value & "") ="") then
                           set rsGetInfo= Conn.execute("Select File_Name,File_Size,Archive_Name,Archive_Size,File_Page_Count from calendar where id=" & Record_ID & " and site_id = " & cint(Site_ID))
                           updateSql=" update calendar set " 
                           if not(isnull(rsGetInfo.fields("File_Name").value)) then
                               updateSql = updateSql & "File_Name='" & rsGetInfo.fields("File_Name").value & "'"
                           else
                               updateSql = updateSql & "File_Name=NULL"
                           end if
                           
                           if not(isnull(rsGetInfo.fields("File_Size").value)) then
                                updateSql = updateSql & ",File_Size=" & rsGetInfo.fields("File_Size").value 
                           else
                                updateSql = updateSql & ",File_Size=0"
                           end if                           
                           
                           if not(isnull(rsGetInfo.fields("Archive_Name").value)) then
                                updateSql = updateSql & ",Archive_Name='" & rsGetInfo.fields("Archive_Name").value & "'"
                           else
                                updateSql = updateSql & ",Archive_Name=NULL" 
                           end if                           
                           
                           if not(isnull(rsGetInfo.fields("Archive_Size").value)) then
                                updateSql = updateSql & ",Archive_Size=" & rsGetInfo.fields("Archive_Size").value
                           else
                                updateSql = updateSql & ",Archive_Size=0"
                           end if
                           if not(isnull(rsGetInfo.fields("File_Page_Count").value)) then
                               updateSql = updateSql & ",File_Page_Count=" & rsGetInfo.fields("File_Page_Count").value 
                           else
                               updateSql = updateSql & ",File_Page_Count=0"
                           end if
                        
                           ' Last Update
                           if isdate(oFileUpEE.form("UDate")) then               
                            updateSql = updateSql & ",UDate="& "'" & killquote(oFileUpEE.form("UDate")) & "'"
                           end if
                        
                           ' Reference Number Revision Code
                           if not isblank(oFileUpEE.form("Revision_Code")) and not isblank(oFileUpEE.form("Item_Number")) then 
                             updateSql = updateSql & ",Revision_Code="& "'" & UCase(ReplaceQuote(oFileUpEE.form("Revision_Code"))) & "'"
                           else  
                             updateSql = updateSql & ",Revision_Code=NULL"
                           end if
                           updateSql = updateSql & ",Status=" & killquote(oFileUpEE.form("Status"))
                           updateSql = updateSql & " where id= " & rsClones.fields("id").value
                           set rsGetInfo = nothing
                           Conn.execute updateSql
                   end if
                   rsClones.movenext
            loop
            set rsClones = nothing
      end if       
      
      
      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      on error goto 0
      
    end if
    
    '>>>>>>>>>>>>>>>>>>>>>>>Modified by zensar to avoid Duplicate item number 12-09-2006>>>>>>>>>
    '>>>>>>>>>>>>>>>>>>>>>>>Modified by zensar for RI 506 12-09-2006>>>>>>>>>
      if isblank(error_msg) then
            ' Price List Save
              'if trim(killquote(oFileUpEE.form("txtAccessCode"))) <> "" then
                  conn.execute "exec PriceList_InsertUpdateAccessCode " & Record_ID & _
                      ",'" & killquote(oFileUpEE.form("txtAccessCode")) & "'"
              'else
                  'conn.execute "exec PriceList_InsertUpdateAccessCode " & Record_ID & _
                   '   ",'" & "FALSE" & "'"
              'end if
      end if
    ' Alert for Duplicate Title
    'if CInt(Show_PID) = CInt(true) then
    '  if CInt(PID_System) = 0 then
    '    if DuplicateTitle = true then
    '      with response
    '      .write vbCrLf
    '      .write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
    '      .write "var Title='Information Only\r\rDuplicate Title found:\r\rThe asset will be saved, however the Asset ID will be prefixed to Title in order to make it unique.\rThe Title was modified from:\r" &_
    '           """" & ReplaceQuote(DuplicateTitleText) & """\r" &_
    '           " to \r" &_
    '           """[" & Record_ID & "] " & ReplaceQuote(DuplicateTitleText) & """';" & vbCrLf
    '      .write "alert(Title);" & vbCrLf
    '      .write "</SCRIPT>" & vbCrLf
    '      end with
    '    end if
    '  elseif CInt(PID_System) = 1 then
    '    'Placeholder
    '  end if    
    'end if
 
  end if
  
  ' Set Status Comment to Current Oracle Deliverables Status for assets that contain Item Numbers
  %>
  <!--#include virtual="/sw-administratorNT/Calendar_Admin_Oracle_Status.asp"-->
  <%
  
  ' Sync MAC Items with Master MAC Data
  
  if Code >= 8000 and Code <= 8999 and Show_Calendar = CInt(False) then
  
    select case Status
      case 0
        sqlu = "UPDATE Calendar SET " & _
           "Status=" & Status & "," & _
           "LDays="  & LDays & ","  & _
           "BDate="  & BDate & ","  & _
           "LDate="  & LDate & ","  & _
           "EDate="  & EDate & ","  & _
           "XDate="  & XDate & ","  & _
           "XDays="  & XDays & ","  & _
           "PEDate=" & PEDate & "," & _
           "Subscription_Early=" & Subscription_Early  & " " &_
           "WHERE Campaign=" & Record_ID & " AND SubGroups NOT LIKE '%nomac%'"
      case 1
        sqlu = "UPDATE Calendar SET " & _
           "Status=" & Status & "," & _
           "LDays="  & LDays & ","  & _
           "BDate="  & BDate & ","  & _
           "LDate="  & LDate & ","  & _
           "EDate="  & EDate & ","  & _
           "XDate="  & XDate & ","  & _
           "XDays="  & XDays & ","  & _
           "PEDate=" & PEDate & "," & _
           "Subscription_Early=" & Subscription_Early  & " " &_
           "WHERE Campaign=" & Record_ID & " AND SubGroups NOT LIKE '%nomac%'"
      case 2
        sqlu = "UPDATE Calendar SET " & _
           "Status=" & Status & "," & _
           "LDays="  & LDays & ","  & _
           "BDate="  & BDate & ","  & _
           "LDate="  & LDate & ","  & _
           "EDate="  & EDate & ","  & _
           "XDate="  & XDate & ","  & _
           "XDays="  & XDays & ","  & _
           "PEDate=" & PEDate & "," & _
           "Subscription_Early=" & Subscription_Early & " " &_
           "WHERE (Campaign=" & Record_ID & " AND Content_Group=2 AND SubGroups NOT LIKE '%nomac%') OR (Campaign=" & Record_ID & " AND Content_Group=4 AND SubGroups NOT LIKE '%nomac%')"
    end select

    on error resume next
'    response.write sqlu
'    response.end
    conn.Execute(sqlu)

    if err.Number <> 0 then
      error_msg = error_msg & "<LI><B>" & Translate("Unable to update master PI/C records in database.",Login_Language,conn) & " " & Translate("PI/C Record",Login_Language,conn) & ": " & Record_ID & "</B><BR>" &_
                              "<INDENT>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & "</INDENT></LI>"
      
      on error goto 0
      
      if Admin_Access = 9 then
        error_msg = error_msg & "<P>" & sqlu & "<P>"
      end if
      
    else

      select case Status
    
        case 0, 1
          sqlu = "UPDATE Calendar SET " & _
                 "Status=" & Status & " " & _
                 "WHERE Campaign=" & Record_ID & " AND SubGroups LIKE '%nomac%'"
          conn.Execute(sqlu)

      end select
    
    end if

  end if

  ' Notify Approver
    
  if isblank(error_msg) then    
    
    if IsItChecked(oFileUpEE,"Send_EMail_Admin") = "on" and isnumeric(oFileUpEE.form("ID")) then
      
      ' --------------------------------------------------------------------------------------
      ' Open Connection to Mail Server
      ' --------------------------------------------------------------------------------------
        
      'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
      'adding new email method
      %>
      <!--#include virtual="/connections/connection_email_new.asp"-->
      <%
      
      'Mailer.ClearAllRecipients 
      'Mailer.ReturnReceipt  = False
      'Mailer.Priority       = 3
      msg.To = ""


      ' test address
      'Commented out by zensar as there is no need to send an email to extranetalerts group as the mail is
      'getting sent to the approver.
      'Mailer.AddBCC "Kelly Whitlock", "Kelly.Whitlock@fluke.com"   ' Domain Admin
        
      ' Get Site Stuff

      SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID

      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQL, conn, 3, 3
              
      if not rsSite.EOF then
        Site_Code          = rsSite("Site_Code")
        Site_Description   = rsSite("Site_Description")
        MailFromName       = rsSite("FromName")
        MailFromAddress    = rsSite("FromAddress")
        'Mailer.FromName    = MailFromName
        'Mailer.FromAddress = MailFromAddress
        msg.From = """" & MailFromName & """" & MailFromAddress
        MailSubject        = "Content Review Request"
      end if           
  
      rsSite.close
      set rsSite = nothing
  
      ' Get Submitter Stuff
      SQL = "SELECT UserData.* FROM UserData WHERE ID=" & oFileUpEE.form("Submitted_By")
      Set rsUser = Server.CreateObject("ADODB.Recordset")
      rsUser.Open SQL, conn, 3, 3
        
      if not rsUser.EOF then
        Submitted_By_Name   = rsUser("FirstName") & " " & rsUser("LastName")
        Submitted_By_Email  = rsUser("EMail")     
        'Mailer.ReplyTo      = Submitted_By_Email
        msg.ReplyTo = Submitted_By_Email
      else  
        Submitted_By_Name   = ""
        Submitted_By_Email  = ""
      end if
    
      rsUser.Close
      set rsUser = nothing
  
      ' Check if Approver Group is selected
      if not isblank(oFileUpEE.form("Review_By_Group")) and oFileUpEE.form("Review_By_Group") <> 0 then
  
        SQLApprovers = "Select Approvers.* FROM Approvers WHERE ID=" & oFileUpEE.form("Review_By_Group")
        Set rsApprovers = Server.CreateObject("ADODB.Recordset")
        rsApprovers.Open SQLApprovers, conn, 3, 3
                  
        ' Group Exists now Get Approver Stuff
        if not rsApprovers.EOF then
  
          SQL = "SELECT UserData.* FROM UserData WHERE ID=" & rsApprovers("Approver_ID")

          Set rsUser = Server.CreateObject("ADODB.Recordset")
          rsUser.Open SQL, conn, 3, 3
  
          if not rsUser.EOF then
            MailSendToName = rsUser("FirstName") & " " & rsUser("LastName")
            MailSendToAddress = rsUser("EMail")
            'Mailer.AddRecipient MailSendToName, MailSendToAddress   ' Approver
            msg.To = """" & MailSendToName & """" & MailSendToAddress
          else  
            'Mailer.AddRecipient FromName, FromAddress               ' Default to Site Administrator
            msg.To = """" & FromName & """" & FromAddress
          end if
          
          rsUser.close
          set rsUser = nothing
  
        end if  
         
      else  
          'Mailer.AddRecipient FromName, FromAddress                 ' Default to Site Administrator        
          msg.To = """" & FromName & """" & FromAddress
      end if  
        
      rsApprovers.close
      set rsApprovers = nothing
  
      ' Compose Advisory
      MailMessage = "This is an automated notification message from the " 
      MailMessage = MailMessage & MailFromName & " Extranet Support Server at:" & vbCrLf
      MailMessage = MailMessage & "http://" & Request.ServerVariables("SERVER_NAME") & "/" & lcase(Site_Code) & vbCrLf & vbCrLf
  
      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf
      MailMessage = MailMessage & "Content submitted by : " & Submitted_By_Name & vbCrLf & vbCrLf        

      SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE ID=" & oFileUpEE.form("Category_ID")

      Set rsCategory = Server.CreateObject("ADODB.Recordset")
      rsCategory.Open SQL, conn, 3, 3
        
      if not rsCategory.EOF then
        MailMessage = MailMessage & "Category : " & rsCategory("Title") & VbCrLf
      end if 
        
      rsCategory.close
      set rsCategory = nothing
     
      MailMessage = MailMessage & "Product or Series : "   

      if not isblank(oFileUpEE.form("Product_New")) then
        MailMessage = MailMessage & oFileUpEE.form("Product_New") & vbCrLf
      else
        MailMessage = MailMessage & oFileUpEE.form("Product") & vbCrLf
      end if  

      MailMessage = MailMessage & "Title : " & oFileUpEE.form("Title") & vbCrLf
      MailMessage = MailMessage & "Description : " & oFileUpEE.form("Description") & VbCrLf & vbCrlf

      MailMessage = MailMessage & "Thumbnail File : "
      if not isblank(oFileUpEE.form("Thumbnail")) or not isblank(oFileUpEE.form("Thumbnail_Existing")) then
        MailMessage = MailMessage & "Yes" & vbCrLf
      else
        MailMessage = MailMessage & "No" & vbCrLf
      end if

      MailMessage = MailMessage & "Content File : "      

      if not isblank(oFileUpEE.Form("File_Name")) or not isblank(oFileUpEE.form("File_Name_Existing")) then
        MailMessage = MailMessage & "Yes" & vbCrLf
      else
        MailMessage = MailMessage & "No" & vbCrLf
      end if

      MailMessage = MailMessage & "Link : "      

      if not isblank(oFileUpEE.Form("Link"))  then
        MailMessage = MailMessage & "Yes" & vbCrLf
      else
        MailMessage = MailMessage & "No" & vbCrLf
      end if
      
      MailMessage = MailMessage & "Proposed Go Live Date : " & GoLiveDate & vbCrLf

      MailMessage = MailMessage & "----------------------------------------------------------" & vbCrLf & vbCrLf
      MailMessage = MailMessage & "To view this Content Review Request, first click on the link below, log on to your administrator's account then return to this email and click again on the link below or check your Approval Queue:" & vbCrLf & vbCrLf
      MailMessage = MailMessage & "http://" & Request.ServerVariables("SERVER_NAME") & "/sw-administratorNT/Calendar_Edit.asp?ID=" & oFileUpEE.form("ID") & "&Site_ID=" & Site_ID & vbCrLf & vbCrLf

      MailMessage = MailMessage & "Note:  Please make sure that the entire URL address above is inserted in your browser. "
      MailMessage = MailMessage & "Sometimes email programs can cut off a long URL at a carriage return, or copy and paste "
      MailMessage = MailMessage & "does not work properly.  This results in a ""404 file not found error"". "
      MailMessage = MailMessage & "If you do receive such an error, please double check that the entire URL was provided to your browser." & vbCrLf & vbCrLf
      
      MailMessage = MailMessage & "If you have comments about this submission, you can reply directly to the submitter, by clicking on the [Reply] button of your email program." & vbCrLf & vbCrLf

      MailMessage = MailMessage & "Sincerely," & vbCrLf & vbCrLf & "The " & MailFromName & " Support Team"

      Call Send_EMail
        
    end if  
    
  end if
    
end if

if not isblank(error_msg) then
  Call Display_Error
else
  if CInt(Show_PID) = CInt(true) then
      if     CInt(PID_System) = 0 then
        %><!--#include virtual="/sw-administratorNT/SW-PCAT_FNET_SAVE.asp"--><%
      elseif CInt(PID_System) = 1 then
        %><!--#include virtual="/sw-administratorNT/SW-PCAT_FIND_SAVE.asp"--><%
      end if
  end if
    
  if lcase(Record_ID) = "add" then
    BackURL= "/sw-administratorNT/default.asp?Site_ID=" & Site_ID & "&ID=edit_record&Category_ID=" & oFileUpEE.form("Category_ID")
  else    
    BackURL= "/sw-administratorNT/Calendar_Edit.asp?ID=" & Record_ID & "&Site_ID=" & Site_ID & "&Show_View=" & oFileUpEE.form("Show_View")
  end if

  response.flush
  %>
  <!--#include virtual="/include/functions_table_border.asp"-->
  <%
  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
  response.write "<TITLE>Calendar Admin Redirect to Edit</TITLE>" & vbCrLf
  response.write "</HEAD>"
  response.write "<BODY BGCOLOR=""White"" onLoad='return document.foo.submit();' LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
  response.write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """>" & vbCrLf
  response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """>" & vbCrLf
  response.write "<DIV ALIGN=CENTER>"
  Call Nav_Border_Begin
  response.write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
  response.write "If your browser does not automatically return to the edit screen<BR>within 5 seconds, click on the [Continue] link below.<P>"
  response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
  response.write "</TD></TR></TABLE>" & vbCrLf
  Call Nav_Border_End
  response.write "</FORM>" & vbCrLf
  response.write "</DIV>"
  response.write "</BODY>"
  response.write "</HTML>"
  response.flush
  
end if

set oFileUpEE = nothing
Set conf = Nothing
Set msg = Nothing

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' IsItChecked - FileUpEE raises and error if you reference a non-exsistant form element
' so this function does that checking and returns the value if valid or "" if the 
' element is not in the collection.
' --------------------------------------------------------------------------------------

function IsItChecked(oFileUpEE,myCheckBoxName)

  Dim oFormItem, sFileUpValue
  
  sFileUpValue = ""
  
  for each oFormItem in oFileUpEE.Form
    if Trim(Lcase(oFormItem.Name)) = Trim(LCase(myCheckBoxName)) then
      sFileUpValue = oFormItem.Value
      exit for
    end if
  next

  IsItChecked = sFileUpValue
      
end function

' --------------------------------------------------------------------------------------
' Returns next available Generic Item_Number within range, based on Site_ID prefix
' --------------------------------------------------------------------------------------
function GetNextGenericNumber(Site_ID)

  sSite_ID = mid("00",1,2 - Len(Trim(CStr(Site_ID)))) & Trim(CStr(Site_ID))

  Start_Number = "9" & sSite_ID & "0000" 
  End_Number   = "9" & sSite_ID & "9999"
  
  SQL = "SELECT TOP 1 L.Item_Number + 1 AS Start " &_
        "FROM dbo.Calendar L LEFT OUTER JOIN " &_
        "     dbo.Calendar R ON L.Item_Number + 1 = R.Item_Number " &_
        "WHERE (L.Item_Number >= " & Start_Number & ") AND (R.Item_Number IS NULL) AND (L.Item_Number <= " & End_Number & ") " &_
        "ORDER BY L.Item_Number"

  Set rsGeneric = Server.CreateObject("ADODB.Recordset")
  rsGeneric.Open SQL, conn, 3, 3
  
  if not rsGeneric.EOF then
    Item_Number = rsGeneric("Start")
  else
    Item_Number = -1
  end if
  
  rsGeneric.close
  set rsGeneric = nothing
  
  GetNextGenericNumber = Item_Number
          
end function

' --------------------------------------------------------------------------------------

sub Display_Error

  Screen_Title   = "Support Extranet - Content / Event Add / Update Error Screen"
  Bar_Title      = Screen_Title
  Top_Navigation = false
  Navigation     = false
  Content_Width  = 95  ' Percent
  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%
  response.write "<B>" & Translate("Invalid format or missing value for",Login_Language,conn) & ":</B><BR>"
  response.write "<FONT COLOR=""Red"">"
  response.write "<UL>"
  response.write error_msg
  
  if IsObject(oFileUpEE.Form("Include")) or IsObject(oFileUpEE.Form("File_Name")) or IsObject(oFileUpEE.Form("File_Name_POD")) or IsObject(oFileUpEE.Form("Thumbnail")) then
    response.write "<BR><LI>" & Translate("Because of the above required field omission, some of the file upload information you supplied may have not be saved.",Login_Language,conn) & "<BR>" & Translate("If you are uploading any Content, Asset, Include or Thumbnail files, you will have to re-specify their source paths, otherwise ignore this portion of the alert.",Login_Language,conn) & "</LI>"
  end if    
    
  response.write "</UL></FONT>"
  response.write "<BR><BR>" & Translate("Please use the [Back] button of your browser to return to the previous screen and correct this information.",Login_Language,conn)
  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%
  
  Call Disconnect_SiteWide
 
  response.flush
  response.end

end sub

' --------------------------------------------------------------------------------------

sub Send_EMail

  'Mailer.QMessage = False
  'Mailer.Subject  = MailSubject
  'Mailer.BodyText = MailMessage

  msg.Subject = MailSubject
  msg.TextBody = MailMessage

  'if Mailer.SendMail then
  'else
  '  error_msg = error_msg & vbCrLf & "<LI>" & Translate("Send eMail Failure.",Login_Language,conn) & "<BR><BR>" & Translate("Error",Login_Language,conn) & ": " & Mailer.Response & ". " & Translate("Report this error to the site Webmaster.",Login_Language,conn) & "</LI>"
  'end if   

  msg.Configuration = conf
  On Error Resume Next
  msg.Send
  If Err.Number = 0 then
    'Success
  Else
    error_msg = error_msg & vbCrLf & "<LI>" & Translate("Send eMail Failure.",Login_Language,conn) & "<BR><BR>" & Translate("Error",Login_Language,conn) & ": " & Err.Description & ". " & Translate("Report this error to the site Webmaster.",Login_Language,conn) & "</LI>"
  End If

end sub

' --------------------------------------------------------------------------------------
%>

<script language=javascript runat=server>
function enflag()
	{	
		document.base64Form.elements["flag"].value='1';		
		alert(document.base64Form.elements["hidtest"].value);		
	}


function deflag()
	{
		document.base64Form.elements["flag"].value='2';
		document.base64Form.elements["hidtest"].value=encodeBase64(document.base64Form.elements["theText"].value);		
		document.base64Form.submit();
	}

function urlDecode(str){
    str=str.replace(new RegExp('\\+','g'),' ');
    return unescape(str);
}
function urlEncode(str){
    str=escape(str);
    str=str.replace(new RegExp('\\+','g'),'%2B');
    return str.replace(new RegExp('%20','g'),'+');
}

var END_OF_INPUT = -1;

var base64Chars = new Array(
    'A','B','C','D','E','F','G','H',
    'I','J','K','L','M','N','O','P',
    'Q','R','S','T','U','V','W','X',
    'Y','Z','a','b','c','d','e','f',
    'g','h','i','j','k','l','m','n',
    'o','p','q','r','s','t','u','v',
    'w','x','y','z','0','1','2','3',
    '4','5','6','7','8','9','+','/'
);

var reverseBase64Chars = new Array();
for (var i=0; i < base64Chars.length; i++){
    reverseBase64Chars[base64Chars[i]] = i;
}

var base64Str;
var base64Count;
function setBase64Str(str){
    base64Str = str;
    base64Count = 0;
}
function readBase64(){    
    if (!base64Str) return END_OF_INPUT;
    if (base64Count >= base64Str.length) return END_OF_INPUT;
    var c = base64Str.charCodeAt(base64Count) & 0xff;
    base64Count++;
    return c;
}
function encodeBase64(str){
    setBase64Str(str);
    var result = '';
    var inBuffer = new Array(3);
    var lineCount = 0;
    var done = false;
    while (!done && (inBuffer[0] = readBase64()) != END_OF_INPUT){
        inBuffer[1] = readBase64();
        inBuffer[2] = readBase64();
        result += (base64Chars[ inBuffer[0] >> 2 ]);
        if (inBuffer[1] != END_OF_INPUT){
            result += (base64Chars [(( inBuffer[0] << 4 ) & 0x30) | (inBuffer[1] >> 4) ]);
            if (inBuffer[2] != END_OF_INPUT){
                result += (base64Chars [((inBuffer[1] << 2) & 0x3c) | (inBuffer[2] >> 6) ]);
                result += (base64Chars [inBuffer[2] & 0x3F]);
            } else {
                result += (base64Chars [((inBuffer[1] << 2) & 0x3c)]);
                result += ('=');
                done = true;
            }
        } else {
            result += (base64Chars [(( inBuffer[0] << 4 ) & 0x30)]);
            result += ('=');
            result += ('=');
            done = true;
        }
        lineCount += 4;
        if (lineCount >= 76){
            result += ('\n');
            lineCount = 0;
        }
    }
    return result;
}
function readReverseBase64(){   
    if (!base64Str) return END_OF_INPUT;
    while (true){      
        if (base64Count >= base64Str.length) return END_OF_INPUT;
        var nextCharacter = base64Str.charAt(base64Count);
        base64Count++;
        if (reverseBase64Chars[nextCharacter]){
            return reverseBase64Chars[nextCharacter];
        }
        if (nextCharacter == 'A') return 0;
    }
    return END_OF_INPUT;
}

function ntos(n){
    n=n.toString(16);
    if (n.length == 1) n="0"+n;
    n="%"+n;
    return unescape(n);
}

function decodeBase64(str){
    setBase64Str(str);
    var result = "";
    var inBuffer = new Array(4);
    var done = false;
    while (!done && (inBuffer[0] = readReverseBase64()) != END_OF_INPUT
        && (inBuffer[1] = readReverseBase64()) != END_OF_INPUT){
        inBuffer[2] = readReverseBase64();
        inBuffer[3] = readReverseBase64();
        result += ntos((((inBuffer[0] << 2) & 0xff)| inBuffer[1] >> 4));
        if (inBuffer[2] != END_OF_INPUT){
            result +=  ntos((((inBuffer[1] << 4) & 0xff)| inBuffer[2] >> 2));
            if (inBuffer[3] != END_OF_INPUT){
                result +=  ntos((((inBuffer[2] << 6)  & 0xff) | inBuffer[3]));
            } else {
                done = true;
            }
        } else {
            done = true;
        }
    }
    return result;
}

var digitArray = new Array('0','1','2','3','4','5','6','7','8','9','a','b','c','d','e','f');
function toHex(n){
    var result = ''
    var start = true;
    for (var i=32; i>0;){
        i-=4;
        var digit = (n>>i) & 0xf;
        if (!start || digit != 0){
            start = false;
            result += digitArray[digit];
        }
    }
    return (result==''?'0':result);
}

function pad(str, len, pad){
    var result = str;
    for (var i=str.length; i<len; i++){
        result = pad + result;
    }
    return result;
}

function encodeHex(str){
    var result = "";
    for (var i=0; i<str.length; i++){
        result += pad(toHex(str.charCodeAt(i)&0xff),2,'0');
    }
    return result;
}

function decodeHex(str){
    str = str.replace(new RegExp("s/[^0-9a-zA-Z]//g"));
    var result = "";
    var nextchar = "";
    for (var i=0; i<str.length; i++){
        nextchar += str.charAt(i);
        if (nextchar.length == 2){
            result += ntos(eval('0x'+nextchar));
            nextchar = "";
        }
    }
    return result;
}

function chr(code)
{
	return String.fromCharCode(code);
}

//returns utf8 encoded charachter of a unicode value.
//code must be a number indicating the Unicode value.
//returned value is a string between 1 and 4 charachters.
function code2utf(code)
{
	if (code < 128) return chr(code);
	if (code < 2048) return chr(192+(code>>6)) + chr(128+(code&63));
	if (code < 65536) return chr(224+(code>>12)) + chr(128+((code>>6)&63)) + chr(128+(code&63));
	if (code < 2097152) return chr(240+(code>>18)) + chr(128+((code>>12)&63)) + chr(128+((code>>6)&63)) + chr(128+(code&63));
}

//it is a private function for internal use in utf8Encode function 
function _utf8Encode(str)
{	
	var utf8str = new Array();
	for (var i=0; i<str.length; i++) {
		utf8str[i] = code2utf(str.charCodeAt(i));
	}
	return utf8str.join('');
}

//Encodes a unicode string to UTF8 format.
function utf8Encode(str)
{	
	var utf8str = new Array();
	var pos,j = 0;
	var tmpStr = '';
	
	while ((pos = str.search(/[^\x00-\x7F]/)) != -1) {
		tmpStr = str.match(/([^\x00-\x7F]+[\x00-\x7F]{0,10})+/)[0];
		utf8str[j++] = str.substr(0, pos);
		utf8str[j++] = _utf8Encode(tmpStr);
		str = str.substr(pos + tmpStr.length);
	}
	
	utf8str[j++] = str;
	return utf8str.join('');
}

//it is a private function for internal use in utf8Decode function 
function _utf8Decode(utf8str)
{	
	var str = new Array();
	var code,code2,code3,code4,j = 0;
	for (var i=0; i<utf8str.length; ) {
		code = utf8str.charCodeAt(i++);
		if (code > 127) code2 = utf8str.charCodeAt(i++);
		if (code > 223) code3 = utf8str.charCodeAt(i++);
		if (code > 239) code4 = utf8str.charCodeAt(i++);
		
		if (code < 128) str[j++]= chr(code);
		else if (code < 224) str[j++] = chr(((code-192)<<6) + (code2-128));
		else if (code < 240) str[j++] = chr(((code-224)<<12) + ((code2-128)<<6) + (code3-128));
		else str[j++] = chr(((code-240)<<18) + ((code2-128)<<12) + ((code3-128)<<6) + (code4-128));
	}
	return str.join('');
}

//Decodes a UTF8 formated string
function utf8Decode(utf8str)
{
	var str = new Array();
	var pos = 0;
	var tmpStr = '';
	var j=0;
	while ((pos = utf8str.search(/[^\x00-\x7F]/)) != -1) {
		tmpStr = utf8str.match(/([^\x00-\x7F]+[\x00-\x7F]{0,10})+/)[0];
		str[j++]= utf8str.substr(0, pos) + _utf8Decode(tmpStr);
		utf8str = utf8str.substr(pos + tmpStr.length);
	}
	
	str[j++] = utf8str;
	return str.join('');
}

</script>
