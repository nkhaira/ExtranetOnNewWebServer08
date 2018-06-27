<%@Language="VBScript" Codepage=65001%>
<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/adovbs.inc"-->
<%

Dim Script_Debug
Script_Debug = false

Dim PostFlag

PostFlag = request.querystring("PostFlag")
if not isblank(PostFlag) then
  if CInt(PostFlag) = -1 then
    PostFlag = true
  else
    PostFlag = false  
  end if  
end if

' --------------------------------------------------------------------------------------
' Setup FileUpEE
' --------------------------------------------------------------------------------------

Dim FileUpEE_Flag, FileUpEE_Remote_Flag, FileUpEE_TempPath, cProgressID, wProgressID, ServerName

ServerName = UCase(request.ServerVariables("SERVER_NAME"))
strPath = server.MapPath("/met-support-gold/download/metcal/procedures/")
FileUpEE_TempPath = Server.MapPath("/SW-FileUp_Temp")

' --------------------------------------------------------------------------------------

Set oFileUpEE = Server.CreateObject("SoftArtisans.FileUpEE")
oFileUpEE.TransferStage                           = saWebServer ' Must be set for WebServer as opposed to FileServer
oFileUpEE.DynamicAdjustScriptTimeout(saWebServer) = true
oFileUpEE.TempStorageLocation(saWebServer)        = FileUpEE_TempPath
oFileUpEE.ProgressIndicator(saClient)             = false
oFileUpEE.ProgressIndicator(saWebServer)          = false

on error resume next
oFileUpEE.ProcessRequest Request, False, False

if err.Number <> 0 then
	response.write "<B>WebServer Process Request Error</B><BR>" & Err.Description & " (" & Err.Source & ")"
  response.flush
 	response.end
end if
on error goto 0

oFileUpEE.OverwriteFiles = true

' --------------------------------------------------------------------------------------
' Posting Form Debug
' --------------------------------------------------------------------------------------

if Script_Debug = true then

  with response
	for each oFormItem in oFileUpEE.Form
		.write oFormItem.Name & ": |"
		if IsObject(oFormItem.Value) then
			for each oSubItem in oFormItem.Value
				.write oSubItem.Value & "| |"
			next
		else
		  .write oFormItem.Value
		end if
		.write "|<P>"
	next
  
  for each oFormItem in oFileUpEE.Files
		.write oFormItem.Name & ": |" & oFileUpEE.Files(oFormItem.Name).ClientFileName & "|<P>"
  next
    
  .flush
  .end
  
  end with

end if  

' --------------------------------------------------------------------------------------
' Get Values
' --------------------------------------------------------------------------------------

Dim cmd, prm
Dim iID, strNewRecord, strSearchKeyword, strSearchCalibrator, strAction, CurrPage
Dim bValidFile, strErr_Msg, bKeepCurrentFileName, bFileNA
Dim UpdateBy

iID                 = Trim(oFileUpEE.form("ID"))
strSearchKeyword    = oFileUpEE.form("Keyword")
strSearchCalibrator = oFileUpEE.form("Calibrator")
strSearchFileName   = oFileUpEE.form("FileName")
CurrPage           = oFileUpEE.form("CurrPage")
strNewRecord        = UCase(Trim(oFileUpEE.form("New")))
strAction           = UCase(Trim(oFileUpEE.form("Action")))


' --------------------------------------------------------------------------------------
' If this is page is being loaded as a result of a form save, process the form
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

if CInt(PostFlag) = CInt(true) then

  ' ------------------------------------------------------------------------------------
  ' Delete Procedure
  ' ------------------------------------------------------------------------------------  
	
  if Trim(UCase(oFileUpEE.form("DoWhat"))) = "DELETE" then
  
  	iID = oFileUpEE.form("ID")
  	
    Set cmd = Server.CreateObject("ADODB.Command")
    with cmd
  		Set .ActiveConnection = conn
  		.CommandType = adCmdStoredProc
  		.CommandText = "Admin_MetCal_Procedure_Delete"
  		.Parameters.Append .CreateParameter("@iID", adInteger, adParamInput, , iID)
  		.execute
  	end with
  	
    if err.number <> 0 then
    	DeleteProcedure = false
  		response.write("Errors found: " & err.Description & "<BR>")
      response.flush
      response.end
    else
    	DeleteProcedure = true
    end if

  ' ------------------------------------------------------------------------------------
  ' Save Procedure
  ' ------------------------------------------------------------------------------------
	
  elseif Trim(UCase(oFileUpEE.form("DoWhat"))) = "SAVE" then

    bValidFile           = false
    bKeepCurrentFileName = false
    bFileNA              = false
  
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Admin_MetCal_Procedure_Edit"
  
    iID = oFileUpEE.form("ID")

    strAction = UCase(oFileUpEE.form("Action"))
    if strAction = "CLONE" then
  	  iID = 0
    end if
  
    if isblank(iID) then
    	iID = 0
    end if

    if IsItChecked(oFileUpEE,"KeepCurrentFileName") = "on" then
  	  bKeepCurrentFileName = true
  	  bValidFile           = true
    end if
  
    if IsItChecked(oFileUpEE,"File_Not_Available") = "on" then
  	  bFileNA              = true
    	bValidFile           = true
    end if

  	if IsObject(oFileUpEE.Files("File_Name")) and (bKeepCurrentFileName = false) and (bFileNA <> true) then
  		if not isblank(oFileUpEE.Files("File_Name").ClientFilename) then

  			strFileName = oFileUpEE.Files("File_Name").ClientFilename

  			' Filter Invalid Characters in File Name
  			strFileName = Replace(strFileName," ","_")  ' Convert spaces to underscores
  
  			if instr(1,strFileName,".") > 0 then
  				bValidFile = true
  			else
  				strErr_Msg = "Not a valid zip file name format."
  			end if
  		else
  			strErr_Msg = "No zip file was loaded or found."
  		end if
  	end if

    Set prm = cmd.CreateParameter("@iID", adInteger, adParamInput, , iID)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@strInstrument", adVarchar, adParamInput, 100, oFileUpEE.form("Instrument"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@strAdjThreshold", adVarchar, adParamInput, 50, oFileUpEE.form("AdjThreshold"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iAuthor", adInteger, adParamInput, , oFileUpEE.form("Author"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iCompany", adInteger, adParamInput, , oFileUpEE.form("Company"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@dDate", adVarchar, adParamInput, 50, oFileUpEE.form("Date"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iPrimaryCalibrator", adInteger, adParamInput, , oFileUpEE.form("PrimCalibrators"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@strRevision", adVarchar, adParamInput, 50, oFileUpEE.form("Revision"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iType", adInteger, adParamInput, , oFileUpEE.form("Types"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@b5500Cal_Ready", adInteger, adParamInput, , oFileUpEE.form("b5500Cal_Ready"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@bRestricted", adInteger, adParamInput, , oFileUpEE.form("Restricted"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@dRestrictedDate", adVarchar, adParamInput, 10, oFileUpEE.form("RestrictedDate"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@strRestrictedNote", adVarchar, adParamInput, 50, oFileUpEE.form("RestrictedNote"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iSource", adInteger, adParamInput, , oFileUpEE.form("Sources"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iPricePoint", adInteger, adParamInput, , oFileUpEE.form("PricePoint"))
    cmd.Parameters.Append prm
  
    ' Save uploaded file now
 
    if ((iID = 0 and bValidFile and UCase(strAction) <> "CLONE") or (iID > 0 and bKeepCurrentFileName = false and bValidFile)) and (bFileNA = false) then
       oFileUpEE.Files("File_Name").SaveAs strPath & "\" & strFileName
    end if

    ' Finish populated command object
    if bKeepCurrentFileName = true then
  	  strFileName = oFileUpEE.form("File_Name_Current")
    end if
  
    if bFileNA = true then
    	strFileName = "NA"
    end if
     
    Set prm = cmd.CreateParameter("@strZipFileName", adVarchar, adParamInput, 255, strFileName)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@strDescription", adVarchar, adParamInput, 6000, left(oFileUpEE.form("Description") & "", 6000))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@iUpdateBy", adInteger, adParamInput, , oFileUpEE.form("UpdateBy"))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@dUpdateDate", adVarchar, adParamInput, 10, oFileUpEE.form("UpdateDate"))
    cmd.Parameters.Append prm
  
    if (bValidFile = true) then
  	  Set rsSubCategory = Server.CreateObject("ADODB.Recordset")
    	rsSubCategory.CursorLocation = adUseClient
    	rsSubCategory.CursorType = adOpenDynamic
    	rsSubCategory.open cmd
    end if
    
    set prm = nothing
    set cmd = nothing

    if err.number <> 0 or strErr_Msg <> "" and UCase(strFileName) <> "NA" then
    	SaveProcedure = false
  		response.write("Errors found: " & strErr_Msg & "<BR>")
      response.flush
      response.end
    else

      if IsItChecked(oFileUpEE,"KeepCurrentFileName") = "on" then
  
        SQL = "SELECT Procedure_ID FROM Metcal_Procedures WHERE ZipFileName='" & oFileUpEE.form("File_Name_Current") & "' AND Procedure_ID <> " & iID
        Set rsMCP = Server.CreateObject("ADODB.Recordset")
        rsMCP.Open SQL, conn, 3, 3
      
        do while not rsMCP.EOF
          SQLU = "UPDATE Metcal_Procedures SET ZipFileName='" & strFileName & "', UpdateBy=" & oFileUpEE.form("UpdateBy") & ", UpdateDate='" & oFileUpEE.form("UpdateDate") & "' WHERE Procedure_ID=" & rsMCP("Procedure_ID")
          conn.execute SQLU
          rsMCP.MoveNext
        loop
      
        rsMCP.close
        set rsMCP = nothing
      
      end if
      
      SaveProcedure = true
      
    end if

    set oFileUpEE = nothing

	end if
  
end if

Call Disconnect_SiteWide

response.redirect "Metcal_Procedures.asp?Mookie=Me&KeyWord=" & strSearchKeyword & "&Calibrator=" & strSearchCalibrator & "&CurrPage=" & CurrPage & "&ID=" & iID & "#PID" & iID

response.flush
response.end

' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------
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
%>
