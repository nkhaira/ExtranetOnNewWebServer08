<%
' --------------------------------------------------------------------------------------
' Translation Function
'
' This function is currently used in THREE places:
'	1) Extranet (\includes)
'	2) WWW (\include)
'	3) evtibg01\d$\nightly\xnet_subscriptions\email_sub_1.vbs
'
' Whenever ANY changes are made to this file, it must be tested and propagated to EACH
' of these files.
' --------------------------------------------------------------------------------------
    
Function Translate(TempString, Login_Language, conn)
  ' Constants for ado - copied from adovbs.inc.
  ' We hard code the ado values for the following reasons:
  ' 	1) If we were to include adovbs.inc WITHIN this function, the adovbs.inc
  '  	   file location would have to be different for each server.
  '	2) We can't include it OUTSIDE the function because although www consistently uses 
  '	   an adovbs.inc include, the extranet does not consistently implement it.
  '	3) The nightly xnet subscription can't use an include
  '
  ' At some point it would be be useful to wrap the contents of this function into a component.

  ' English Language Only
  
  Dim Active_Language
  
  ' English Language Only Check
  if LCase(Login_Language) = "elo" then
    Active_Language = "eng"    ' Europe wants English Only Sites for Euro PartnerPortal
  else
    Active_Language = LCase(Login_Language)
  end if
  
  Dim adInteger
  Dim adVarChar
  Dim adLongVarWChar 
  Dim adParamInput 
  Dim adCmdStoredProc 

  adInteger = 3
  adCmdStoredProc = &H0004
  adVarChar = 200
  adLongVarWChar = 203
  adParamInput = &H0001

  Dim Translated_String
  Dim Translated_ID
  Dim cmd
  Dim prm
  Dim rsTranslate
  Dim bGetRecordset

  if not isblank(TempString) then
    Translated_String = Trim(TempString)
  else
    Translated_String = ""
  end if
  Translated_ID     = ""
  
  set cmd = Server.CreateObject("ADODB.Command")
  set prm = Server.CreateObject("ADODB.Parameter")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  
  ' Look Up by Phrase ID - not currently used
  if not IsBlank(Translated_String) and isnumeric(Translated_String) then
   	cmd.CommandText = "Translations_Get_Translation_By_ID"
   	Set prm = cmd.CreateParameter("@iTranslationID", adInteger,adParamInput ,, CInt(Translate_String))
   	cmd.Parameters.Append prm
	  bGetRecordset = true
  ' Look Up by Phrase
  else
    if (LCase(Active_Language) <> "eng" and not isblank(Active_Language)) and not isblank(TempString) then
  	  cmd.CommandText = "Translations_Get_Translation_ByString"
     	Set prm = cmd.CreateParameter("@iSiteID", adInteger, adParamInput, , Site_ID)
  	  cmd.Parameters.Append prm
     	Set prm = cmd.CreateParameter("@sSearch_String", adLongVarWChar,adParamInput ,255, left(ReplaceQuote(TempString), 255))
     	cmd.Parameters.Append prm
     	Set prm = cmd.CreateParameter("@iOriginalSearch_String_Length", adInteger, adParamInput, , len(TempString))
     	cmd.Parameters.Append prm
	    bGetRecordset = true
    end if
  end if

  if bGetRecordset = true then
    Set prm = cmd.CreateParameter("@sLanguage_String", adVarchar,adParamInput ,3, uCase(Active_Language))
    cmd.Parameters.Append prm
    set rsTranslate = Server.CreateObject("ADODB.Recordset")
    set rsTranslate = cmd.execute

  	'response.write("eof: " & rsTranslate.eof & "<BR>")
  	'response.write("len: " & len(TempString) & "<BR>")

    if not rsTranslate.EOF then
      Translated_ID = rsTranslate("translation_id")
      Translated_Grouping = rsTranslate("grouping")      
      Translated_String = Trim(rsTranslate("translation"))

      rsTranslate.close
      set rsTranslate = nothing
      set prm = nothing
      set cmd = nothing
    else  
      Translated_ID="New"  
  	  ' If the len of the string was larger than 255 then we'll have to go ahead and insert it ourselves
  	  if len(TempString) > 255 and (Login_Language <> "eng" or Login_Language <> "elo") then
  	    set cmd = Server.CreateObject("ADODB.Command")
  	    Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdStoredProc
  		  cmd.CommandText = "Translations_Insert_Translation"
  	   	Set prm = cmd.CreateParameter("@sSearchString", adLongVarWChar, adParamInput, 4000, left(TempString, 4000))
  		  cmd.Parameters.Append prm
  	   	Set prm = cmd.CreateParameter("@sLanguageString", adLongVarWChar, adParamInput, 4000, "")
  		  cmd.Parameters.Append prm
  	   	Set prm = cmd.CreateParameter("@iSiteID", adInteger, adParamInput, , Site_ID)
    		cmd.Parameters.Append prm
    		cmd.execute
  	  end if
    end if
  end if      
    
'response.write("Translated_String: " & Translated_String & "<BR>")

  if Session("ShowTranslation") = True and LCase(Active_Language) <> "eng" and LCase(Active_Language) <> "elo" then
    Translate = "<SPAN CLASS=ShowTranslation>" & Translated_String & " [" & Trim(CStr(Translated_Grouping)) & "]</SPAN>"
  elseif Session("ShowTranslation") = True then
    Translate = "<SPAN CLASS=ShowTranslation>" & Translated_String & "</SPAN>"  
  else
    Translate = Translated_String
  end if

end Function      

%>