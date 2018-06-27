<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/5/2001
' Name:       Asset File Find It
' Purpose:    Log to SiteWide DB, Activity Table any Asset View, Download or Send
' --------------------------------------------------------------------------------------
' SW-Find_It.asp?Locator=[Site_ID]O[Accout_ID]O[Asset_ID]O[Key;Method]O[Expiration_Date]O[[Language]O[Session_ID]O[CID]O[SCID]O[PCID]O[CIN]O[CINN]
' SW-Find_It.asp?Document=[7-Digit Oracle Item Number of User Viewable PDF document]
'
' Delimiter = O   (Letter UCase("O") - Successive delimeters without data, required for optional parameters if parameters are used beyond required fields)
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_DB.asp"-->
<!--#include virtual="/connections/connection_EMail.asp"--> 
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

' --------------------------------------------------------------------------------------

Dim Script_Debug
Dim bShowPerms

Dim Login_Language
Dim Alt_Language

Dim Close_Window
Dim ErrMessage
Dim ErrType

Dim Site_ID
Dim Site_Code
Dim Site_Description

Dim Document_Site_ID
Document_Site_ID = 0

Dim File_Name
Dim Link_Name
Dim Thumbnail

Dim Mailer
Dim Verify

Dim Activity_Log
Activity_Log = True

Login_Language = "eng"
Alt_Language   = "eng"
ErrMessage     = ""
ErrType        = 0

bShowPerms     = True           ' bShowPerms allows the hidden link to subgroup display

if LCase(request("Debug")) = "on" then
  Script_Debug = True
else
  Script_Debug = False
end if

Dim Locator
Dim Document

Close_Window   = False

if not isblank(Request("SW-Locator")) then
  Locator    = Request("SW-Locator")
  Close_Window   = False
else
  Locator    = Request("Locator")
end if

Document   = Request("Document")
Verify     = Request("Verify")

set Session("Asset_ID") = nothing

Randomize()
RandomFile = Int(1000 * Rnd())

' Parameters passed from www.fluke.com to record site/category activity

Dim CMS_Site, CMS_Path, CMS_ID

CMS_Site = Request("CMS_Site")
CMS_Path = Request("CMS_Path")

' --------------------------------------------------------------------------------------
' Method Key
' --------------------------------------------------------------------------------------

xOLView         = 0  ' On-Line View (Default)
xOLDownLoad     = 1  ' On-Line Download
xOLSend         = 2  ' On-Line Send
xSSView         = 3  ' Subscription Service View
xSSDownload     = 4  ' Subscription Service Download
xSSSend         = 5  ' Subscription Service Send
xOLSendNoZip    = 6  ' On-Line Send Non-Zip File Version
xOLLink         = 7  ' On-Line Link
xOLLinkNoPop    = 8  ' On-Line Link No Pop-Up
xOLGateway      = 9  ' On-Line Gateway to Site
xOLGatewayNoPop = 10 ' On-Line Gateway to Site No Pop-Up
xOLViewPOD      = 11 ' On-Line Download Print on Demand Doccument
xOLDownLoadPOD  = 12 ' On-Line Download Print on Demand Doccument

' --------------------------------------------------------------------------------------
' Document - This just converts a Document ID (7-Digit Oracle Item Number) to the format
'            Used by Locator with presets.
' --------------------------------------------------------------------------------------

if not isblank(Document) and isnumeric(Document) then

  bShowPerms = False
  
  select case len(Document)
    case 6,7
      Asset  = Document
      Method = xOLView
    case 8
      Asset  = Mid(Document,1,7)
      Method = Mid(Document,8,1)
      
      select case Method
        case 0
          Method = xOLViewPOD
        case 1  
          Method = xOLDownLoadPOD
        case else
          ErrMessage = "<LI>Invalid Document ID Locator</LI>"
      end select   
    case else
      ErrMessage = "<LI>Invalid Document ID Locator</LI>"
  end select
  
  if isblank(ErrMessage) then

    SQL = "SELECT * FROM Calendar WHERE Calendar.Item_Number='" & Asset & "' AND File_Name IS NOT NULL AND Calendar.SubGroups LIKE '%view%' ORDER BY Calendar.UDate DESC"
    Set rsAsset = Server.CreateObject("ADODB.Recordset")
    rsAsset.Open SQL, conn, 3, 3
 
    if not rsAsset.EOF then
  
      Locator = CStr(rsAsset("Site_ID"))                 & "o" & _
                "1"                                      & "o" & _
                CStr(rsAsset("ID"))                      & "o" & _
                CStr(Method)                             & "o" & _
                CStr(Encode_Key(CStr(rsAsset("Site_ID")),"1",rsAsset("ID"))) & "o" & _             
                CStr(CLng(Date))                         & "o" & _
                "0"
       
      Document_Site_ID = 99
                
    else
    
      if LCase(Verify) <> "on" then
        Call Send_Invalid_Document_Email
      end if  

      ErrMessage = "<LI>We are sorry but the Document that you have requested is currently not available.</LI>"
      ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
      ErrMessage = ErrMessage & "<LI>Since this system does not have your Email address, please send an email to: Fluke-Info@fluke.com to enquire if you can be sent this document by another method.</LI>"
      ErrMessage = ErrMessage & "<LI>The Document Number you had requested was: " & Mid(Document,1,7) & "</LI>"

      Document_Site_ID = 99
       
    end if
  
    rsAsset.close
    set rsAsset = nothing

  else
  
    if LCase(Verify) <> "on" then
      Call Send_Invalid_Document_Email
    end if  

    ErrMessage = "<LI>We are sorry but the Document that you have requested is currently not available.</LI>"
    ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
    ErrMessage = ErrMessage & "<LI>Since this system does not have your Email address, please send an email to: Fluke-Info@fluke.com to enquire if you can be sent this document by another method.</LI>"
    ErrMessage = ErrMessage & "<LI>The Document Number you had requested was: " & Mid(Document,1,7) & "</LI>"

    Document_Site_ID = 99
  
  end if  
  
end if              

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

if not isblank(Locator) and isblank(ErrMessage) then

  Dim ErrString
  Dim MailMessage
  Dim MailSubject 

  %>
  <!--#include virtual="/include/functions_locator.asp"-->
  <%
  
  if Script_Debug then
  
    response.write "<FONT FACE=""Arial"" SIZE=2>"
    response.write "Locator: " & Replace(Locator,"O","<FONT COLOR=""Red"">O</FONT>") & "<BR>"
    response.write "Document: " & Document & "<BR>"
    response.write "Today's Serial Date: " & CLng(Date) & " " & CDate(CLng(Date)) & "<BR>"  
    response.write "Encoded Key: " & Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID)) & "<BR>"
    response.write "Decoded Key: " & Decode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID),Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID))) & "<BR><BR>"
  
    response.write "<TABLE>"

    for i = 0 to Parameter_Max
      response.write "<TR>"
      response.write "<TD>" & i & "</TD>"
      response.write "<TD>" & Parameter_Key(i) & "</TD>"
      response.write "<TD>" & Parameter(i) & "</TD>"
      response.write "</TR>"      
    next

    response.write "</TABLE>"
    
    with response
      .write "<BR>"
      .write "Short Format Test URL: "    & "http://" & request.ServerVariables("SERVER_NAME") & "/Find_It.asp?Locator=" 
      .write CInt(Parameter(xSite_ID))    & "o"
      .write CInt(Parameter(xAccount_ID)) & "o"
      .write CInt(Parameter(xAsset_ID))   & "o"
      .write CInt(Parameter(xMethod))     & "o"
      .write CInt(Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID))) & "o"
      .write CLng(Date)                   & "o"
      .write "0"
      .write "&Debug=on<BR><BR>"
    end with

    with response
      .write "Long&nbsp;&nbsp;Format Test URL: " & "http://" & request.ServerVariables("SERVER_NAME") & "/Find_It.asp?Locator=" 
      .write CInt(Parameter(xSite_ID))    & "o"
      .write CInt(Parameter(xAccount_ID)) & "o"
      .write CInt(Parameter(xAsset_ID))   & "o"
      .write CInt(Parameter(xMethod))     & "o"
      .write CInt(Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID))) & "o"
      .write CLng(Date)                   & "o"
      .write "0"                          & "o"
      .write CLng(Parameter(xSession_ID)) & "o"
      .write CInt(Parameter(xCID))        & "o"
      .write CInt(Parameter(xSCID))       & "o"
      .write CInt(Parameter(xPCID))       & "o"
      .write CInt(Parameter(xCIN))        & "o"
      .write CInt(Parameter(xCINN))
      .write "&Debug=on<BR><BR>"
    end with

    response.write CDate(Parameter(xExpiration_Date)) & " " & CDate(Date)
    if CDate(Parameter(xExpiration_Date)) >= CDate(Date) then response.write " True"
    response.write "<BR><BR>"
    
  end if
    
  ' --------------------------------------------------------------------------------------
  ' Verify Expiration Date to ensure that link to asset should not automatically expire (Subscription Service)
  ' --------------------------------------------------------------------------------------
  
  if CDate(Parameter(xExpiration_Date)) >= CDate(Date) _
     and Decode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID),Parameter(xKey)) = True then

    ' --------------------------------------------------------------------------------------
    ' Validate User
    ' --------------------------------------------------------------------------------------
    
    if CInt(Parameter(xAccount_ID)) > 1 then

      SQL = "SELECT UserData.ID, UserData.Site_ID, UserData.NTLogin, UserData.Password, UserData.ExpirationDate, UserData.SubGroups FROM UserData WHERE UserData.Site_ID=" & CInt(Parameter(xSite_ID)) & " AND UserData.ID=" & CInt(Parameter(xAccount_ID))
  
      Set rsValidate = Server.CreateObject("ADODB.Recordset")
      rsValidate.Open SQL, conn, 3, 3
  
      Validated = False

      if not rsValidate.EOF then
        if CDate(Date) < CDate(rsValidate("ExpirationDate")) then
          Validated = True
          Session("LOGON_USER") = rsValidate("NTLogin")
          Session("Password")   = rsValidate("Password")
          Session("Site_ID")    = Parameter(xSite_ID)
          if instr(1,rsValidate("SubGroups"),"submitter")     = 0 and _
             instr(1,rsValidate("SubGroups"),"content")       = 0 and _
             instr(1,rsValidate("SubGroups"),"account")       = 0 and _
             instr(1,rsValidate("SubGroups"),"administrator") = 0 and _             
             instr(1,rsValidate("SubGroups"),"domain")        = 0 then
             Activity_Log = False   ' Do not log Admin Actions
          else
             Activity_Log = True
          end if
        end if
      end if
  
      rsValidate.close
      set rsValidate = nothing
      
    else                  ' Bypass check for DOCUMENT

      Validated    = True
      Activity_Log = True

    end if
      
    ' --------------------------------------------------------------------------------------
    ' Validated User - Continue with Site Code and Asset Lookup
    ' --------------------------------------------------------------------------------------
  
    if Validated then
  
      ' --------------------------------------------------------------------------------------
      ' Get Site Information
      ' --------------------------------------------------------------------------------------
      
      SQL = "SELECT Site.Site_Code, Site_Description FROM Site WHERE ID=" & CInt(Parameter(xSite_ID))      
      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQL, conn, 3, 3
  
      if not rsSite.EOF then
          Site_Code        = rsSite("Site_Code")
          Site_Description = rsSite("Site_Description")
          Site_ID          = CInt(Parameter(xSite_ID))
      end if
      
      rsSite.close
      set rsSite = nothing    
          
      ' Use FulFillment Center Description    

      if Document_Site_ID = 99 then

        SQL = "SELECT Site.Site_Code, Site_Description FROM Site WHERE ID=" & CInt(Document_Site_ID)
        Set rsSite = Server.CreateObject("ADODB.Recordset")
        rsSite.Open SQL, conn, 3, 3
  
        if not rsSite.EOF then
          Site_Description = rsSite("Site_Description")
        else
          Site_Description = ""
        end if

        rsSite.close
        set rsSite = nothing    
          
      end if

      ' --------------------------------------------------------------------------------------
      ' Get Asset Path
      ' --------------------------------------------------------------------------------------
  
      if CInt(Parameter(xAccount_ID)) > 1 then
        SQL = "SELECT * FROM Calendar WHERE Site_ID=" & CInt(Parameter(xSite_ID)) & " AND ID=" & CInt(Parameter(xAsset_ID))
      else
        SQL = "SELECT * FROM Calendar WHERE ID=" & CInt(Parameter(xAsset_ID))
      end if
        
      Set rsAsset = Server.CreateObject("ADODB.Recordset")
      rsAsset.Open SQL, conn, 3, 3
  
      File_Name = ""
      Link_Name = ""
      Thumbnail = ""              

      if not rsAsset.EOF then
      
        select case CInt(Parameter(xMethod))

          case xOLView, xOLSendNoZip, xSSView

            if not isblank(rsAsset("File_Name")) then
              File_Name = Trim(rsAsset("File_Name"))
            end if

          case xOLDownLoad, xOLSend, xSSDownload, xSSSend

            if not isblank(rsAsset("Archive_Name")) then
              File_Name = Trim(rsAsset("Archive_Name"))              
            else  
              File_Name = Trim(rsAsset("File_Name"))
            end if
  
          case xOLLink, xOLLinkNoPop

            Link_Name    = Replace(Trim(rsAsset("Link")),"https://support.fluke.com","https://" & request.ServerVariables("SERVER_NAME"))
            Link_Name    = Replace(Link_Name,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))
  
          case xOLGateway, xOLGatewayNoPop

            Link_Name    = "http://" & request.ServerVariables("SERVER_NAME")

        end select
        
        Thumbnail    = Trim(rsAsset("Thumbnail"))
        Title        = Trim(rsAsset("Title"))
        Category     = Trim(rsAsset("Sub_Category"))
        if not isnull(rsAsset("File_Size")) then
          File_Size    = FormatNumber(CDbl(CDbl(rsAsset("File_Size") / 1024)),0)
        else
          File_Size = 0
        end if    
        if not isnull(rsAsset("Archive_Size")) then
          Archive_Size = FormatNumber(CDbl(CDbl(rsAsset("Archive_Size") / 1024)),0)
        else  
          Archive_Size = 0
        end if  
      end if
  
      rsAsset.close
      set rsAsset = nothing
  
      if not isblank(File_Name) or not isblank(Link_Name) then
      
        if not isblank(Thumbnail) and instr(1,LCase(Thumbnail),LCase(request.ServerVariables("SERVER_NAME"))) = 0 then
          Thumbnail = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(Site_Code) & "/" &  Thumbnail
        end if  
      
        if not isblank(File_Name) then
        
          ' Needs code added here using CreateObject("Scripting.FileSystemObject") to ensure that file has not been deleted
        
          File_Redirect = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(Site_Code) & "/" & File_Name
        
          ' --------------------------------------------------------------------------------------
          ' Convert Language Code
          ' --------------------------------------------------------------------------------------
 
          if CInt(Parameter(xLanguage)) > 0 then
  
            SQL = "SELECT Language.ID, Language.Code FROM Language WHERE Language.ID=" & CInt(Parameter(xLanguage))
            Set rsLanguage = Server.CreateObject("ADODB.Recordset")
            rsLanguage.Open SQL, conn, 3, 3
  
            if not rsLanguage.EOF then
              Login_Language = rsLanguage("Code")
            end if
  
            rsLanguage.close
            Set rsLanguage = nothing
  
          end if
        
          select case LCase(Login_Language)
            case "chi", "zho", "thi", "jpn", "kor"
              Alt_Language = "eng"
            case else
              Alt_Language = LCase(Login_Language)
          end select
        
        elseif not isblank(Link_Name) then

          File_Redirect = Link_Name

        end if        
  
        ' --------------------------------------------------------------------------------------
        ' Log Activity of Download to Activity Table
        ' --------------------------------------------------------------------------------------
        
        if CInt(Activity_Log) = CInt(True) then                             ' Do not Log Fluke Entity Users Access to Assets

          ' Retrieve/Add Cross Reference to CMS_Path
          
          if not isblank(CMS_Site) then
          
            if not isblank(CMS_Path) then
            
              CMS_SQL = "SELECT ID, CMS_Path FROM CMS_XReference WHERE CMS_Path='" & CMS_Path & "'"
              Set rsCMS = Server.CreateObject("ADODB.Recordset")
              rsCMS.Open CMS_SQL, conn, 3, 3
 
              if not rsCMS.EOF then
                CMS_ID = rsCMS("ID")
              else
                CMS_ID = 0
              end if

              rsCMS.close
              set rsCMS = nothing

              if CInt(CMS_ID) = 0 then
                CMS_ID = Get_New_Record_ID ("CMS_XReference", "CMS_Path", "", conn)
                CMS_SQL = "UPDATE CMS_XReference SET CMS_Path='" & CMS_Path & "' WHERE ID=" & CMS_ID
                conn.Execute(CMS_SQL)
                set CMS_SQL = nothing
              end if
              
            else
              CMS_ID = 0
            end if  
          
          end if      
                
          ' Update User's Last Logon Date/Time since User is Accessing an Asset at the Site.
          
          if Parameter(xAccount_ID) > 1 then
            ActivitySQL = "UPDATE UserData SET UserData.Logon='" & Now() & "' WHERE UserData.NTLogin='" & Session("LOGON_USER") & "'"
            conn.Execute (ActivitySQL)
          end if  

          ActivitySQL = "INSERT INTO Activity"               & _
                        " ( "                                & _
                          "Account_ID,"                      & _
                          "Site_ID,"                         & _
                          "Session_ID,"                      & _
                          "View_Time,"                       & _
                          "CID,"                             & _
                          "SCID,"                            & _
                          "PCID,"                            & _
                          "CIN,"                             & _
                          "CINN,"                            & _
                          "Language,"                        & _
                          "Method,"                          & _
                          "Calendar_ID,"                     & _
                          "CMS_Site,"                        & _
                          "CMS_ID"                           & _
                        " ) "                                & _
                        "VALUES"                             & _
                        " ( "                                & _
                          CInt(Parameter(xAccount_ID)) & "," & _
                          CInt(Parameter(xSite_ID))    & "," & _
                          CLng(Parameter(xSession_ID)) & "," & _
                          "'" & Now & "'"              & "," & _
                          CInt(Parameter(xCID))        & "," & _
                          CInt(Parameter(xSCID))       & "," & _
                          CInt(Parameter(xPCID))       & "," & _
                          CInt(Parameter(xCIN))        & "," & _
                          CInt(Parameter(xCINN))       & "," & _
                          "'" & Login_Language & "'"   & "," & _
                          CInt(Parameter(xMethod))     & "," & _
                          CInt(Parameter(xAsset_ID))   & ","
                          
                          if isblank(CMS_Site) then
                            ActivitySQL = ActivitySQL & "NULL," & CInt(CMS_ID) & " ) "
                          else
                            ActivitySQL = ActivitySQL & "'" & CMS_Site & "'," & CInt(CMS_ID) & " ) "
                          end if

          if Script_Debug then
            response.write "<BR>Activity SQL: " & ActivitySQL & "<BR><BR>"
            response.write "File Path: " & File_Redirect & "<BR>"
            response.write "</FONT>"
          end if
          
          conn.Execute (ActivitySQL)
                   
        end if
        
       ' Send User the File Requested         
       
        if Script_Debug = False then

          select case CInt(Parameter(xMethod))
           
            case xOLView, xSSView
          
              Session("Asset_ID") = CInt(Parameter(xAsset_ID))

              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0>"
                ErrMessage = ErrMessage & "<TR>"
                ErrMessage = ErrMessage & "<TD WIDTH=80>"
                ErrMessage = ErrMessage & "<IMG SRC=""" & Thumbnail & """ WIDTH=80 BORDER=1>"
                ErrMessage = ErrMessage & "</TD>"
                ErrMessage = ErrMessage & "<TD CLASS=Medium><UL>"
              end if          
              
              ErrMessage = ErrMessage & "<LI>" & Translate("The document you have requested is currently loading to your browser to view.",Login_Language,conn) & "</LI><BR><BR>"

              ErrMessage = ErrMessage & "<LI>" & Translate("Title",Login_Language,conn) & ": "

      			  if bShowPerms then
      			  	ErrMessage = ErrMessage & "<A HREF=""/SW-Administrator/SubGroup_Codes.asp?asset_id=" &_
                    				 CInt(Parameter(xAsset_ID)) & "&site_id=" & Cint(Parameter(xSite_ID)) &_
                             "&Language=" & Login_Language &_
                   				   """><span style=""text-decoration:none;"">" & Title & "</span></a>"
      			  else
			        	ErrMessage = ErrMessage & Title
      			  end if
			  
              if not isblank(Category) then
                ErrMessage = ErrMessage & " - " & Translate(Category,Login_Language,conn)
              end if
              ErrMessage = ErrMessage & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI>" & Translate("This window is no longer required and can be closed at any time.",Login_Language,conn) & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI><FONT COLOR=""Red"">" & Translate("If the document you have requested does not appear",Login_Language,conn) & ", " & "<A HREF=""" & File_Redirect & """>" & Translate("click here",Login_Language,conn) & "</A>.</FONT><BR><BR></LI>"

              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "</UL></TD></TR></TABLE>"
              end if
              
              ' Uses same window as status screen for viewing only
              if CInt(Parameter(xMethod)) = xOLView then
                ErrMessage = ErrMessage & vbCrLf              
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.location.href='" & File_Redirect & "','File_View_" & RandomFile & "','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=no,links=no,history=no';" & vbCrLf
                ErrMessage = ErrMessage & "window.title='Electronic File Fulfillment';"
                ErrMessage = ErrMessage & "window.focus();" & vbCrLf
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf

              else
                ErrMessage = ErrMessage & vbCrLf              
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "wFile_View = window.open('" & File_Redirect & "','File_View_" & RandomFile & "','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=no,links=no');" & vbCrLf
                if instr(1,LCase(Link_Name),LCase(request.ServerVariables("SERVER_NAME"))) > 0 then
                  ErrMessage = ErrMessage & "wSite_Link.focus();" & vbCrLf
                end if  
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
                if Close_Window = False then
                  ErrMessage = ErrMessage & vbCrLf
                  ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                  ErrMessage = ErrMessage & "<!--" & vbCrLf
                  ErrMessage = ErrMessage & "window.blur();" & vbCrLf              
                  ErrMessage = ErrMessage & "// -->" & vbCrLf
                  ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
                else 
                  ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                  ErrMessage = ErrMessage & "<!--" & vbCrLf
                  ErrMessage = ErrMessage & "window.blur();" & vbCrLf
                  ErrMessage = ErrMessage & "window.close();" & vbCrLf
                  ErrMessage = ErrMessage & "// -->" & vbCrLf
                  ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf                
                end if
              end if  
                           
              ErrType = 1
              
              Call Status_Screen
             
            case xOLDownload, xSSDownload

              Session("Asset_ID") = CInt(Parameter(xAsset_ID))
              
              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0>"
                ErrMessage = ErrMessage & "<TR>"
                ErrMessage = ErrMessage & "<TD WIDTH=80>"
                ErrMessage = ErrMessage & "<IMG SRC=""" & Thumbnail & """ WIDTH=80 BORDER=1>"
                ErrMessage = ErrMessage & "</TD>"
                ErrMessage = ErrMessage & "<TD CLASS=Medium><UL>"
              end if          

              ErrMessage = ErrMessage & "<LI>" & Translate("The document or file you have requested is currently loading. Once loading has completed, a pop-up dialog box should appear asking you if to 'Open' or 'Save File As...' to your local PC.",Login_Language,conn) & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI>" & Translate("Title",Login_Language,conn) & ": "

      			  if bShowPerms then
      			  	ErrMessage = ErrMessage & "<A HREF=""/SW-Administrator/SubGroup_Codes.asp?asset_id=" &_
                  				   CInt(Parameter(xAsset_ID)) & "&site_id=" & Cint(Parameter(xSite_ID)) &_
                  				   "&Language=" & Login_Language &_
                    				 """><span style=""text-decoration:none;"">" & Title & "</span></a>"
        		  else
			        	ErrMessage = ErrMessage & Title
      			  end if
			  
              if not isblank(Category) then
                ErrMessage = ErrMessage & " - " & Translate(Category,Login_Language,conn)
              end if
              ErrMessage = ErrMessage & "</LI><BR><BR>"              
              ErrMessage = ErrMessage & "<LI>" & Translate("This window is no longer required and can be closed at any time.",Login_Language,conn) & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI><FONT COLOR=""Red"">" & Translate("If the document or file you have requested does not appear",Login_Language,conn) & ", " & "<A HREF=""" & File_Redirect & """>" & Translate("click here",Login_Language,conn) & "</A>.</FONT></LI>"

              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "</UL></TD></TR></TABLE>"
              end if
              
              ErrMessage = ErrMessage & vbCrLf              
              ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
              ErrMessage = ErrMessage & "<!--" & vbCrLf
              ErrMessage = ErrMessage & "wFile_View = window.open('" & File_Redirect & "','File_View_" & RandomFile & "','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=no,links=no');" & vbCrLf
              ErrMessage = ErrMessage & "// -->" & vbCrLf
              ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf

              if Close_Window = False then
                ErrMessage = ErrMessage & vbCrLf
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf              
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
              else
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf
                ErrMessage = ErrMessage & "window.close();" & vbCrLf
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf                
              end if

              ErrType = 1
              
              Call Status_Screen                        
              
            case xOLSend, xSSSend, xOLSendNoZip
                                               
              Session("Asset_ID") = CInt(Parameter(xAsset_ID))
              
              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0>"
                ErrMessage = ErrMessage & "<TR>"
                ErrMessage = ErrMessage & "<TD WIDTH=80>"
                ErrMessage = ErrMessage & "<IMG SRC=""" & Thumbnail & """ WIDTH=80 BORDER=1>"
                ErrMessage = ErrMessage & "</TD>"
                ErrMessage = ErrMessage & "<TD CLASS=Medium><UL>"
              end if
              
              ErrMessage = ErrMessage & "<LI>" & Translate("The document or file that you have requested has been sent to you by email.",Login_Language,conn) & "<BR><BR></LI>"
              ErrMessage = ErrMessage & "<LI>" & Translate("Title",Login_Language,conn) & ": " & Title
              if not isblank(Category) then
                ErrMessage = ErrMessage & " - " & Translate(Category,Login_Language,conn)
              end if
              ErrMessage = ErrMessage & "<BR><BR></LI>"
              ErrMessage = ErrMessage & "<LI>" & Translate("This window is no longer required and can be closed at any time.",Login_Language,conn) & "</LI>"
        
              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "</UL></TD></TR></TABLE>"
              end if

              if Close_Window = False then
                ErrMessage = ErrMessage & vbCrLf
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf              
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
              else
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf
                ErrMessage = ErrMessage & "window.close();" & vbCrLf
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf                
              end if

              ErrType = 1
                           
              Call Status_Screen

              ErrType = 0
              Call Send_EMail_File
              
            ' Link
            case xOLLink, xOLLinkNoPop
            
              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0>"
                ErrMessage = ErrMessage & "<TR>"
                ErrMessage = ErrMessage & "<TD WIDTH=80>"
                ErrMessage = ErrMessage & "<IMG SRC=""" & Thumbnail & """ WIDTH=80 BORDER=1>"
                ErrMessage = ErrMessage & "</TD>"
                ErrMessage = ErrMessage & "<TD CLASS=Medium><UL>"
              end if          
              
              ErrMessage = ErrMessage & "<LI>" & Translate("The site you have requested is currently loading to your browser to view.",Login_Language,conn) & "</LI><BR><BR>"

              ErrMessage = ErrMessage & "<LI>" & Translate("Title",Login_Language,conn) & ": " & Title
              if not isblank(Category) then
                ErrMessage = ErrMessage & " - " & Translate(Category,Login_Language,conn)
              end if
              ErrMessage = ErrMessage & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI>" & Translate("This window is no longer required and can be closed at any time.",Login_Language,conn) & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI><FONT COLOR=""Red"">" & Translate("If the site you have requested does not appear",Login_Language,conn) & ", " & "<A HREF=""" & File_Redirect & """>" & Translate("click here",Login_Language,conn) & "</A>.</FONT><BR><BR></LI>"

              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "</UL></TD></TR></TABLE>"
              end if

              ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
              ErrMessage = ErrMessage & "<!--" & vbCrLf
              ErrMessage = ErrMessage & "wSite_Link = window.open('" & Link_Name & "','Site_Link_" & RandomFile & "');" & vbCrLf
              if instr(1,LCase(Link_Name),LCase(request.ServerVariables("SERVER_NAME"))) > 0 then
                ErrMessage = ErrMessage & "wSite_Link.focus();" & vbCrLf
              end if  
              ErrMessage = ErrMessage & "// -->" & vbCrLf
              ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf

              if Close_Window = False then
                ErrMessage = ErrMessage & vbCrLf
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf              
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
              else
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf
                ErrMessage = ErrMessage & "window.close();" & vbCrLf
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf                
              end if

              Session("Asset_ID") = CInt(Parameter(xAsset_ID))
                            
              ErrType = 1
                           
              Call Status_Screen

            ' Gateway
            case xOLGateway, xOLGatewayNoPop

              if not isblank(Thumbnail) and instr(1,LCase(Thumbnail),LCase(request.ServerVariables("SERVER_NAME"))) = 0 then
                Thumbnail = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(Site_Code) & "/" &  Thumbnail
              elseif not isblank(Thumbnail) then
            
                ErrMessage = ErrMessage & "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0>"
                ErrMessage = ErrMessage & "<TR>"
                ErrMessage = ErrMessage & "<TD WIDTH=80>"
                ErrMessage = ErrMessage & "<IMG SRC=""" & Thumbnail & """ WIDTH=80 BORDER=1>"
                ErrMessage = ErrMessage & "</TD>"
                ErrMessage = ErrMessage & "<TD CLASS=Medium><UL>"

              end if          
              
              ErrMessage = ErrMessage & "<LI>" & Translate("The site you have requested is currently loading to your browser to view.",Login_Language,conn) & "</LI><BR><BR>"

              ErrMessage = ErrMessage & "<LI>" & Translate("Title",Login_Language,conn) & ": "
              
      			  if bShowPerms then
      			  	ErrMessage = ErrMessage & "<A HREF=""/SW-Administrator/SubGroup_Codes.asp?asset_id=" &_
      					CInt(Parameter(xAsset_ID)) & "&site_id=" & Cint(Parameter(xSite_ID)) &_
      					"&Language=" & Login_Language &_
      					"""><span style=""text-decoration:none;"">" & Title & "</span></a>"
      			  else
      			  	ErrMessage = ErrMessage & Title
      			  end if
			  
              if not isblank(Category) then
                ErrMessage = ErrMessage & " - " & Translate(Category,Login_Language,conn)
              end if
              ErrMessage = ErrMessage & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI>" & Translate("This window is no longer required and can be closed at any time.",Login_Language,conn) & "</LI><BR><BR>"
              ErrMessage = ErrMessage & "<LI><FONT COLOR=""Red"">" & Translate("If the site you have requested does not appear",Login_Language,conn) & ", " & "<A HREF=""" & File_Redirect & """>" & Translate("click here",Login_Language,conn) & "</A>.</FONT><BR><BR></LI>"

              if not isblank(Thumbnail) then
                ErrMessage = ErrMessage & "</UL></TD></TR></TABLE>"
              end if

              Link_Name = Link_Name & "/" & Site_Code & "/Default.asp?Site_ID=" & parameter(xSite_ID) & "&Language=eng&NS=False&CID=" & Parameter(xCID) & "&SCID=" & Parameter(xSCID) & "&PCID=" & Parameter(xPCID) & "&CIN=" & Parameter(xCIN) & "&CINN=" & Parameter(xCINN)

              ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
              ErrMessage = ErrMessage & "<!--" & vbCrLf
              ErrMessage = ErrMessage & "wSite_Gateway = window.open('" & Link_Name & "','Site_Gateway_" & RandomFile & "');" & vbCrLf
              if instr(1,LCase(Link_Name),LCase(request.ServerVariables("SERVER_NAME"))) > 0 then
                ErrMessage = ErrMessage & "wSite_Gateway.focus();" & vbCrLf
              end if  
              ErrMessage = ErrMessage & "// -->" & vbCrLf
              ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf

              if Close_Window = False then
                ErrMessage = ErrMessage & vbCrLf
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf              
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
              else
                ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
                ErrMessage = ErrMessage & "<!--" & vbCrLf
                ErrMessage = ErrMessage & "window.blur();" & vbCrLf
                ErrMessage = ErrMessage & "window.close();" & vbCrLf
                ErrMessage = ErrMessage & "// -->" & vbCrLf
                ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf                
              end if

              Session("Asset_ID") = CInt(Parameter(xAsset_ID))
                            
              ErrType = 1
                           
              Call Status_Screen

            case else 

              ' Invalid Method
            
              Call Send_Invalid_SiteWide_Email
              
              ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but the link provided to you to access this information is not valid.",Login_Language,conn) & "</LI>"
              ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
              ErrMessage = ErrMessage & "<LI>Invalid Locator Format 3" & Parameter(xMethod) & "</LI>"
              ErrMessage = ErrMessage & "<LI>" & Mid(Locator,instr(1,Locator,"=") + 1)
  
              ErrType = 0
              
          end select
                      
        end if
        
      else
      
        ' Display message for invalid File Name or EMail to Site Administrator
        
        Call Send_Invalid_SiteWide_Email        

        ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but the link provided to you to access this information is not valid.",Login_Language,conn) & "</LI>"
        ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
        ErrMessage = ErrMessage & "<LI>Invalid Locator Format 2</LI>"
        ErrMessage = ErrMessage & "<LI>" & Mid(Locator,instr(1,Locator,"=") + 1)

        ErrType = 0
       
      end if
      
    else
      
      ' User Not Validated or improper querystring parameters

      ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but this link has expired.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>" & Translate("We expire links after a period of time to ensure that the user is getting the most up-to-date version of this item.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>" & Translate("Please visit the",Login_Language,conn) & " " & Translate(Site_Description,Login_Language,conn) & Translate(" - Extranet Support Site to get the latest version of this item.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>Invalid Locator Format 1</LI>"

      ErrType = 0
      
    end if
   
  else
  
    ' Expired Link Date

    ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but this link has expired.",Login_Language,conn) & "</LI>"
    ErrMessage = ErrMessage & "<LI>" & Translate("We expire links after a period of time to ensure that the user is getting the most up-to-date version of this item.",Login_Language,conn) & "</LI>"
    ErrMessage = ErrMessage & "<LI>" & Translate("Please visit the",Login_Language,conn) & " " & Translate(Site_Description,Login_Language,conn) & Translate(" - Extranet Support Site to get the latest version of this item.",Login_Language,conn) & "</LI>"
    ErrMessage = ErrMessage & "<LI>Invalid Locator Format 0</LI>"    

    ErrType = 0
    
  end if
  
end if

' --------------------------------------------------------------------------------------

if Script_Debug = True then

  response.end 

elseif not isblank(ErrMessage) then

  if ErrType = 1 then

    if Close_Window = True then
      ErrMessage = ErrMessage & vbCrLf
      ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
      ErrMessage = ErrMessage & "<!--" & vbCrLf
      ErrMessage = ErrMessage & "window.blur();" & vbCrLf
      ErrMessage = ErrMessage & "window.close();" & vbCrLf
      ErrMessage = ErrMessage & "// -->" & vbCrLf
      ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf
    end if    
  
  end if
  
  Call Status_Screen
      
elseif Script_Debug = False and Close_Window = True then

  response.write vbCrLf
  response.write "<SCRIPT LANGUAGE='JAVASCRIPT'>" & vbCrLf
  response.write "<!--" & vbCrLf
  response.write "window.blur();" & vbCrLf
  response.write "window.close();" & vbCrLf
  response.write "// -->" & vbCrLf
  response.write "</SCRIPT>" & vbCrLf

else  

  'response.redirect "http://" & Request("SERVER_NAME") & "/register/default.asp"    ' Redirect to a safe place

end if  

Call Disconnect_SiteWide

' --------------------------------------------------------------------------------------
' Functions or Subroutines
' --------------------------------------------------------------------------------------

sub Status_Screen

  Dim Top_Navigation        ' True / False
  Dim Side_Navigation       ' True / False
  Dim Screen_Title          ' Window Title
  Dim Bar_Title             ' Black Bar Title

  Site_ID_Save = Site_ID
  if Document_Site_ID = 99 then
    Site_ID = 99
  end if
    
  %>
  <!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
  <%

  if isblank(Site_Description) then
    Screen_Title    = Translate("Fluke",Alt_Language,conn) & " - " & Translate("Electronic Document / File Fulfillment Center",Alt_Language,conn)
    Bar_Title       = Translate("Fluke",Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "</FONT>"
  else
    Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Electronic Document / File Fulfillment Center",Alt_Language,conn)
    Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "</FONT>"
  end if
      
  Top_Navigation  = False 
  Side_Navigation = False
  Content_Width   = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%

  Site_ID = Site_ID_Save
  
  response.write "<FONT CLASS=Heading3>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn)
  response.write "</FONT>"
  response.write "<BR><BR>"

  response.write "<FONT CLASS=Medium>"

  response.write ErrMessage & vbCrLf
  
  ErrType    = 0
  ErrMessage = ""
  
  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%
 
end sub

' --------------------------------------------------------------------------------------

sub Send_Email_Header

  Set Mailer = Server.CreateObject("SMTPsvg.Mailer") 
 
  %>
  <!--#include virtual="/connections/connection_EMail_Timeout.asp"-->  
  <%
  
  Mailer.QMessage      = False
  
end sub

' --------------------------------------------------------------------------------------
 
sub Send_EMail_File

  Call Send_Email_Header
  
  Mailer.ReturnReceipt = False
  Mailer.Priority      = 1
  
  ' Get User Info
    
  SQL = "SELECT UserData.* FROM UserData WHERE ID=" & CInt(Parameter(xAccount_ID))
  Set rsUser = Server.CreateObject("ADODB.Recordset")
  rsUser.Open SQL, conn, 3, 3

  Mailer.AddRecipient   rsUser("FirstName") & " " & rsUser("LastName"), rsUser("EMail") 
'  Mailer.AddBCC         "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"

  rsUser.Close
  set rsUser = nothing
  
  ' Get Site Info  

  SQL = "SELECT Site.* FROM Site WHERE ID=" & CInt(Parameter(xSite_ID))
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3
  
  Mailer.FromName       = rsSite("FromName")
  Mailer.FromAddress    = rsSite("FromAddress")
  Mailer.ReplyTo        = rsSite("ReplyTo")
  
  rsSite.close
  set rsSite = nothing
  
  ' Get Language Info

  SQL = "SELECT Language.* FROM Language WHERE Language.Code='" & Login_Language & "'"
  Set rsLanguage = Server.CreateObject("ADODB.Recordset")
  rsLanguage.Open SQL, conn, 3, 3
  
  Mailer.CustomCharSet = rsLanguage("Name_CHARSET")
   
  rsLanguage.close
  set rsLanguage = nothing 
  
  ' EMail Body
  
  MailSubject     = Translate("Information Requested",Alt_Language,conn) & " - " & Translate("Electronic Document / File Fulfillment",Alt_Language,conn)
    
  MailMessage     = Translate("This is an automated notification message from the",Login_Language,conn) & " "
  MailMessage     = MailMessage & RestoreQuote(Translate(Site_Description,Login_Language,conn)) & " " & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "." & vbCrLf & vbCrLf
  MailMessage     = MailMessage & "Requestor's IP Address: " & request.ServerVariables("REMOTE_ADDR") & vbCrLf
  
  ' Asset Data
  
  SQL = "SELECT Calendar.* FROM Calendar WHERE ID=" & CInt(Parameter(xAsset_ID))
  Set rsAsset = Server.CreateObject("ADODB.Recordset")
  rsAsset.Open SQL, conn, 3, 3
   
  MailMessage = MailMessage & UCase(Translate("Product or Product Series",Login_Language,conn)) & ":" & vbCrLf & rsAsset("Product") & vbCrLf & vbCrLf
  MailMessage = MailMessage & UCase(Translate("Title",Login_Language,conn)) & ": " & vbCrLf & rsAsset("Title") & vbCrLf & vbCrLf

  ' Attach File
  
  Attach_File = Server.MapPath(Replace(File_Redirect,"http://" & Request("SERVER_NAME"),""))
  
  if Script_Debug then
    response.write "Server Path to File Attachment: " & Attach_File & "<BR>"
  end if
  
  Mailer.AddAttachment Attach_File
   
  if CInt(rsAsset("Confidential")) = True then
    if isdate(rsAsset("PEDate")) then
      if CDate(rsAsset("PEDate")) > CDate(Date) then
        MailMessage = MailMessage & vbCrLf & UCase(Translate("Restrictions",Login_Language,conn)) & ":" & vbCrLf
      end if  
    end if
    if CInt(rsAsset("Confidential")) = True then
      MailMessage = MailMessage & Translate("Confidential Information - Not for Public Release",Login_Language,conn) & "." & vbCrLf
    end if  
    if CDate(rsAsset("PEDate")) > CDate(Date) then
      MailMessage = MailMessage & Translate("Embargoed Information - Not for Public Release until",Login_Language,conn) & ": "
      MailMessage = MailMessage & Day(rsAsset("PEDate")) & " " & Translate(MonthName(Month(rsAsset("PEDate"))),Login_Language,conn) & ", " & Year(rsAsset("PEDate")) & vbCrLf
    end if  
  end if

  rsAsset.close
  set rsAsset = nothing
  
  MailMessage = MailMessage & vbCrLf & Translate("The document or file that you have requested is attached to this email.",Login_Language,conn) & vbCrLf & vbCrLf
  MailMessage = MailMessage & Translate("Sincerely",Login_Language,conn) & "," & vbCrLf & vbCrLf & RestoreQuote(Translate(Site_Description,Login_Language,conn)) & " - " & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn)
   
  Call Send_EMail
  
end sub

' --------------------------------------------------------------------------------------

sub Send_Invalid_SiteWide_Email

  Call Send_Email_Header
    
  Mailer.ReturnReceipt = False
  Mailer.Priority      = 1
  Mailer.AddBCC         "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"

  ' Get Site Info  

  SQL = "SELECT Site.* FROM Site WHERE ID=" & Parameter(xSite_ID)
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  Mailer.FromName       = rsSite("FromName")
  Mailer.FromAddress    = rsSite("FromAddress")
  Mailer.ReplyTo        = rsSite("ReplyTo")
  Mailer.AddRecipient     rsSite("FromName"), rsSite("FromAddress")
  
  Site_Description      = rsSite("Site_Description") 
  Site_Code             = rsSite("Site_Code")
  
  rsSite.close
  set rsSite = nothing

  MailSubject = "Invalid Document Number Requested"   
  MailMessage = "This is an automated notification message from the " & Site_Description & "." & vbCrLf & vbCrLf
  MailMessage = MailMessage & "The following locator number: " & Mid(Document,1,7) & " was not found." & vbCrLf
  MailMessage = MailMessage & "Site Wide Asset ID Number: " & Parameter(xAsset_ID) & vbCrLf & vbCrLf
  MailMessage = MailMessage & "Requestor's IP Address: " & request.ServerVariables("REMOTE_ADDR") & vbCrLf
  MailMessage = MailMessage & "Sincerely" & "," & vbCrLf & vbCrLf
  MailMessage = MailMessage & Site_Description & " at Support.Fluke.com"

  Call Send_EMail

end sub

' --------------------------------------------------------------------------------------

sub Send_Invalid_Document_Email

  Call Send_Email_Header
    
  Mailer.ReturnReceipt = False
  Mailer.Priority      = 1
' Mailer.AddBCC         "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"

  ' Get Site Info  

  SQL = "SELECT Site.* FROM Site WHERE ID=99"
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  Mailer.FromName       = rsSite("FromName")
  Mailer.FromAddress    = rsSite("FromAddress")
  Mailer.ReplyTo        = rsSite("ReplyTo")
  Mailer.AddRecipient     rsSite("FromName"), rsSite("FromAddress")
  Mailer.AddRecipient     "Webmaster", "Webmaster@Fluke.com"
  
  Site_Description      = rsSite("Site_Description") 
  Site_Code             = rsSite("Site_Code")
  
  rsSite.close
  set rsSite = nothing

  MailSubject = "Invalid Document Number Requested"   
  MailMessage = "This is an automated notification message from the " & Site_Description & "." & vbCrLf & vbCrLf
  MailMessage = MailMessage & "The following document number: " & Mid(Document,1,7) & " was not found." & vbCrLf & vbCrLf
  MailMessage = MailMessage & "Requestor's IP Address: " & request.ServerVariables("REMOTE_ADDR") & vbCrLf
  MailMessage = MailMessage & "Click for Domain Lookup: http://support.dev.fluke.com/sw-common/sw-dns_lookup.asp?ipAddress=" & request.ServerVariables("REMOTE_ADDR") & vbCrLf & vbCrLf
  MailMessage = MailMessage & "Sincerely" & "," & vbCrLf & vbCrLf
  MailMessage = MailMessage & Site_Description & " at Support.Fluke.com"
  
  Call Send_EMail

end sub

' --------------------------------------------------------------------------------------

sub Send_EMail

  Mailer.Subject  = MailSubject
  Mailer.BodyText = MailMessage

  if Mailer.SendMail then
'    ErrMessage = ErrMessage = "<LI>" & Translate("The document or file you have requested has been successfully sent to you by email.",Login_Language,conn) & "</LI>"
  else
    ErrType    = 0
    ErrMessage = ErrMessage & "<LI>" & Translate("Send Email Failure",Login_Language,conn) & ".<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & Mailer.Response & ". " & Translate("Send any errors noted to be reported to the Webmaster",Login_Language,conn) & ". </LI>"   
  end if   

end sub

' --------------------------------------------------------------------------------------
%>
