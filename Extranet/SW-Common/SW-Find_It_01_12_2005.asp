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
<!--#include virtual="/include/functions_table_border.asp"-->  
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

Dim Script_Debug, Email_Debug

Email_Debug = false   ' Set True to disable distribution emails except to Kelly Whitlock

Dim bShowPerms

Dim Login_Language
Dim Alt_Language

Dim Close_Window
Dim ErrMessage
Dim ErrType
Dim LogFlag

Dim Site_ID
Dim Site_Code
Dim Site_Description
Dim Status_Comment
Dim Submitted_By
Submitted_By = ""

Dim Document_Site_ID, Style_Site_ID
Document_Site_ID = 0

Dim Calendar_Title, Oracle_Title
Calendar_Title = ""
Oracle_Title   = ""

Dim File_Name
Dim Link_Name
Dim Thumbnail

Dim Mailer
Dim Verify

Dim SQL
Dim rsOwner

Dim Activity_Log
Activity_Log = True

'if not isblank(Session("Language")) then
'  Login_Language = Session("Language")
'  Alt_Language   = "eng"
'else
  Login_Language = "eng"
  Alt_Language   = "eng"
'end if
  
ErrMessage     = ""
ErrType        = 0

bShowPerms     = True           ' bShowPerms allows the hidden link to subgroup display

Dim Style(3)
Style(0) = "<FONT STYLE=""font-size:10pt;font-weight:Bold;color:Black;background:#FFFF99;text-decoration:none;font-family:Arial;"">"
Style(1) = "<FONT STYLE=""font-size:10pt;font-weight:Normal;color:Black;background:#FFFFFF;text-decoration:none;font-family:Arial;"">"
Style(2) = "<FONT STYLE=""font-size:10pt;font-weight:Normal;color:Black;background:#FF99CC;text-decoration:none;font-family:Arial;"">"

if LCase(request("Debug")) = "on" then
  Script_Debug = True
else
  Script_Debug = False
end if

Dim Locator
Locator = ""
Dim Invalid_Locator
Invalid_Locator = false
Dim Document
Document = ""

Close_Window   = False

if not isblank(Request("SW-Locator")) then
  Locator    = replace(Request("SW-Locator")," ","")
  Close_Window = False
elseif not isblank(Request("Locator")) then
  Locator    = replace(Request("Locator")," ","")
end if

Dim Referer
if not isblank(Request("Referer")) then
  Referer    = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Request("Referer"),"%3A",":"),"%2F","/"),"%2E","."),"%5F","_"),"%3F","?"),"%3D","="),"%26","&"),"%20","+")
else
  Referer    = ""
end if

if not isblank(Request("Document")) then
  Document   = replace(Request("Document")," ","")
  Verify     = Request("Verify")
end if

' Source of request such as:  EFF=Electronic File Fulfillment (Dennis Sims), FDL=Fluke Digital Library (www.fluke.com)

Dim SRC
if not isblank(Request("SRC")) then
  SRC = UCase(Request("SRC"))
  select case SRC
    case "EFF"
      Document_Site_ID = 99
    case else
      Document_Site_ID = 98
  end select
else
  SRC = ""
  Document_Site_ID = 98
end if  

if not isblank(Request("Style")) then
  Style_Site_ID   = replace(Request("Style")," ","")
end if  

' Check for Invalid Document or Locator Value (0-9) attempt to filter out invalid characters.

New_Document = ""
for x = 1 to len(Document)
  select case asc(mid(UCase(document),x,1))
    case 48,49,50,51,52,53,54,55,56,57
    New_Document = New_Document & mid(Document,x,1)
    case else
      ErrMessage = "<LI>Syntax Error: An Invalid Character was found in the Document or Locator parameter supplied by the application sending the request to Find_IT at character position: " & x & " Key |Document=|" & Document & "|.  An attempt was made to filter out these invalid character(s) in an attempt to continue with the fulfillment request.  The information related to the asset below may not accurately reflect the original Document or Locator parameter.</LI>"
  end select  
next

Document = New_Document

set Session("Asset_ID") = nothing

Randomize()
RandomFile = Int(1000 * Rnd())

' Parameters passed from www.fluke.com to record site/category activity

Dim CMS_Site, CMS_Path, CMS_ID

CMS_Site = Request("CMS_Site")
CMS_Path = Request("CMS_Path")
if not isblank(CMS_SITE) or not isblank(CMS_Path) then
  SRC = "WWW"
end if

' --------------------------------------------------------------------------------------
' Method Key
' --------------------------------------------------------------------------------------

Dim xOLView, xOLDownLoad, xOLSend, xSSView, xSSDownload, xSSSend, xOLSendIt
Dim xOLLink, xOLLinkNoPop, xOLGateway, xOLGatewayNoPop
Dim xOLViewPOD, xOLDownLoadPOD, xOLSentIt, xOLSentItNoZip

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
xOLSendItNoZip  = 13 ' On-Line Send to an Associate
xOLSendIt       = 14 ' On-Line Send to an Associate Non-ZIP File Version

' --------------------------------------------------------------------------------------
' Document - This just converts a Document ID (7-Digit Oracle Item Number) to the format
'            Used by Locator with presets.
' --------------------------------------------------------------------------------------

if not isblank(Document) then

  if isnumeric(Document) then   ' Oracle Item Numbers

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
            ErrMessage = ErrMessage & "<LI>Method - Invalid Document ID Locator (Oracle) reset to xOLView</LI>"
            Method = xOLView
            Invalid_Locator = true
        end select   
      case else
        ErrMessage = ErrMessage & "<LI>Length - Invalid Document ID Locator (Oracle)</LI>"
        Invalid_Locator = true
    end select
    
  else                          ' European Item Numbers

    bShowPerms = False
    
    select case len(Document)
      case 9
        Asset  = Document
        Method = xOLView
      case 10
        Asset  = Mid(Document,1,9)
        Method = Mid(Document,10,1)
        
        select case Method
          case 0
            Method = xOLViewPOD
          case 1  
            Method = xOLDownLoadPOD
          case else
            ErrMessage = ErrMessage & "<LI>Method - Invalid Document ID Locator (European) Reset to xOLView</LI>"
            Method = xOLView
            Invalid_Locator = true
        end select   
      case else
        ErrMessage = ErrMessage & "<LI>Length - Invalid Document ID Locator (European)</LI>"
        Invalid_Locator = true
    end select
  
  end if
  
  if isblank(ErrMessage) then

    SQL = "SELECT * " &_
          "FROM   dbo.Calendar " &_
          "WHERE (Item_Number = '" & Asset & "') " &_
          "       AND (File_Name IS NOT NULL) " &_
          "       AND (SubGroups LIKE '%view%' OR SubGroups LIKE '%fedl%') " &_
          "ORDER  BY Revision_Code DESC, UDate DESC"
          
    Set rsAsset = Server.CreateObject("ADODB.Recordset")
    rsAsset.Open SQL, conn, 3, 3
 
    ' response.write SQL
    ' response.flush
    ' response.end
 
    if not rsAsset.EOF then
  
      Locator = CStr(rsAsset("Site_ID"))                 & "o" & _
                "1"                                      & "o" & _
                CStr(rsAsset("ID"))                      & "o" & _
                CStr(Method)                             & "o" & _
                CStr(Encode_Key(CStr(rsAsset("Site_ID")),"1",rsAsset("ID"))) & "o" & _             
                CStr(CLng(Date))                         & "o" & _
                "0"
       
      if rsAsset("Status") = 2 then

        Status_Comment = rsAsset("Status_Comment")
        Submitted_By   = CStr(rsAsset("Submitted_By"))
        
        if LCase(Verify) <> "on" then
          if Script_Debug = false then
            Call Send_Retired_Document_Email
          end if  
        end if  

        ErrMessage = "<LI>We are sorry but the Document that you have requested is not available.</LI>"
        ErrMessage = ErrMessage & "<LI>The Document Number you had requested was: " & Mid(Document,1,9) & "</LI>"

      end if  
                
    else
      if LCase(Verify) <> "on" then
        if Script_Debug = false then
          Status_Comment = "No Asset Container exists for this Item Number."
          Call Send_Invalid_Document_Email
        end if  
      end if  

      ErrMessage = "<LI>We are sorry but the Document that you have requested is not available.</LI>"
      ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
      ErrMessage = ErrMessage & "<LI>The Document Number you had requested was: " & Mid(Document,1,9) & "</LI>"

    end if
  
    rsAsset.close
    set rsAsset = nothing

  else
  
    if LCase(Verify) <> "on" then
      if Script_Debug = false then
        Status_Comment = ErrMessage
        Call Send_Invalid_Document_Email
      end if  
    end if  

    if not Script_Debug then
      ErrMessage = "<LI>We are sorry but the Document that you have requested is not available.</LI>"
      ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
      ErrMessage = ErrMessage & "<LI>The Document Number you had requested was: " & Mid(Document,1,9) & "</LI>"
    end if  

  end if  
  
end if

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

if (not isblank(Locator) and isblank(ErrMessage)) or Script_Debug = true then
  
  Dim ErrString
  Dim MailMessage
  Dim MailSubject 
  
  %>
  <!--#include virtual="/include/functions_locator.asp"-->
  <%
    
  if Script_Debug then
    
    Dim Top_Navigation  ' True / False
    Dim Side_Navigation ' True / False
    Dim Screen_Title    ' Window Title
    Dim Bar_Title       ' Black Bar Title
    Dim Content_width	  ' Percent
      
    Screen_Title    = "Electronic Document Fulfillment - Locator and Parameter Decode Utility"
    Bar_Title       = "Electronic Document Fulfillment" &_
                      "<BR><SPAN CLASS=MediumBoldGold>" & _
                      "Locator and Parameter Decode Utility</SPAN>"
    Top_Navigation  = False
    Side_Navigation = True
    Content_Width   = 95
  
    %>
    <!--#include virtual="/SW-Common/SW-Header.asp"-->
    <!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
    <%
      
    SQLDebug = "SELECT dbo.Calendar.*, dbo.Site.Site_Description AS Site_Description, dbo.Site.Site_Code AS Site_Code, dbo.Calendar_Category.Title AS Category,  " &_
               "       dbo.UserData.FirstName AS FirstName, dbo.UserData.LastName AS LastName, dbo.UserData.Email AS Email,  " &_
               "       dbo.UserData.Business_Phone AS Phone " &_
               "FROM   dbo.Calendar LEFT OUTER JOIN " &_
               "       dbo.UserData ON dbo.Calendar.Submitted_By = dbo.UserData.ID LEFT OUTER JOIN " &_
               "       dbo.Calendar_Category ON dbo.Calendar.Category_ID = dbo.Calendar_Category.ID LEFT OUTER JOIN " &_
               "       dbo.Site ON dbo.Calendar.Site_ID = dbo.Site.ID "
               
    if Parameter(xAsset_ID) = 0 then
      SQLDebug = SQLDebug & "WHERE  (dbo.Calendar.Item_Number = " & Asset & ") AND dbo.Calendar.Status <> 2 "
    else
      SQLDebug = SQLDebug & "WHERE  dbo.Calendar.ID = " & Parameter(xAsset_ID) & " "
      'SQLDebug = SQLDebug & " AND dbo.Calendar.Status <> 2 "
    end if

    SQLDebug = SQLDebug & "ORDER BY dbo.Calendar.Revision_Code DESC, dbo.Calendar.UDate DESC"

    'response.write SQLDebug & "<P>"
    'response.flush
    'response.end
    
    response.write "<FORM ACTION=""" & request.ServerVariables("Script_Name") & """   METHOD=""Post"">" & vbCrLf
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Debug"" VALUE=""on"">" & vbCrLf
    call Nav_Border_Begin
    response.write "<TABLE BORDER=0 BGCOLOR=""White""><TR>" & vbCrLf
    response.write "<TD CLASS=SmallBold>Find Literature Item Number:&nbsp;</TD>" & vbCrLf
    response.write "<TD CLASS=Small><INPUT TYPE=""TEXT"" NAME=""Document"" WIDTH=""20"" MAXLENGTH=""7"" CLASS=Small VALUE=""" & Document & """></TD>" & vbCrLf
    response.write "<TD CLASS=Navlefthighlight1><INPUT TYPE=""Submit"" NAME=""Submit"" VALUE=""&nbsp;Go&nbsp;""></TD>" & vbCrLf
    response.write "</TR></TABLE>" & vbCrLf
    call Nav_Border_End
    response.write "</FORM>" & vbCrLf
    
    if not isblank(ErrMessage) then
      if Script_Debug then
        call Nav_Border_Begin
        response.write "<TABLE BGCOLOR=""WHITE"" Border=0><TR><TD CLASS=Small>"
        response.write "<SPAN CLASS=SmallBold>Message displayed to the requester of this Asset:</SPAN><P>"
      end if
      response.write ErrMessage
      if Script_Debug then
        response.write "<BR>&nbsp;</TD></TR></TABLE>"
        call Nav_Border_End
      end if
    response.write "<P>"
    end if  
                       
    if not isblank(Document) then
    
      Set rsDebug = Server.CreateObject("ADODB.Recordset")
      rsDebug.Open SQLDebug, conn, 3, 3
     
      Table_Flag = false
  
      Flag_PDF  = false
      File_PDF  = 0
      Flag_POD  = false
      Item_Rev  = 0
   
      if not rsDebug.EOF then
        response.write "<SPAN CLASS=SMALLBOLD>"
        response.write "Asset Record Details on Support.Fluke.com"
        response.write "</SPAN><BR>"
         
        if not rsDebug.EOF then
          
          Last_Rev  = ""
          Last_PDF  = ""
          Last_POD  = ""
      
          do while not rsDebug.EOF
          
            if not isblank(rsDebug("Revision_Code")) then
              if Last_Rev <> LCase(rsDebug("Revision_Code")) then
                Last_Rev = LCase(rsDebug("Revision_Code"))
                Item_Rev = Item_Rev + 1
              end if  
            end if
            if not isblank(rsDebug("File_Name")) then
              if Last_PDF <> LCase(rsDebug("File_Name")) then
                Last_PDF = LCase(rsDebug("File_Name"))
                File_PDF = File_PDF + 1
              end if  
            end if
            if not isblank(rsDebug("File_Name_POD")) then
              File_POD = File_POD + 1
            end if
              
            rsDebug.MoveNext
              
          loop
  
          if Item_Rev > 1 or File_PDF > 1 or File_POD > 1 then
            response.write "<P><SPAN CLASS=SmallBold>Please Check the following Assets for these problems:</SPAN><BR><SPAN CLASS=SmallRed>"
            if Item_REV > 1 then
              response.write "<LI>Item Revision Code - There are multiple assets with the same Item Number, but different Revision Codes.</LI>"
            end if
            if File_PDF > 1 then
              response.write "<LI>Asset (Low Resolution) PDF - There are different file names associated with this Item Number.</LI>"
            end if
            if File_POD > 1 then
              response.write "<LI>Asset (High Resolution) POD - There are multiple POD file names associated with this Item Number.  There should be only 1 occurrence.</LI>"
            end if
            response.write "</SPAN><P>"
          end if  
              
        end if  
    
        rsDebug.MoveFirst        
    
        Table_Flag = true
        call Nav_Border_Begin
        response.write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 BGCOLOR=""White"">"
      else  
        response.write "<SPAN CLASS=SMALLBOLDRED>"
        response.write "There are no active asset containers availble for ID or Item Number: " & Asset
        response.write "</SPAN><P>"
      end if
        
      do while not rsDebug.EOF
    
        response.write "<TR><TD CLASS=SMALL>Asset ID</TD><TD CLASS=SMALL BGCOLOR="
          
        Edit_Content = "/sw-administrator/Calendar_Edit.asp?ID=" & rsDebug("ID") & "&Site_ID=" & rsDebug("Site_ID")
        select case rsDebug("Status")
          case 0
            response.write """Yellow""><B>"
            response.write "<A HREF=""javascript:void(0);"" onclick=""var MyPop1 = window.open('" & Edit_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""" & Translate("Click to Edit this Item",Login_Language,conn) & """>" & rsDebug("ID") & "</A>"    
            response.write " - REVIEW</B>"
          case 1
            response.write """#00CC00""><B>"
            response.write "<A HREF=""javascript:void(0);"" onclick=""var MyPop1 = window.open('" & Edit_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""" & Translate("Click to Edit this Item",Login_Language,conn) & """>" & rsDebug("ID") & "</A>"    
            response.write " - LIVE</B>"
          case 2
            response.write """#AAAACC""><B>"
            response.write "<A HREF=""javascript:void(0);"" onclick=""var MyPop1 = window.open('" & Edit_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""" & Translate("Click to Edit this Item",Login_Language,conn) & """>" & rsDebug("ID") & "</A>"    
            response.write " - ARCHIVED</B>"
        end select
        if rsDebug("Campaign") > 0 then
          response.write "&nbsp;&nbsp;&nbsp;Associated with PI/C Container "
          Edit_Content = "/sw-administrator/Calendar_Edit.asp?ID=" & rsDebug("Campaign") & "&Site_ID=" & rsDebug("Site_ID")
          response.write "<A HREF=""javascript:void(0);"" onclick=""var MyPop1 = window.open('" & Edit_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""" & Translate("Click to Edit this Item",Login_Language,conn) & """>" & rsDebug("Campaign") & "</A>"            
        else
          response.write "&nbsp;&nbsp;&nbsp;Individual"
        end if
        response.write "</TD></TR>" & vbCrLf
    
        response.write "<TR><TD CLASS=SMALL>Item Number</TD><TD CLASS=SMALLBold>" & rsDebug("Item_Number") & "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL>Revision</TD><TD CLASS=SMALLBOLD>" & rsDebug("Revision_Code")
        if CInt(rsDebug("Status")) = 2 then
          SQLRev = "SELECT Revision FROM Literature_Items_US WHERE Active_Flag = -1 AND Item=" & rsDebug("Item_Number") & " ORDER BY Revision DESC"
          Set rsRev = Server.CreateObject("ADODB.Recordset")
          rsRev.Open SQLRev, conn, 3, 3
          if not rsRev.EOF then
            response.write "&nbsp;&nbsp;&nbsp;"
            if UCase(rsRev("Revision")) <> UCase(rsDebug("Revision_Code")) then
              response.write "<SPAN CLASS=SmallRed>(Oracle Revision: " & rsRev("Revision") & ")</SPAN>"
            else
              response.write "(Oracle Revision: " & rsRev("Revision") & ")"
            end if
          end if
          rsRev.close
          set reRev = nothing
        end if
        response.write "</TD></TR>" & vbCrLf
  
        if not isblank(rsDebug("Status_Comment")) then
        response.write "<TR><TD CLASS=SMALL>Status Notes:</TD><TD CLASS=SMALLBOLD>" & rsDebug("Status_Comment") & "</TD></TR>" & vbCrLf
        end if
        response.write "<TR><TD CLASS=SMALL NOWRAP>Enabled for EEF</TD><TD CLASS=SMALL BGCOLOR="
        if instr(1,LCase(rsDebug("SubGroups")),"view") > 0 then
          response.write """#00CC00""><B>YES</B>"
        else
          response.write """Yellow"">NO"
        end if          
        response.write "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL NOWRAP>Enabled for FDL</TD><TD CLASS=SMALL BGCOLOR="
        if instr(1,LCase(rsDebug("SubGroups")),"fedl") > 0 then
          response.write """#00CC00""><B>YES</B>"
        else
          response.write """Yellow"">NO"
        end if          
        response.write "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL NOWRAP>Enabled for Shopping Cart</TD><TD CLASS=SMALL BGCOLOR="
        if instr(1,LCase(rsDebug("SubGroups")),"shpcrt") = 0 then
          response.write """#00CC00""><B>YES</B>"
        else
          response.write """YELLOW"">NO"
        end if          
        response.write "</TD></TR>" & vbCrLf
        
        response.write "<TR><TD CLASS=SMALL>Title</TD><TD CLASS=SMALL>" & rsDebug("Title") & "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL>Library Category</TD><TD CLASS=SMALL>" & rsDebug("Category") & "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL>Last Update</TD><TD CLASS=SMALL>" & rsDebug("UDate") & "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL>Last Update By</TD><TD CLASS=SMALL>" & rsDebug("FirstName") & " " & rsDebug("LastName") & ", Phone: " & rsDebug("Phone") & ", Email: " & rsDebug("Email") & "</TD></TR>" & vbCrLf            
        response.write "<TR><TD CLASS=SMALL>Site ID</TD><TD CLASS=SMALL>" & rsDebug("Site_ID") & "</TD></TR>" & vbCrLf      
        response.write "<TR><TD CLASS=SMALL>Site Name</TD><TD CLASS=SMALL>" & rsDebug("Site_Description") & "</TD></TR>" & vbCrLf
        response.write "<TR><TD CLASS=SMALL>Asset_File</TD><TD CLASS=SMALL>"
        if not isblank(rsDebug("File_Name")) then
          response.write "http://Support.Fluke.com/" & rsDebug("Site_Code") & "/" & rsDebug("File_Name")
        else
          response.write "&nbsp;"
        end if
        response.write "</TD></TR>" & vbCrLf
          
        response.write "<TR><TD CLASS=SMALL>POD_File</TD><TD CLASS=SMALL"
        if not isblank(rsDebug("File_Name_POD")) then
          response.write " BGCOLOR=""#FF6633"">"
          response.write "http://Support.Fluke.com/" & rsDebug("Site_Code") & "/" & rsDebug("File_Name_POD")
        else
          response.write ">&nbsp;"
        end if
        response.write "</TD></TR>"
                
        response.write "<TR><TD CLASS=SMALL>Thumbnail File</TD><TD CLASS=SMALL>"
        if not isblank(rsDebug("Thumbnail")) then
          response.write "http://Support.Fluke.com/" & rsDebug("Site_Code") & "/" & rsDebug("Thumbnail")
        else
          response.write "&nbsp;"
        end if
        response.write "</TD></TR>"
          
        response.write "<TR><TD CLASS=SMALL>Usage Notes:</TD><TD CLASS=SMALL"
        Marker = ">&nbsp;"
        if Flag_PDF = false and not isblank(rsDebug("File_Name")) and (instr(1,LCase(rsDebug("SubGroups")),"view") > 0 or instr(1,LCase(rsDebug("SubGroups")),"fedl") > 0) then
          response.write " BGCOLOR=""Yellow"">This file associated with this asset is used as the (Low Resolution) PDF file for Electronic Document Fulfillment (EEF/Digital Library) for this Item Number"
          Flag_PDF = True
          Flag_SME = True
          Marker = ""
        end if
        if Flag_POD = false and not isblank(rsDebug("File_Name")) and Flag_SME = True and not isblank(rsDebug("File_Name_POD")) and instr(1,LCase(rsDebug("SubGroups")),"view") > 0 then
         response.write " and for Print-On-Demand Fulfillment (POD)."
          Flag_POD = True
          Marker = ""
        elseif Flag_POD = false and not isblank(rsDebug("File_Name_POD")) and instr(1,LCase(rsDebug("SubGroups")),"view") > 0 then
          response.write " BGCOLOR=""Yellow"">This file associated with asset is used as the (High Resolution) POD file for Print-On-Demand (POD) Fulfillment."
          Flag_POD = True        
          Marker = ""
        end if
        response.write Marker
        response.write "</TD></TR>" & vbCrLf
          
        Flag_SME = False
    
        rsDebug.MoveNext
          
        if not rsDebug.EOF then
          response.write "<TR><TD COLSPAN=2 CLASS=SMALL BGCOLOR=""#CCCCCC"">&nbsp;</TD></TR>" & vbCrLf      
        end if  
        
      loop
        
      rsDebug.close
      set rsDebug  = nothing
      set SQLDebug = nothing
   
      if Table_Flag then
        response.write "</TABLE>" & vbCrLf & vbCrLf
        call Nav_Border_End
      end if  
        
      response.write "<P>"
      
      response.write "<SPAN CLASS=SMALLBOLD>"
      response.write "Internal Find_It Parameters (For Webmaster Use Only)"
      response.write "</SPAN><BR>"
        
      call Nav_Border_Begin
      response.write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 BGCOLOR=""White"">" & vbCrLf
      response.write "<TR>" & vbCrLf
      response.write "<TD CLASS=Small>Locator:</TD><TD CLASS=SMALL COLSPAN=3>" & Replace(Locator,"O","<FONT COLOR=""Red"">O</FONT>") & "</TD></TR>" & vbCrLf
      response.write "<TD CLASS=Small>Document:</TD><TD CLASS=SMALL COLSPAN=3>" & Document & "</TD></TR>" & vbCrLf
      response.write "<TD CLASS=Small>Today's Serial Date:</TD><TD CLASS=SMALL COLSPAN=3>" & CLng(Date) & " " & CDate(CLng(Date)) & "</TD></TR>" & vbCrLf
      response.write "<TD CLASS=Small>Encoded Key:</TD><TD CLASS=SMALL COLSPAN=3>" & Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID)) & "</TD></TR>" & vbCrLf
      response.write "<TD CLASS=Small>Decoded Key:</TD><TD CLASS=SMALL COLSPAN=3>" & Decode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID),Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID))) & "</TD></TR>" & vbCrLf
   
      for i = 0 to Parameter_Max
        response.write "<TR>" & vbCrLf
        response.write "<TD CLASS=SMALL>" & i & "</TD>" & vbCrLf
        response.write "<TD CLASS=SMALL WIDTH=""1%"">" & Parameter_Key(i) & ":</TD>" & vbCrLf
        response.write "<TD CLASS=SMALL WIDTH=""1%"">" & Parameter(i) & "</TD>" & vbCrLf
        response.write "<TD CLASS=SMALL>&nbsp;</TD>" & vbCrLf      
        response.write "</TR>" & vbCrLf & vbCrLf      
      next
        
      with response
        .write "<TD CLASS=Small>Short Format Test URL:</TD><TD CLASS=SMALL COLSPAN=3>"    & "http://" & request.ServerVariables("SERVER_NAME") & "/Find_It.asp?Locator="
        .write CInt(Parameter(xSite_ID))    & "o"
        .write CInt(Parameter(xAccount_ID)) & "o"
        .write CInt(Parameter(xAsset_ID))   & "o"
        .write CInt(Parameter(xMethod))     & "o"
        .write CInt(Encode_Key(Parameter(xSite_ID),Parameter(xAccount_ID),Parameter(xAsset_ID))) & "o"
        .write CLng(Date)                   & "o"
        .write "0"
        .write "&Debug=on</TD></TR>" & vbCrLf
      end with
    
      with response
        .write "<TD CLASS=Small>Long&nbsp;&nbsp;Format Test URL:</TD><TD CLASS=SMALL COLSPAN=3>" & "http://" & request.ServerVariables("SERVER_NAME") & "/Find_It.asp?Locator="
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
        .write "&Debug=on</TD></TR>" & vbCrLf
      end with
    
      response.write "<TD CLASS=Small>Expiration Date Match:<TD CLASS=SMALL COLSPAN=3>" & CDate(Parameter(xExpiration_Date)) & " " & CDate(Date)
      if CDate(Parameter(xExpiration_Date)) >= CDate(Date) then response.write " True"
      response.write "</TD></TR>" & vbCrLf
      response.write "</TABLE>"
      call Nav_Border_End
  
    end if

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

      if isblank(SRC) then
        SRC = "PPS"
      end if

    else                  ' Bypass check for DOCUMENT
      if isblank(SRC) then
        SRC = "WWW"
      end if
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
  
      if Document_Site_ID > 90 then
      
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
  
          case xOLView, xOLSendNoZip, xSSView, xOLSendItNoZip
  
            if not isblank(rsAsset("File_Name")) then
              File_Name = Trim(rsAsset("File_Name"))
            end if
  
          case xOLDownLoad, xOLSend, xSSDownload, xSSSend, xOLSentIt
  
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
      
      if not isblank(File_Name) or not isblank(Link_Name) then
        
        if not isblank(Thumbnail) and instr(1,LCase(Thumbnail),LCase(request.ServerVariables("SERVER_NAME"))) = 0 then
          Thumbnail = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(Site_Code) & "/" &  Thumbnail
        end if  
        
        if not isblank(File_Name) then
          
          ' Needs code added here using CreateObject("Scripting.FileSystemObject") to ensure that file has not been deleted
          
          File_Redirect = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(Site_Code) & "/" & File_Name
         
        elseif not isblank(Link_Name) then
  
          File_Redirect = Link_Name
  
        end if        
    
        ' --------------------------------------------------------------------------------------
        ' Log Activity of Download to Activity Table
        ' --------------------------------------------------------------------------------------
          
        if CInt(Activity_Log) = CInt(True) then   ' Do not Log Fluke Entity Users Access to Assets
  
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
                if Script_Debug = false then
                  conn.Execute(CMS_SQL)
                end if  
                set CMS_SQL = nothing
              end if
                
            else
              CMS_ID = 0
            end if  
            
          end if      
                  
          ' Update User's Last Logon Date/Time since User is Accessing an Asset at the Site.
            
          if Parameter(xAccount_ID) > 1 and Script_Debug = false then
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
            response.write "<P>"
            response.write "<SPAN CLASS=SmallBold>Activity Log Information (For Webmaster Use Only)</SPAN>"
            response.write "<BR>"
            call Nav_Border_Begin
            response.write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 BGCOLOR=""White"">" & vbCrLf
            response.write "<TR>" & vbCrLf
            response.write "<TD CLASS=SMALL NOWRAP>Activity SQL:</TD><TD CLASS=SMALL COLSPAN=3>" & ActivitySQL & "</TD></TR>"
            response.write "<TD CLASS=SMALL>File Path:</TD><TD CLASS=SMALL COLSPAN=3>" & File_Redirect & "</TD></TR>"
            response.write "</TABLE>"
            call Nav_Border_End
          end if
  
          if Script_Debug = false then
            conn.Execute (ActivitySQL)
          end if  
            
        else
          if Script_Debug then
            call Nav_Border_Begin
            response.write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 BGCOLOR=""White"">" & vbCrLf
            response.write "<TR>" & vbCrLf
            response.write "<TD CLASS=SMALL>Activity SQL:</TD><TD CLASS=SMALL COLSPAN=3>&nbsp;</TD></TR>"
            response.write "<TD CLASS=SMALL>File Path:</TD><TD CLASS=SMALL COLSPAN=3>&nbsp;</TD></TR>"
            response.write "</TABLE>"
            call Nav_Border_End
          end if  
        end if
      end if

      ' Send User the File Requested         
       
      if Script_Debug = False then
      
        if not isblank(Thumbnail) and instr(1,LCase(Thumbnail),LCase(request.ServerVariables("SERVER_NAME"))) = 0 then
          Thumbnail = "http://" & request.ServerVariables("SERVER_NAME") & "/" & LCase(Site_Code) & "/" &  Thumbnail
        end if  
    
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
              
          ' Send File to the User
          
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
              
          ' Send File to an Associate
          
          case xOLSendIt, xOLSendItNoZip
                                 
            Session("Asset_ID") = CInt(Parameter(xAsset_ID))
              
            if not isblank(Thumbnail) then
              ErrMessage = ErrMessage & vbCrLf
              ErrMessage = ErrMessage & "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=0 BORDER=0>" & vbCrLf
              ErrMessage = ErrMessage & "<TR>" & vbCrLf
              ErrMessage = ErrMessage & "<TD WIDTH=80>" & vbCrLf
              ErrMessage = ErrMessage & "<IMG SRC=""" & Thumbnail & """ WIDTH=80 BORDER=1>" & vbCrLf
              ErrMessage = ErrMessage & "</TD>" & vbCrLf
              ErrMessage = ErrMessage & "<TD CLASS=Medium VALIGN=TOP>" & vbCrLf
            end if

            ErrMessage = ErrMessage & "<FORM NAME=""Send_It"" ACTION=""/SW-Common/SW-Send_It_Associate.asp"" METHOD=""POST"">" & vbCrLf              
            ErrMessage = ErrMessage & "<INPUT TYPE=""HIDDEN"" NAME=""SendItAccountID"" VALUE=""" & Parameter(xAccount_ID) & """>" & vbCrLf
            ErrMessage = ErrMessage & "<INPUT TYPE=""HIDDEN"" NAME=""SendItMethod"" VALUE=""" & Parameter(xMethod) & """>" & vbCrLf
            ErrMessage = ErrMessage & "<INPUT TYPE=""HIDDEN"" NAME=""SendItAssetID"" VALUE=""" & Parameter(xAsset_ID) & """>" & vbCrLf            
            ErrMessage = ErrMessage & "<INPUT TYPE=""HIDDEN"" NAME=""SendItLanguage"" VALUE=""" & Login_Language & """>" & vbCrLf                        
            ErrMessage = ErrMessage & "<INPUT TYPE=""HIDDEN"" NAME=""SendItSiteCode"" VALUE=""" & Site_Code & """>" & vbCrLf                        
            ErrMessage = ErrMessage & "<INPUT TYPE=""HIDDEN"" NAME=""SendItSiteID"" VALUE=""" & Site_ID & """>" & vbCrLf                                    
            ErrMessage = ErrMessage & "<UL>" & vbCrLf
            ErrMessage = ErrMessage & "<LI>" & Translate("The document or file that you have requested is ready to be sent by email to the person you will note below.",Login_Language,conn) & " " 
            ErrMessage = ErrMessage & Translate("Please enter their Name and Email Address, then click on the [Send Document] button.",Login_Language,conn) & "&nbsp;&nbsp;<SPAN CLASS=NORMALRED>*</SPAN> "
            ErrMessage = ErrMessage & Translate("Required Field",Login_Language,conn) & "</LI><P>"
            ErrMessage = ErrMessage & "<LI>" & Translate("Title",Login_Language,conn) & ": <SPAN CLASS=MediumRed>" & Title
            if not isblank(Category) then
              ErrMessage = ErrMessage & " - " & Translate(Category,Login_Language,conn)
            end if
            ErrMessage = ErrMessage & "</SPAN>&nbsp;&nbsp;" & Translate("File Size",Login_Language,conn) & ": "             
            select case CInt(Parameter(xMethod))
              case xOLSendIt
                ErrMessage = ErrMessage & Archive_Size
              case xOLSendItNoZip
                ErrMessage = ErrMessage & File_Size              
            end select
            ErrMessage = ErrMessage & " KBytes</LI>"

            ErrMessage = ErrMessage & "</UL>"
            ErrMessage = ErrMessage & "<TABLE BORDER=0 WIDTH=""90%"" ALIGN=""RIGHT"">" & vbCrLf
            ErrMessage = ErrMessage & "<TR><TD CLASS=SMALL WIDTH=""1%"" VALIGN=TOP>" & Translate("Send to Name",Login_Language,conn) & " :</TD><TD CLASS=SMALL VALIGN=TOP><INPUT SIZE=38 CLASS=SMALL TYPE=""TEXT"" NAME=""SendItName"" VALUE=""""></TD></TR>" & vbCrLf
            ErrMessage = ErrMessage & "<TR><TD CLASS=SMALL WIDTH=""1%"" VALIGN=TOP>" & Translate("Send to Email",Login_Language,conn) & "<SPAN CLASS=SmallRED>*</SPAN>" & ":</TD><TD CLASS=SMALL VALIGN=TOP><INPUT SIZE=38 CLASS=SMALL TYPE=""TEXT"" NAME=""SendItEmail"" VALUE=""""></TD></TR>" & vbCrLf
            ErrMessage = ErrMessage & "<TR><TD CLASS=SMALL WIDTH=""1%"" VALIGN=TOP>" & Translate("Select Method ",Login_Language,conn) & "<SPAN CLASS=NORMALRED>*</SPAN>" & ":</TD><TD CLASS=SMALL VALIGN=TOP>" & vbCrLf
            ErrMessage = ErrMessage & "<SELECT NAME=""SendItHow"" CLASS=Small>" & vbCrLf
            ErrMessage = Errmessage & "<OPTION CLASS=Region1NavSmall VALUE=0 SELECTED>" & Translate("Send Document/File as an Attachment",Login_Language,conn) & "</OPTION>" & vbCrLf
            ErrMessage = Errmessage & "<OPTION CLASS=Region2NavSmall VALUE=1>" & Translate("Send Document/File as a Website Link",Login_Language,conn) & "</OPTION>" & vbCrLf
            ErrMessage = ErrMessage & "</SELECT>" & vbCrLf            
            ErrMessage = ErrMessage & "</TD></TR>" & vbCrLf           
            ErrMessage = ErrMessage & "<TR><TD CLASS=SMALL WIDTH=""1%"" VALIGN=TOP>" & Translate("Subject",Login_Language,conn) & ":</TD><TD CLASS=SMALL VALIGN=TOP><INPUT SIZE=38 CLASS=SMALL TYPE=""TEXT"" NAME=""SendItSubject"" VALUE=""" & Translate("Document",Login_Language,conn) & ": " & Title & """></TD></TR>" & vbCrLf           
            ErrMessage = ErrMessage & "<TR><TD CLASS=SMALL WIDTH=""1%"" VALIGN=TOP>" & Translate("Message",Login_Language,conn) & ":</TD><TD CLASS=SMALL VALIGN=TOP><TEXTAREA CLASS=SMALL ROWS=4 COLS=40 MAXLENGTH=1000 NAME=""SendItMessage"">" & Translate("You may find the information contained in this document useful.",Login_Language,conn) & "</TEXTAREA></TD></TR>" & vbCrLf
            
            ErrMessage = ErrMessage & "<TR><TD VALIGN=TOP>&nbsp;</TD><TD CLASS=SMALL VALIGN=TOP>" & "<INPUT CLASS=NavLeftHighlight1 TYPE=""Button"" NAME=""Doit"" VALUE="" " & Translate("Send Email",Login_Language,conn) & " "" ONCLICK=""return ckMyEmail();""></TD></TR>" & vbCrLf
            ErrMessage = ErrMessage & "</TABLE></FORM>" & vbCrLf
            if not isblank(Thumbnail) then
              ErrMessage = ErrMessage & "</TD></TR></TABLE>" & vbCrLf
            end if
            
            ErrMessage = ErrMessage & "<A NAME=""SEND_DOCUMENT""></A>" & vbCrLf
            ErrMessage = ErrMessage & "<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'>" & vbCrLf
            ErrMessage = ErrMessage & "<!--" & vbCrLf
            ErrMessage = ErrMessage & "window.focus();" & vbCrLf              
            ErrMessage = ErrMessage & "// -->" & vbCrLf
            ErrMessage = ErrMessage & "</SCRIPT>" & vbCrLf

            ErrType = 0
                           
            Call Status_Screen
           
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
            ErrMessage = ErrMessage & "var wSite_Gateway = window.open('" & Link_Name & "','Site_Gateway_" & RandomFile & "');" & vbCrLf
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
            ErrMessage = ErrMessage & "<LI>Invalid Locator Format 3 - Invalid Method" & Parameter(xMethod) & "</LI>"
            ErrMessage = ErrMessage & "<LI>" & Mid(Locator,instr(1,Locator,"=") + 1)
    
            ErrType = 0
              
          end select
                      
        end if
        
      else
      
        ' Display message for invalid File Name or EMail to Site Administrator
        
        if Script_Debug = false then
          Call Send_Invalid_SiteWide_Email        
          ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but the link provided to you to access this information is not valid.",Login_Language,conn) & "</LI>"
          ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
          ErrMessage = ErrMessage & "<LI>Invalid Locator Format 2 - Invalid Document Number</LI>"
        end if
        
        ErrType = 0
     
      end if
      
    else
    
      ' Expired Link Date
  
      ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but this link has expired.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>" & Translate("We expire links after a period of time to ensure that the user is getting the most up-to-date version of this item.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>" & Translate("Please visit the",Login_Language,conn) & " " & Translate(Site_Description,Login_Language,conn) & Translate(" - Extranet Support Site to get the latest version of this item.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>Invalid Locator Format 0 - Link Expired</LI>"    
  
      ErrType = 0
      
    end if
  
else  

    if not Script_Debug and isblank(ErrMessage) then
      ErrMessage = ErrMessage & "<LI>" & Translate("We are sorry, but the link provided to you to access this information is not valid.",Login_Language,conn) & "</LI>"
      ErrMessage = ErrMessage & "<LI>An automatic notification message detailing this problem was sent to the site administrator.</LI>"
      ErrMessage = ErrMessage & "<LI>Invalid Locator Format 2 - Invalid Document Number</LI>"
    end if
  
    ErrType = 0
  
end if

' --------------------------------------------------------------------------------------

if Script_Debug = True then

  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%
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

' --------------------------------------------------------------------------------------
' Functions or Subroutines
' --------------------------------------------------------------------------------------

sub Status_Screen

  Dim Top_Navigation        ' True / False
  Dim Side_Navigation       ' True / False
  Dim Screen_Title          ' Window Title
  Dim Bar_Title             ' Black Bar Title

  Site_ID_Save = Site_ID
  if not isblank(Style_Site_ID) then
    Site_ID = Style_Site_ID
  elseif Document_Site_ID > 90 then
    Site_ID = 98
  end if
  
  ServerName = LCase(request.ServerVariables("SERVER_NAME"))
        
  if Instr(1,ServerName,"metermantesttools") > 0 then
    Site_ID = 16
  end if  
    
  %>
  <!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
  <%

  if Instr(1,ServerName,"metermantesttools") > 0 then
    Screen_Title    = Translate("Meterman Test Tools",Alt_Language,conn) & " - " & Translate("Electronic Document / File Fulfillment Center",Alt_Language,conn)
    Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><SPAN CLASS=SmallBoldGold>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "</SPAN>"
  elseif isblank(Site_Description) then
    Screen_Title    = Translate("Fluke",Alt_Language,conn) & " - " & Translate("Electronic Document / File Fulfillment Center",Alt_Language,conn)
    Bar_Title       = Translate("Fluke",Login_Language,conn) & "<BR><SPAN CLASS=SmallBoldGold>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "</SPAN>"
  else
    Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Electronic Document / File Fulfillment Center",Alt_Language,conn)
    Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><SPAN CLASS=SmallBoldGold>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "</SPAN>"
  end if

  Top_Navigation  = False 
  Side_Navigation = False
  Content_Width   = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%

  Site_ID = Site_ID_Save
  
  response.write "<SPAN CLASS=Heading3>" & Translate("Electronic Document / File Fulfillment Center",Login_Language,conn) & "</SPAN>"
  response.write "<BR><BR>"

  response.write "<SPAN CLASS=Medium>"

  Script_Locator = instr(1,UCase(ErrMessage),"<SCRIPT")
  if Script_Locator > 0 then
    response.write Mid(ErrMessage,1,Script_Locator -1)
    response.flush
    %>
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
    <%
    response.write Mid(ErrMessage,Script_Locator) & vbCrLf
  else
    response.write ErrMessage & vbCrLf
    response.flush
    %>
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
    <%
  end if  
  
  ErrType    = 0
  ErrMessage = ""
  
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
' Logs Bad Item Number fetch, and allows only one email per item number per day advisory
' --------------------------------------------------------------------------------------

function LogBadDocument(Document)

  SQLBad = "SELECT * FROM Literature_Items_US_Log WHERE Item_Number='" & Document & "' ORDER BY Log_Date_Current DESC"
  Set rsBad = Server.CreateObject("ADODB.Recordset")
  rsBad.Open SQLBad, conn, 3, 3
  
  if rsBad.EOF then
    SQLLog = "INSERT INTO Literature_Items_US_Log (Item_Number,Status,Log_Date_First,Log_Date_Current) VALUES ('" & Document & "'," & "1" & ",'" & Date() & "','" & Date() & "')"
    conn.execute SQLLog
    LogBadDocument = true
  elseif CDate(rsBad("Log_Date_Current")) <> CDate(Date()) then
    SQLLog = "INSERT INTO Literature_Items_US_Log (Item_Number,Status,Log_Date_First,Log_Date_Current) VALUES ('" & Document & "'," & "1" & ",'" & Date() & "','" & Date() & "')"
    conn.execute SQLLog
    LogBadDocument = true
  elseif CDate(rsBad("Log_Date_Current")) = CDate(Date()) then
    SQLLog = "UPDATE Literature_Items_US_Log SET Log_Date_Current='" & Date() & "',Status=" & CInt(rsBad("Status")) + 1 & " WHERE Item_Number='" & Document & "'"
    conn.execute SQLLog
    LogBadDocument = false
  else
    LogBadDocument = false
  end if
  
  rsBad.close
  set rsBad  = nothing
  set SQLLog = nothing
  
end function

' --------------------------------------------------------------------------------------

sub Send_Invalid_SiteWide_Email

  LogFlag = LogBadDocument(Document)

  if CInt(LogFlag) = CInt(true) and script_debug = false then

    Call Send_Email_Header
      
    SQL = "SELECT Site.* FROM Site WHERE ID=" & Parameter(xSite_ID)
    Call GetSiteEmail
  
    Call GetItemEmail
    
    MailSubject = "Invalid Document Number Requested"   
    MailMessage = "This is an automated notification message from the " & Site_Description & ".<P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & "The following locator number: " & Mid(Document,1,9) & " was not found.<BR>" & vbCrLf
    MailMessage = MailMessage & UCase(request.ServerVariables("Server_Name")) & " Asset ID Number: " & Parameter(xAsset_ID) & "<P>" & vbCrLf & vbCrLf
    select case Document_Site_ID
      case 0, 99
        MailMessage = MailMessage & "Fulfillment requested by: <B>Fluke Email Fulfillment</B> "
        if not isblank(SRC) then
          MailMessage = MailMessage & "(" & SRC & ")"
        end if
        MailMessage = MailMessage & "<P>" & vbCrLf & vbCrLf
      case 98
        MailMessage = MailMessage & "Fulfillment requested by: <B>Fluke Digital Library</B> "
        if not isblank(SRC) then
          MailMessage = MailMessage & "(" & SRC & ")" & vbCrLf
        end if
        MailMessage = MailMessage & "<P>" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "WWW PCatalog DB Lookup:  <A HREF=""http://us.fluke.com/pcatalog/admin/LookUpAppnoteByDocID.asp?DocID=" & Mid(Document,1,9) & """>View</A><P>" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Requestor's Web Page Referencing this Asset (if known): " & Referer  & "<P>" & vbCrLf & vbCrLf
      case else
        MailMessage = MailMessage & "Fulfillment requested by: Partner Portal Site " & Site_Description  & "<P>" & vbCrLf & vbCrLf
    end select      
    MailMessage = MailMessage & "Requestor's IP Address: " & request.ServerVariables("REMOTE_ADDR")  & "<BR>" & vbCrLf
    MailMessage = MailMessage & "Portal Debug Information at : " & "<A HREF=""http://" & UCase(request.ServerVariables("Server_Name")) & "/Find_It.asp?Document=" & replace(Parameter(xAsset_ID)," ","") & "&Debug=on" & vbCrLf & """>View</A><P>" & vbCrLf
    MailMessage = MailMessage & "Sincerely" & ",<P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & Site_Description & " at " & UCase(request.ServerVariables("Server_Name"))

    MailMessage = "<HTML>" & vbCrLf &_
                  "<BODY>" & vbCrLf &_
                  "<FONT Style=""font-weight:Normal; font-size:10pt;font-family:Arial;color:#000000;font-style:Normal;"">" & vbCrLf &_
                  MailMessage &_
                  "</FONT></BODY>" & vbCrLf &_
                  "</HTML>" & vbCrLf

    Call Send_EMail

  end if

end sub

' --------------------------------------------------------------------------------------

sub Send_Invalid_Document_Email

  LogFlag = LogBadDocument(Document)

  if CInt(LogFlag) = CInt(true) and script_debug = false then

    Call Send_Email_Header
    
    select case Document_Site_ID
      case 0,99
        SQL = "SELECT Site.* FROM Site WHERE ID=99"
      case else
        SQL = "SELECT Site.* FROM Site WHERE ID=98"
        'if not Email_Debug then
        '  Mailer.AddRecipient     "Webmaster", "Webmaster@Fluke.com"  
        'end if  
    end select
  
    Call GetSiteEmail
    Call GetItemEmail
    
    MailSubject = "Invalid Document Number Requested"
    Status_Comment = Replace(Replace(Status_Comment,"<LI>",""),"</LI>","")
    MailMessage = "This is an automated notification message from the " & Site_Description & "." & vbCrLf & vbCrLf
    MailMessage = MailMessage & "The following document Item Number: <B>" & Trim(Mid(Document,1,9)) & " was not found</B> on " & UCase(request.ServerVariables("Server_Name")) & " because of the following reason:<P>" & vbCrLf & vbCrLf & Style(0) & """" & Status_Comment & """" & "</FONT><P>" &  vbCrLf & vbCrLf  

    SQLOracle = "Select Literature_Type, Efulfillment FROM Literature_Items_US WHERE Item=" & Trim(Mid(Document,1,9)) & " ORDER BY REVISION DESC"
    Set rsOracle = Server.CreateObject("ADODB.Recordset")
    rsOracle.Open SQLOracle, conn, 3, 3
    
    if not rsOracle.EOF then
      MailMessage = MailMessage & rsOracle("Literature_Type") & ": " & rsOracle("Efulfillment") & "<P>" & vbCrLf & vbCrLf
      rsOracle.close
      set rsOracle = nothing
    else
      rsOracle.close
      set rsOracle = nothing
      
      SQLOracle = "Select Product, Title FROM Calendar WHERE Item_Number='" & Trim(Mid(Document,1,9)) & "' ORDER BY Revision_Code DESC"
      Set rsOracle = Server.CreateObject("ADODB.Recordset")
      rsOracle.Open SQLOracle, conn, 3, 3
  
      if not rsOracle.EOF then
        MailMessage = MailMessage & "Document Description: " & rsOracle("Product") & ", " & rsOracle("Title") & "<P>" & vbCrLf & vbCrLf
      end if  
      rsOracle.close
      set rsOracle = nothing
    end if

    select case Document_Site_ID
      case 0, 99
        MailMessage = MailMessage & "Fulfillment requested by: <B>Oracle Email Fulfillment</B> "
        if not isblank(SRC) then
          MailMessage = MailMessage & "(" & SRC & ")"
        end if
        MailMessage = MailMessage & "<P>" & vbCrLf & vbCrLf
      case 98
        MailMessage = MailMessage & "Fulfillment requested by: <B>Fluke Digital Library</B> "
        if not isblank(SRC) then
          MailMessage = MailMessage & "(" & SRC & ")"
        end if
        MailMessage = MailMessage & "<P>" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "WWW PCatalog DB Lookup:  <A HREF=""http://us.fluke.com/pcatalog/admin/LookUpAppnoteByDocID.asp?DocID=" & Mid(Document,1,9) & """>View</A><P>" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Requestor's Web Page Referencing this Asset (if known): " & Referer & "<P>" & vbCrLf & vbCrLf   
      case else
        MailMessage = MailMessage & "Fulfillment requested by: <B>Partner Portal Site " & Site_Description & "</B><P>" & vbCrLf & vbCrLf
    end select      
    MailMessage = MailMessage & "Requestor's IP Address: " & request.ServerVariables("REMOTE_ADDR") & "<BR>" & vbCrLf
    MailMessage = MailMessage & "Domain Lookup: <A HREF=""http://" & UCase(request.ServerVariables("Server_Name")) & "/sw-common/sw-dns_lookup.asp?ipAddress=" & request.ServerVariables("REMOTE_ADDR") & """>View</A><P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & "Portal Debug Information: " & "<A HREF=""http://" & UCase(request.ServerVariables("Server_Name")) & "/Find_It.asp?Document=" & Trim(Mid(Document,1,9)) & "&Debug=on" & """>View</A><P>" & vbCrLf & vbCrLf 
    MailMessage = MailMessage & "Sincerely" & ",<P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & Site_Description & " at " & UCase(request.ServerVariables("Server_Name"))

    MailMessage = "<HTML>" & vbCrLf &_
                  "<BODY>" & vbCrLf &_
                  "<FONT Style=""font-weight:Normal; font-size:10pt;font-family:Arial;color:#000000;font-style:Normal;"">" & vbCrLf &_
                  MailMessage &_
                  "</FONT></BODY>" & vbCrLf &_
                  "</HTML>" & vbCrLf

    Call Send_EMail

  end if

end sub

' --------------------------------------------------------------------------------------

sub Send_Retired_Document_Email

  LogFlag = LogBadDocument(Document)

  if CInt(LogFlag) = CInt(true) and script_debug = false then
  
    Call Send_Email_Header
      
    select case Document_Site_ID
      case 0,99
        SQL = "SELECT Site.* FROM Site WHERE ID=99"
      case else
        SQL = "SELECT Site.* FROM Site WHERE ID=98"
        if not Email_Debug then
          Mailer.AddRecipient     "Shelly Carothers", "Shelly.Carothers@Fluke.com"  
        end if
    end select
  
    Call GetSiteEmail
  
    Call GetItemEmail
  
    MailSubject = "Retired Document Number Requested"   
    MailMessage = "This is an automated notification message from the " & Site_Description & ".<P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & "The following document Item Number: <B>" & Trim(Mid(Document,1,9)) & " was not found</B> on " & UCase(request.ServerVariables("Server_Name")) & " because of the following reason:<P>" & vbCrLf & vbCrLf & Style(0) & """" & Status_Comment & """" & "</FONT><P>" &  vbCrLf & vbCrLf  

    SQLOracle = "Select Literature_Type, Efulfillment FROM Literature_Items_US WHERE Item=" & Trim(Mid(Document,1,9)) & " ORDER BY REVISION DESC"
    Set rsOracle = Server.CreateObject("ADODB.Recordset")
    rsOracle.Open SQLOracle, conn, 3, 3
    
    if not rsOracle.EOF then
      MailMessage = MailMessage & "Document Description: " & rsOracle("Literature_Type") & ", " & rsOracle("Efulfillment") & "<P>" & vbCrLf & vbCrLf
      rsOracle.close
      set rsOracle = nothing
    else
      rsOracle.close
      set rsOracle = nothing
      
      SQLOracle = "Select Product, Title FROM Calendar WHERE Item_Number='" & Trim(Mid(Document,1,9)) & "' ORDER BY Revision_Code DESC"
      Set rsOracle = Server.CreateObject("ADODB.Recordset")
      rsOracle.Open SQLOracle, conn, 3, 3
  
      if not rsOracle.EOF then
        MailMessage = MailMessage & rsOracle("Product") & ": " & rsOracle("Title") & "<P>" & vbCrLf & vbCrLf
      end if  
      rsOracle.close
      set rsOracle = nothing
    end if

    select case Document_Site_ID
      case 0, 99
        MailMessage = MailMessage & "Fulfillment requested by: <B>Oracle Email Fulfillment</B> "
        if not isblank(SRC) then
          MailMessage = MailMessage & vbCrLf & Style(0) & "(" & SRC & ")</FONT>" & vbCrLf
        end if
        MailMessage = MailMessage & "<P>" & vbCrLf & vbCrLf
      case 98
        MailMessage = MailMessage & "Fulfillment requested by: <B>Fluke Digital Library</B> "
        if not isblank(SRC) then
          MailMessage = MailMessage & "(" & SRC & ")"
        end if
        MailMessage = MailMessage & "<P>" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "WWW PCatalog Lookup:  <A HREF=""http://us.fluke.com/pcatalog/admin/LookUpAppnoteByDocID.asp?DocID=" & Mid(Document,1,9) & """>View</A><P>" & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Requestor's Web Page Referencing this Asset (if known): " & Referer & "<P>" & vbCrLf & vbCrLf
      case else
        MailMessage = MailMessage & "Fulfillment requested by: <B>Partner Portal Site " & Site_Description & "</B><P>" & vbCrLf & vbCrLf
    end select      
    MailMessage = MailMessage & "Requestor's IP Address: " & request.ServerVariables("REMOTE_ADDR") & "<BR>" & vbCrLf
    MailMessage = MailMessage & "Domain Lookup: <A HREF=""http://" & UCase(request.ServerVariables("Server_Name")) & "/sw-common/sw-dns_lookup.asp?ipAddress=" & request.ServerVariables("REMOTE_ADDR") & """>View</A><P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & "Portal Debug Information: " & "<A HREF=""http://" & UCase(request.ServerVariables("Server_Name")) & "/Find_It.asp?Document=" & Trim(Mid(Document,1,9)) & "&Debug=on" & """>View</A><P>" & vbCrLf & vbCrLf 
    MailMessage = MailMessage & "Sincerely" & ",<P>" & vbCrLf & vbCrLf
    MailMessage = MailMessage & Site_Description & " at " & UCase(request.ServerVariables("Server_Name"))

    MailMessage = "<HTML>" & vbCrLf &_
                  "<BODY>" & vbCrLf &_
                  "<FONT Style=""font-weight:Normal; font-size:10pt;font-family:Arial;color:#000000;font-style:Normal;"">" & vbCrLf &_
                  MailMessage &_
                  "</FONT></BODY>" & vbCrLf &_
                  "</HTML>" & vbCrLf

    Call Send_EMail

  end if

end sub

' --------------------------------------------------------------------------------------

sub GetItemEmail

  ' Find Partner Portal Site Owner
  
  SQL = "SELECT DISTINCT dbo.Literature_Items_US.ITEM, dbo.Literature_Items_US.COST_CENTER, dbo.Site.ID, dbo.Site.MailToName, dbo.Site.MailTo, dbo.Site.MailCCName, dbo.Site.MailCC " &_
        "FROM   dbo.Lit_Cost_Center LEFT OUTER JOIN " &_
        "       dbo.Site ON dbo.Lit_Cost_Center.Site_ID = dbo.Site.ID RIGHT OUTER JOIN " &_
        "       dbo.Literature_Items_US ON dbo.Lit_Cost_Center.Cost_Center = dbo.Literature_Items_US.COST_CENTER " &_
        "WHERE (dbo.Site.MailTo NOT LIKE '%webmaster%') AND (dbo.Site.MailTo NOT LIKE '%webmail%') AND (dbo.Literature_Items_US.ITEM=" & Trim(Mid(Document,1,9)) & ") " &_
        "ORDER BY MailTo"

  'response.write SQL & "<P>"

  Set rsOwner = Server.CreateObject("ADODB.Recordset")
  rsOwner.Open SQL, conn, 3, 3
  
  if not Email_Debug then
    if not rsOwner.EOF then
      Mailer.AddCC   rsOwner("MailToName"), rsOwner("MailTo")
      if not isblank(rsOwner("MailCC")) then
        Mailer.AddCC   rsOwner("MailCCName"), rsOwner("MailCC")
      end if
    end if
  end if
  
  rsOwner.close
  set SQL = nothing    

  ' Find Item Owner from Oracle Deliverables

  if len(Document) = 7 and isnumeric(Document) then
  
    SQL = "SELECT  MARCOM_MANAGER " &_
          "FROM    dbo.Literature_Items_US " &_
          "WHERE   ITEM='" & Trim(Mid(Document,1,9)) & "' " &_
          "ORDER BY REVISION DESC"
  
    'response.write SQL & "<P>"
  
    Set rsOwner = Server.CreateObject("ADODB.Recordset")
    rsOwner.Open SQL, conn, 3, 3
    
    if not rsOwner.EOF then
    
      OwnerName  = split(rsOwner("Marcom_Manager"),", ")
      OwnerLast  = OwnerName(0)
      OwnerFirst = OwnerName(1)
      if instr(1,rsOwner("Marcom_Manager"),", ") > 0 then
        OwnerInitial = Mid(OwnerFirst,1,2)
      else
        OwnerInitial = OwnerFirst
      end if
            
      SQL = "SELECT DISTINCT Email FROM UserData WHERE (Lastname='" & OwnerLast & "' AND FirstName like '" & OwnerInitial & "%') AND (SUBGROUPS Like '%administrator%' OR SUBGROUPS Like '%content%')"
      Set rsEmail = Server.CreateObject("ADODB.Recordset")
      rsEmail.Open SQL, conn, 3, 3
      
      'response.write SQL & "<P>"
      
      if not Email_Debug then
        OwnerEmail = rsEmail("Email")
        if not rsEmail.EOF then
          Mailer.AddRecipient   Trim(OwnerFirst & " " & OwnerLast), OwnerEmail
        end if
      end if
      
      rsEmail.close
      set rsEmail = nothing
      set SQL     = nothing
      
    end if

    rsOwner.close
    set rsOwner = nothing
    set SQL     = nothing
    
    ' Find Asset Container owner
    
    if not isblank(Submitted_By) then
    
      SQL = "SELECT Firstname, Lastname, email FROM dbo.UserData WHERE ID=" & Submitted_By
      Set rsEmail = Server.CreateObject("ADODB.Recordset")
      rsEmail.Open SQL, conn, 3, 3
      
      if not rsEmail.EOF then
        if LCase(OwnerEmail) <> LCase(rsEmail("Email")) then
          Mailer.AddRecipient   Trim(rsEmail("Firstname") & " " & rsEmail("Lastname")), rsEmail("Email")
        end if
      end if

      rsEmail.close
      set rsEmail = nothing
      set SQL     = nothing
    end if

  end if

end sub

' --------------------------------------------------------------------------------------

sub GetSiteEmail

  Mailer.ReturnReceipt = False
  Mailer.Priority      = 1
  Mailer.AddBCC         "Kelly Whitlock", "Kelly.Whitlock@Fluke.com"

  ' Get Site Info  

  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  Mailer.FromName       = rsSite("FromName")
  Mailer.FromAddress    = rsSite("FromAddress")
  Mailer.ReplyTo        = rsSite("ReplyTo")

  if not Email_Debug then
    Mailer.AddCC     rsSite("MailToName"), rsSite("MailTo")
    if not isblank(rsSite("MailCC")) then
      Mailer.AddCC   rsSite("MailCCName"), rsSite("MailCC")
    end if
  end if
  
  Site_Description      = rsSite("Site_Description") 
  Site_Code             = rsSite("Site_Code")
  
  rsSite.close
  set rsSite = nothing

end sub  

' --------------------------------------------------------------------------------------

sub Send_EMail

  Mailer.ContentType = "text/html; charset=us-ascii"
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
<SCRIPT LANGUAGE="JavaScript">
<!--
  
  function ckMyEmail() {
  
    var dff = document.Send_It.SendItEmail;
  
    if (dff.value.length == 0) {
      alert("<%=Translate("You must supply a valid Send to Email Address.",Alt_Language,conn)%>");
        dff.style.backgroundColor = "#FFB9B9";
        dff.focus();              
      return true;
    }
    else if (!dff.value.match(/^[\w]{1}[\w\.\-_]*@[\w]{1}[\w\-_\.]*\.[\w]{2,6}$/i)) {    
        alert("<%=Translate("Invalid Send to Email Address.",Alt_Language,conn)%>");
        dff.style.backgroundColor = "#FFB9B9";          
        dff.focus();              
        return true;
    }
    else {  
      alert("<%=Translate("Your Email has been sent.",Alt_Language,conn)%>");
      window.blur();
      document.Send_It.submit();
      return false;
    }
    return false;
  }
//-->
</SCRIPT>   

<%
Call Disconnect_SiteWide
%>