'<%   ' Rem out for VBS version (also see end of script)
Option Explicit

Dim DevEnv
DevEnv = false   ' Set to True for Debuging with browser on EVTIBG03
Dim Server       ' Rem Out for DevEnv = True

' --------------------------------------------------------------------------------------
' Title:  Externet Thumbnail Requests Email
'
' Author: Kelly Whitlock
' Date:   01/16/2004
'
' --------------------------------------------------------------------------------------

'On Error Resume Next

' --------------------------------------------------------------------------------------

' Globals
Dim dbConnSiteWide

' Objects

Dim Mailer, rsSite, rsThumbs, rsEmail

' Strings
Dim Site_Root, SQLSite, SQLThumbs, SQLEmail, strBody
Dim ErrorMsg

' Arrays
Dim Status(3)
Status(0) = "<FONT STYLE=""font-size:8.5pt;font-weight:Normal;color:Black;background:#FFFF00;text-decoration:none;font-family:Arial,Verdana;"">&nbsp;In Review&nbsp;</FONT>"
Status(1) = "<FONT STYLE=""font-size:8.5pt;font-weight:Normal;color:Black;background:#00FF33;text-decoration:none;font-family:Arial,Verdana;"">&nbsp;Live&nbsp;</FONT>"
Status(2) = "<FONT STYLE=""font-size:8.5pt;font-weight:Normal;color:Black;background:#FF99CC;text-decoration:none;font-family:Arial,Verdana;"">&nbsp;Archived&nbsp;</FONT>"

' Numbers

' True/false
Dim Old_Site_ID, Site_ID, Site_Status

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

Call Connect_SiteWideDatabase

Call Main

if err.number <> 0 then
  ErrorMsg = ErrorMsg = "A fatal error has occured in Nightly\XNet_Email_Thumbnail_Requests.vbs<BR>" & vbCrLf &_
                        "Error Number: " & err.number & "<BR>" & vbCrLf &_
                        "Error Description: " & err.description & "<BR>" & vbcrlf
	err.clear
end if

if not isblank(ErrorMsg) then  
  Call Send_Error_Email
end if

Call Disconnect_SiteWideDatabase

' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------

sub Main

  SQLThumbs = "SELECT  dbo.Calendar.*, dbo.UserData.FirstName + ' ' + dbo.UserData.LastName AS Owner, " &_
              "        dbo.UserData.Email AS Owner_Email, dbo.UserData.ID AS Owner_ID " &_
              "FROM    dbo.Calendar LEFT OUTER JOIN " &_
              "        dbo.UserData ON dbo.Calendar.Submitted_By = dbo.UserData.ID " &_
              "WHERE  (dbo.Calendar.File_Name LIKE '%.pdf%') " &_
              "        AND (dbo.Calendar.Thumbnail IS NULL) " &_
              "        AND (dbo.Calendar.Status <= 1) " &_
              "        AND (LEN(dbo.Calendar.Item_Number) = 7) " &_
              "ORDER BY dbo.Calendar.Site_ID, dbo.UserData.ID"

  if DevEnv then
    Set rsThumbs = Server.CreateObject("ADODB.Recordset")
  else
    Set rsThumbs = CreateObject("ADODB.Recordset")
  end if      
  rsThumbs.Open SQLThumbs, dbConnSiteWide, 3, 3

  if not rsThumbs.EOF then
  
    Site_ID = -1

    strBody = ""
    
    do while not rsThumbs.EOF
    
      if CInt(Site_ID) <> CInt(rsThumbs("Site_ID")) then
      
        SQLSite = "SELECT dbo.Site.Site_Description, dbo.Site.Enabled, dbo.Site.Site_Code FROM Site WHERE dbo.Site.ID=" & rsThumbs("Site_ID")
        
        if DevEnv then
          Set rsSite = Server.CreateObject("ADODB.Recordset")
        else
          Set rsSite = CreateObject("ADODB.Recordset")
        end if  
        rsSite.Open SQLSite, dbConnSiteWide, 3, 3
        
        Site_Root   = rsSite("Site_Code")
        Site_Status = rsSite("Enabled")
        
        if Site_ID <> -1 and CInt(Site_Status) = CInt(True) then
          strBody = strBody & vbCrLf
        end if
        
        if CInt(Site_Status) = CInt(True) then
          strBody = strBody & "<FONT Style=""font-weight:bold;font-size:10pt;font-family: Arial,Verdana;"">" &_
                              "Extranet Site: " & rsSite("Site_Description") &_
                              "</FONT><BR><BR>"
        end if                      

        rsSite.Close
        set rsSite  = nothing
        set SQLSite = nothing
      
        Site_ID = rsThumbs("Site_ID")
          
      end if  
        
      do while not rsThumbs.EOF
      
        if (CInt(rsThumbs("Site_ID")) = CInt(Site_ID)) and (CInt(Site_Status) = CInt(True)) then
          strBody = strBody &_
                    "<FONT Style=""font-weight:Normal; font-size:8.5pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">" &vbCrLf &_
                    "Item Number: <B>" & rsThumbs("Item_Number") & "</B>&nbsp;&nbsp;&nbsp;" & vbCrLf &_
                    "<A HREF=""http://Support.Fluke.com/Find_It.asp?Document=" & rsThumbs("Item_Number") & """>" & vbCrLf &_
                    "View</A>&nbsp;&nbsp;&nbsp;" & vbCrLf &_
                    "<A HREF=""http://Support.Fluke.com/SW-Administrator/Calendar_Edit.asp?ID=" & rsThumbs("ID") & "&Site_ID=" & Site_ID & """>" &_
                    "Edit Asset</A><BR>" & vbCrLf &_                    
                    "Title: " & rsThumbs("Title") & "<BR>" & vbCrLf &_
                    "Status: " & Status(rsThumbs("Status")) & "<BR>" & vbCrLf &_
                    "Owner: " & rsThumbs("Owner") & "<BR>" & vbCrLf &_
                    "Asset ID: " & rsThumbs("ID") & vbCrLf &_                                    
                    "</FONT><BR><BR>" & vbCrLf

           rsThumbs.MoveNext         

        else
          exit do
        end if

      loop

    loop              
        
    rsThumbs.close
    set rsThumbs = nothing
    
    if not isblank(strBody) then
    
      strBody = "<HTML>" & vbCrLf & "<BODY>" & vbCrLf &_
                "<FONT Style=""font-weight:Normal; font-size:10pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">" &_
                "The following is an automated weekly report generated by the Extranet Sites at Support.Fluke.com that lists publications with item numbers that do not have thumbnail images.&nbsp;&nbsp;" & vbCrLf &_
                "To view the corresponding PDF file, click on View.  To edit the asset container, click on the Edit Asset.  The first time you click on the Edit Asset link. You may have to log into your account, then click on the Edit Asset link again." & vbCrLf &_
                "<P></FONT>" &_
                strBody &_
                "</BODY>" & vbCrLf & "</HTML>" & vbCrLf
                
      Call Connect_Mailer

        'Mailer.FromName = "Portal Sites at Support.Fluke.com"
        'Mailer.FromAddress = "webmaster@fluke.com"

        msg.From = """Portal Sites at Support.Fluke.com""" & "webmaster@fluke.com"

        ' Notify Graphics Department
        if not DevEnv then
          SQLEmail = "SELECT  dbo.UserData.FirstName + ' ' + dbo.UserData.LastName AS Owner, " &_
                     "        dbo.UserData.Email AS Owner_Email " &_
                     "FROM    dbo.Userdata " &_
                     "WHERE   dbo.UserData.Site_ID=102 "
  
          if DevEnv then
            Set rsEmail = Server.CreateObject("ADODB.Recordset")
          else
            Set rsEmail = CreateObject("ADODB.Recordset")
          end if              
          rsEmail.Open SQLEmail, dbConnSiteWide, 3, 3
          
          do while not rsEmail.EOF
            'Mailer.AddRecipient rsEmail("Owner"), rsEmail("Owner_Email")

            msg.To = msg.To & ";" & """" & rsEmail("Owner") & """" & rsEmail("Owner_Email")

            rsEmail.MoveNext
          loop
          
          rsEmail.close
          set rsEmail  = nothing
          set SQLEmail = nothing

        end if
        
        'Mailer.AddBCC "Kelly Whitlock","Kelly.Whitlock@Fluke.com"
        msg.Bcc = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"
        'Mailer.Subject = "Extranet Thumbnail Request Report"
        'Mailer.BodyText =  strBody

        msg.Subject = "Extranet Thumbnail Request Report"
        msg.TextBody = strBody

        'if Mailer.SendMail then
        'else
        '  ErrorMsg = ErrorMsg & "Send Email Failure<BR><BR>" & "Error Description: " & Mailer.Response & ". "
        'end if

        msg.Configuration = conf
        On Error Resume Next
        msg.Send
        If Err.Number = 0 then
          'Success
        Else
          ErrorMsg = ErrorMsg & "Send Email Failure<BR><BR>" & "Error Description: " & Err.Description & ". "
        End If
        
        if DevEnv then
          response.write ErrorMsg & "<P>"     
          response.write strBody
        end if
          
      Call Disconnect_Mailer
        
    end if
    
  end if                  

end sub

' --------------------------------------------------------------------------------------

sub Connect_SiteWideDatabase()

	Dim strConnectionString_SiteWide
	
	set dbConnSiteWide = CreateObject("ADODB.Connection")
	
	if DevEnv then
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=EVTIBG03; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	else
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=FLKPRD18.DATA.IB.FLUKE.COM; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	end if
	
	dbConnSiteWide.ConnectionTimeOut = 120
	dbConnSiteWide.CommandTimeout = 120
	dbConnSiteWide.Open strConnectionString_SiteWide

end sub

' --------------------------------------------------------------------------------------

sub Disconnect_SiteWideDatabase()

	if IsObject(dbConnSiteWide) then
		dbConnSiteWide.Close
		set dbConnSiteWide = nothing
	end if

end sub

' --------------------------------------------------------------------------------------

sub Connect_Mailer

  'Set Mailer = CreateObject("SMTPsvg.Mailer")
  'adding new email method
  %>
  <!--#include virtual="/connections/connection_email_new.asp"-->
  <%

  if DevEnv then
    'Mailer.RemoteHost = "mailhost.tc.fluke.com"
  else
    'Mailer.RemoteHost = "mail.evt.danahertm.com:25"
  end if

  'Mailer.ReturnReceipt = false
  'Mailer.ConfirmRead = false
  'Mailer.WordWrap = True
  'Mailer.WordWrapLen = 85
  'Mailer.QMessage = True
  'Mailer.ClearAttachments
  'Mailer.ContentType = "text/html; charset=us-ascii"
 
end sub

' --------------------------------------------------------------------------------------

sub Disconnect_Mailer

  'set Mailer = Nothing

end sub

' --------------------------------------------------------------------------------------

sub Send_Error_Email

  if not isblank(ErrorMsg) then
  
    Call Connect_Mailer

      'Mailer.FromName    = "EVTIBG01 admin"
      'Mailer.FromAddress = "webmaster@fluke.com"

      msg.From = """EVTIBG01 admin""" & "webmaster@fluke.com"
  
      'Mailer.AddRecipient "Kelly Whitlock","Kelly.Whitlock@fluke.com"
      msg.To = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"

    	'Mailer.Subject     = "Error: Nightly - Extranet Thumbnail Requests " & Date()
    	'Mailer.BodyText    = ErrorMsg
      
      msg.Subject = "Error: Nightly - Extranet Thumbnail Requests " & Date()
      msg.TextBody = ErrorMsg

      'Mailer.SendMail
      msg.Configuration = conf
      On Error Resume Next
      msg.Send
      If Err.Number = 0 then
        'Success
      Else
        'Fail
      End If
    
    Call Disconnect_Mailer    
  
  end if
  
end sub

' --------------------------------------------------------------------------------------

function IsBlank(MyString)

  if isnull(MyString) then
    IsBlank = True
  elseif not isnull(MyString) then
    if Len(Trim(MyString)) = 0 then
      IsBlank = True
    else
      IsBlank = False
    end if
  else
    IsBlank = False
  end if
  
end function

' --------------------------------------------------------------------------------------
'%>