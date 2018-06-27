'<%   ' Rem out for VBS version (also see end of script)
Option Explicit

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
Dim Server, JMailer, rsSite, rsThumbs, rsEmail

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
Dim DevEnv, Old_Site_ID, Site_ID, Site_Status

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

DevEnv = false                 ' Set to True for Debuging with browser on EVTIBG03

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

  ' All Sites
  SQLThumbs = "SELECT  dbo.Calendar.*, dbo.UserData.FirstName + ' ' + dbo.UserData.LastName AS Owner, " &_
              "        dbo.UserData.Email AS Owner_Email, dbo.UserData.ID AS Owner_ID " &_
              "FROM    dbo.Calendar LEFT OUTER JOIN " &_
              "        dbo.UserData ON dbo.Calendar.Submitted_By = dbo.UserData.ID " &_
              "WHERE  (dbo.Calendar.File_Name LIKE '%.pdf%') " &_
              "        AND (dbo.Calendar.Thumbnail IS NULL) " &_
              "        AND (dbo.Calendar.Status <= 1) " &_
              "        AND (LEN(dbo.Calendar.Item_Number) = 7) " &_
              "ORDER BY dbo.Calendar.Site_ID, dbo.UserData.ID, dbo.Calendar.ID DESC"
              
  ' Site 82 Excluded (Fluke Networks)
  SQLThumbs = "SELECT  dbo.Calendar.*, dbo.UserData.FirstName + ' ' + dbo.UserData.LastName AS Owner, " &_
              "        dbo.UserData.Email AS Owner_Email, dbo.UserData.ID AS Owner_ID " &_
              "FROM    dbo.Calendar LEFT OUTER JOIN " &_
              "        dbo.UserData ON dbo.Calendar.Submitted_By = dbo.UserData.ID " &_
              "WHERE  (dbo.Calendar.File_Name LIKE '%.pdf%') " &_
              "        AND (dbo.Calendar.Thumbnail IS NULL) " &_
              "        AND (dbo.Calendar.Status <= 1) " &_
              "        AND (LEN(dbo.Calendar.Item_Number) = 7) " &_
              "        AND (dbo.Calendar.Site_ID <> 82) " &_
              "ORDER BY dbo.Calendar.Site_ID, dbo.UserData.ID, dbo.Calendar.ID DESC"

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
    
        if Site_ID <> -1 then
        strBody = strBody & vbCrLf
        end if
        
        strBody = strBody & "<FONT Style=""font-weight:bold;font-size:10pt;font-family: Arial,Verdana;"">" &_
                            "Extranet Site: " & rsSite("Site_Description") &_
                            "</FONT><BR><BR>"

        Site_ID     = rsThumbs("Site_ID")
        
        rsSite.Close
        set rsSite  = nothing
        set SQLSite = nothing
      
        Site_ID = rsThumbs("Site_ID")
          
      end if  
        
      do while not rsThumbs.EOF
      
        if (CInt(rsThumbs("Site_ID")) = CInt(Site_ID)) then
          strBody = strBody &_
                    "<FONT Style=""font-weight:Normal; font-size:8.5pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">" &vbCrLf &_
                    "Item Number: <B>" & rsThumbs("Item_Number") & "</B>&nbsp;&nbsp;&nbsp;" & vbCrLf &_
                    "<A HREF=""http://Support.Fluke.com/Find_It.asp?Document=" & rsThumbs("Item_Number") & """>" & vbCrLf &_
                    "View</A>&nbsp;&nbsp;&nbsp;" & vbCrLf &_
                    "<A HREF=""http://Support.Fluke.com/SW-Administrator/Calendar_Edit.asp?ID=" & rsThumbs("ID") & "&Site_ID=" & Site_ID & """>" &_
                    "Edit Asset</A><BR>" & vbCrLf &_                    
                    "Title: " & rsThumbs("Title") & "<BR>" & vbCrLf &_
                    "Language: " & UCase(rsThumbs("Language")) & "<BR>" & vbCrLf &_                    
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
                
		Call Connect_JMailer

        JMailer.SenderName = "Portal Sites at Support.Fluke.com"
        JMailer.Sender = "webmaster@fluke.com"

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
            JMailer.AddRecipient rsEmail("Owner_Email")
            rsEmail.MoveNext
          loop
          
          rsEmail.close
          set rsEmail  = nothing
          set SQLEmail = nothing

        end if
        JMailer.Charset="UTF-8"
        JMailer.ContentType="text/html;ContentType=utf-8"
        JMailer.AddRecipientBCC "Kelly.Whitlock@Fluke.com"        
        JMailer.Subject = "Extranet Thumbnail Request Report"
        JMailer.Body =  strBody
		
		if JMailer.Execute =true then
			
		else
			ErrorMsg = ErrorMsg & "Send Email Failure<BR><BR>" & "Error Description: " & JMailer.ErrorMessage & ". "
		end if
      
        if DevEnv then
          response.write ErrorMsg & "<P>"     
          response.write strBody
        end if
          
      Call Disconnect_JMailer
        
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

sub Connect_JMailer

  Set JMailer = CreateObject("JMail.SMTPMail")

  if DevEnv then
    JMailer.ServerAddress = "mailhost.tc.fluke.com"
  else
    JMailer.ServerAddress = "mail.fluke.com"
  end if
  JMailer.Silent=true
  JMailer.ReturnReceipt = false
  JMailer.ClearAttachments
  JMailer.ContentType = "text/html; Content-Transfer-Encoding: quoted-printable; charset=""UTF-8"""
'  JMailer.AddHeader  "Content-Transfer-Encoding: quoted-printable"  

end sub

' --------------------------------------------------------------------------------------

sub Disconnect_JMailer

  set JMailer = Nothing

end sub

' --------------------------------------------------------------------------------------

sub Send_Error_Email

  if not isblank(ErrorMsg) then
  
    Call Connect_JMailer

      JMailer.SenderName   = "EVTIBG01 admin"
      JMailer.ServerAddress = "webmaster@fluke.com"
  
      JMailer.AddRecipient "Kelly.Whitlock@fluke.com"
  
      JMailer.Subject     = "Error: Nightly - Extranet Thumbnail Requests " & Date()
      JMailer.Body    = ErrorMsg
        
      JMailer.Execute
    
    Call Disconnect_JMailer    
  
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