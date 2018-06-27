Option Explicit
On Error Resume Next
' these really are globals
Dim dbConnSiteWide,logfile,Session, Site_id,Tclass,fso
' these I'm dim'ing for speed reasons
' objects
Dim objArgs,Mailer,dbRS1,dbRS,objfolder,cmd,cmd1,prm
' arrays
Dim sites(),site_Nusers(),Rusers,CMS(3),site_emails()
' strings (related to application, mail_stuff, user, site, asset respectively)
Dim site_str,user_str,prev_site,prev_lang,Locator1,Locator2
Dim sBody,sPBody,sHBody,sMSubject,sFtr3,sFtr4,sFtr5
Dim sMsg1,sMsg2,sMsgP,sMsgH,sql
Dim email,Llang,username,mname,dbl_byte
Dim sSiteName,Logo,sSiteTitle,sSiteDesc,sCompany,unt_sSiteTitle
Dim arg_error,fyr,fday,fmon,ffdate,lfilename,abody,err_str
' numbers (not necessarily Cint, but could be)
Dim Dcode,I,j,mail_type,Userid,aid,tot_sites,doitnow
Dim l_at,dot_at,K,eDays,lang_id,iCMS_1,iCMS_2,iCMS_3
' True/false
Dim have_assets,have_site,DevEnv


' Constants for ado
Const adInteger = 3
Const adVarChar = 200
Const adParamInput = &H0001
Const adCmdStoredProc = &H0004

tot_sites = -1

Set Session = CreateObject("Scripting.Dictionary")
Set Tclass = CreateObject("Scripting.Dictionary")
Session.Add "ShowTranslation",FALSE   ' this is a hack enabling use of copied Translate function
Set objArgs = WScript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")

site_str = ""		' optionally used (with command line arg) in limiting out-going emails
user_str = ""		'  ditto
prev_site = "none"	' update a bunch of values when a new site (or lang) is reached in user table
prev_lang = "none"	' in conjunction with prev_site 
Dcode = CLng(DateAdd("d",6,Now)) ' for use with Find_it.asp (links, etc.)
Locator1 = "http://support.fluke.com/register/login.asp?Locator="
Locator2 = "O0O0O0O0O0O0O0O9025O0O0O0O0"
dbl_byte = "-chi-zho-tha-jpn-kor-"
lfilename = "Logs\renew_"
DevEnv = False

arg_error = ""
' parse command line args
if objArgs.Count >= 1 then
	For I = 0 to objArgs.Count - 1
		if objArgs(I) = "-site" then
			I = I + 1
			site_str = objArgs(I)
			if Not IsNumeric(site_str) then
				arg_error = "Command Line error: -site must be numeric, not """ & site_str & """"
			End if
		Elseif objArgs(I) = "-user" then
			Rusers = ParseArgs(I,objArgs)
			user_str = "'" & join(Rusers,"','") & "'"
		Elseif objArgs(I) = "-test" then
			mail_type = 4
		Elseif objArgs(I) = "-tlog" then
			lfilename = "D:\Nightly\Test\AutoRenew\Logs\renew_"
		Elseif objArgs(I) = "-dev" then
			DevEnv = True
		End if
	Next
End if

' create a new logfile
fyr = Year(now)
fmon = Month(now)
if fmon < 10 then fmon = "0" & fmon
fday = Day(now)
if fday < 10 then fday = "0" & fday
ffdate = fyr & "_" & fmon &  "_" & fday
lfilename = lfilename & ffdate & ".log"
Set logfile = fso.OpenTextFile(lfilename, 2, TRUE) ' write to the log file
logfile.write vbcrlf & "----- " & now & vbcrlf

if len(site_str) > 0 then
	logfile.write "Restricted Sites: " & site_str & vbcrlf
end if

if len(user_str) > 0 then
	logfile.write "Restricted Users: " & user_str & vbcrlf
end if

if mail_type = 4 then
	logfile.write "Test option" & vbcrlf
end if

' setup the Mailer object

'Set Mailer = CreateObject("SMTPsvg.Mailer")

'adding new email method
%>
<!--#include virtual="/connections/connection_email_new.asp"-->
<%
'Mailer.ReturnReceipt = false
'Mailer.ConfirmRead = false
'Mailer.WordWrapLen = 80
''Mailer.CharSet = 2  ' 2 = ISO-8859-1 --- now using .CustomCharSet
'Mailer.QMessage = True
'Mailer.ClearAttachments
'Mailer.RemoteHost = "mailhost.tc.fluke.com"
'Mailer.AddExtraHeader "Content-Transfer-Encoding: quoted-printable"
'Mailer.WordWrapLen = 120
'Mailer.WordWrap = True

if len(arg_error) = 0 then
	Real_thing
	if err.number <> 0 then
		arg_error = err.description & vbcrlf & vbcrlf
		err.clear
	end if
	Disconnect_SiteWideDatabase
end if

' send mail to Peter and Kelly
'Mailer.ClearRecipients
''Mailer.AddRecipient "Kelly Whitlock","Kelly.Whitlock@fluke.com"
''Mailer.AddRecipient "Tom Bettenhausen","Tom.Bettenhausen@fluke.com"
'Mailer.AddRecipient "Extranet Group","ExtranetAlerts@fluke.com"
'Mailer.FromName = "EVTIBG01 admin"
'Mailer.FromAddress = "webmaster@fluke.com"

msg.To = """Extranet Group""" & "ExtranetAlerts@fluke.com"
msg.From = """EVTIBG01 admin""" & "webmaster@fluke.com"

' final stats for logfile & email text
logfile.write "Site" & vbtab & vbtab & "Expirers" & vbtab & "Emails" & vbcrlf
abody = abody & "Site" & vbtab & vbtab & "Expirers" & vbtab & "Emails" & vbcrlf

if have_site then
	for i = 0 to ubound(sites)
		logfile.write sites(i) & vbtab & vbtab
		logfile.write site_Nusers(i) & vbtab & site_emails(i) & vbcrlf
		abody = abody & sites(i) & vbtab & vbtab
		abody = abody & site_Nusers(i) & vbtab & site_emails(i) & vbcrlf
	Next
end if
tot_sites = tot_sites + 1
logfile.write "Total sites = " & tot_sites & vbcrlf
	
' check for failed email messages in aspqmail folder
' note: this is setup specific...
set objfolder = fso.GetFolder("C:\AspQMail\Queue\Failed")
if objfolder.Files.Count > 0 then
	logfile.write (objfolder.Files.Count / 2) & " Emails in Failed Queue" & vbcrlf
	abody = abody & (objfolder.Files.Count / 2) & " Emails in Failed Queue" & vbcrlf
end if
set objfolder = Nothing
	
'Mailer.ClearBodyText
'Mailer.ContentType = "text/plain"

if len(arg_error) > 0 then
	'Mailer.BodyText = arg_error & vbcrlf & abody
	'Mailer.Subject = "ERROR: Extranet Auto Renew " & ffdate
	'Mailer.AddRecipient "Webmaster","Webmaster@fluke.com"

	msg.TextBody = arg_error & vbcrlf & abody
	msg.Subject = "ERROR: Extranet Auto Renew " & ffdate
	msg.To = """Webmaster""" & "Webmaster@fluke.com"
else
	'Mailer.BodyText = abody
	'Mailer.Subject = "Extranet Auto Renew " & ffdate

	msg.TextBody = abody
	msg.Subject = "Extranet Auto Renew " & ffdate
end if

'Mailer.SendMail
msg.Configuration = conf
On Error Resume Next
msg.Send
If Err.Number = 0 then
	'Success
Else
	'Fail
End If


logfile.close
'set Mailer = Nothing
Set conf = Nothing
Set msg = Nothing
' -------------------------- this is the end of main ---------------------------------------------

sub Real_thing
	Connect_SiteWideDatabase
	
	' this is the list of sites against which we run this code.
	
	sql = "select id from site where RENEW_DAYS > 0"
	set dbRS1 = dbConnSiteWide.Execute(sql)
	ReDim Preserve asset_sites(0)
	I = 0
	
	if not dbRS1.EOF then
		do until dbRS1.EOF
			ReDim Preserve asset_sites(I)
			asset_sites(I) = dbRS1("id")
			I = I + 1
			dbRS1.MoveNext
		loop
	else
		exit sub
	end if
	
	if len(site_str) > 0 then
		ReDim asset_sites(0)
		asset_sites(0) = site_str
	end if
	
	' get the site information
	
	' (first create the 2 command objects we will use)
	set cmd = CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = dbConnSiteWide
  	cmd.CommandType = adCmdStoredProc
   	cmd.CommandText = "Renewal_GetSiteInfo"
   	cmd.Parameters.Append cmd.CreateParameter("@site_id", adInteger,adParamInput)
	
	' for the users
	set cmd1 = CreateObject("ADODB.Command")
	Set cmd1.ActiveConnection = dbConnSiteWide
  	cmd1.CommandType = adCmdStoredProc
   	cmd1.CommandText = "Renewal_GetUsers"
   	cmd1.Parameters.Append cmd1.CreateParameter("@site_id", adInteger,adParamInput)
	
	for each Site_id in asset_sites
		cmd.Parameters("@site_id").Value = CInt(Site_id)
		
		set dbRS1 = cmd.execute
		
		if not dbRS1.EOF then
			' pure site attributes
			sSiteName = dbRS1("site_code")
			unt_sSiteTitle = dbRS1("site_description")
			sCompany = dbRS1("Company")
			CMS(1) = dbRS1("CMS_Region_1")
			CMS(2) = dbRS1("CMS_Region_2")
			CMS(3) = dbRS1("CMS_Region_3")
			
			logfile.write vbcrlf & "Handling site " & sSiteName & vbcrlf
			
			' site - email attributes
			'Mailer.FromName = Trim(dbRS1("FromName"))
			'Mailer.FromAddress = Trim(dbRS1("FromAddress"))
			'Mailer.ReplyTo = Trim(dbRS1("ReplyTo"))

			msg.From = """" & Trim(dbRS1("FromName")) & """" & Trim(dbRS1("FromAddress"))
			msg.ReplyTo = Trim(dbRS1("ReplyTo"))
			
			' site - display attributes
			if Len(dbRS1("Logo")) > 0 then
				Logo = Trim(dbRS1("Logo"))
			Else
				Logo = "/images/FlukeLogo3.gif"
			End if
			
			' Get_headers() builds the Tclass dictionary with style definitions
			Tclass.RemoveAll()
			Get_headers "SW-Common"
			Get_headers sSiteName
			
			' re-set & set counter arrays
			tot_sites = tot_sites + 1
			ReDim Preserve sites(tot_sites)
			ReDim Preserve site_Nusers(tot_sites)
			ReDim Preserve site_emails(tot_sites)
			sites(tot_sites) = sSiteName
			site_Nusers(tot_sites) = 0
			site_emails(tot_sites) = 0
			
			'logfile.write " Sites(i) = " & sites(tot_sites)
			
			dbRS1.Close
			have_site = True
		else
			have_site = False
		end if
		
		set dbRS1 = nothing
			
		' get the users for this site
		
		if len(user_str) = 0 then
			cmd1.Parameters("@site_id").Value = CInt(Site_id)
			set dbRS1 = cmd1.execute
		else ' (this branch is for when we have a restricted user list)
			sql = "SELECT u.id" & vbcrlf &_
					",u.NTLogin" & vbcrlf &_
					",u.email" & vbcrlf &_
					",u.email_method" & vbcrlf &_
					",u.FirstName" & vbcrlf &_
					",u.MiddleName" & vbcrlf &_
					",u.LastName" & vbcrlf &_
					",u.language" & vbcrlf &_
					",u.site_id" & vbcrlf &_
					",u.newflag" & vbcrlf &_          
					",u.region" & vbcrlf &_
					",l.id as Lang_id" & vbcrlf &_
					",l.name_charset" & vbcrlf &_
					",l.description as Lang_desc" & vbcrlf &_
					", datediff(d,getdate(),u.expirationdate) as EDays" & vbcrlf &_
				"FROM userdata u" & vbcrlf &_
					"INNER JOIN Language l on u.language = l.code" & vbcrlf &_
				"WHERE" & vbcrlf &_
					"datediff(d,getdate(),u.expirationdate) in (14,30,60)" & vbcrlf &_
					"and u.site_id = " & site_id & vbcrlf &_
					"and u.NTlogin in (" & user_str & ") " & vbcrlf &_
          "and u.newflag = 0 " & vbcrlf &_
					"ORDER BY Language"
					
			set dbRS1 = dbConnSiteWide.Execute(sql)
		end if

		Do Until dbRS1.EOF
			
			eDays = dbRS1("EDays")
			site_Nusers(tot_sites) = site_Nusers(tot_sites) + 1
			
			'Mailer.ContentType = "multipart/alternative; boundary=""----xxxxxx"""
			'; charset=""" 
			' &_				Trim(dbRS1("name_charset")) & """"
			
			
			username = dbRS1("FirstName")
			mname = dbRS1("MiddleName")
			if len(mname) > 1 AND mname <> " " then
				username = username & " " & mname
			End if
			
			username = username & " " & dbRS1("LastName")
			email = dbRS1("Email")
			Userid = dbRS1("id")
			
			Llang = dbRS1("Language")
			Lang_id = dbRS1("Lang_id")
			
			' at this point we know everything we need to know - go ahead and 
			' create the email...
			
			logfile.write vbtab & "User: " & dbRS1("NTLogin") & " " & eDays & " " & email
			
			' mail_type is a critical parameter
			' mail_type	= 0 is plain text
			'			= antyhing else is HTML
			'			= 4 is a test mode
			if mail_type < 4 then
				'Mailer.CustomCharSet = Trim(dbRS1("name_charset"))
				'Mailer.ClearRecipients
				'Mailer.AddRecipient username, email
				
				msg.BodyPart.CharSet = Trim(dbRS1("name_charset"))
				msg.To = """" & username & """" & email
				
				mail_type = Clng(dbRS1("email_method"))
				' now build the email header stuff (once per email)
				
				if prev_site <> sSiteName OR Llang <> prev_lang then 
					sSiteTitle = Translate(unt_sSiteTitle,Llang,dbConnSiteWide)
					
					' there is an issue with encoding subjects in double byte
					if Instr(dbl_byte,Llang)>0 then
						'Mailer.Subject = unt_sSiteTitle & " - Account Information"
						msg.Subject = unt_sSiteTitle & " - Account Information"
					else
						sMSubject =  sSiteTitle & " - " &_
							Translate("Account Information",Llang,dbConnSiteWide)
						'Mailer.Subject = Mailer.EncodeHeader(sMSubject)
						msg.Subject = sMSubject
					end if
					
					sSiteDesc = Translate("Extranet Support Site",Llang,dbConnSiteWide)
					
					sMsg1 = "Your SITETITLE Account is set to expire in NN days.  To " &_
						"renew your account, please hit the ""Click Here"" link found below, " &_
						"enter your user name and password, then click ""Ok"". This will take " &_
						"you to the SITETITLE. Look at the navigation buttons on the left " &_
						"side of your screen, and click ""Profile"" to go to your account " &_
						"profile, revise any of your account profile information that may " &_
						"have changed, and the click the ""Update"" button at the bottom " &_
						"of the form. You must click ""Update"" even if you do not change " &_
						"your account profile information."
						
					sMsg1 = Translate(sMsg1,Llang,dbConnSiteWide)
					sMsg1 = Replace(sMsg1,"SITETITLE",sSiteTitle)
						
					sMsg2 = Translate("to go to your account profile",Llang,dbConnSiteWide)
					
					sMsgP = Translate("Use this link",Llang,dbConnSiteWide)
					
					sMsgH = Translate("Click here",Llang,dbConnSiteWide)
					
          sFtr3 = ""
					'sFtr3 = Translate("Renewing your account ensures your continued access to information from COMPANYNAME to help you sell more COMPANYNAME products!",Llang,dbConnSiteWide)
					'sFtr3 = Replace(sFtr3,"COMPANYNAME",sCompany)

					sFtr4 = Translate("Sincerely",Llang,dbConnSiteWide) & ","
					sFtr5 = sSiteTitle & " - " & Translate("Support Team",Llang,dbConnSiteWide)

				End if
				
				sPBody = "------xxxxxx" & VBCrLf & "Content-Type: text/plain;" & VBCrLf & _
					"Content-Transfer-Encoding: quoted-printable" & VBCrLf & VBCrLf
				sHBody = "------xxxxxx" & VBCrLf & "Content-Type: text/html;" & VBCrLf & _
					"Content-Transfer-Encoding: quoted-printable" & VBCrLf & VBCrLf
				
				'Mailer.ContentType = "text/plain;charset="""  & Trim(dbRS1("name_charset")) & """"
				sPBody = sPBody & username & "," & vbcrlf & vbcrlf & Replace(sMsg1,"NN",eDays) & "  " &_
					sMsgP & " " & sMsg2 & "." & vbcrlf & Locator1 & site_id & Locator2 &_
					vbcrlf & vbcrlf & sFtr3 & vbcrlf & vbcrlf & sFtr4 & vbcrlf & sFtr5 & vbcrlf
		
				'Mailer.ContentType = "text/html;charset="""  & Trim(dbRS1("name_charset")) & """"
				
				sHBody = sHBody & "<html>" & vbcrlf & "<body>" & vbcrlf &_
					"<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"" " &_
					"BGCOLOR=""#000000"" FGCOLOR=""FFFFFF"">" & vbcrlf &_
					"   <TR>" & vbcrlf &_
					"     <TD WIDTH=""12"" HEIGHT=""75"">&nbsp;</TD>" & vbcrlf &_
					"     <TD><FONT STYLE=""" & Tclass("heading3fluke") & """>" & sSiteTitle &_
							"</FONT>" & vbcrlf &_
					"       <BR><FONT STYLE=""" & Tclass("smallboldgold") & """>" & sSiteDesc &_
							"</FONT>" & vbcrlf &_
					"     </TD>" & vbcrlf &_
					"     <TD ALIGN=""RIGHT"">" & vbcrlf &_
					"       <IMG SRC=""http://support.fluke.com" & Logo & """ WIDTH=134 " &_
								"HEIGHT=44 BORDER=0>" & vbcrlf &_
					"     </TD>" & vbcrlf &_
					"   </TR>" & vbcrlf &_
					"   <TR>" & vbcrlf &_
					"     <TD COLSPAN=""10"" STYLE=""" & Tclass("linebackground") &_
					""" VSPACE=""0"" HEIGHT=""6""></TD>" & vbcrlf &_
					"   </TR>" & vbcrlf & "</TABLE>" & vbcrlf &_
					"<P><FONT STYLE=""" & Tclass("normal") & """>" &_
					username & ",<P>" & vbcrlf &_
					Replace(sMsg1,"NN",eDays) & "<P>" & "<A HREF=""" & Locator1 & site_id &_
					Locator2 & """>" & sMsgH & "</a> " & sMsg2 & ".<P>" & vbcrlf & sFtr3 &_
					"<P>" & vbcrlf & sFtr4 &_
					"<BR>" & vbcrlf & sFtr5 & vbcrlf & "</font></BODY></HTML>" & vbcrlf
			
				'Mailer.ClearBodyText
				
				' validate that email address looks viable and that the user's CMS_Region is good
				doitnow = 0
				l_at = instr(email,"@")	
				dot_at = instrRev(email, ".")
				
				if CMS(Cint(dbRS1("Region"))) <> 0 then 
					doitnow = 2
				elseif (not l_at > 0) and (dot_at > l_at) then
					doitnow = 1
				end if
				
				if doitnow = 0 then
					'Mailer.BodyText = sBody
					'Mailer.BodyText = sPBody & vbcrlf & sHBody & vbcrlf & "------xxxxxx--" & vbcrlf
					msg.TextBody = sPBody & vbcrlf & sHBody & vbcrlf & "------xxxxxx--" & vbcrlf
					
					'if Mailer.SendMail then
					'	logfile.Write " Mail sent" & vbcrlf
					'	site_emails(tot_sites) = site_emails(tot_sites) + 1
					'else
					'	logfile.Write " Mail failure " & Mailer.Response & vbcrlf
					'end if

					msg.Configuration = conf
					On Error Resume Next
					msg.Send
					If Err.Number = 0 then
						logfile.Write " Mail sent" & vbcrlf
						site_emails(tot_sites) = site_emails(tot_sites) + 1
					Else
						logfile.Write " Mail failure " & Err.description & vbcrlf
					End If

				elseif doitnow = 1 then
					logfile.Write "Bad email" & vbcrlf
				elseif doitnow = 2 then
					logfile.write "conflict on CMS_Region" & vbcrlf
				end if
			else
				logfile.write " test" & vbcrlf
			End if
			
			dbRS1.MoveNext
			prev_site = sSiteName
			prev_lang = Llang
		Loop
		set dbRS1 = Nothing
	Next 'end of "foreach site_id in ..."	
	set cmd = nothing
	set cmd1 = nothing
end sub

Sub Connect_SiteWideDatabase()
	Dim strConnectionString_SiteWide
	
	set dbConnSiteWide = CreateObject("ADODB.Connection")
	
	if DevEnv then
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=EVTIBG18; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	else
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=FLKPRD18.DATA.IB.FLUKE.COM; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	end if
	
	dbConnSiteWide.ConnectionTimeOut = 120
	dbConnSiteWide.CommandTimeout = 120
	dbConnSiteWide.Open strConnectionString_SiteWide
End Sub

Sub Disconnect_SiteWideDatabase()
	if IsObject(dbConnSiteWide) then
		dbConnSiteWide.Close
		set dbConnSiteWide = nothing
	end if
End Sub

' copied from support.fluke.com/Include/functions_translate.asp
' with one tiny edit: remove "Server." from CreateObject method (several places)
' also added thired "Dim SQLtranslate" line to support option explicit
' removed conditional include of adovbs.inc and added them explicitly...

' --------------------------------------------------------------------------------------
    
' --------------------------------------------------------------------------------------
' Translation Function
'
' This function is currently used in THREE places:
'	1) Extranet (\includes)
'	2) WWW (\include)
'	3) evtibg02\d$\nightly\xnet_subscriptions\email_sub_1.vbs
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

  Dim adInteger
  Dim adVarChar 
  Dim adParamInput 
  Dim adCmdStoredProc 

  adInteger = 3
  adCmdStoredProc = &H0004
  adVarChar = 200
  adParamInput = &H0001

  Dim Translated_String
  Dim Translated_ID
  Dim cmd
  Dim prm
  Dim rsTranslate
  Dim bGetRecordset       
  Translated_String = TempString
  Translated_ID     = ""
  
  set cmd = CreateObject("ADODB.Command")
  set prm = CreateObject("ADODB.Parameter")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  
'response.write("sSearch_String: " & left(ReplaceQuote(TempString), 255) & "<BR>")
'response.write("sLanguage_String: " & Login_Language & "<BR>")

  ' Look Up by Phrase ID - not currently used
  if not IsBlank(Translated_String) and isnumeric(Translated_String) then
   	cmd.CommandText = "Translations_Get_Translation_By_ID"
   	Set prm = cmd.CreateParameter("@iTranslationID", adInteger,adParamInput ,, CInt(Translate_String))
   	cmd.Parameters.Append prm
	bGetRecordset = true
  ' Look Up by Phrase
  else
    if (LCase(Login_Language) <> "eng" and not isblank(Login_Language)) and not isblank(TempString) then
	cmd.CommandText = "Translations_Get_Translation_ByString"
   	Set prm = cmd.CreateParameter("@iSiteID", adInteger, adParamInput, , Site_ID)
	cmd.Parameters.Append prm
   	Set prm = cmd.CreateParameter("@sSearch_String", adVarchar,adParamInput ,255, left(ReplaceQuote(TempString), 255))
   	cmd.Parameters.Append prm
   	Set prm = cmd.CreateParameter("@iOriginalSearch_String_Length", adInteger, adParamInput, , len(TempString))
   	cmd.Parameters.Append prm
	bGetRecordset = true
    end if
  end if

  if bGetRecordset = true then
    Set prm = cmd.CreateParameter("@sLanguage_String", adVarchar,adParamInput ,3, uCase(Login_Language))
    cmd.Parameters.Append prm
    set rsTranslate = CreateObject("ADODB.Recordset")
    set rsTranslate = cmd.execute

    if not rsTranslate.EOF then
      Translated_ID = rsTranslate("translation_id")
      Translated_String = Trim(rsTranslate("translation"))

      rsTranslate.close
      set rsTranslate = nothing
      set prm = nothing
      set cmd = nothing
    else  
      Translated_ID="New"  
  	  ' If the len of the string was larger than 255 then we'll have to go ahead and insert it ourselves
  	  if len(TempString) > 255 then
    '    on error resume next
        set cmd = nothing
  	    set cmd = CreateObject("ADODB.Command")
  	    Set cmd.ActiveConnection = conn
          cmd.CommandType = adCmdStoredProc
    		cmd.CommandText = "Translations_Insert_Translation"
    	   	Set prm = cmd.CreateParameter("@sSearchString", adVarChar, adParamInput, 4000, left(TempString, 4000))
    		cmd.Parameters.Append prm
    	   	Set prm = cmd.CreateParameter("@sLanguageString", adVarChar, adParamInput, 4000, "")
    		cmd.Parameters.Append prm
    	   	Set prm = cmd.CreateParameter("@iSiteID", adInteger, adParamInput, , Site_ID)
    		cmd.Parameters.Append prm
    		cmd.execute
    '    if err.number <> 0 then
    '      msgbox err.Description & vbcrlf & vbcrlf & TempString,vbOKOnly
    '    end if
  	  end if
    end if
  end if      
    
'response.write("Translated_String: " & Translated_String & "<BR>")

  if Session("ShowTranslation") = True and LCase(Login_Language) <> "eng" then
    Translate = "<FONT CLASS=ShowTranslation>" & Translated_String & " [" & Trim(CStr(Translated_ID)) & "]</FONT>"
  elseif Session("ShowTranslation") = True then
    Translate = "<FONT CLASS=ShowTranslation>" & Translated_String & "</FONT>"  
  else
    Translate = Translated_String
  end if

end Function      
'-----------------------------end of new function

Sub E_handler(eflag,e_state,e_num,edesc)
	logfile.write "Error: " & e_state & e_num & edesc & vbcrlf
End Sub

Function ParseArgs(j,objArgs)
	Dim i,k,myarray()
	k=0
	j = j + 1
	For i = j to objArgs.Count - 1
		if Instr(objArgs(i),"-") = 1 then
			ParseArgs = myarray
			Exit Function
		End if
		ReDim Preserve myarray(k)
		myarray(k) = objArgs(i)
		k = k + 1
	Next
	ParseArgs = myarray
End Function

' copied from support.fluke.com/Include/functions_string.asp
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

' copied from support.fluke.com/Include/functions_string.asp
Function ReplaceQuote(str)
  Dim tempstr
  if not isnull(str) and not isblank(str) then
    tempstr      = replace(str,"'","&acute;")
    str 		 = replace(tempstr,"""","&quot;")
    tempstr      = replace(str,"&rsquo;","&acute;")
    ReplaceQuote = replace(tempstr,"&rdquo","&quot;")
  else
    ReplaceQuote = str
  end if   
End Function

' copied from support.fluke.com/Include/functions_string.asp
Function RestoreQuote(str)
  Dim tempstr
  if not isnull(str) and not isblank(str) then
    tempstr      = replace(str,"&acute;","'")
    str      = replace(tempstr,"&rsquo;","'")
    tempstr      = replace(str,"&rdquo;","""")
    RestoreQuote = replace(tempstr,"&quot;","""")
  else
    RestoreQuote = str
  end if  
End Function

sub Get_Headers(sitename)
	Dim aryline,key,rootfile,my_file,dFile,Tstrm,thisline,l1,my_val
	
	' this is broken!!!!!
	rootfile = "D:\InetPub\Extranet"
	'rootfile = "W:\web\grponly\support_pb"
	
	my_file = rootfile & chr(92) & sitename & chr(92) & "SW-Style.css"
	
	if fso.FileExists(my_file) then
		Set dFile = fso.GetFile(my_file)
		Set Tstrm = dFile.OpenAsTextStream(1)
		
		Do Until Tstrm.AtEndOfStream
			thisline = Ltrim(Tstrm.Readline)
			
			' we only want lines out of the class file that start with a "."
			if Instr(thisline,".") = 1 then
				aryline = split(thisline,"{")
				key = Trim(aryline(0))
				l1 = Len(key)-1
				key = Lcase(Right(key,l1))
				Select case key
					Case "heading3","heading3red","heading3fluke","product","normalbold", _
						 "smallbold","normal","smallboldgold","small","linebackground"
    					my_val = Trim(aryline(1))
						l1 = Instr(my_val,"}") - 1
						my_val = Left(my_val,l1)
						Tclass.Item(key) = my_val
				End Select
			End if
		Loop
		Tstrm.close
		set Tstrm = Nothing
		set dFile = Nothing
	End if
End sub

function Un_HTML(str)
	' we need to:
	' replace multiple space with one space
	' replace "vbcrlf " with vbcrlf
	Dim st1,st2,tstr
	
	if isblank(str) OR isnull(str) then
		Un_HTML = str
		Exit Function
	End if
	
	str = RestoreQuote(str)
	
	' replace vbcrlf with space (assume wrapped text)
	str = Replace(str,vbcrlf," ")
	' replace each <BR> with placeholeer
	str = Replace(str,"<BR>","$$")
	' replace each <P> with placeholeer
	str = Replace(str,"<P>","$$$$")
	' force all combinations of "li>" to uppercase
	str = Replace(str,"li>","LI>")
	str = Replace(str,"Li>","LI>")
	str = Replace(str,"lI>","LI>")
	' replace <LI> with "vbcrlf- "
	str = Replace(str,"<LI>",vbcrlf & "- ")
	' replace </LI> with vbcrlf
	str = Replace(str,"</LI>",vbcrlf)
	
	' now get the all the other tags (replace with nothing)
	st1 = Instr(str,"<")	
	if st1 > 0 then
		st2 = Instr(st1,str,">")
	Else
		st2 = 0
	End if
	While st2 > st1
		tstr = Mid(str,st1,st2-st1+1)
		str = Replace(str,tstr,"")
		st1 = Instr(str,"<")
		if st1 > 0 then
			st2 = Instr(st1,str,">")
		Else
			st2 = 0
		End if
	Wend
	
	' eliminate leading spaces on lines
	while Instr(str,vbcrlf & " ") > 0
		str = Replace(str,vbcrlf & " ",vbcrlf)
	Wend
	
	' eliminate double vbcrlf
	while Instr(str,vbcrlf & vbcrlf) > 0
		str = Replace(str,vbcrlf & vbcrlf,vbcrlf)
	Wend
	
	' replace each placeholeer with vbcrlf
	str = Replace(str,"$$",vbcrlf)
	
	' some leading spaces are creeping in
	str = Replace(str,vbcrlf & " ",vbcrlf)
	
	' eliminate multiple spaces
	while Instr(str,"  ") > 0
		str = Replace(str,"  "," ")
	Wend
	Un_HTML = str
End function