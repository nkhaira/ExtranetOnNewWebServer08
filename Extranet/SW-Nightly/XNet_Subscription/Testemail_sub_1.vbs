Option Explicit
' Debug (Copy Line to CMD: email_sub_1.vbs -user whitlock -interval [+# days back from todays date] 
On Error Resume Next
' these really are globals
Dim dbConnSiteWide,logfile,Session, Site_id,Tclass,fso
' these I'm dim'ing for speed reasons
' objects
Dim objArgs,JMail,dbRS1,dbRS2,dbRS,objfolder,cmd,prm
' arrays
Dim Rsites,Rusers,ugroups,Subscription_Early,Subscription_Early_Temp ',sites,site_users,site_assets,site_emails
ReDim asset_sites(0)
' strings (related to application, mail_stuff, user, site, asset respectively)
Dim site_str,user_str,prev_site,prev_lang,user_sql,l_sql,prod_sql,Locator,link_mesg
Dim sHeader,sPageName,sBody,sEmbDef,sMSubject,sFtr1,sFtr2,sFtr3,sFtr4,sFtr5,sFtr6
Dim hBody,tBody
Dim sFtrA,sFtrB,sFtrC,sCompany
Dim email,Llang,ucountry,strSubgrp,username,mname,lang_id,dbl_byte
Dim sSiteName,Logo,sSiteTitle,sSiteDesc,sMhder,sMfter
Dim unt_sMfter,unt_sSiteTitle,unt_sMhder
Dim sConfMesg,sCategory,sProduct,sSubCat,sTitle,sDescription,sEmbMesg,sDate,eDate,sALang
Dim arg_error,fyr,fday,fmon,ffdate,lfilename,abody,err_str,sLogMesg
' numbers (not necessarily Cint, but could be)
Dim new_assets,interval,Dcode,I,j,mail_type,Userid,bps,aid,f_size,z_size,tot_sites,tot_users
Dim l_at,dot_at,acode,acatid,K,lastlogon
' True/false
Dim embargo,Confidential,have_file,have_zip,have_link,conf_link,have_container
Dim have_assets,have_site,DevEnv
Dim sites(),site_users(),site_assets(),site_emails(),site_Nusers()

' Constants for ado
Const adInteger = 3
Const adVarChar = 200
Const adParamInput = &H0001
Const adCmdStoredProc = &H0004

tot_sites = -1
tot_users = 0

Set Session = CreateObject("Scripting.Dictionary")
Set Tclass = CreateObject("Scripting.Dictionary")
Session.Add "ShowTranslation",FALSE   ' this is a hack enabling use of copied Translate function
Set objArgs = WScript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")

Subscription_Early = 0    ' if set to -1, subscription will go at early interval 12:00pm PST, else 0 = 9:00PST
interval = 1		' how many days back we'll go
site_str = ""		' optionally used (with command line arg) in limiting out-going emails
user_str = ""		' ditto
prev_site = "none"	' update a bunch of values when a new site (or lang) is reached in user table
prev_lang = "none"	' in conjunction with prev_site 
Dcode = CLng(DateAdd("d",31,Now)) ' Expire Links 30-days after issue - for use with Find_it.asp (links, etc.)
Locator = "http://support.fluke.com/Find_it.asp?SRC=PPSS&Locator="
dbl_byte = "-chi-zho-tha-jpn-kor-"
lfilename = "D:\Extranet\SW-Nightly\XNet_Subscription\Logs\sub_"
DevEnv = False

arg_error = ""
' parse command line args
if objArgs.Count >= 1 then
	For I = 0 to objArgs.Count - 1
		If objArgs(I) = "-site" then
			I = I + 1
			site_str = objArgs(I)
			if Not IsNumeric(site_str) then
				arg_error = "Command Line error: -site must be numeric, not """ & site_str & """"
			End if
		Elseif objArgs(I) = "-interval" then
			I = I + 1
			interval = objArgs(I)
			if Not IsNumeric(interval) then
				arg_error = "Command Line error: -interval must be numeric, not """ & interval & """"
			End if
		Elseif objArgs(I) = "-early" then
			Subscription_Early = -1
		Elseif objArgs(I) = "-user" then
			Rusers = ParseArgs(I,objArgs)
			user_str = "'" & join(Rusers,"','") & "'"
		Elseif objArgs(I) = "-test" then
			mail_type = 4
		Elseif objArgs(I) = "-tlog" then
			lfilename = "D:\Nightly\Test\Xnet_subscription\Logs\sub_"
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

lfilename = lfilename & ffdate
if CInt(Subscription_Early) = CInt(-1) then
  lfilename = lfilename & "_1200PMPST" & ".log"
else
  lfilename = lfilename & "_0900PMPST" & ".log"
end if

Set logfile = fso.OpenTextFile(lfilename, 2, TRUE) ' write to the log file
logfile.write vbcrlf & "----- " & now & vbcrlf

if len(site_str) > 0 then
	logfile.write "Restricted Sites: " & site_str & vbcrlf
end if

if len(user_str) > 0 then
	logfile.write "Restricted Users: " & user_str & vbcrlf
end if

if interval <> 1 then
	logfile.write "Interval: " & interval & vbcrlf
end if

if mail_type = 4 then
	logfile.write "Test Option" & vbcrlf
end if

' setup the JMail object

Set JMail = CreateObject("JMail.SMTPMail")
JMail.ReturnReceipt = false
JMail.Silent=true
JMail.ClearAttachments
JMail.ServerAddress = "mail.fluke.com"

if len(arg_error) = 0 then
	Real_thing
	if err.number <> 0 then
		arg_error = err.description & vbcrlf & vbcrlf
		err.clear
	end if
	Disconnect_SiteWideDatabase
end if

'Send mail to Peter and Kelly

JMail.ClearRecipients
'JMail.AddRecipient "Kelly.Whitlock@fluke.com"
JMail.SenderName = "EVTIBG01 Nightly"
JMail.Sender = "WebMail@fluke.com"
JMail.ContentType = "text/plain"
'JMail.AddHeader  "Content-Transfer-Encoding: quoted-printable"
JMail.ContentTransferEncoding ="quoted-printable"
JMail.AddNativeHeader "X-MimeOLE:Produced fluke.com"
JMail.ISOEncodeHeaders= true
'Final stats for logfile & email text

if CInt(Subscription_Early) = CInt(-1) then
  logfile.write "Subscription Broadcast Time: 12:05PM PST" & vbcrlf
  abody = abody & "Subscription Broadcast Time: 12:05PM PST" & vbcrlf
else
  logfile.write "Subscription Broadcast Time: 09:05PM PST" & vbcrlf
  abody = abody & "Subscription Broadcast Time: 09:05PM PST" & vbcrlf
end if

logfile.write vbcrlf & "Total Accounts = " & tot_users & vbcrlf
logfile.write "Site" & vbtab & vbtab & "Subscribers" & vbtab & vbtab
logfile.write "Emails" & vbtab & "Assets" & vbtab & "Non-Subscribers" & vbcrlf
abody = abody & vbcrlf & "Total Accounts = " & tot_users & vbcrlf
abody = abody & "Site" & vbtab & vbtab & "Subscribers" & vbtab & vbtab
abody = abody & "Emails" & vbtab & "Assets" & vbtab & "Non-Subscribers" & vbcrlf

if have_assets then
	for i = 0 to ubound(sites)
		logfile.write sites(i)
    if len(sites(i)) > 8 then
      logfile.write vbtab & vbtab
    else
      logfile.write vbtab   
    end if  
		logfile.write site_users(i) & vbtab & vbtab
		logfile.write site_emails(i) & vbtab
		logfile.write site_assets(i) & vbtab
		logfile.write site_Nusers(i) & vbcrlf
		abody = abody & sites(i)
    if len(sites(i)) > 8 then
      abody = abody & vbtab & vbtab
    else
      abody = abody & vbtab
    end if  
		abody = abody & site_users(i) & vbtab & vbtab
		abody = abody & site_emails(i) & vbtab
		abody = abody & site_assets(i) & vbtab
		abody = abody & site_Nusers(i) & vbcrlf
	Next
end if
tot_sites = tot_sites + 1
logfile.write vbcrlf & "Total Sites = " & tot_sites & vbcrlf
	
JMail.Body=""

if len(arg_error) > 0 then
	
  JMail.Body = arg_error & abody
  
  if CInt(Subscription_Early) = CInt(-1) then
  	JMail.Subject = "ERROR: Partner Portal Subscription " & Replace(ffdate,"_"," ") & " 12:05:00 PM"
  else
  	JMail.Subject = "ERROR: Partner Portal Subscription " & Replace(ffdate,"_"," ") & " 09:05:00 PM"
  end if
'  JMail.AddRecipient "kelly.whitlock@fluke.com"
else
	JMail.Body = abody
  if CInt(Subscription_Early) = CInt(-1) then
  	JMail.Subject = "Partner Portal Subscription " & Replace(ffdate,"_"," ") & " 12:05:00 PM"
  else
  	JMail.Subject = "Partner Portal Subscription " & Replace(ffdate,"_"," ") & " 09:05:00 PM"
  end if
end if
'JMail.Execute

logfile.close
set JMail = Nothing

' -------------------------- this is the end of main ---------------------------------------------
sub Real_thing

	Connect_SiteWideDatabase
	
	' build the temp table and distinct sites
	if interval <> 0 then
		interval = Cint(interval) * -1
		set cmd = CreateObject("ADODB.Command")
		cmd.ActiveConnection = dbConnSiteWide
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "Subscription_NewAssets"
		cmd.Parameters.Append cmd.CreateParameter("@days", adInteger,adParamInput ,, interval)
		cmd.Parameters.Append cmd.CreateParameter("@early", adInteger,adParamInput ,, Subscription_Early)    
		set dbRS1 = cmd.execute
		set cmd = nothing
	else
		set dbRS1 = dbConnSiteWide.Execute("select distinct(site_id) from New_assets")
	end if
		
	' create a string containing site_ids which have new assets
	if dbRS1.EOF then
		' there are no new assets in the entire site
		logfile.writeline "No new assets" & vbCrLf
		abody = "No new assets" & vbCrLf
		have_assets = False
		set dbRS1 = Nothing
		exit sub
	end if
	have_assets = True
	
	K = -1
	logfile.write "Sites with Assets: "
	abody = "Sites with Assets: "
	
	do until dbRS1.EOF
		site_id = dbRS1("site_id")
		logfile.write site_id & ", "
		abody = abody & site_id & ", "
		
		' if site_str has non-zero length then it contains the only site we'll handle
		' but other sites might be in the recordset
		
		if len(site_str) = 0 then
			K = K + 1
			ReDim Preserve asset_sites(K)
			asset_sites(K) = site_id
		elseif cint(trim(site_str)) = cint(site_id) then
			ReDim Preserve asset_sites(0)
			asset_sites(0) = site_id
		end if
		dbRS1.MoveNext
	Loop
	dbRS1.Close
	set dbRS1 = Nothing
	logfile.write vbcrlf & "Sites with Assets: " & join(asset_sites,"-") & vbcrlf
	abody = abody & vbcrlf
	
	' create the recordset of users to send email
	' site information comes out of this recordset
	' this is a mildly wasteful join in that detailed site information is duplicated into
	' many more rows than it needs to be.  My thinking is that it is more efficient to simply
	' bring that data in because we need to do the join for "site.Subscription_Enabled" anyway
	
	for each Site_id in asset_sites
		set cmd = CreateObject("ADODB.Command")
		set dbRS1 = CreateObject("ADODB.Recordset")
		
		Set cmd.ActiveConnection = dbConnSiteWide
	  	cmd.CommandType = adCmdStoredProc
	   	cmd.CommandText = "Subscription_GetSiteInfo"
	   	cmd.Parameters.Append cmd.CreateParameter("@site_id", adInteger,adParamInput ,, CInt(Site_id))
		
		set dbRS1 = cmd.execute
		set cmd = nothing
		
		' REMEMBER about restricted sites and users
		
		if not dbRS1.EOF then
			' pure site attributes
			sSiteName = dbRS1("site_code")
			unt_sSiteTitle = dbRS1("site_description")
			unt_sMhder = dbRS1("Subscription_Header")	
			unt_sMfter = dbRS1("subscription_footer")
			sCompany = dbRS1("Company")
			
			logfile.write vbcrlf & "Handling site " & sSiteName
			
			' site - email attributes
			JMail.SenderName = Trim(dbRS1("FromName"))
			JMail.Sender = Trim(dbRS1("FromAddress"))
			JMail.ReplyTo  = Trim(dbRS1("ReplyTo"))

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
			
			' re-set counter arrays
			tot_sites = tot_sites + 1
			ReDim Preserve sites(tot_sites)
			ReDim Preserve site_users(tot_sites)
			ReDim Preserve site_Nusers(tot_sites)
			ReDim Preserve site_assets(tot_sites)
			ReDim Preserve site_emails(tot_sites)
			sites(tot_sites) = sSiteName
			site_users(tot_sites) = 0
			site_Nusers(tot_sites) = 0
			site_assets(tot_sites) = 0
			site_emails(tot_sites) = 0
			
			logfile.write " Sites(i) = " & sites(tot_sites) & vbcrlf
			
			dbRS1.Close
			have_site = True
		else
			have_site = False
		end if
		
		set dbRS1 = nothing
		
		' in case the site associated with the new asset isn't associated with 
		' subscriptions we needed to be able to skip the next processing
		if have_site then
			
			' get the users for this site
			if len(user_str) = 0 then
				set cmd = CreateObject("ADODB.Command")
				set dbRS1 = CreateObject("ADODB.Recordset")
				
				Set cmd.ActiveConnection = dbConnSiteWide
			  	cmd.CommandType = adCmdStoredProc
				
				logfile.writeline "Using GetUsers"
			   	cmd.CommandText = "Subscription_GetUsers"
			   	cmd.Parameters.Append cmd.CreateParameter("@site_id", adInteger,adParamInput _
														,, CInt(Site_id))
			
				set dbRS1 = cmd.execute
				set cmd = nothing
			else
				' we need the capability of testing multiple users so this is inline SQL
				user_sql = "SELECT u.id" & vbcrlf &_
						",u.NTLogin" & vbcrlf &_
						",u.email" & vbcrlf &_
						",u.email_method" & vbcrlf &_
						",u.FirstName" & vbcrlf &_
						",u.MiddleName" & vbcrlf &_
						",u.LastName" & vbcrlf &_
						",u.language" & vbcrlf &_
						",u.business_country" & vbcrlf &_
						",u.groups" & vbcrlf &_
						",u.subgroups" & vbcrlf &_
						",u.site_id" & vbcrlf &_
						",u.logon" & vbcrlf &_
						",dt.bps" & vbcrlf &_
						",l.id as Lang_id" & vbcrlf &_
						",l.name_charset" & vbcrlf &_
						",l.description as Lang_desc" & vbcrlf &_
					"FROM userdata u" & vbcrlf &_
						"inner join download_time dt on dt.id = u.connection_speed" & vbcrlf &_
						"inner join language l on l.code = u.language" & vbcrlf &_
					"WHERE" & vbcrlf &_
						"u.site_id = " & Site_id & vbcrlf &_
						"and u.NTLogin in (" & user_str & ")" & vbcrlf &_
						"and len(u.email) > 1 " & vbcrlf &_
						"and u.subscription = -1 " & vbcrlf &_
						"and u.newflag = 0 " & vbcrlf &_
						"and l.enable  = -1 " & vbcrlf &_
						"and u.expirationdate > getdate()" ' & vbcrlf &_

				logfile.writeline "Using GetSingleUser"
				set dbRS1 = dbConnSiteWide.Execute(user_sql)
				logfile.writeline user_sql
			end if
			
			' note: there is no dbRS1.Close...
			Do Until dbRS1.EOF
				logfile.write "User: " & dbRS1("NTLogin") & " assets: "
				site_users(tot_sites) = site_users(tot_sites) + 1
				tot_users = tot_users + 1
	
				ucountry = dbRS1("Business_Country")
				
				strSubgrp = dbRS1("SubGroups")
				if len(strSubgrp) > 1 then
					ugroups = split(strSubgrp,", ")   ' Note: ugroups is an array
				else
					redim ugroups(0)
					ugroups(0) = "none"
				End if
				
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
				
				bps = Clng(dbRS1("bps"))
				
				lastlogon = DateDiff("d",dbRS1("logon"),Now)
				if  lastlogon < 30 then
					lastlogon = 0
				else
					site_Nusers(tot_sites) = site_Nusers(tot_sites) + 1
				end if
				
				' mail_type is a critical parameter
				' mail_type	= 0 is plain text
				'			= antyhing else is HTML
				'			= 4 is a test mode

				if mail_type < 4 then
					JMail.Charset = Trim(dbRS1("name_charset"))
					JMail.ClearRecipients
					'Modified by zensar on 10/04/2006
					'JMail.AddRecipient email
					mail_type = Clng(dbRS1("email_method"))
					' now build the email header stuff (once per email)
					Setup_Email
				End if
				
				' get the assets this user can see and loop through them
				' this would be a select * except that Description must be cast
				' this isn't a stored proc because we're going to "hand build"
				' the sql which does the SubGroup matching
				
				prod_sql = "select na.id" & vbcrlf &_
						",na.category" & vbcrlf &_
						",na.category_id" & vbcrlf &_
						",na.sub_category" & vbcrlf &_
						",na.Product" & vbcrlf &_
						",na.Title" & vbcrlf &_
						",cast(na.Description as nvarchar(4000)) as mDescription" & vbcrlf &_
						",na.Ldate" & vbcrlf &_
						",na.Bdate" & vbcrlf &_
						",na.Edate" & vbcrlf &_
						",na.PEDate" & vbcrlf &_
						",na.Thumbnail" & vbcrlf &_
						",na.Link" & vbcrlf &_
						",na.File_Name" & vbcrlf &_
						",na.File_Size" & vbcrlf &_
						",na.Location" & vbcrlf &_
						",na.Archive_name" & vbcrlf &_
						",na.Archive_size" & vbcrlf &_
						",na.Language" & vbcrlf &_
						",na.Confidential" & vbcrlf &_
						",l.eng" & vbcrlf &_
						"From New_Assets na" & vbcrlf &_
						"inner join Language l on na.Language = l.code" & vbcrlf &_
						"where site_id = " & Cint(Site_id) & vbcrlf &_
					        "AND l.enable = -1" & vbCrLf &_
          					"AND (Country = 'none' OR " & vbcrlf &_
          					"(Country LIKE '%0%' AND Country NOT LIKE '%" & ucountry & "%')" & vbcrlf &_
          					"OR (Country NOT LIKE '%0%' AND Country LIKE '%" & ucountry & "%'))" ' & vbcrlf &_

				' we need to deal with the special case of users being an administrator
        
				if not (Instr(Lcase(strSubgrp),"administrator") >0  or _
						Instr(Lcase(strSubgrp),"domain") >0 ) then
					prod_sql = prod_sql & "and (na.SubGroups LIKE '%all%'" & vbcrlf
					for i = 0 to ubound(ugroups)
						prod_sql = prod_sql & " OR na.SubGroups LIKE '%" & ugroups(i) & "%'" & vbcrlf
				    next
					prod_sql = prod_sql & ")"  & vbcrlf
				end if
					
				prod_sql = prod_sql & "Order by na.Sort, na.Category, na.Product, na.Sub_Category"
				
				set dbRS = dbConnSiteWide.execute(prod_sql)
				'logfile.write "Executing:" & vbcrlf & prod_sql & vbcrlf
				
				if Err.number <> 0 then
					E_handler 1,prod_sql,Err.number,Err.Description
				End if
				'on error goto 0
				
				' pre-setting these 3 triggers "do something special when they change"
				sCategory = "foobar"
				sProduct  = "foobar"
				sSubCat   = "foobar"
				new_assets = 0
				
				if not dbRS.EOF then
					do Until dbRS.EOF
						new_assets = new_assets + 1
						aid = dbRS("id")
						logfile.write aid & " "
						sTitle = dbRS("title")
						sDescription = dbRS("mDescription")
						
						' is this under Public Embargo?
						if len(dbRS("PEDate")) > 0 then
							if CDate(dbRS("PEDate")) > CDate(dbRS("LDate")) then
								embargo = TRUE
								sEmbMesg = sEmbDef & " " & Fdate(dbRS("PEDate"),Llang,dbConnSiteWide)
							Else
								embargo = FALSE
							End if
						End if
						
						' is it Confidential?
						if dbRS("Confidential") = -1 then
							Confidential = TRUE
						Else 
							Confidential = FALSE
						end if
						
						' do we have file_name or archive_name
						if len(dbRS("file_name")) > 1 then
							have_file = TRUE
							f_size = dbRS("file_size")
						Else
							have_file = FALSE
						End if
						
						if len(dbRS("archive_name")) > 1 then
							have_zip = TRUE
							z_size = dbRS("archive_size")
						Elseif Instr(1,Lcase(dbRS("file_name")),".zip") > 0 or _
							Instr(1,Lcase(dbRS("file_name")),".exe") > 0 then
								have_zip = TRUE
								z_size = f_size
						Else
							have_zip = FALSE
						End if
						
						' investigate the link
						if len(dbRS("link")) > 1 then
							have_link = TRUE
							if InStr(lcase(dbRS("link")),"sw-gateway") > 0 then
							    conf_link = 2
							elseif InStr(lcase(dbRS("link")),"support.fluke.com") > 0 OR _
									InStr(lcase(dbRS("link")),".tc.fluke.com") > 0 then
							    conf_link = 1
							Else
								conf_link = 0
							End if
						else
							have_link = FALSE
						End if
						
						' if there isn't a link, zip, or file then maybe we should build a link
						' this is a rare occurence so there is no stored proc
						have_container = False
						if (not have_link) and (not have_zip) and (not have_file) then
							l_sql = "select code,category_id from calendar where id = " & aid
							set dbRS2 = dbConnSiteWide.execute(l_sql)
							acode = Cint(dbRS2("Code"))
							acatid = dbRS2("category_id")
							if acode > 7999 and acode < 9001 then
								have_container = True
							end if
							dbRS2.Close
							set dbRS2 = Nothing
						end if 
						
						'format Ldate
						sDate = FDate(dbRS("bdate"),Llang,dbConnSiteWide)
						eDate = FDate(dbRS("edate"),Llang,dbConnSiteWide)
						
						if sDate <> eDate then
							sDate = sDate & " - " & eDate
						end if
						
						' get the asset language in the user's language
						sALang = Translate(dbRS("eng"),Llang,dbConnSiteWide)
						
						' build the body section of the email for this asset
						if mail_type < 4 then  ' HTML & text
							Plain_mail_asset
							HTML_mail_asset
						End if
						
						dbRS.MoveNext
					Loop
					dbRS.Close
				end if ' end of "if not dbRS.EOF"
				set dbRS = Nothing
				
				' end the email
				tBody = tBody & vbcrlf & vbcrlf
				
				if len(sMfter) > 0 then
					tBody = tBody & sMfter & vbcrlf & vbcrlf
				End if
				tBody = tBody & sFtrA & " " & sCompany & " " & sFtrB & ":" & vbcrlf &_
					"http://support.fluke.com/" & sSiteName & vbcrlf & vbcrlf &_
					sFtrC & vbcrlf & vbcrlf &_
					sFtr1 & vbcrlf & vbcrlf &_
					sFtr2 & " " & sSiteTitle & ". " &_
					sFtr3 & vbcrlf & vbcrlf &_
					sFtr4 & vbcrlf &_
					sFtr5 & vbcrlf & vbcrlf &_
					Translate("Copyright",Llang,dbConnSiteWide) & " " & sFtr6 & vbcrlf
			
				hBody = hBody & "</TABLE>" & vbcrlf &_
					"<FONT STYLE=""" & Tclass("normal") & """>" & vbcrlf
					
				if len(sMfter) > 0 then
					hBody = hBody & "<P>" & sMfter & vbcrlf
				End if
				
				hBody = hBody & "<P><HR><BR>" & sFtrA & " " & sCompany & " " & sFtrB &_
					": <a href=""http://support.fluke.com/" & sSiteName & """>" &_
					sSiteTitle & "</A>.  " & sFtrC & vbcrlf & "<P>" &_
					sFtr1 & "<P>" & vbcrlf & sFtr2 & " " & sSiteTitle & ". " &_
					sFtr3 & "<P>" & vbcrlf &_
					sFtr4 & "<BR>" & vbcrlf &_
					sFtr5 & "<P><P>" & vbcrlf &_
					"<CENTER><FONT STYLE=""" & Tclass("small") & """>&copy; " & sFtr6 &_
					"</FONT></CENTER>" & vbcrlf &_
					"</BODY></HTML>" & vbcrlf
				
				' without ClearBodyText the call to BodyText is additive
				' there could be a speed trade-off between the two methods:
				'	1) setting and adding to internal var (my method, sBody is that var)
				'	2) simply adding to Mailer.BodyText again and again
				' I chose to communicate with the Mailer object fewer times, didn't do research
				JMail.Body =""
				JMail.Body = tbody & vbcrlf & hbody & vbcrlf & "------xxxxxx--" & vbcrlf
				
				' mail_type = 4 means we don't actually send mail
				if mail_type < 4 AND new_assets > 0 then
					' validate that email looks viable
					l_at = instr(email,"@")	
					dot_at = instrRev(email, ".")
					
					if (l_at > 0) and (dot_at > l_at) then
						JMail.ClearRecipients()
						if username="scotton" then
						    msgbox JMail.SenderName
						    msgbox JMail.ServerAddress
						    MsgBox JMail.Sender
						end if
						JMail.AddRecipient "parag.deshpande@fluke.com"
						if JMail.Execute = true then
							logfile.Write " Mail Sent To: " & email & vbcrlf
							site_emails(tot_sites) = site_emails(tot_sites) + 1
							JMail.Close()
						else
							logfile.Write "Mail Failure: " & email & "-"
							logfile.Write JMail.ErrorMessage & vbcrlf
							logfile.Write JMail.Log & vbcrlf
							logfile.Write JMail.ErrorSource & vbcrlf
							logfile.Write JMail.ErrorCode & vbcrlf
							MsgBox err.Description
							msgbox err.number
						end if
					else
						logfile.Write "Bad Email: " & email & vbcrlf
					end if
				Else
					logfile.write " Type = 4 or No New Assets" & vbcrlf
				End if
				if new_assets > site_assets(tot_sites) then site_assets(tot_sites) = new_assets
				dbRS1.MoveNext
				prev_site = sSiteName
				prev_lang = Llang
			Loop
			set dbRS1 = Nothing
		end if 'end of "if have_site..."
	Next 'end of "foreach site_id in ..."
	
end sub

Sub Connect_SiteWideDatabase()
	Dim strConnectionString_SiteWide
	
	set dbConnSiteWide = CreateObject("ADODB.Connection")
	
	if DevEnv then
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=EVTIBG03; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	else
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=EVTIBG18.tc.fluke.com; " &_
			"UID=marcomweb;DATABASE=fluke_SiteWide;pwd=!?wwwProd1"
		'strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=FLKPRD18.DATA.IB.FLUKE.COM; " &_
			'"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
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

Function DTime(size,speed)
	Dim hr,min,sec,x
	' size is in bytes; speed is bps; x is, at this point, in hours (thus /3600)
	' it could be that I should do a *7 to convert from bytes to bits
	x = (Cdbl(size) / Cdbl(speed)) / 3600.00
	hr = x
	' get the whole hours
	hr = Fix(hr)
	' convert fractional hrs to minutes
	x = Cdbl((x - hr) * 60.0)
	min = x
	' get the whole minutes
	min = Fix(min)
	' convert fractional minutes to seconds and get whole secs
	sec = Fix((x - min) * 60.0)
	
	' do zero-padding on 3 segments
	if hr = 0 and min = 0 and sec = 0 then
		Dtime = " "
	Else
		if hr < 10 then
			hr =  "0" & hr
		End if
		if min < 10 then
			min =  "0" & min
		End if
		if sec < 10 then
			sec =  "0" & sec
		End if
		Dtime = hr & ":" & min & ":" & sec
	End if
End Function

Function SizeFormat(fsize,ll,conn)
	if fsize > 1000000 then
		fsize = Cdbl(fsize) / 1000000
		SizeFormat = FormatNumber(fsize,2) & " MB"
	elseif fsize > 1000 then
		fsize = Cdbl(fsize) / 1000
		SizeFormat = FormatNumber(fsize,2) & " KB"
	elseif fsize = 0 then
		SizeFormat = Translate("Unknown size",ll,conn)
	else
		SizeFormat = fsize & " bytes"
	End if
End Function

Function FDate(xBDATE,Llang,conn)
	FDate = Day(xBDATE) & " " &_
		Translate(MonthName(Month(xBDATE)),Llang,conn) &_
		" " & Year(xBDATE)
End Function

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
  Translated_ID = ""
  
  Dim LLang
  ' Europe wants only English Language Only so change special code "elo" to "eng"
  if LCase(Login_Language) = "elo" then LLang = "eng" else LLang = Login_Language
    
  set cmd = CreateObject("ADODB.Command")
  set prm = CreateObject("ADODB.Parameter")
  Set cmd.ActiveConnection = conn
  cmd.CommandType = adCmdStoredProc
  
 'response.write("sSearch_String: " & left(ReplaceQuote(TempString), 255) & "<BR>")
 'response.write("sLanguage_String: " & LLang & "<BR>")

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
    Set prm = cmd.CreateParameter("@sLanguage_String", adVarchar,adParamInput ,3, uCase(LLang))
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

  Dim tmpType

  IsBlank = false

  select case VarType(MyString)

    case 0, 1                       ' Empty & Null
      IsBlank = true
    case 8                          ' String
      if Len(MyString) = 0 then
        IsBlank = true
      elseif UCase(MyString) = "NULL" or UCase(MyString) = "NO VALUE" or Trim(MyString & " ") = "" then
        IsBlank = true
      end if
    case 9                          ' Object  
      tmpType = TypeName(MyString)
      if (tmpType = "Nothing") or (tmpType = "Empty") then
        IsBlank = True
      end if
    case 8192, 8204, 8209            ' Array
      ' Does it have at least one element?
      if UBound(MyString) = 0 then
        IsBlank = True
      end if
  end select
  
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
	
	rootfile = "D:\Nightly\XNet_StyleSheets"
	
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
					Case "topblackbar", "heading3","heading3red","heading3fluke","product","normalbold", _
						 "smallbold","smallboldbar","topcolorbar","normal","smallboldgold","small","linebackground","tablebackground"
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

function Mk_Locator(User,Asset,Method,Dcode,Lid)

	Dim key
	' Site_id is a global
	
    key = ((2 * CInt(Site_id) + (CLng(CLng(User)/2)) + (CLng(CLng(Asset)/3))) + Asset)
    Mk_Locator = Site_id & "O" & User & "O" & Asset & "O" & Method & "O" &_
		key & "O" & Dcode & "O" & Lid
    
end function

function Lg_Locator(User,Asset,Method,Dcode,Lid,Code,CatID)

	Dim key
	' Site_id is a global
	
    key = ((2 * CInt(Site_id) + (CLng(CLng(User)/2)) + (CLng(CLng(Asset)/3))) + Asset)
	
    Lg_Locator = Site_id & "O" & User & "O" & Asset & "O" & Method & "O" &_
		key & "O" & Dcode & "O" & Lid & "O0O9003O" & Asset & "O1O" & Code & "O" & CatID

end function

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
  
  str = Replace(str,"ul>","UL>")
  str = Replace(str,"Ul>","UL>")
  str = Replace(str,"uL>","uL>")
  
  str = Replace(str,"ol>","OL>")
  str = Replace(str,"Ol>","OL>")
  str = Replace(str,"oL>","OL>")
  
	str = Replace(str,"li>","LI>")
	str = Replace(str,"Li>","LI>")
	str = Replace(str,"lI>","LI>")
	
  ' Substitutions
  
	str = Replace(str,"<LI>",vbcrlf & "- ")
	str = Replace(str,"</LI>",vbcrlf)
	str = Replace(str,"<UL>",vbcrlf)
	str = Replace(str,"</UL>",vbcrlf)
	str = Replace(str,"<OL>",vbcrlf)
	str = Replace(str,"</OL>",vbcrlf)
	
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

sub Setup_email
	JMail.ContentType = "multipart/alternative; boundary=""----xxxxxx""; charset=""" & Trim(dbRS1("name_charset")) & """"
	
	tBody = "------xxxxxx" & VBCrLf & "Content-Type: text/plain;" & VBCrLf & _
		"Content-Transfer-Encoding: quoted-printable" & VBCrLf & VBCrLf
	hBody = "------xxxxxx" & VBCrLf & "Content-Type: text/html;" & VBCrLf & _
		"Content-Transfer-Encoding: quoted-printable" & VBCrLf & VBCrLf
	
	if prev_site <> sSiteName OR Llang <> prev_lang then
		sSiteTitle = Translate(unt_sSiteTitle,Llang,dbConnSiteWide)
		
		' there is an issue with encoding subjects in double byte
		if Instr(dbl_byte,Llang)>0 then
			JMail.Subject = unt_sSiteTitle & " - Today's News"
		else
			sMSubject =  sSiteTitle & " - " & Translate("Today's News",Llang,dbConnSiteWide)
			JMail.Subject = sMSubject
		end if
		
		sSiteDesc = Translate("Extranet Support Site",Llang,dbConnSiteWide)
		sPageName = Translate("What's New",Llang,dbConnSiteWide)
		
		sHeader = Translate("New information that you may be interested in viewing or downloading by clicking on the links below.",Llang,dbConnSiteWide)
		
    sFtrA = Translate("For security purposes, the above links expire 30-days after the date of this notice.",Llang,dbConnSiteWide)
    
    if sSiteName = "Educators" then
  		sFtrA = sFtrA & " " & Translate("For more information on education resources from",Llang,dbConnSiteWide)
    else
  		sFtrA = sFtrA & " " & Translate("For much more information about products, support and other resources from",Llang,dbConnSiteWide)
    end if
		
		sFtrB = Translate("visit",Llang,dbConnSiteWide)
		
		sFtrC = Translate("You will need to enter your account User Name and Password to gain access to the site.  If you've forgotten your user name and password enter what you think it might be.  After your second failed attempt to log on to the site, you will get an opportunity to have your account information sent to your email account that you registered with this site.",Llang,dbConnSiteWide)
		
		sFtr1 = Translate("How to subscribe/unsubscribe to our subscription service",Llang,dbConnSiteWide) & ":"
		
		sFtr2 = Translate("You are receiving this notification because we have received a subscription request from you. To change any of your subscription options, please visit:",Llang,dbConnSiteWide)
		
		sFtr3 = Translate("Modify your Account Profile by unselecting the Subscription Service.",Llang,dbConnSiteWide) & "  " &_
			Translate("Please note, that by unsubscribing from the Subscription Service you will no longer receive advance notice of any new items or events that have been added to this site.",Llang,dbConnSiteWide) 
	
		sFtr4 = Translate("Sincerely",Llang,dbConnSiteWide) & ","
		sFtr5 = sSiteTitle & " - " & Translate("Support Team",Llang,dbConnSiteWide)
		
		sFtr6 = "1995-" & Year(Now) & " " & Translate("Fluke Corporation",Llang,dbConnSiteWide) &_
			" - " & Translate("All rights reserved",Llang,dbConnSiteWide)
		
		sConfMesg = Translate("Confidential Information - Not for Public Release",Llang,dbConnSiteWide)
		sEmbDef = Translate("Embargoed Information - Not for Public Release until",Llang,dbConnSiteWide)
		
    if 1=2 then
		sLogMesg = "We have noticed that you have not logged on to your extranet account for " &_
			"NN days.  We encourage you to logon at least once a month to view more detailed " &_
			"information than the limited information we send you in this email.  " &_
			"See the link in the footer below for easy access to the site."
			
		sLogMesg = Translate(sLogMesg,Llang,dbConnSiteWide)
    end if
			
		if len(sMhder) > 0 then
			sMhder = Translate(unt_sMhder,Llang,dbConnSiteWide)
		End if
		
		if len(sMfter) > 0 then
			sMfter = Translate(unt_sMfter,Llang,dbConnSiteWide)
		End if
	End if
	
	tBody = tBody & sHeader & vbcrlf & vbcrlf
		
	hBody = hBody & "<HTML>" & vbcrlf & "<BODY>" & vbcrlf &_
	"<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"" STYLE=""" & Tclass("topblackbar") & """>" & vbcrlf &_
	"   <TR>" & vbcrlf &_
	"     <TD WIDTH=""12"" HEIGHT=""75"">&nbsp;</TD>" & vbcrlf &_
	"     <TD><FONT STYLE=""" & Tclass("heading3fluke") & """>" & sSiteTitle & "</FONT>" & vbcrlf &_
	"       <BR><FONT STYLE=""" & Tclass("smallboldgold") & """>" & sSiteDesc & "</FONT>" & vbcrlf &_
	"     </TD>" & vbcrlf &_
	"     <TD ALIGN=""RIGHT"">" & vbcrlf &_
	"       <IMG SRC=""http://support.fluke.com" & Logo & """ WIDTH=134 HEIGHT=44 BORDER=0>" & vbcrlf &_
	"     </TD>" & vbcrlf &_
	"   </TR>" & vbcrlf &_
	"   <TR>" & vbcrlf &_
	"     <TD COLSPAN=""10"" STYLE=""" & Tclass("topcolorbar") & """ VSPACE=""0"" HEIGHT=""6""></TD>" & vbcrlf &_
	"   </TR>" & vbcrlf &_
  "</TABLE>" & vbcrlf

	hBody = hBody & "<P><FONT STYLE=""" & Tclass("heading3") & """>" &_
		sPageName & "</font><P>" & vbcrlf &_
		"<FONT STYLE=""" & Tclass("normal") & """>" &_
		sHeader & vbcrlf
	
	if len(sMhder) > 0 then
		tBody = tBody & sMhder & vbcrlf & vbcrlf
		
		hBody = hBody & "<P>" & sMhder & vbcrlf
	End if
	
	hBody = hBody & "</FONT><P>" & vbcrlf &_
		"<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbcrlf
end sub

sub Plain_mail_asset
	' Category
	if sCategory <> dbRS("Category") then
		sCategory = dbRS("Category")
		
		tBody = tBody & "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=" &_
			vbcrlf & Translate(sCategory,Llang,dbConnSiteWide) & vbcrlf &_
			"=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=" & vbcrlf & vbcrlf
	End if
	
	' Product and Sub-Category
	if sProduct <> dbRS("Product") or sSubCat <> dbRS("Sub_Category") then
		sProduct = dbRS("Product")
		sSubCat = dbRS("Sub_Category")
		
		tBody = tBody & "-------------------------------------------" & vbcrlf &_
			Translate(sProduct,Llang,dbConnSiteWide) & vbTab & vbTab &_
			Translate(sSubCat,Llang,dbConnSiteWide) & vbcrlf
	End if
	
	' Title and Description
	tBody = tBody & "-------------------------------------------" & vbcrlf & vbcrlf &_
		Ucase(Translate("Title",Llang,dbConnSiteWide)) & ":" & vbcrlf &_
		Un_HTML(sTitle) & vbcrlf & vbcrlf &_
		Ucase(Translate("Description",Llang,dbConnSiteWide)) & ":" & vbcrlf &_
		Un_HTML(sDescription) & vbcrlf & vbcrlf
	
	' Confidential and Embargo
	if Confidential then
		tBody = tBody & " " & sConfMesg & vbcrlf
	end if
	
	if embargo then
		tBody = tBody & " " & sEmbMesg & vbcrlf
	End if
	
	' Language
	tBody = tBody & vbcrlf & Ucase(Translate("Language",Llang,dbConnSiteWide)) &_
		": " & sAlang & vbcrlf &_
		Ucase(Translate("Date",Llang,dbConnSiteWide)) & ": " & sDate & vbcrlf & vbcrlf
	
	' file_name, archive_name and link are all very similar
	if have_file then
		tBody = tBody & Ucase(Translate("View",Llang,dbConnSiteWide)) & ": " &_
      Locator & Mk_Locator(Userid,aid,3,Dcode,Lang_id) & vbcrlf &_
			Ucase(Translate("Size",Llang,dbConnSiteWide)) & ": " &_
			SizeFormat(Clng(f_size),Llang,dbConnSiteWide) & vbcrlf &_
			Ucase(Translate("Transfer Time",Llang,dbConnSiteWide)) & ": " &_
			DTime(f_size,bps) & vbcrlf & vbcrlf
	End if
	
	if have_zip then
		tBody = tBody & Ucase(Translate("Download",Llang,dbConnSiteWide)) & ": " &_
			Locator & Mk_Locator(Userid,aid,4,Dcode,Lang_id) & vbcrlf &_
			Ucase(Translate("Size",Llang,dbConnSiteWide)) & ": " &_
			SizeFormat(Clng(z_size),Llang,dbConnSiteWide) & vbcrlf &_
			Ucase(Translate("Transfer Time",Llang,dbConnSiteWide)) & ": " &_
			DTime(z_size,bps) & vbcrlf & vbcrlf
	End if
	
	' give them an "Email to Me" choice, favor the file
	if have_file OR have_zip then
		tBody = tBody & Translate("Send this to me by email",Llang,dbConnSiteWide) & ": " &_
			Locator & Mk_Locator(Userid,aid,5,Dcode,Lang_id) & vbcrlf & vbcrlf
	End if
	
	if have_link then
		if conf_link = 1 then
			tBody = tBody & Ucase(Translate("Link",Llang,dbConnSiteWide)) & ": " &_
				Locator & Mk_Locator(Userid,aid,4,Dcode,Lang_id) & vbcrlf & vbcrlf
		elseif conf_link = 2 then
			tBody = tBody & Ucase(Translate("View",Llang,dbConnSiteWide)) & ": " &_
				Locator & Mk_Locator(Userid,aid,8,Dcode,Lang_id) & vbcrlf & vbcrlf
		Else
			tBody = tBody & Ucase(Translate("Link",Llang,dbConnSiteWide)) & ": " &_
				dbRS("link") & vbcrlf & vbcrlf
		End if
	End if
	
	if have_container then
		tBody = tBody & Ucase(Translate("Link",Llang,dbConnSiteWide)) & ": " &_
			Locator & Lg_Locator(Userid,aid,9,Dcode,Lang_id,acode,acatid) & vbcrlf & vbcrlf
	End if
	
	' Location also includes the idea of Begin and End dates
	if len(dbRS("location")) > 1 then
		tBody = tBody & Ucase(Translate("Location",Llang,dbConnSiteWide)) & ": " &_
			dbRS("location") & vbcrlf & vbcrlf
	End if
end sub

sub HTML_mail_asset
	' Category
	if sCategory <> dbRS("Category") then
		sCategory = dbRS("Category")
		
		hBody = hBody & "  <TR>" & vbcrlf &_
			"    <TD COLSPAN=5 HEIGHT=32 VALIGN=MIDDLE> " & vbcrlf &_
			"<FONT Style=""" & Tclass("heading3red") & """>" &_
			Translate(sCategory,Llang,dbConnSiteWide) & "</FONT>    </TD>" &_
			vbcrlf & "    </TR>" & vbcrlf
	End if
	
	' Product and Sub-Category
	if sProduct <> dbRS("Product") or sSubCat <> dbRS("Sub_Category") then
		sProduct = dbRS("Product")
		sSubCat = dbRS("Sub_Category")
		
		hBody = hBody & "  <TR>" & vbcrlf & "    <TD HEIGHT=16 COLSPAN=3 Style=""" &_
			Tclass("product") & """>&nbsp;" & Translate(sProduct,Llang,dbConnSiteWide) &_
			"</TD>" & vbcrlf & "    <TD HEIGHT=16 COLSPAN=2 Style=""" &_
			Tclass("product") & """ ALIGN=RIGHT>" &_
			Translate(sSubCat,Llang,dbConnSiteWide) & "&nbsp;</TD>" & vbcrlf &_
			"    </TR>" & vbcrlf & "  <TR>" & vbcrlf &_
			"    <TD HEIGHT=1 COLSPAN=5><HR></TD>" & vbcrlf & "    </TR>" & vbcrlf
	End if
	
	' Title and Date
	hBody = hBody & "  <TR>" & vbcrlf & "    <TD COLSPAN=2 WIDTH=""50%"" VALIGN=TOP><BR>" &_
		"<FONT Style=""" & Tclass("normalbold") & """>" &_
		sTitle & "</FONT></TD>" & vbcrlf &_
		"    <TD COLSPAN=3 WIDTH=""25%"" VALIGN=TOP ALIGN=RIGHT><BR>" &_
		"<FONT Style=""" & Tclass("smallbold") & """>" &_
		sDate & " </FONT></TD>" & vbcrlf & "  </TR>" & vbcrlf & "  <TR>" & vbcrlf &_
		"    <TD COLSPAN=""10"" VALIGN=TOP>" & vbcrlf &_
		"      <TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR=""#F3F3F3"">" & vbcrlf &_
		"        <TR>" & vbcrlf
		
	
	' Thumbnail (note this is totally absent in text/plain)
	if len(dbRS("thumbnail")) > 1 then
		hBody = hBody & "          <TD WIDTH=""1%"" ALIGN=""CENTER"" VALIGN=""MIDDLE"">" &_
			vbcrlf & "                <IMG SRC=""http://support.fluke.com/" &_
			sSiteName & "/" & dbRS("thumbnail") & """ WIDTH=80 BORDER=1>" & vbcrlf
	Else
		hBody = hBody & "          <TD WIDTH=""80"" ALIGN=""CENTER"" VALIGN=""MIDDLE"">" &_
			"&nbsp;" & vbcrlf
	End if
	
	' Description
	hBody = hBody & "          </TD>" & vbcrlf &_
		"          <TD WIDTH=""90%"" VALIGN=""TOP"">" & vbcrlf &_
		"<FONT Style=""" & Tclass("small") & """>" &_
		sDescription
	
	' Confidential and Embrago
	if Confidential then
		hBody = hBody & "<BR><BR><FONT COLOR=""red"">" & sConfMesg & "</FONT>" & vbcrlf
	end if
	
	if embargo then
		hBody = hBody & "<BR><BR><FONT COLOR=""red"">" & sEmbMesg & "</FONT>" & vbcrlf
	End if

	' file_name, archive_name, and link are all similar
	if have_file then
		hBody = hBody & "<BR><BR><A HREF=""" & Locator & Mk_Locator(Userid,aid,3,Dcode,Lang_id) & """>" &_
      Translate("View",Llang,dbConnSiteWide) & "</A>&nbsp;&nbsp;" &_
      "(" &_
      Translate("File Size",Llang,dbConnSiteWide) & ": " &	SizeFormat(Clng(f_size),Llang,dbConnSiteWide) & ", " &_
      Translate("Download Time",Llang,dbConnSitewide) & ": " & DTime(f_size,bps) & ", " &_
			Translate("Language",Llang,dbConnSiteWide) & ": " & sALang & ")" & vbcrlf
	End if
		
	if have_zip then
		hBody = hBody & "<BR><BR><A HREF=""" & Locator & Mk_Locator(Userid,aid,4,Dcode,Lang_id) & """>" &_
			Translate("Download",Llang,dbConnSiteWide) & "</A>&nbsp;&nbsp;" &_
      "(" &_
      Translate("File Size",Llang,dbConnSiteWide) & ": " &	SizeFormat(Clng(z_size),Llang,dbConnSiteWide) & ", " &_
      Translate("Download Time",Llang,dbConnSitewide) & ": " & DTime(z_size,bps) & ", " &_
			Translate("Language",Llang,dbConnSiteWide) & ": " & sALang & ")" & vbcrlf      
	End if
	
	' "Email to Me" choice, favor the file
	if have_file OR have_zip then
		hBody = hBody & "<BR><BR><A HREF=""" & Locator &_
			Mk_Locator(Userid,aid,5,Dcode,Lang_id) & """>" &_
			 Translate("Send this to me by email",Llang,dbConnSiteWide) &_
			 "</A>" & vbcrlf
	End if
	
	if have_link then
		if conf_link = 1 then
			hBody = hBody & "<BR><BR><A HREF=""" & Locator &_
				Mk_Locator(Userid,aid,7,Dcode,Lang_id) & """>" &_
				Translate("Link",Llang,dbConnSiteWide) & "</A>" & vbcrlf
		elseif conf_link = 2 then
			hBody = hBody & "<BR><BR><A HREF=""" & Locator &_
				Mk_Locator(Userid,aid,8,Dcode,Lang_id) & """>" &_
				Translate("View",Llang,dbConnSiteWide) & "</A>" & vbcrlf
		Else
			hBody = hBody & "<BR><BR><A HREF=""" & dbRS("link") & """>" &_
				Translate("Link",Llang,dbConnSiteWide) & "</A>" & vbcrlf
		End if
	End if
	
	if have_container then				
		hBody = hBody & "<BR><BR><A HREF=""" & Locator &_
		    Lg_Locator(Userid,aid,9,Dcode,Lang_id,acode,acatid) & """>" &_
			Translate("View",Llang,dbConnSiteWide) & "</A>" & vbcrlf
	End if
	
	' Location includes printing begin and end dates
	if len(dbRS("location")) > 1 then
		hBody = hBody & "<B>" &_
			Translate("Location",Llang,dbConnSiteWide) & "</B>: " &_
			dbRS("location") & vbcrlf
	End if
  
	hBody = hBody & "          </TD>" & vbcrlf &_
		"        </TR>" & vbcrlf &_
		"      </TABLE></TD>" & vbcrlf & "     </TR>" & vbcrlf
    
  ' do a little trick to help ensure <A> tags don't get mangled
  hBody = Replace(hBody,"</A>",vbcrlf & "</A>")
end sub