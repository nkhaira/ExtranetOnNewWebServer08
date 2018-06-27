<% 

' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/1/2000
'             The Big Momma
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

'Response.Buffer = True

' --------------------------------------------------------------------------------------
' For non-SSL Servers, we do not need to switch to SSL mode for certain navigation items
' --------------------------------------------------------------------------------------

Dim Border_Toggle
Border_Toggle = 0
Dim HomeURL
Dim BackURL
Dim BackURLSecure

Dim Page_Timer_Begin

Dim sDomain, sNodes, sMyEnvironment

sDomain        = UCase(request.ServerVariables("SERVER_NAME"))
sNodes         = split(sDomain,".")
sMyEnvironment = ""



for each node in sNodes
 	select case node
		case "DEV", "TST", "TEST", "PRD", "DTMEVTVSDV15", "DTMEVTVSDV18"
			sMyEnvironment = node
			exit for
	end select
next

DomainURL = request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("Query_String")

if isblank(Session("SSL")) then
  Session("SSL") = "off"
end if

select case sMyEnvironment
	case "DTMEVTVSDV15", "DEV", "PRD"
    HomeURL            = "Default.asp"
    BackURL            = "http://"  & DomainURL
    BackURLSecure      = "http://"  & DomainURL
	case else
    HomeURL            = "Default.asp"
    BackURL            = "http://"  & DomainURL
    BackURLSecure      = "https://" & DomainURL
end select

' Contact Us, Profile - switch to SSL

if (request("CID") = "9024" or request("CID") = "9025" or request("CID") = "9026") and Session("SSL") = "off" then
  Session("SSL") = "on"
  response.redirect BackURLSecure & "&SSL=" & Session("SSL")
elseif (request("CID") <> "9024" and request("CID") <> "9025" and request("CID") <> "9026") and Session("SSL") = "on" then
  Session("SSL") = "off"
  response.redirect BackURL & "&SSL=" & Session("SSL")
end if

Session("BackURL")       = BackURL
Session("BackURLSecure") = BackURLSecure
Session("ErrorString")   = ""
Session("Server_Name")   = request.ServerVariables("Server_Name")

' --------------------------------------------------------------------------------------

Set Session("rs")  = nothing
Set Session("EquivRS") = nothing

Dim Site_ID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSite_ID_File = objFSO.OpenTextFile(Server.MapPath("Site_ID.dat"))
Site_ID = CInt(objSite_ID_File.ReadLine)
objSite_ID_File.close

' --------------------------------------------------------------------------------------
' Used by the Translation Function to highlight translations.
' --------------------------------------------------------------------------------------

if Request("Language") = "XON" then
  Session("ShowTranslation") = True
elseif Request("Language")="XOF" then
  Session("ShowTranslation") = False
end if  
' --------------------------------------------------------------------------------------

Dim Page_QS
Dim Record_Count
Dim Record_Limit
Dim Record_Pages
Dim ltEnabled

Record_Limit  = 10

Dim SortBy
Dim Show_Detail
Dim Show_Thumbnail
Dim Show_Title_Append
Dim Icon_Type

Dim Show_Days             ' Headlines
Show_Days   = 30          ' Headline History

Dim Product
Dim Category

Dim Access_Level
Dim Access_Level_Title(9) ' Value Should be Highest Access Level Account (i.e., Domain Administrator=9)

Access_Level = 0          ' Determined from Login and SiteWide-DB

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Dim Show_Date             ' System Date PST
Show_Date = True

Dim Site_Code
Dim Site_Timeout
Dim Site_Description
Dim Shopping_Cart         ' Site Shopping Cart Enabled Flag
Dim Shopping_Cart_R1      ' Regional Shopping Cart Enabled Flag
Dim Shopping_Cart_R2      ' Regional Shopping Cart Enabled Flag
Dim Shopping_Cart_R3      ' Regional Shopping Cart Enabled Flag                                                               
Dim Shopping_Cart_Country ' Regional Shopping Cart Disable Flag by ISO Country Code [Array]
Dim Order_Inquiry         ' Site Order Inquiry Enabled Flag
Dim Order_Entry           ' Site Order Entry Enabled Flag
Dim Price_Delivery        ' Site Order Inquiry Enabled Flag

Dim TempString
Dim KeySearch
Dim BolSearch
Dim ItemSearch
Dim SearchDB
KeySearch  = ""
BolSearch  = 0
SearchDB   = 0
ItemSearch = False

Dim CID
Dim SCID
Dim SCID_Campaign
Dim PCID
Dim CIN
Dim CINN
Dim SAID        ' Asset ID Number used with Send_IT

CID      = CInt(request("CID"))  ' Category ID Number


if CID   = Null or CID = 0 then CID = 9000
SCID    = CInt(request("SCID")) ' Sub-Category String
if SCID = Null or SCID = 0 then
  SCID = 0
  SCID_Campaign = 0
else
  SCID_Campaign = SCID
  SCID=0
end if    
PCID     = CInt(request("PCID")) ' Page Sequence
if PCID  = Null or PCID = 0 then PCID = 0

if not isblank(request("CIN")) and isnumeric(request("CIN")) then
  CIN = CInt(request("CIN"))    ' Content Category Code
else
  CIN = 0
end if

if not isblank(request.form("CINN")) and isnumeric(request.form("CINN")) then
  CINN = CInt(request.form("CINN"))  ' Content Category ID Number
elseif not isblank(request("CINN")) and isnumeric(request("CINN")) then
  CINN = CInt(request("CINN"))  ' Content Category ID Number
else
  CINN = 0
end if

if not isblank(request.form("KeySearch")) then
  KeySearch = request.form("KeySearch")
  BolSearch = request.form("BolSearch")
elseif not isblank(request("KeySearch")) then
  KeySearch = request("KeySearch")
  BolSearch = request("BolSearch")
else
  KeySearch = ""
  BolSearch = 0
end if

if not isblank(request.form("SearchDB")) then
  SearchDB = request.form("SearchDB")
elseif not isblank(request("SearchDB")) then
  SearchDB = request("SearchDB")
else
  SearchDB = 0
end if

if not isblank(request("Show_Detail")) and isnumeric(request("Show_Detail")) then
  Show_Detail = CInt(request("Show_Detail"))  ' Current or Archive
else
  Show_Detail = -1
end if    

if not isblank(request("Show_Days")) and isnumeric(request("Show_Days")) then
  Show_Days = Abs(CInt(request("Show_Days")))
  if Show_Days < 7 or Show_Days > 60 then
    Show_Days = 7
  end if
else
  Show_Days = 7
end if

' Audio / Video Media

if not isblank(request("Media")) then
  Dim Media
  Dim Media_Title
  Dim Media_Description
  Dim Media_Extension

  Media = request("Media")
  Media_Title = request("Media_Title")
  Media_Title = request("Media_Description")
  Media_Type = UCase(Mid(Media, InstrRev(Media, ".") + 1))
end if

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_FormData.asp"-->
<%

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Check for Automatic BUTTON_URL Navigation Redirects
' --------------------------------------------------------------------------------------

if not isblank(request("SortBy")) and isnumeric(request("SortBy")) then
  SortBy = CInt(request("SortBy"))' Sort By
elseif CIN > 0 then
  SQLSort = "SELECT SortBy FROM Calendar_Category WHERE Site_ID=" & Site_ID & " AND Code=" & CIN
  Set rsSort = Server.CreateObject("ADODB.Recordset")
  rsSort.Open SQLSort, conn, 3, 3
  if not rsSort.EOF then
    SortBy = rsSort("SortBy")
  else
    SortBy = 1    
  end if
  rsSort.close
  set rsSort = nothing
elseif CID = 9004 and isblank(SortBy) then
  SortBy = 2   
else
  SortBy = 1
end if    

SQL = "Select Navigation.* from Navigation WHERE Navigation.Site_ID=" & Site_ID
Set rsRedirect = Server.CreateObject("ADODB.Recordset")
rsRedirect.Open SQL, conn, 3, 3

do while not rsRedirect.EOF
  if CInt(CID) = CInt(rsRedirect("Button")) and not isblank(rsRedirect("Button_URL")) and rsRedirect("Button_Enable") = CInt(True) then
    Auto_Redirect = rsRedirect("Button_URL")
    rsRedirect.close
    set rsRedirect = nothing
    Call Disconnect_SiteWide
    response.redirect Auto_Redirect
  end if
  rsRedirect.MoveNext
loop

rsRedirect.close
set rsRedirect = nothing

' --------------------------------------------------------------------------------------
' Determine Login Credentials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

if isblank(Session("LOGON_USER")) and not isblank(Request.ServerVariables("LOGON_USER")) then
  Login_Name = Request.ServerVariables("LOGON_USER")
elseif not isblank(Session("LOGON_USER")) then
  Login_Name = Session("LOGON_USER")
else
  Login_Name = ""
end if

do while instr(1,Login_Name,"\") > 0  %><%
  Login_Name = mid(Login_Name,instr(1,Login_Name,"\")+1) %><%
loop  

if isblank(Login_Name) then 
  if UCase(Site_Logon_Method) = "DB" and not isblank(CStr(Site_ID)) then
    Session("ErrorString") = "<LI>" & Translate("Your session has expired.",Login_Language,conn) & " " & Translate("For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please logon below.",Login_Language,conn) & "</LI>"
    with response
      .write "<HTML>" & vbCrLf
      .write "<HEAD>" & vbCrLf
      .write "<TITLE>Account Not Verified</TITLE>" & vbCrLf
      .write "</HEAD>" & vbCrLf
      .write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
      .write "<FORM ACTION=""" & "/register/login.asp"" METHOD=""POST"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & BackURL & """>" & vbCrLf
    end with
    if Login_Language <> "eng" and Login_Language <> Session("Language") then
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
    end if
    with response
      .write "</FORM>" & vbCrLf
      .write "</BODY>" & vbCrLf
      .write "</HTML>" & vbCrLf
    end with
  else   
    Session("ErrorString") = "<LI>" & Translate("Your session has expired.",Login_Language,conn) & " " & Translate("For your protection, you have been automatically logged off of your extranet site account.",Login_Language,conn) & "</LI><LI>" & Translate("To establish another session, please type in the site's code in the &quot;Name of the Site where you want to go&quot;, then click on [ Login ] or",Login_Language,conn) & "</LI><LI>" & Translate("Use the Site Search feature below.",Login_Language,conn) & "</LI>"
    Call Disconnect_SiteWide
    response.redirect "/register/default.asp?backurl=" & BackURL
  end if
  
elseif CID = 9999 then

  SQLLogoff = "SELECT URL_Page_Logoff FROM Site WHERE ID=" & Site_ID

  Set rsLogoff = Server.CreateObject("ADODB.Recordset")
  rsLogoff.Open SQLLogoff, conn, 3, 3
  
  URL_Page_Logoff = ""
  
  if not rsLogoff.EOF then
    if not isblank(rsLogoff("URL_Page_Logoff")) then
      URL_Page_Logoff = rsLogoff("URL_Page_Logoff")
      rsLogoff.close
      set rsLogoff = nothing
      Call Disconnect_SiteWide
      response.redirect URL_Page_Logoff
    end if
  end if  
  
  rsLogoff.close
  set rsLogoff = nothing
  
  Session("ErrorString") = Site_Description & " - " & Translate("Logoff Completed",Login_Language,conn)
  Call Disconnect_SiteWide
  response.redirect "/register/default.asp?backurl=" & BackURL

end if  
'-----------------------------
if cint(site_id) =3 then
	if cstr(trim(request("CINN"))) = "117" then
		PriceListLink=false
		session("PriceListLink")= false
		if isblank(Login_Name)= false then
			sqlRegion = "select  region from userdata where ntlogin like '%" & Login_Name & "%' and site_id=" & Site_ID & " and region=1"
			Set rsRegion = Server.CreateObject("ADODB.Recordset")
			rsRegion.Open sqlRegion, conn, 3, 3
				
			if (rsRegion.EOF=false) then
				showPriceListLink = true
				session("PriceListLink")= true
			end if  
			set  rsRegion = nothing
		end if
		if session("PriceListLink") = false or session("PriceListLink")="" then
			if isblank(Login_Name)= false then
				response.write "<HTML>" & vbCrLf
				response.write "<HEAD>" & vbCrLf
				response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
				response.write "<TITLE>Error</TITLE>" & vbCrLf
				response.write "</HEAD>"
				response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
				response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """>" & vbCrLf
				response.write "<DIV ALIGN=CENTER>"
				Call Nav_Border_Begin
				response.write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
				Response.Write "You are not authorized to view this page."
				response.write "</TD></TR></TABLE>" & vbCrLf
				Call Nav_Border_End
				response.write "</DIV>"
				response.write "</BODY>"
				response.write "</HTML>"
				on error goto 0
			end if
		end if
	end if 
end if
'-----------------------------

SQL =  "SELECT UserData.* FROM UserData WHERE UserData.NTLogin='" & Login_Name & "' AND Site_ID=" & Site_ID & " AND NewFlag=0"

Set rsLogin = Server.CreateObject("ADODB.Recordset")
rsLogin.Open SQL, conn, 3, 3
  
if rsLogin.EOF then

  if UCase(Site_Logon_Method) = "DB" and not isblank(CStr(Site_ID)) then
    with response
      .write "<HTML>" & vbCrLf
      .write "<HEAD>" & vbCrLf
      .write "<TITLE>Account Not Verified</TITLE>" & vbCrLf
      .write "</HEAD>" & vbCrLf
      .write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
      .write "<FORM ACTION=""" & "/register/login.asp"" METHOD=""POST"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & BackURL & """>" & vbCrLf
    end with
    if Login_Language <> "eng" and Login_Language <> Session("Language") then
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
    end if
    with response
      .write "</FORM>" & vbCrLf
      .write "</BODY>" & vbCrLf
      .write "</HTML>" & vbCrLf
    end with
  else   
    Call Disconnect_SiteWide
    response.redirect "/register/default.asp?backurl=" & BackURL
  end if

elseif not rsLogin.EOF then		

  Session("Site_Code") = Site_Code
  Session("Site_Description") = Site_Description
  Session("Site_ID") = Site_ID
  Session("LOGON_USER") = Login_Name

  Login_ID             = rsLogin("ID")
  Login_FirstName      = rsLogin("FirstName")
  Login_MiddleName     = rsLogin("MiddleName")
  Login_LastName       = rsLogin("LastName")
  Login_Company        = rsLogin("Company")
  Login_City           = rsLogin("Business_City")
  Login_Country        = rsLogin("Business_Country")
  Login_Region         = rsLogin("Region")
  Login_EMail          = rsLogin("EMail")
  Login_Business_Phone = rsLogin("Business_Phone")
  Login_Business_Phone_Extension = rsLogin("Business_Phone_Extension")
  Login_SubGroups      = rsLogin("SubGroups")
  Login_Language       = rsLogin("Language")
  Login_Fcm            = rsLogin("Fcm")
  Login_Fcm_ID         = rsLogin("Fcm_ID")
  Login_Groups_Aux     = rsLogin("Groups_Aux")
  Login_Type_Code      = rsLogin("Type_Code")
  
  Session("Login_Region")  = Login_Region
  Session("Login_Country") = Login_Country
  'Modified by Zensar for RI 506.
  Session("PriceListCode") = rsLogin("Pricelist_Code") & ""
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  ' Modify Shopping Cart Accessibility based on user's region or country exclusion
  if Shopping_Cart = CInt(True) then
     
    select case Login_Region
      case 1  ' US
        Shopping_Cart = Shopping_Cart_R1
      case 2  ' Europe
        Shopping_Cart = Shopping_Cart_R2
      case 3  ' Intercon
        Shopping_Cart = Shopping_Cart_R3
    end select

    if instr(1,LCase(Shopping_Cart_Country),LCase(Login_Country)) > 0 then
      Shopping_Cart = 0
    end if
  
  end if

  ' Preferred Language
  %>
  <!--#include virtual="/SW-Common/Preferred_Language.asp"-->
  <%

  ' Check Account Expiration if expired redirect
  
  if isdate(rsLogin("ExpirationDate")) then

    if rsLogin("ExpirationDate") <= Date then
      
      ' Check to see if there is an Account Administrator Assigned  

      SQL =       "SELECT Approvers_Account.* "
      SQL = SQL & "FROM Approvers_Account "
      SQL = SQL & "WHERE Approvers_Account.Site_ID=" & Site_ID & " "
      SQL = SQL & "AND Approvers_Account.Region=" & Login_Region & " "
      SQL = SQL & "AND Approvers_Account.Approver_ID<>0"

      Set rsApprovers = Server.CreateObject("ADODB.Recordset")
      rsApprovers.Open SQL, conn, 3, 3

      AA_Flag = False
      if not rsApprovers.EOF then
        SQL       = "SELECT UserData.* "
        SQL = SQL & "FROM UserData "
        SQL = SQL & "WHERE ID=" & rsApprovers("Approver_ID")
        Set rsUser = Server.CreateObject("ADODB.Recordset")
        rsUser.Open SQL, conn, 3, 3

        if not rsUser.EOF then
          AA_Flag  = True
          AA_Name  = rsUser("FirstName") & " " & rsUser("LastName")
          AA_Email     = rsUser("Email")
        end if  

        rsUser.Close
        set rsUser = nothing
      end if
      
      rsApprovers.Close
      set rsApprovers = nothing
    
      Session("ErrorString") = "<B>" & Translate("Your account has expired.",Login_Language,conn) & "</B>&nbsp;&nbsp;" & Translate("Please contact your Account Administrator:",Login_Language,conn) & " "

      ' No Approvers then Default to Site Admin
      
      if not AA_Flag then       ' No Approvers then Default to Site Admin
        Session("ErrorString") = Session("ErrorString") & "<A HREF=""mailto:" & Site_Admin_Email & Server.URLEncode(Replace("?Subject=Account Expiration - User Re-Instatement Request" & "&body=Site: " & Replace(Site_Description,"Partner Portal - ","") & "%0D%0A" & "ID: " & Login_ID & "%0D%0A" & "Name: " & Login_FirstName & " " & Login_LastName & "%0D%0A" & "Company: " & Login_Company," ","%20")) & """>" & Site_Admin_Name & "</A> " & Translate("by clicking on the Site Administrator&acute;s name.",login_language,conn)
      else
        Session("ErrorString") = Session("ErrorString") & "<A HREF=""mailto:" & AA_Email & Server.URLEncode(Replace("?Subject=Account Expiration - User Re-Instatement Request" & "&body=Site: " & Replace(Site_Description,"Partner Portal - ","") & "%0D%0A" & "ID: " & Login_ID & "%0D%0A" & "Name: " & Login_FirstName & " " & Login_LastName & "%0D%0A" & "Company: " & Login_Company," ","%20")) & """>" & AA_Name & "</A> " & Translate("by clicking on the Account Administrator&acute;s name.",login_language,conn)
      end if  
      response.redirect "/register/default.asp?backurl=" & BackURL  
    end if
  end if          
  
  if CInt(rsLogin("FCM")) = true then
    Access_Level = 1
    Access_Level_Title(Access_Level) = "Account Manager"
  end if
  
  if instr(lcase(rsLogin("SubGroups")),"submitter") > 0 then
    Access_Level = 2
    Access_Level_Title(Access_Level) = "Content Submitter"
  end if
  
  if instr(lcase(rsLogin("SubGroups")),"content") > 0 then
    Access_Level = 4
    Access_Level_Title(Access_Level) = "Content Administrator"
  end if
    
  if instr(lcase(rsLogin("SubGroups")),"account") > 0 then
    Access_Level = 6
    Access_Level_Title(Access_Level) = "Account Administrator"
  end if
  
  if instr(lcase(rsLogin("SubGroups")),"administrator") > 0 then
    Access_Level = 8
    Access_Level_Title(Access_Level) = "Site Administrator"
  end if
  
  if instr(lcase(rsLogin("SubGroups")),"domain") > 0 or lcase(Login_Name) = "whitlock" then
    Access_Level = 9
    Access_Level_Title(Access_Level) = "Domain Administrator"    
  end if 
  
  ' Get FCM Information for Contact Us EMail
  
  if not isblank(Login_Fcm_ID) then
    
    SQL =  "SELECT UserData.* FROM UserData WHERE UserData.ID=" & Login_Fcm_ID & ""
    Set rsLogin = Server.CreateObject("ADODB.Recordset")
    rsLogin.Open SQL, conn, 3, 3
    if not rsLogin.EOF then
      Fcm_Name  = rsLogin("FirstName") & " " & rsLogin("LastName")
      Fcm_EMail = rsLogin("EMail")
    else
      Fcm_Name  = ""
      Fcm_EMail = ""
      Fcm_ID    = 0
    end if
        
    rsLogin.close
    set rsLogin = nothing
  
  end if
    
  if Access_Level >= 2 then
    Session.Timeout = CInt(Site_Timeout) * 2
  else
    Session.Timeout = Site_Timeout
  end if

  ' --------------------------------------------------------------------------------------
  ' Filter out Content Items that are in Review for Regular User
  ' --------------------------------------------------------------------------------------

  if (Access_Level <= 2 or Access_Level = 6) then ' Filter
    Select Case Show_Detail
      case -2, -1, 1, 2
      case else
        Show_Detail = -1
    end select    
  end if
  
  ' --------------------------------------------------------------------------------------
  ' Get User's SubGroup Membership Array
  ' --------------------------------------------------------------------------------------
  
  Dim UserSubGroups
  Dim UserSubGroups_Max
  
  UserSubGroups = Split(Login_SubGroups,", ")
  UserSubGroups_Max = Ubound(UserSubGroups)
  
  for x = 0 to UserSubGroups_Max
    UserSubGroups(x) = Trim(UserSubGroups(x))
  next
  
  ' --------------------------------------------------------------------------------------
  ' Configuration Site Title and Navigation Bar in Preferred Lanaguage or default to English
  ' --------------------------------------------------------------------------------------
  
  Dim Button_Max
  Dim Button(30)            ' Change value to match SITEWIDE - Navigation Table Record total
  Dim Button_Title(30)      ' Change value to match SITEWIDE - Navigation Table Record total
  Dim Button_Help(30)       ' Change value to match SITEWIDE - Navigation Table Record total
  
  SQL =       "SELECT Navigation.Site_ID, Navigation.Order_Num, Navigation.Button, Navigation.Button_Enable, Navigation.Button_URL, Navigation.Button_ENG "
  SQL = SQL & "FROM Navigation "
  SQL = SQL & "WHERE Navigation.Site_ID=" & Site_ID & " "
  SQL = SQL & "ORDER BY Navigation.Order_Num"
  
  'response.write("sql: " & sql & "<BR>")
  Set rsNavigation = Server.CreateObject("ADODB.Recordset")
  rsNavigation.Open SQL, conn, 3, 3
  
  i = -1
  
  if not rsNavigation.EOF then 
  	do while not rsNavigation.eof
  	  i = i + 1
  	  if  CInt(rsNavigation("Button_Enable")) = False  or isblank(rsNavigation("Button"))  or isblank(rsNavigation("Button_ENG")) then
  	    Button(i)       = ""
  	    Button_Title(i) = ""
  	    Button_Help(i)  = ""
  	  elseif i = 29 then
        Button(i)       = rsNavigation("Button")
        Button_Title(i) = Translate(Access_Level_Title(Access_Level),Login_Language,conn)
  	    Button_Help(i)  = Access_Level_Title(Access_Level)
      else    
  	    Button(i)       = rsNavigation("Button")
        Button_Title(i) = Translate(rsNavigation("Button_ENG"),Login_Language,conn)
  	    Button_Help(i)  = rsNavigation("Button_ENG")
  	  end if  
  	  
  	  rsNavigation.MoveNext
  	  
  	loop
  end if 
  
  Button_Max = i
    
  rsNavigation.close
  set rsNavigation = nothing
  
  ' --------------------------------------------------------------------------------------
  ' Set initial Status of Shopping Cart
  ' --------------------------------------------------------------------------------------
  
  if CInt(Shopping_Cart) = CInt(True) then
    if CInt(Session("Cart_Active")) = -2 then   ' Initally set in Global.ASA - Do once for Session
      SQL = "SELECT Count(*) AS Items FROM Shopping_Cart_Lit WHERE Account_NTLogin='" & Login_Name & "' AND Submit_Date IS NULL"
      Set rsItems = Server.CreateObject("ADODB.Recordset")
      rsItems.Open SQL, conn, 3, 3
      if rsItems("Items") > 0 then
        Session("Cart_Active") = CInt(True)
      else
        Session("Cart_Active") = CInt(False)
      end if
      rsItems.close
      set rsItems = nothing
    end if
  end if  
  
  ' --------------------------------------------------------------------------------------
  ' Fetch Content Table (SiteWide.Calendar) Field Names Array
  ' --------------------------------------------------------------------------------------  
  %>
  <!--#include virtual="/SW-Common/SW-Field_Names.asp"-->
  <%
  
  ' --------------------------------------------------------------------------------------
  
  Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Extranet Support Site", Alt_Language, conn)
  Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><SPAN CLASS=SmallBoldBar>" & Translate("Extranet Support Site", Login_Language, conn) & "</SPAN>"
  
  if lcase(request("NS")) = "true" then
      Top_Navigation  = True
  elseif lcase(request("NS")) = "false" then
      Top_Navigation  = False
  else
    Top_Navigation = False
  end if
    
  Side_Navigation = True
  Content_Width   = 95  ' Percent
  
  ' Check to see if Site is Open or Closed
  
  if Site_Closed = True then
    Call Disconnect_SiteWide
    response.redirect ("/Register/Site_Closed.asp?Language=" & Login_Language & "&Site_ID=" & Site_ID )
  end if
  
  ' --------------------------------------------------------------------------------------
  ' Begin Main
  ' --------------------------------------------------------------------------------------

  %>  
  <!--#include virtual="/SW-Common/SW-Header.asp"-->  
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%

  ' --------------------------------------------------------------------------------------
  ' Home Page - 9000

	 ' --------------------------------------------------------------------------------------
	 
	
  
  if CID = Button(0) then
  
    ' --------------------------------------------------------------------------------------
    ' Update Records for Archive
    ' --------------------------------------------------------------------------------------

    SQL = "UPDATE dbo.Calendar " &_
          "SET    Status = 2, Status_Comment = 'An asset container archive date has been reached for this Item Number.' " &_
          "WHERE (Status = 1 OR Status = 2) AND (BDate = EDate) AND (XDays > 0) AND (XDate < '" & Date() & "') " &_
          "   OR (Status = 1 OR Status = 2) AND (BDate <> EDate) AND (XDate < '" & Date() & "')"
    conn.execute SQL

    ' --------------------------------------------------------------------------------------

    Home_Page_Text = ""

    ' Special Condition -- If Home Page Text exists by Regional Definition then select from Home_Page Table, else from Standard Calendar_Category description

    SQL = "SELECT Text_Home_Page.Text_Region_" & Trim(CStr(Login_Region)) & " FROM Text_Home_Page WHERE Text_Home_Page.Site_ID=" & CInt(Site_ID)
    Set rsHome_Page = Server.CreateObject("ADODB.Recordset")
    rsHome_Page.Open SQL, conn, 3, 3

    if not rsHome_Page.EOF then
      Home_Page_Text = RestoreQuote(rsHome_Page("Text_Region_" & Trim(CStr(Login_Region))))
    end if

    rsHome_Page.close
    set rsHome_Page = nothing
    
    SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Code=9999"
    Set rsCategory = Server.CreateObject("ADODB.Recordset")
    rsCategory.Open SQL, conn, 3, 3

    response.write "<SPAN CLASS=Heading3>" & Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</SPAN><BR>"
    if not isblank(Login_MiddleName) then
      response.write Login_FirstName & " " & Login_MiddleName & " " & Login_LastName
    else
    response.write Login_FirstName & " " & Login_LastName
    end if
    response.write "<BR><BR>"

    if isblank(Home_Page_Text) then   
      Home_Page_Text = RestoreQuote(rsCategory("Description"))
    end if  

    rsCategory.close
    set rsCategory = nothing
  
    Headlines = True
          
    if Headlines = True then
      call Translate_Include(Translate_Embedded(Home_Page_Text,Login_Language,conn))
    else  
      response.write Translate_Embedded(Home_Page_Text,Login_Language,conn)
    end if
  
  ' --------------------------------------------------------------------------------------
  ' What's New Page - 9001
  ' -------------------------------------------------------------------------------------- 
    
  elseif CID = Button(1) then
  
    response.write "<SPAN CLASS=Heading3>"
    response.write Button_Title(1)
    response.write "</SPAN><BR><BR>"

  
    %>
    <!--#include virtual="SW-Common/SW-WhatsNew.asp"-->
    <%
  
  ' --------------------------------------------------------------------------------------
  ' Calendar and Library Page - 9002 and 9003
  ' --------------------------------------------------------------------------------------

  elseif CID = Button(2) or CID = Button(3) then
   
      if CIN = 0 Then
  
        response.write "<SPAN CLASS=Heading3>"
        select case CID
          case Button(2)
            response.write Button_Title(2)
          case Button(3)
            response.write Button_Title(3)
        end select
        response.write "</SPAN><BR><BR>"
       
        if CID = Button(3) then
          'response.write Translate("Please Select one of the Categories from the Navigation Buttons on the left",Login_Language,conn) & ".<P>"
          %><!--#include virtual="/sw-common/SW-Library_Text.asp"--><%
        else      
          response.write Translate("Please Click on a Calendar Item or Select one of the Categories from the Navigation Buttons on the left",Login_Language,conn) & ".<P>"
        %>     
        <!--#include virtual="/include/Include_Event_Calendar.asp"-->
        <%
        end if
        
      else
        if SCID_Campaign > 0 then
          %><!--#include virtual="/SW-Common/SW-Campaigns.asp"--><%             
        else         
          if CIN >=8000 and CIN <= 8999 then
            if isblank(request("SortBy")) then
              SortBy=2
            end if
          end if    
          %><!--#include virtual="/SW-Common/SW-Content.asp"--><%
        end if
        
      end if
  
  ' --------------------------------------------------------------------------------------
  ' Site Search - 9004
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(4) then
  
        if isblank(KeySearch) then
          CINN = 0
          BolSearch = 0
        end if

        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)
        if CINN > 0 then response.write "&nbsp;" & Translate("Results",Language,conn)
        response.write "</SPAN><P>"

        if CINN > 0 then
          response.write "<SPAN CLASS=SMALLBOLD>" & Translate("Search Keywords",Login_Language,conn) & "</SPAN>: " & "<SPAN CLASS=SMALLBOLDRED>" & Replace(KeySearch,","," ") & "</SPAN>"
          response.write "&nbsp;&nbsp;&nbsp;<SPAN CLASS=Small>["
          select case bolSearch
            case 0
              response.write Translate("All Words",Login_Language,conn)
            case 1  
              response.write Translate("Any Word",Login_Language,conn)
            case 2
              response.write Translate("Exact Phrase",Login_Language,conn)
            case else
              response.write "Invalid bolSearch value"  
          end select
          response.write "]</SPAN>"
          response.write "&nbsp;&nbsp;&nbsp;"
          response.write "<SPAN CLASS=SMALLBOLD>" & Translate("Database",Login_Language,conn) & "</SPAN>: "
          response.write "<SPAN CLASS=SMALL>"
          select case CINN
            case 0, 1
              response.write Translate(Site_Description,Login_Language,conn)
            case 2
              response.write Translate("Literature",Login_Language,conn)
            case 3
              response.write Translate("Manuals",Login_Language,conn)
          end select    
          response.write "</SPAN>" & vbCrLf
          
          if not isblank(Language_Filter) then
            response.write "&nbsp;&nbsp;&nbsp;"
            response.write "<SPAN CLASS=SMALLBOLD>" & Translate("Filter by Language",Login_Language,conn) & "</SPAN>: "
            response.write "<SPAN CLASS=SMALL>"
            SQLLanguage = "SELECT     Description AS Description FROM dbo.[Language] WHERE (Code = '" & Language_Filter & "')"
            Set rsLanguage = Server.CreateObject("ADODB.Recordset")
            rsLanguage.Open SQLLanguage, conn, 3, 3
            if not isblank(rsLanguage("Description")) then
              response.write Translate(rsLanguage("Description"),Login_Language,conn)
            end if
            response.write "</SPAN>" & vbCrLf
          end if
          response.write "<P>" & vbCrLf    
          
        end if  

        for x = 0 to 9                        ' Max Types of Category Searches
          if instr(1,BackURL,"CINN=" & trim(cstr(x))) > 0 then
            BackURL = Replace(BackURL,"CINN=" & trim(cstr(x)),"CINN=" & CINN)
            exit for
          end if  
        next
               
        if CINN = 0 then
          %><!--#include virtual="/SW-Common/SW-Search_Form.asp"--><%
        elseif CINN = 1 then
          %><!--#include virtual="/SW-Common/SW-Content.asp"--><%
        elseif CINN = 2 then
          %><!--#include virtual="/SW-Common/SW-Content.asp"--><% ' Place Holder for future Search Feature
        elseif CINN = 3 then
          %><!--#include virtual="/SW-Common/SW-Content.asp"--><% ' Place Holder for future Search Feature
        elseif CINN = 4 then
          %><!--#include virtual="/SW-Common/SW-Content.asp"--><% ' Place Holder for future Search Feature
        elseif CINN = 5 then
          %><!--#include virtual="/SW-Common/SW-Content.asp"--><% ' Place Holder for future Search Feature
        end if  
      
  ' --------------------------------------------------------------------------------------
  ' Brand Sites - 9005
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(5) then
  
        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)
        response.write "</SPAN><BR><BR>"
  
        SQL =       "SELECT Site_Aux.Site_ID, Site.Site_Code, Site.Enabled, Site.Site_Description, Site.URL, Site.ID, Site.Logo "
        SQL = SQL & "FROM Site_Aux LEFT JOIN Site ON Site_Aux.Site_ID_Aux = Site.ID "
        SQL = SQL & "WHERE Site_Aux.Site_ID=" & Site_ID & " AND Site.Enabled=" & CInt(True) & " ORDER BY Site.Site_Description"
  
        SQL = "SELECT  dbo.UserData.Site_ID, dbo.UserData.NTLogin, dbo.UserData.NewFlag, dbo.UserData.ExpirationDate, dbo.Site.Enabled, " &_
              "        dbo.Site.Site_Description, dbo.Site.Logo, dbo.Site.URL, dbo.Site.Site_Code, dbo.Site.ID " &_
              "FROM    dbo.UserData LEFT OUTER JOIN " &_
              "        dbo.Site ON dbo.UserData.Site_ID = dbo.Site.ID " &_
              "WHERE   (dbo.Site.Enabled = -1) AND (dbo.UserData.NewFlag = 0) AND (dbo.UserData.NTLogin = '" & Login_Name & "') AND " &_
              "        (dbo.UserData.ExpirationDate >= CONVERT(DATETIME, '" & Date() & "', 102)) " &_
              "ORDER BY dbo.Site.Site_Description"
  
        Set rsSite_Aux = Server.CreateObject("ADODB.Recordset")
        rsSite_Aux.Open SQL, conn, 3, 3
            
        if not rsSite_Aux.EOF then
        
          response.write Translate("Your Fluke Support Extranet Passport allows you to login to the following site(s) using your same account user name and password.  Just click one of the links below.",Login_Language,conn) & "<BR><BR>"      
        
          Call Nav_Border_Begin

          response.write "<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=4 BGCOLOR=#666666>" & vbCrLf
          
         	Set fso = CreateObject("Scripting.FileSystemObject")
          
          do while not rsSite_Aux.EOF

            Recriprocity_Site = Replace(LCase(rsSite_Aux("URL")),"http://support.fluke.com","")          

            response.write "<TR><TD CLASS=NavLeftHighlight1>"
            response.write "<A HREF=""" & Recriprocity_Site & """ CLASS=NavLeftHighlight1>"
            ' Use Text Link
            response.write "" & rsSite_Aux("Site_Description")
            response.write "</A>"
            response.write "</TD>" & vbCrLf

            Graphic_Logo = false
            
            if Graphic_Logo then
            
              response.write "<TD CLASS=NavLeftHighlight1>"
              response.write "<A HREF=""" & Recriprocity_Site & """ CLASS=NavLeftHighlight1>"                                      

              ' Use Graphic Logo Link 

              if isblank(rsSite_Aux("Logo")) then
                response.write "<IMG SRC=""/images/FlukeLogo3.gif"" WIDTH=80 HEIGHT=44 BORDER=0>"
              else
                response.write "<IMG SRC=""" & rsSite_Aux("Logo") & """ WIDTH=80 BORDER=0>"
              end if
              response.write "</A>"
              response.write "</TD>" & vbCrLf

            end if
            
            response.write "</TR>" & vbCrLf

            rsSite_Aux.MoveNext

          loop
          
          response.write "</TABLE>" & vbCrLf & vbCrLf

          Call Nav_Border_End
          
        end if
        
        rsSite_Aux.Close
        set rsSite_Aux = nothing    
  
  ' --------------------------------------------------------------------------------------
  ' Forums - 9006
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(6) then
  
        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)
        response.write "</SPAN><BR><BR>"
        
' --------------------------------------------------------------------------------------
  ' WTB Distributor StoreFront - 9009
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(9) then
  
        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)  
        response.write "</SPAN><BR><BR>"        
  
  
  ' --------------------------------------------------------------------------------------
  ' Help - 9023
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(23) then
  
        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)
        response.write "</SPAN><BR><BR>"
        
        %>
        <!--#include virtual="/SW-Common/SW-Help.asp"-->
        <%
  
  ' --------------------------------------------------------------------------------------
  ' Contact Us - 9024
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(24) then
   
      if CIN = 0 or CIN = 1 then
  
        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)
        response.write "</SPAN><BR><BR>"
        
      end if
      
      if CIN = 0 then
      
        ' Edit Email
        
        %>
        <!--#include virtual="/SW-Common/SW-Contact_Us_Edit.asp"-->
        <%
       
      elseif CIN = 1 then
      
      ' Send Email
      
' --------------------------------------------------------------------------------------
' Configure EMail Header Information
' --------------------------------------------------------------------------------------

        Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
        %>
        <!--#include virtual="/connections/connection_EMail_Timeout.asp"-->        
        <!--#include virtual="/SW-Common/SW-Contact_Us_Admin.asp"-->
        <%
      
      end if     
  
  ' --------------------------------------------------------------------------------------
  ' Update Profile - 9025
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(25) then
  
        response.write "<SPAN CLASS=Heading3>"
        response.write Button_Title(CID-9000)
        response.write "</SPAN><BR><BR>"
        
        if CINN=1 then
          response.write "<FONT COLOR=""Red"">" & translate("Your Account Profile has been updated.",Login_Language,conn) & "</FONT><BR>"
        end if
        
        if request("PF") = "-1" then
          response.write "<FONT COLOR=""Red"">" & translate("Your Password Change becomes effective the next time you Logon.",Login_Language,conn) & "</FONT><BR>"
        end if
        response.write "<BR>"  
        
        %>
        <!--#include virtual="/SW-Common/SW-Profile_Edit.asp"-->
        <%
  
  ' --------------------------------------------------------------------------------------
  ' Messages - 9026
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(26) then
  
    response.write "<SPAN CLASS=Heading3>"
    response.write Button_Title(CID-9000)
    response.write "</SPAN><BR><BR>"
    
    ' Hard code something dummy for now and complete it in Phase two.
    
    response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;&nbsp;" & Translate("Message Summary",Login_Language,conn) & "&nbsp;&nbsp;&nbsp;</SPAN><BR><BR>"
    response.write "&nbsp;&nbsp;" & Translate("Inbox",Login_Language,conn) & " 0&nbsp;&nbsp;[<B>0 " & Translate("New",Login_Language,conn) & "</B>]"
  
  ' --------------------------------------------------------------------------------------
  ' Site Statistics - 9028
  ' --------------------------------------------------------------------------------------
  
  elseif CID = Button(28) then
    with Response
      .write "<SPAN CLASS=Heading3>"
      .write Button_Title(CID-9000)
      .write "</SPAN><P>"

      .write Translate("Omiture reports for Fluke websites.",Login_Language,conn) & "<P>"

      Call Nav_Border_Begin
        .write "<A HREF="""" onclick=""Admin_Window=window.open('http://my.omniture.com','Admin_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Admin_Window.focus();return false;"" CLASS=NavLeftHighlight1>"
  	    .write "&nbsp;&nbsp;" & Translate("View",Login_Language,conn) & "&nbsp;&nbsp;</A>"
      Call Nav_Border_End
  	end with

  ' --------------------------------------------------------------------------------------
  ' Administration Pages - 9029
  ' --------------------------------------------------------------------------------------
  
  elseif Access_Level > 0 and CID = Button(29) then
  
    with response
      
      if Access_Level > 0 then
      
        .write "<SPAN CLASS=Heading3>"
        .write Button_Title(CID-9000)
        .write "</SPAN><BR><BR>"
        
        .write Translate("After selecting one of this site&acute;s administration tools listed below, a new browser window will appear. At that point you may be prompted for your user name and password before allowed access to these tools.",Login_Language,conn) & "<BR><BR>"
        .write Translate("Use can use the [Alt] + [Tab] keyboard combination to toggle back and forth beween the live site browser window and the administrator&acute;s window.",Login_Language,conn) & "<BR><BR>"
        .write Translate("To logoff after you are done using the administration tools, just close the browser window.",Login_Language,conn) & "<BR><BR>"
    
        ' Administrator's Tool Kit
         
        select case Access_Level
    
          case 2, 4, 6, 8, 9
    
            ' Link to SW-Administration Tools
    
            .write "<UL>"
            .write "<LI>"
            .write "<SPAN CLASS=NORMALBOLD>" & Translate("Site Administrator",Login_Language,conn) & " - " & Translate("Tool Kit",Login_Language,conn) & "</SPAN><BR><BR>" 
              
            Call Nav_Border_Begin
            .write "<A HREF="""" onclick=""Admin_Window=window.open('/sw-administrator/default.asp?Site_ID=" & Site_ID & "&Language=" & Login_Language & "&Logon_User=" & Login_Name &  "','Admin_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Admin_Window.focus();return false;"" CLASS=NavLeftHighlight1>"
            .write "&nbsp;&nbsp;" & Translate(Access_Level_Title(Access_Level),Login_language,conn) & " - " & Translate("Tool Kit",Login_Language,conn) & "&nbsp;&nbsp;</A>"
            Call Nav_Border_End
            .write "</LI>"
  
          end select
              
          ' Gateway Administrator's Tool Kit
          select case Access_Level
              
            case 2, 4, 8, 9
          
              SQLGateway = "SELECT * FROM Gateway_Applications ORDER BY Gateway_Title"
              Set rsGateway = Server.CreateObject("ADODB.Recordset")
              rsGateway.Open SQLGateway, conn, 3, 3
              
              Gateway_Flag = false
              
              do while not rsGateway.EOF
                if instr(1,Login_SubGroups,lcase(Trim(rsGateway("Gateway_Code")))) > 0 or Access_Level = 9 then
                  Gateway_Flag = true
                  exit do
                end if
                rsGateway.MoveNext
              loop
              
              if Gateway_Flag = true then
                
                .write "<BR><BR>"
                .write "<LI>"
                .write "<SPAN CLASS=NORMALBOLD>" & Translate("Gateway Application Administrator",Login_Language,conn) & " - " & Translate("Tool Kit(s)",Login_Language,conn) & "</SPAN><BR><BR>" 
                
                rsGateway.MoveFirst
                
                do while not rsGateway.EOF
                  Call Nav_Border_Begin
                  .write "<A HREF="""" onclick=""Admin_Window=window.open('" & rsGateway("Gateway_URL") & "?Site_ID=" & Site_ID & "&Language=" & Login_Language & "&Logon_User=" & Login_Name &  "','Admin_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Admin_Window.focus();return false;"" CLASS=NavLeftHighlight1>"
                  .write "&nbsp;&nbsp;" & Translate(Trim(rsGateway("Gateway_Title")),Login_language,conn) & "&nbsp;&nbsp;</A>"
                  Call Nav_Border_End
                  .write "<BR>"
                  rsGateway.MoveNext
                loop
                  
                .write "</LI>"
                
              end if
              
              rsGateway.close
              set rsGateway = nothing
                
          end select
    
          ' Link to Big Brother
          select case Access_Level
            case 8, 9
              .write "<BR>"
              .write "<LI>"
              .write "<SPAN CLASS=NORMALBOLD>" & Translate("IT Portal",Login_Language,conn) & "</SPAN><BR><BR>"
              .write Translate("Use the following link to report bugs, work requests, improvement suggestions, statistics, server monitoring, etc.",Login_Language,conn) & "&nbsp;&nbsp;"               
              .write Translate("To review this information, you must be logged into the Fluke Wide Area Network (WAN). This information is not available when accessing this site through the internet.",Login_Language,conn) & "<BR><BR>"

              'Big_Brother = "http://it.intranet.danahertm.com/default.aspx"
	       Big_Brother = "http://itglobal.intranet.danahertm.com/web/ReleaseManagement/default.aspx"
    
              Call Nav_Border_Begin
              .write "<A HREF="""" onclick=""openit_maxi('" & Big_Brother & "','Vertical');return false;"" CLASS=NavLeftHighlight1>"
              .write "&nbsp;&nbsp;" & Translate("Site Services",Login_language,conn) & "&nbsp;&nbsp;</A>"
              Call Nav_Border_End
              .write "</LI>"
          end select
            
          .write "</UL>"                

      else
        
        response.write Translate("The Site Feature",Login_Language,conn) & " &quot;" & Button_Title(CID-9000) & "&quot; " & Translate("is currently under development",Login_Language,conn) & "."
    
      end if
      
    end with
        
  ' --------------------------------------------------------------------------------------
  ' Out of Range Pages - < 9000 or Greater than ????
  ' --------------------------------------------------------------------------------------
  
  elseif CID < 9000 or CID > (Button_Max + 9000) then
  
    response.write "<SPAN CLASS=Heading3>"
    response.write CID
    response.write "</SPAN><BR><BR>"

    response.write Translate("Unknown Navigation Request.",Login_Language,conn)
    
    CID = 9000
         
  ' --------------------------------------------------------------------------------------
  ' All Other Page's Not Completed
  ' --------------------------------------------------------------------------------------
  
  else
  
    response.write "<SPAN CLASS=Heading3>"
    response.write Button_Title(CID-9000) & " " & CID
    response.write "</SPAN><BR><BR>"
      
    response.write Translate("The Site Feature",Login_Language,conn) & " &quot;" & Button_Title(CID-9000) & "&quot; " & Translate("is currently under development",Login_Language,conn) & "."
  
  end if
  
  ' --------------------------------------------------------------------------------------
  ' Error Handler
  ' --------------------------------------------------------------------------------------
  
  if Err.Number then
    response.write "Error       : " & Err.Number & "<BR>"
    response.write "Line        : " & Err.Line & "<BR>"
    response.write "Description : " & Err.Description & "<BR>"
  end if
  
  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%

  ' --------------------------------------------------------------------------------------
  ' Instant Message Alert
  ' --------------------------------------------------------------------------------------

  SQL = "SELECT * FROM Messages WHERE NTLogin='" & Session("Logon_User") & "' ORDER BY Message_Date"
  Set rsMessage = Server.CreateObject("ADODB.Recordset")
  rsMessage.Open SQL, conn, 3, 3
  
  Message = False
  if not rsMessage.EOF then
          Message = True
      response.write vbCrLf & vbCrLf
      response.write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
      response.write "<!-- //" & vbCrLf
  
      do while not rsMessage.EOF
        response.write "alert('"
        response.write "Message to    : " & rsMessage("To_tName") & "\r\n"
        response.write "Message from: " & rsMessage("Fm_Name") & "\r\n"
        response.write "Sent on: " & rsMessage("Message_Date") & "\r\n\n"
        strMessage = replace(rsMessage("Message"),"&quot;","\""")%><%
        strMessage = replace(strMessage,"'","\'")
        strMessage = replace(strMessage,"&acute;","\'")
        strMessage = replace(strMessage,"<BR>","\r\n")
        response.write strMessage & "\r\n')" & vbCrLf
        rsMessage.MoveNext
      loop
      response.write "// -->" & vbCrLf
      response.write "</SCRIPT>" & vbCrLf & vbCrLf

  end if
  rsMessage.close
  set rsMessage = nothing
  
  if Message = True then
    SQL = "DELETE FROM Messages WHERE NTLogin='" & Session("Logon_User") & "'"
    response.write SQL
    conn.execute SQL
  end if
    
  ' --------------------------------------------------------------------------------------
  ' Send User Requested Content file by email as an attachment if requested from previous screen.
  ' --------------------------------------------------------------------------------------
  
  if not isblank(request("SAID")) then
    Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
    %>
    <!--#include virtual="/connections/connection_EMail_Timeout.asp"-->
    <!--#include virtual="/SW-Common/SW-Send_It.asp"-->
    <%
    Call Send_It
  end if

  Select Case CID
    Case 9001,9002,9003,9004,9029
      %>
      <!--#include virtual="/SW-Common/SW-Content_Subroutines.asp"-->
      <!--#include virtual="/include/Pop-Up.asp"-->
      <!--#include virtual="/include/Pop-Up_Media_Player.asp"-->
      <%
  end select

  %>
  <!--#include virtual="/include/functions_date_formatting.asp"-->
  <!--#include virtual="/include/functions_string.asp"-->
  <!--#include virtual="/include/functions_file.asp"-->
  <!--#include virtual="/include/functions_translate.asp"-->
  <!--#include virtual="/include/functions_table_border.asp"-->  
  <!--#include virtual="/include/functions_preprocess.asp"-->
  <!--#include virtual="/connections/connection_EMail.asp"-->  
  <%
  
  ' Use Session ID to Store Logon Date
  
  SQL = "SELECT Logon.* FROM Logon WHERE Logon.Session_ID=" & Session.SessionID & " AND Logon.Site_ID=" & Site_ID 
  Set rsLogon = Server.CreateObject("ADODB.Recordset")
  rsLogon.Open SQL, conn, 3, 3
  
  if not rsLogon.EOF then
    Session("Session_ID") = Session.SessionID
  end if
  
  rsLogon.close
  set rsLogon = nothing  
  
  if isblank(Session("Session_ID")) then
    Session("Session_ID") = Session.SessionID
    Session("Logon_Date") = Now
    SQL = "INSERT INTO Logon ( Account_ID, Session_ID, Site_ID, Logon, Logoff ) SELECT "
    SQL = SQL & Login_ID & " AS Account_ID, "
    SQL = SQL & session.sessionID & " AS Session_ID, "
    SQL = SQL & Site_ID & " AS Site_ID, "
    SQL = SQL & "'" & Session("Logon_Date") & "'" & " AS Logon, "
    SQL = SQL & "'" & Session("Logon_Date") & "'" & " AS Logoff"
    
    conn.Execute (SQL)
    
    SQL = "Update UserData SET "
    SQL = SQL & "UserData.Logon='" & Session("Logon_Date") & "'"
    SQL = SQL & " WHERE UserData.ID=" & Login_ID
   
    conn.Execute (SQL) 
  
  ' --------------------------------------------------------------------------------------  
  ' Update Logoff Date for current session
  ' --------------------------------------------------------------------------------------
  
  else
  
    Session("Logoff_Date") = Now  
    SQL = "UPDATE Logon SET "
    SQL = SQL & "Logon.Logoff=" & "'" & Session("Logoff_Date") & "'"
    SQL = SQL & " WHERE Logon.Session_ID=" & Session.SessionID
  
    conn.Execute (SQL)
  
  end if
  
  ' --------------------------------------------------------------------------------------
  ' Log Pages Viewed by Accounts
  ' --------------------------------------------------------------------------------------

  ActivitySQL =                                       "INSERT INTO Activity ( Account_ID, Site_ID, Session_ID, View_Time, CID, SCID, PCID, CIN, CINN, Language, Region, Country ) "
  ActivitySQL = ActivitySQL &                         "SELECT "
  ActivitySQL = ActivitySQL & Login_ID              & " AS Account_ID, "
  ActivitySQL = ActivitySQL & Site_ID               & " AS Site_ID, "
  ActivitySQL = ActivitySQL & Session("Session_ID") & " AS Session_ID, "
  ActivitySQL = ActivitySQL & "'" & Now            & "' AS View_Time, "
  ActivitySQL = ActivitySQL & CID                   & " AS CID, "
  ActivitySQL = ActivitySQL & SCID                  & " AS SCID, "
  ActivitySQL = ActivitySQL & PCID                  & " AS PCID, "
  ActivitySQL = ActivitySQL & CIN                   & " AS CIN, "
  ActivitySQL = ActivitySQL & CINN                  & " AS CINN, "
  ActivitySQL = ActivitySQL & "'" & Login_Language  & "' AS Language,"
  ActivitySQL = ActivitySQL & Login_Region          & " AS Region,"
  ActivitySQL = ActivitySQL & "'" & Login_Country   & "' AS Country"
  conn.Execute (ActivitySQL)
  
end if

Call Disconnect_SiteWide

%>
