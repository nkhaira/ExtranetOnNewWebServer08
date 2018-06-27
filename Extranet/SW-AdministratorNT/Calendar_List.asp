<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       06/1/2000
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim HomeURL
Dim BackURL

Dim Site_ID
Site_ID        = request("Site_ID")

Dim SC0
Dim SC1
Dim SC2
Dim SC3
Dim SC4
Dim SC5
Dim Updated_ID
Dim Export_Delimited

Dim Page                           ' Current Display Page
Dim Record_Count                   ' Current Record Count
Dim Record_Pages                   ' Total Pages
Dim Record_Number                  ' Current Record Number
Dim Record_Limit                   ' Maximum Records per Page

Record_Limit = 50
Export_Delimited = 50              ' Action ID Number

' --------------------------------------------------------------------------------------
' Decode QueryString Parameters
' --------------------------------------------------------------------------------------

SC0 = 1
if isnumeric(Request("SC0")) then     ' Utility ID
  if CInt(request("SC0")) > 0 then
    SC0 = CInt(Request("SC0"))
  end if
end if

SC1 = 0
if isnumeric(Request("SC1")) then    ' Sort Criteria 1
  SC1 = CInt(Request("SC1"))
end if

SC2=0
if isnumeric(Request("SC2")) then    ' Sort Criteria 2
  SC2 = CInt(Request("SC2"))
end if

SC3=0
if isnumeric(Request("SC3")) then    ' Sort Criteria 3
  SC3 = CInt(Request("SC3"))
end if

SC4 = ""
if (SC2 = 90 or SC2 = 91) and not isblank(Request("SC4")) then  ' Sort Criteria 4 (Search Key Word)
  SC2 = 91
  SC4 = Trim(Request("SC4"))
elseif SC2 = 90 and isblank(Request("SC4")) then
  SC2 = 90
end if

SC5 = 0
if isnumeric(Request("SC5")) then    ' Sort Criteria 4 (Show Optional Fields)
  SC5 = CInt(Request("SC5"))
end if

Updated_ID = 0
if isnumeric(request("Z")) and instr(1,request("Z"),"#") = 0 then      ' Last Update ID Number
  if CInt(request("Z")) > 0 then
    Updated_ID = CInt(request("Z"))
  end if
elseif instr(1,request("Z"),"#") > 1 then
  Updated_ID = CInt(mid(request("Z"),1,instr(1,request("Z"),"#") - 1))
end if

if isnumeric(Request("Page")) then   ' Page
    Page = CInt(Request("Page"))
end if
if Page < 1 then Page = 1    

' --------------------------------------------------------------------------------------
' Redirect for Special Events
' --------------------------------------------------------------------------------------

select case SC2
  case 11   ' Content Administrator Matrix
    if SC0 <> 3 then
      response.redirect "/SW-Administrator/Account_List.asp?Site_ID=" & Site_ID & "&SC0=3" & "&SC2=" & SC2
    end if
  case 13   ' Account Administrator Matrix
    if SC0 <> 4 then
      response.redirect "/SW-Administrator/Account_List.asp?Site_ID=" & Site_ID & "&SC0=4" & "&SC2=" & SC2
    end if
  case 99   ' Add New
     response.redirect "/SW-Administrator/Account_Edit.asp?ID=edit_account&Site_ID=" & Site_ID & "&Account_ID=add" & "&SC0=1" & "&SC1=0" & "&SC2=1" & "&SC3=0" & "&SC4=" & "&SC5=" & SC5 & "&Page=" & Page & "&Z=0"
end select

' --------------------------------------------------------------------------------------
' Page Header Information
' --------------------------------------------------------------------------------------

Dim Bar_Tag
select case SC2
  case 0
    Bar_Tag = Translate("Account",Login_Language,conn) & " - " & Translate("Pending Approval",Login_Language,conn) & " - " & Translate("New",Login_Language,conn)
  case 1
    Bar_Tag = Translate("Account",Login_Language,conn) & " - " & Translate("Updated Today",Login_Language,conn)
  case 7
    Bar_Tag = Translate("Account",Login_Language,conn) & " - " & Translate("Expired",Login_Language,conn)
  case 2
    Bar_Tag = Translate("Account",Login_Language,conn) & " - " & Translate("Fluke",Login_Language,conn) & " - " & Translate("Excluded",Login_Language,conn)
  case 3
    Bar_Tag = Translate("Account",Login_Language,conn) & " - " & Translate("Fluke",Login_Language,conn) & " - " & Translate("Only",Login_Language,conn)
  case 4
    Bar_Tag = Translate("Region 1 - US Only",Login_Language,conn)
  case 5
    Bar_Tag = Translate("Region 2 - Europe Only",Login_Language,conn)
  case 6
    Bar_Tag = Translate("Region 3 - Intercon Only",Login_Language,conn)
  case 8
    Bar_Tag = Translate("Account Manager",Login_Language,conn)
  case 9
    Bar_Tag = Translate("Content Submitter",Login_Language,conn)
  case 10
    Bar_Tag = Translate("Content Administrator",Login_Language,conn)
  case 11
    Bar_Tag = Translate("Content Administrator",Login_Language,conn) & " - " & Translate("Matrix",Login_Language,conn)
  case 12
    Bar_Tag = Translate("Account Administrator",Login_Language,conn)
  case 13
    Bar_Tag = Translate("Account Administrator",Login_Language,conn) & " - " & Translate("Matrix",Login_Language,conn)
  case 14
    Bar_Tag = Translate("Site Administrator",Login_Language,conn)
  case 90
    Bar_Tag = Translate("Search",Login_Language,conn)
  case 91
    Bar_Tag = Translate("Search Results",Login_Language,conn)
  case 99
    Bar_Tag = Translate("Add New Account",Login_Language,conn)
end select

select case SC0           ' Speciality Lists
  case 1
    Bar_Tag = Bar_Tag
  case 2
    Bar_Tag = Bar_Tag
  case 3
    Bar_Tag = Translate("Content Administrator",Login_Language,conn) & " - " & Translate("Matrix",Login_Language,conn)  
  case 4
    Bar_Tag = Translate("Account Administrator",Login_Language,conn) & " - " &     Bar_Tag = Translate("Matrix",Login_Language,conn)  
end select  

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number
' --------------------------------------------------------------------------------------

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Screen_Title     = rsSite("Site_Description") & " - " & Translate("Account Administrator",Alt_Language,conn)
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Administrator",Login_Language,conn) & "</FONT>"
Bar_Title        = Bar_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("List",Login_Language,conn) & " / " & Translate("Edit",Login_Language,conn) & " " & Translate("Group",Login_Language,conn) & ": " & Bar_Tag & "</FONT>"

Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

rsSite.close
set rsSite=nothing

' --------------------------------------------------------------------------------------
' Account Listing
' --------------------------------------------------------------------------------------

if SC0 = 1 then

  SQL = "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID

  ' --------------------------------------------------------------------------------------
  ' Query Criteria - Groups
  ' --------------------------------------------------------------------------------------

  if SC2 = 0 then
      SQL = SQL & " AND UserData.NewFlag=" & CInt(True)
  else
      SQL = SQL & " AND UserData.NewFlag=" & CInt(False)
  end if

  select case SC2
    case 0  ' Pending Accounts (See Above)
      SC1 = 4
    case 1  ' Approved Today
      SQL = SQL & " AND Reg_Approval_Date='" & Date & "'"
    case 2  ' Non-Fluke - All
      SQL = SQL & " AND UserData.Type_Code<>5"
    case 3  ' Fluke - All        
      SQL = SQL & " AND UserData.Type_Code=5"
    case 4  ' Region 1 - US Only
      SQL = SQL & " AND UserData.Region=1"
    case 5  ' Region 2 - Europe Only        
      SQL = SQL & " AND UserData.Region=2"
    case 6  ' Region 3 - Intercon Only
      SQL = SQL & " AND UserData.Region=3"
    case 7  ' Expired Accounts
      SQL = SQL & " AND UserData.ExpirationDate <= '" & Date & "'"
    case 8  ' Account Managers
      SQL = SQL & " AND UserData.FCM=" & CInt(True)
    case 9  ' Content Submitters
      SQL = SQL & " AND UserData.Subgroups LIKE '%submitter%'"
    case 10 ' Content Administrators
      SQL = SQL & " AND UserData.Subgroups LIKE '%content%'"
    case 11 ' Content Administrators Matrix (See Above Redirect SC0=3
    case 12 ' Account Administrators
      SQL = SQL & " AND UserData.Subgroups LIKE '%account%'"
    case 13 ' Account Administrators Matrix (See Above Redirect SC0=4
    case 14 ' Site Administrators       
      SQL = SQL & " AND UserData.Subgroups LIKE '%administrator%'"
    case 90, 91 ' Search and Search Results
      if SC2 = 90 and not isblank(SC4) then
        if isnumeric(SC4) then
          SQL = SQL & " AND (UserData.ID=" & SC4 & ")"
        else
          SQL = SQL & " AND (UserData.LastName LIKE '%" & SC4 & "%' OR UserData.Company LIKE '%" & SC4 & "%')"
        end if
        SC2 = 91    ' Reset to Search Results
        SC3 = 0     ' Reset
        Bar_Tag = Translate("Search",Login_Language,conn) & " - " & Translate("Results",Login_Language,conn)
      elseif SC2 = 91 and not isblank(SC4) then
        if isnumeric(SC4) then
          SQL = SQL & " AND (UserData.ID=" & SC4 & ")"
        else
          SQL = SQL & " AND (UserData.LastName LIKE '%" & SC4 & "%' OR UserData.Company LIKE '%" & SC4 & "%')"
        end if
        SC3 = 0     ' Reset
        Bar_Tag = Translate("Search",Login_Language,conn) & " - " & Translate("Results",Login_Language,conn)
      else
        SC2 = 90
        SC3 = 0
        SC4 = ""
      end if
    case 99 ' Add New Account (See Above Redirect)
  end select

  ' --------------------------------------------------------------------------------------
  ' Query Criteria - Begins With
  ' --------------------------------------------------------------------------------------

  if SC3 > 0 then
    select case SC1
      case 0
        SQL = SQL & " AND UserData.LastName LIKE '"         & chr(SC3) & "%'"
      case 1  
        SQL = SQL & " AND UserData.Company LIKE '"          & chr(SC3) & "%'"
      case 2
        SQL = SQL & " AND UserData.Business_State LIKE '"   & chr(SC3) & "%'"
      case 3
        SQL = SQL & " AND UserData.Business_Country LIKE '" & chr(SC3) & "%'"
    end select
  end if

  ' --------------------------------------------------------------------------------------
  ' Query Criteria - Order By
  ' --------------------------------------------------------------------------------------

  if SC2 = 0 then
    SQL = SQL & " ORDER BY UserData.Reg_Request_Date"
  else
    select case SC1
      case 1      
          SQL = SQL & " ORDER BY UserData.Company, UserData.LastName, UserData.FirstName"
      case 2      
          SQL = SQL & " ORDER BY UserData.Business_Country, UserData.Company, UserData.LastName, UserData.FirstName"
      case 3
          SQL = SQL & " ORDER BY UserData.Business_State, UserData.Company, UserData.LastName, UserData.FirstName"
      case else  
          SQL = SQL & " ORDER BY UserData.LastName, UserData.FirstName"
    end select
  end if  

  'response.write SQL
  'response.end

  ' --------------------------------------------------------------------------------------
  ' Get UserData
  ' --------------------------------------------------------------------------------------

  if SC2 <> 90 then
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3
  end if

  if SC2 = 90 then

    Call Main_Menu
    response.write "&nbsp;&nbsp;&nbsp;"
    response.write "<BR><BR>"
    Call Query_Criteria

    response.write "<LI>" & Translate("Enter your search keyword above, the click on [Search] to begin.",Login_Language,conn) & ".</LI>"
    response.write "<LI>" & Translate("Search keyword can be one or more characters that could appear anywhere in the account's last name or company name, or the numeric ID number of the account.",Login_Language,conn) & "</LI><BR><BR>"
    TableOn = false

  elseif rsUser.EOF and rsUser.BOF  then

    Call Main_Menu
    response.write "&nbsp;&nbsp;&nbsp;"
    response.write "<BR><BR>"
    Call Query_Criteria

    response.write Translate("There are no User Accounts for this site that match your criteria.",Login_Language,conn) & "<BR><BR>"
    TableOn = false

  else

    if SC5 = Export_Delimited then
    
      SQLField_Names = "SELECT * from Field_Names WHERE Field_Names.Table_Name='UserData' AND Field_Names.Enabled=" & CInt(True) & " ORDER BY Field_Names.ID"
      Set rsField_Names = Server.CreateObject("ADODB.Recordset")
      rsField_Names.Open SQLField_Names, conn, 3, 3

      response.write "<FORM ""Field_Names"">"
      response.write "<TABLE WIDTH=""50%"">"
      
      do while not rsField_Names.EOF
      
        response.write "<TR>"
        response.write "<TD WIDTH=""5%"" CLASS=SMALL>"
        response.write "<INPUT TYPE=""CHECKBOX"" VALUE=""" & rsField_Names("ID") & """"
        if CInt(rsField_Names("Default_Select")) = CInt(True) then
          response.write " CHECKED"
        end if
        response.write " CLASS=Small>"
        response.write "</TD>"
        response.write "</TR>" & vbCrLf

        response.write "<TR>"
        response.write "<TD CLASS=SMALL>"       
        response.write rsField_Names("Description")
        response.write "</TD>"
        response.write "</TR>" & vbCrLf
                        
      loop
      
      response.write "<INPUT TYPE="Submit" NAME="Export" VALUE=""" & Translate("Submit",Login_Language,conn) & """>"
      
      response.write "</TABLE>"

      rsField_Names.close
      set rsField_Names = nothing

      rsUser.close
      set rsUser = nothing
      Call Disconnect_SiteWide      
      response.end

    else
    end if
    
    Record_Count  = rsUser.RecordCount
    Record_Pages  = Record_Count \ Record_Limit
    if Record_Count mod Record_Limit > 0 then Record_Pages = Record_Pages + 1
    Page_QS = "Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 &  "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&SC5=" & SC5 & "&Page=" & Page & "&Z=0"

  	rsUser.MoveFirst
    if Record_Limit * (Page - 1) > 0 then
     	rsUser.Move (Record_Limit * (Page - 1))
    end if     

    Record_Number = 1 

    TableOn = true

    Call Main_Menu
    response.write "&nbsp;&nbsp;&nbsp;"
    Call Group_Code_Table
    response.write "<BR><BR>"

    Call Query_Criteria
    Call Page_Navigation
    Call Count_Records

    response.write "<FORM NAME=""Display_Accounts"">" & vbCrLf

    response.write "<TABLE WIDTH=""100%"" BORDER=1 CELLPADDING=0 CELLSPACING=0 BORDERCOLOR=""#666666"" BGCOLOR=""#666666"">" & vbCrLF
    response.write "<TR>" & vbCrLF
    response.write "<TD>" & vbCrLF
    response.write "<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLF  
    response.write "<TR>" & vbCrLF

    if Admin_Access >=6 then
      response.write "<TD BGCOLOR=""Red"" ALIGN=CENTER CLASS=SmallBoldWhite>" & Translate("Action",Login_Language,conn) & "</TD>" & vbCrLF
    end if
    response.write "<TD BGCOLOR=""#000000"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("ID",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Users Name",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Company",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("City",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("State",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Country",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Phone",Login_Language,conn) & "</TD>" & vbCrLF
    select case SC5
      case 0
        response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Group Affiliations",Login_Language,conn) & "</TD>" & vbCrLF
      case 1
        response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("EMail",Login_Language,conn) & "</TD>" & vbCrLF
      case 2
        response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Postal Code",Login_Language,conn) & "</TD>" & vbCrLF
      case 3
        response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Customer Number",Login_Language,conn) & "</TD>" & vbCrLF
      case 9
        if Admin_Access >=8 then
          response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("User Name",Login_Language,conn) & " &amp; " & Translate("Password",Login_Language,conn) & "</TD>" & vbCrLF
        else
          response.write "&nbsp;"
        end if
      case Export_Delimited
        if Admin_Access >=6 then
          response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT   CLASS=SmallBoldGold>" & Translate("Export to Excel",Login_Language,conn) & "</TD>" & vbCrLF
        else
          response.write "&nbsp;"
        end if
    end select
    response.write "<TD BGCOLOR=""#000000"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("Status",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("Logon",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "<TD BGCOLOR=""#000000"" ALIGN=CENTER CLASS=SmallBoldGold>" & Translate("Expiration",Login_Language,conn) & "</TD>" & vbCrLF
    response.write "</TR>" & vbCrLF & vbCrLF

   end if

   if SC2 <> 90 then

     Do while not rsUser.EOF and Record_Number <= Record_Limit

      response.write "<TR>" & vbCrLF

      ' Edit Button
      response.write "<TD BGCOLOR="""
      if CInt(Updated_ID) = CInt(rsUser("ID")) then
        response.write "Green"
      else
        response.write "Silver"
      end if
      response.write """ ALIGN=CENTER CLASS=Small>" & vbCrLf

      if Admin_Access = 6 then
        select case SC2
          case 0,1,2,3,4,5,6,7,8,9,10,11,91
            if (instr(1,LCase(rsUser("SubGroups")),"administrator") = 0 and instr(1,LCase(rsUser("SubGroups")),"account") = 0) or isblank(rsUser("SubGroups")) then
              response.write "<A NAME=" & rsUser("ID") & "></A>" & vbCrLf
              response.write "<A HREF=""/SW-Administrator/Account_Edit.asp"
              response.write "?Site_ID=" & Site_ID
              response.write "&ID=edit_account&Account_ID=" & rsUser("ID")
              response.write "&SC0="  & SC0
              response.write "&SC1="  & SC1
              response.write "&SC2="  & SC2
              response.write "&SC3="  & SC3
              response.write "&SC4="  & SC4
              response.write "&SC5="  & SC5
              response.write "&Page=" & Page
              response.write "&Z=0"
              response.write """ "
              response.write "CLASS=NavLeftHighlight1 onClick=""location.href='/SW-Administrator/Account_Edit.asp"
              response.write "?Site_ID=" & Site_ID
              response.write "&ID=edit_account&Account_ID=" & rsUser("ID")
              response.write "&SC0="  & SC0
              response.write "&SC1="  & SC1
              response.write "&SC2="  & SC2
              response.write "&SC3="  & SC3
              response.write "&SC4="  & SC4
              response.write "&SC5="  & SC5
              response.write "&Page=" & Page
              response.write "&Z=0"
              response.write "'"" VALUE="" Edit "">&nbsp;&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;&nbsp;</A>"
            else  
              response.write "&nbsp;"
            end if  
          case else
            response.write "&nbsp;"
        end select

      elseif Admin_Access >= 8 then
            response.write "<A NAME=" & rsUser("ID") & "></A>" & vbCrLf
            response.write "<A HREF=""/SW-Administrator/Account_Edit.asp"
            response.write "?Site_ID=" & Site_ID
            response.write "&ID=edit_account&Account_ID=" & rsUser("ID")
            response.write "&SC0="  & SC0
            response.write "&SC1="  & SC1
            response.write "&SC2="  & SC2
            response.write "&SC3="  & SC3
            response.write "&SC4="  & SC4
            response.write "&SC5="  & SC5
            response.write "&Page=" & Page
            response.write "&Z=0"
            response.write """ "
            response.write "CLASS=NavLeftHighlight1 onClick=""location.href='/SW-Administrator/Account_Edit.asp"
            response.write "?Site_ID=" & Site_ID
            response.write "&ID=edit_account&account_ID=" & rsUser("ID")
            response.write "&SC0="  & SC0
            response.write "&SC1="  & SC1
            response.write "&SC2="  & SC2
            response.write "&SC3="  & SC3
            response.write "&SC4="  & SC4
            response.write "&SC5="  & SC5
            response.write "&Page=" & Page
            response.write "&Z=0"
            response.write "'"" VALUE="" Edit "">&nbsp;&nbsp;" & Translate("Edit",Login_Language,conn) & "&nbsp;&nbsp;</A>"
      end if

      response.write "</TD>" & vbCrLf

      ' Account ID
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=RIGHT CLASS=Small>"
      if Admin_Access = 6 then        
        select case SC2
          case 0,1,2,3,4,5,6,7,8,9,10,11,91
            if (instr(1,LCase(rsUser("SubGroups")),"administrator") = 0 and instr(1,LCase(rsUser("SubGroups")),"account") = 0) or isblank(rsUser("SubGroups")) then
              response.write rsUser("ID")
            else
              response.write "&nbsp;"
            end if  
          case else
            response.write "&nbsp;"
        end select
      elseif Admin_Access >= 8 then
            response.write rsUser("ID")
      end if      
      response.write "</TD>" & vbCrLF

      ' Name
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=LEFT CLASS=Small>"
      response.write "<B>" & Highlight_Keyword(rsUser("LastName"),SC4,"Region4NavSmall") & "</B>, "
      response.write rsUser("FirstName")
      if not isblank(rsUser("MiddleName")) then response.write " " & rsUser("MiddleName")
      if not isblank(rsUser("Prefix")) then response.write " " & rsUser("Prefix") & ". "
      response.write "</TD>" & vbCrLF

      ' Company
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=LEFT CLASS=Small>"
      response.write Highlight_Keyword(rsUser("Company"),SC4,"Region4NavSmall")
      response.write "</TD>" & vbCrLF

      ' Business City
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=LEFT CLASS=Small>"
      response.write rsUser("Business_City")
      response.write "</TD>" & vbCrLF

      ' Business State
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=LEFT CLASS=Small>"
      response.write rsUser("Business_State")
      response.write "</TD>" & vbCrLF

      ' Business Country
      response.write "<TD CLASS=Region" & Trim(CStr(rsUser("Region"))) & "NavSmall ALIGN=""LEFT"">"
      response.write rsUser("Business_Country")
      response.write "</TD>" & vbCrLF

      ' Business Phone
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=RIGHT CLASS=Small>"
      response.write FormatPhone(rsUser("Business_Phone"))
      response.write "</TD>" & vbCrLF

      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=LEFT CLASS=Small>"
      select case SC5
        case 0
          Call Write_SubGroups
        case 1
          if isblank(rsUser("EMail")) then
            response.write "&nbsp;"
          else  
            response.write Lcase(rsUser("EMail"))
          end if
        case 2
          if isblank(rsUser("Business_Postal_Code")) then
            response.write "&nbsp;"
          else
            response.write rsUser("Business_Postal_Code")
          end if
        case 3
          if isblank(rsUser("Fluke_ID")) then
            response.write "&nbsp;"
          else
            response.write rsUser("Fluke_ID")
          end if
        case 9
          if Admin_Access >=8 and instr(1,rsUser("Subgroups"),"administrator") = 0 then
            response.write "[" & rsUser("NTLogin") & "] [" & rsUser("Password") & "]"
          else
            response.write "[--------] [--------]"
          end if
        case else
          response.write "&nbsp;"
      end select
      response.write "</TD>" & vbCrLF

      ' Account Status
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=CENTER CLASS=Small>"
      Call Write_Account_Status
      response.write "</TD>" & vbCrLF

      ' Last Logon
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=CENTER CLASS=Small>"
      Call Write_Last_Logon
      response.write "</TD>" & vbCrLF

      ' Expiration Date
      response.write "<TD BGCOLOR="
      if isdate(rsUser("ExpirationDate")) then
        if DateDiff("d",CDate(rsUser("ExpirationDate")),Date) >=20 then
          response.write """Red"""
        else
          response.write """#FFFFFF"""
        end if
      end if  
      response.write " ALIGN=""CENTER"" CLASS=Small>"

      if CInt(rsUser("NewFlag")) = True then
        response.write "&nbsp;" 
      elseif CDate(rsUser("ExpirationDate")) = CDate("09/09/9999") then                
        response.write "Never"
      else  
        response.write FormatDate(1,rsUser("ExpirationDate"))
      end if
      response.write "</TD>" & vbCrLF

      response.write "</TR>" & vbCrLF & vbCrLF

      if Record_Number > Record_Limit then
        Page = Page + 1
        Record_Number = 1
      else
        Record_Number = Record_Number + 1
      end if

      rsUser.MoveNext

    loop

    rsUser.close
    set rsUser = nothing

  end if

  if TableOn then
    response.write "</TABLE>" & vbCrLF
    response.write "</TD>"    & vbCrLF
    response.write "</TR>"    & vbCrLF
    response.write "</TABLE>" & vbCrLF
    response.write "</FORM>"  & vbCrLF & vbCrLf
  end if

' --------------------------------------------------------------------------------------  
' Account Search
' --------------------------------------------------------------------------------------

elseif SC0 = 2 then

    response.write "<BR><BR>"
  
    Call Query_Criteria    
     
' --------------------------------------------------------------------------------------
' Content Administrator's Matrix
' --------------------------------------------------------------------------------------

elseif SC0 = 3 then

  if CInt(request("Toggle")) = CInt(True) and isnumeric(request("Approver_ID")) then
  
    SQL = "UPDATE Approvers SET Approvers.Approver_ID =" & request("Approver_ID") & " WHERE (((Approvers.ID)=" & request("Group_ID") & "))"
 	  conn.Execute(SQL)

  end if
  
  Toggle  = False    
  TableOn = True
  
  SQL = "SELECT Approvers.* FROM Approvers WHERE Approvers.Site_ID=" & CInt(Site_ID) & " ORDER BY Approvers.Order_Num, Approvers.Description"
  
  Set rsApproverGroups = Server.CreateObject("ADODB.Recordset")
  rsApproverGroups.Open SQL, conn, 3, 3    
  
  if rsApproverGroups.EOF then
    response.write Translate("There are no Content Administrator Groups established for this site.",Login_Language,conn) & "<BR><BR>"
    TableOn = false
  end if
 
  SQL = "Select UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND (UserData.Subgroups LIKE '%content%' OR UserData.Subgroups LIKE '%administrator%') ORDER BY UserData.LastName"
  Set rsApproverNames = Server.CreateObject("ADODB.Recordset")
  rsApproverNames.Open SQL, conn, 3, 3
  
  if rsApproverNames.EOF then
    response.write Translate("There are no Content Administrators established for this site.",Login_Language,conn) & "<BR><BR>"
    TableOn = false
  end if

  if TableOn = True then  

    Call Main_Menu
    response.write "<BR><BR>"
    Call Query_Criteria
    
    response.write "<FORM NAME=""CA-Matrix"">" & vbCrLf
    response.write "<TABLE WIDTH=""100%"" BORDER=1 CELLPADDING=0 CELLSPACING=0 BORDERCOLOR=""#666666"" BGCOLOR=""#666666"">" & vbCrLf
    response.write "<TR>" & vbCrLf
    response.write "<TD>" & vbCrLf
    response.write "<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
    response.write "<TR>" & vbCrLf
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT CLASS=SmallBoldGold WIDTH=""50%"">" & vbCrLf
    response.write Translate("Region",Login_Language,conn) & " / "
    response.write Translate("Group",Login_Language,conn) & " / "
    response.write Translate("Sub-Region",Login_Language,conn) & " " & Translate("or",Login_Language,conn) & " " & Translate("Description",Login_Language,conn)
    response.write "</TD>" & vbCrLf
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT CLASS=SmallBoldGold WIDTH=""50%"">"
    response.write Translate("Content Administrator",Login_Language,conn) & " " & Translate("Name",Login_Language,conn)
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf

    Do while not rsApproverGroups.EOF
     
      response.write "<TR>" & vbCrLf
      response.write "<TD CLASS=Region" & Trim(CStr(rsApproverGroups("Region"))) & "NavSmall ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"
      response.write rsApproverGroups("Description")
      response.write "</TD>" & vbCrLf
  
      rsApproverNames.MoveFirst
      
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Small VALIGN=MIDDLE>" & vbCrLf

      response.write "<SELECT CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='/SW-Administrator/Account_List.asp?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 & "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&SC5=" & SC5 & "&Toggle=" & CInt(True) & "&Group_ID=" & rsApproverGroups("ID") & "&Approver_ID='+this.options[this.selectedIndex].value"" NAME=""Approver_ID"">" & vbCrLf

      response.write "<OPTION VALUE=""0"" CLASS=NavLeftHighlight1>" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
      
      Do while not rsApproverNames.EOF    
        response.write "<OPTION CLASS=Region" & Trim(CStr(rsApproverNames("Region"))) & "NavSmall VALUE=""" & rsApproverNames("ID") & """"
        if isnumeric(rsApproverGroups("Approver_ID")) then
          if CInt(rsApproverNames("ID")) = CInt(rsApproverGroups("Approver_ID")) then
            response.write " SELECTED"
          end if
        end if    
        response.write ">" & rsApproverNames("LastName") & ", " & rsApproverNames("FirstName") & "</OPTION>" & vbCrLf
        rsApproverNames.MoveNext
      Loop
      
      response.write "</SELECT>" & vbCrLf
      response.write "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
      
      rsApproverGroups.MoveNext
     
    Loop
          
  end if  
  
  rsApproverGroups.close
  set rsApproverGroups = nothing
  
  rsApproverNames.close
  set rsApproverNames = nothing  
  
  if TableOn then

    response.write "</TABLE>" & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf
    response.write "</FORM>" & vbCrLf & vbCrLf
            
  end if
 
' --------------------------------------------------------------------------------------
' Account Administrator's Matrix
' --------------------------------------------------------------------------------------

elseif SC0 = 4 then

  if CInt(request("Toggle")) = CInt(True) and isnumeric(request("Approver_ID")) then
  
    SQL = "UPDATE Approvers_Account SET Approvers_Account.Approver_ID =" & request("Approver_ID") & " WHERE (((Approvers_Account.ID)=" & request("Group_ID") & "))"
 	  conn.Execute(SQL)

  end if
  
  Toggle  = false    
  TableOn = True
  
  SQL = "SELECT Approvers_Account.* FROM Approvers_Account WHERE Approvers_Account.Site_ID=" & CInt(Site_ID) & " ORDER BY Approvers_Account.Order_Num, Approvers_Account.Description"
  
  Set rsApproverGroups = Server.CreateObject("ADODB.Recordset")
  rsApproverGroups.Open SQL, conn, 3, 3    
  
  if rsApproverGroups.EOF then
    response.write Translate("There are no Account Administrator Groups established for this site.",Login_Language,conn) & "<BR><BR>"
    TableOn = false
  end if
 
  SQL = "Select UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND (UserData.Subgroups LIKE '%account%' OR UserData.Subgroups LIKE '%administrator%') ORDER BY UserData.LastName"
  Set rsApproverNames = Server.CreateObject("ADODB.Recordset")
  rsApproverNames.Open SQL, conn, 3, 3
  
  if rsApproverNames.EOF then
    response.write Translate("There are no Account Administrators established for this site.",Login_Language,conn) & "<BR><BR>"
    TableOn = false
  end if

  if TableOn = True then  

    Call Main_Menu
    response.write "<BR><BR>"
    Call Query_Criteria    

    response.write "<FORM NAME=""AA-Matrix"">"
    response.write "<TABLE WIDTH=""100%"" BORDER=1 CELLPADDING=0 CELLSPACING=0 BORDERCOLOR=""#666666"" BGCOLOR=""#666666"">"
    response.write "<TR>"
    response.write "<TD>"
    response.write "<TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH=""100%"">"
    response.write "<TR>"
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT CLASS=SmallBoldGold WIDTH=""50%"">" & Translate("Region",Login_Language,conn) & "</TD>"
    response.write "<TD BGCOLOR=""#000000"" ALIGN=LEFT CLASS=SmallBoldGold WIDTH=""50%"">" & Translate("Account Administrator",Login_Language,conn) & " " & Translate("Name",Login_Language,conn) & "</TD>"
    response.write "</TR>"
      
    Do while not rsApproverGroups.EOF
     
      response.write "<TR>"
      response.write "<TD CLASS=Region" & Trim(CStr(rsApproverGroups("Region"))) & "NavSmall ALIGN=""LEFT"" VALIGN=MIDDLE>"
      response.write rsApproverGroups("Description")
      response.write "</TD>"
  
      rsApproverNames.MoveFirst
      
      response.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" CLASS=Medium VALIGN=MIDDLE>"

      response.write "<SELECT CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='/SW-Administrator/Account_List.asp?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 & "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&SC5=" & SC5 & "&Toggle=" & CInt(True) & "&Group_ID=" & rsApproverGroups("ID") & "&Approver_ID='+this.options[this.selectedIndex].value"" NAME=""Approver_ID"">"

      response.write "<OPTION VALUE=""0"" CLASS=NavLeftHighlight1>" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
      
      Do while not rsApproverNames.EOF    
        response.write "<OPTION CLASS=Region" & Trim(CStr(rsApproverNames("Region"))) & "NavSmall VALUE=""" & rsApproverNames("ID") & """"
        if isnumeric(rsApproverGroups("Approver_ID")) then
          if CInt(rsApproverNames("ID")) = CInt(rsApproverGroups("Approver_ID")) then
            response.write " SELECTED"
          end if
        end if    
        response.write ">" & rsApproverNames("LastName") & ", " & rsApproverNames("FirstName") & "</OPTION>" & vbCrLf
        rsApproverNames.MoveNext
      Loop
      
      response.write "</SELECT>" & vbCrLf
      response.write "</TD>" & vbCrLf
      response.write "</TR>" & vbCrLf
      
      rsApproverGroups.MoveNext
     
    Loop
          
  end if  
  
  rsApproverGroups.close
  set rsApproverGroups = nothing
  
  rsApproverNames.close
  set rsApproverNames = nothing  
  
  if TableOn then

    response.write "</TABLE>"
    response.write "</TD>"
    response.write "</TR>"
    response.write "</TABLE>"
    response.write "</FORM>"
            
  end if                     
  
else

  response.write Translate("Invalid Account List Option Number",Login_Language,conn) & "<BR><BR>"
  
end if  

Call Main_Menu
response.write "<BR><BR>"

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"--> 
<%

' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------

sub Query_Criteria()
  
  response.write "<FORM NAME=""Sort_Criteria"" ACTION=""/SW-Administrator/Account_List.asp"" METHOD=""GET"">" & vbCrLf
  response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=4 CELLSPACING=0>" & vbCrLf
  response.write "<TR>" & vbCrLf

  ' Groups
  response.write "<TD CLASS=NavLeftHighlight1>" & vbCrLf
  response.write Translate("Groups",Login_Language,conn) & ": "
  %>
  <!--#include virtual="/sw-administrator/Account_List_Query_Criteria.asp"-->
  <%
  response.write "</TD>" & vbCrLf

  ' Sort By
  if SC0 = 1 then
    response.write "<TD CLASS=NavLeftHighlight1>" & vbCrLf
    response.write Translate("Sort By",Login_Language,conn) & ": " & vbCrLf
    
    response.write "<SELECT NAME=""SC1"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Account_List.asp?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC2=" & SC2 & "&SC3=0" & "&SC4=" & SC4 & "&SC5=" & SC5 & "&Z=" & Z & "&Page=1" & "&SC1=" & "'+this.options[this.selectedIndex].value"">" & vbCrLf

    response.write "<OPTION CLASS=Small VALUE=""0"""
    if SC1 = 0 then response.write " SELECTED"
    response.write ">" & Translate("Last Name",Login_Language,conn) & "</OPTION>" & vbCrLf
  
    response.write "<OPTION CLASS=Small VALUE=""1"""
    if SC1 = 1 then response.write " SELECTED"
    response.write ">" & Translate("Company",Login_Language,conn) & ", " & Translate("Last Name",Login_Language,conn) & "</OPTION>" & vbCrLf
  
    response.write "<OPTION CLASS=Small VALUE=""2"""
    if SC1 = 2 then response.write " SELECTED"
    response.write ">" & Translate("State",Login_Language,conn) & ", " & Translate("Company",Login_Language,conn) & ", " & Translate("Last Name",Login_Language,conn) & "</OPTION>" & vbCrLf
  
    response.write "<OPTION CLASS=Small VALUE=""3"""
    if SC1 = 3 then response.write " SELECTED"
    response.write ">" & Translate("Country",Login_Language,conn) & ", " & Translate("Company",Login_Language,conn) & ", " & Translate("Last Name",Login_Language,conn) & "</OPTION>" & vbCrLf
  
    response.write "<OPTION CLASS=Small VALUE=""4"""
    if SC1 = 4 then response.write " SELECTED"
    response.write ">" & Translate("Aging",Login_Language,conn) & "</OPTION>" & vbCrLf
  
    response.write "</SELECT>" & vbCrLf
    response.write "</TD>" & vbCrLf
  end if
  
  ' Search
  if SC0=1 and (SC2 = 90 or SC2 = 91) then
    response.write "<TD CLASS=NavLeftHighlight1>" & vbCrLf
    response.write Translate("Keyword",Login_Language,conn) & ": " & vbCrLf & vbCrLf
      response.write "<INPUT TYPE=""TEXT"" NAME=""SC4"" SIZE=""20"" MAXLENGTH=""50"" CLASS="
    if not isblank(SC4) then
      response.write "Region4NavSmall"
    else
      response.write "SMALL"
    end if
    response.write " VALUE=""" & SC4 & """>"
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE="""  & Site_ID  & """>"
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""SC3"" VALUE="""      & SC3  & """>"
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Page"" VALUE="""     & Page & """>"
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Z"" VALUE="""        & ""   & """>&nbsp;&nbsp;"
    response.write "<INPUT TYPE=""SUBMIT"" NAME=""Search"" VALUE="" "  & Translate("Search",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>"
  end if

  ' Begins With
  if SC0 = 1 then
  
    select case SC2
      case 1,2,3,4,5,6,7,8,9,10,11,12,13,14
        response.write "<TD CLASS=NavLeftHighlight1>" & vbCrLf
        response.write Translate("Begins With",Login_Language,conn) & ": " & vbCrLf & vbCrLf
  
        response.write "<SELECT NAME=""SC3"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Account_List.asp?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 & "&SC2=" & SC2 & "&SC4=" & SC4 & "&SC5=" & SC5 & "&Page=1" & "&SC3=" & "'+this.options[this.selectedIndex].value"">" & vbCrLf
'        response.write "<OPTION VALUE=""0"">" & Translate("Any Character",Login_Language,conn) & "</OPTION>" & vbCrLf
        response.write "<OPTION VALUE=""0""></OPTION>" & vbCrLf      
        ' A - Z
        for i = 65 to 65 + 25
          response.write "<OPTION CLASS=Small VALUE=""" & i & """"
          if SC3 = i then
            response.write " SELECTED"
          end if
          response.write ">" & chr(i) & "</OPTION>" & vbCrLf
        next
        
        ' 0 - 9
        for i = 48 to 48 + 9
          response.write "<OPTION VALUE=""" & i & """"
          if SC3 = i then
            response.write " SELECTED"
          end if
          response.write ">" & chr(i) & "</OPTION>" & vbCrLf
        next
      
        response.write "</SELECT>" & vbCrLf
        response.write "</TD>" & vbCrLf & vbCrLf  
    end select
  end if
  
  ' Show Alternate Field Information
  if SC0 = 1 then
    response.write "<TD CLASS=NavLeftHighlight1>" & vbCrLf
    response.write Translate("View",Login_Language,conn) & ": " & vbCrLf
    response.write "<SELECT NAME=""SC5"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Account_List.asp?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 & "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&Page=" & Page & "&Z=" & Z & "&SC5=" & "'+this.options[this.selectedIndex].value"">" & vbCrLf

    response.write "<OPTION CLASS=Small VALUE=""0"""
    if SC5 = 0 then response.write " SELECTED"
    response.write ">" & Translate("Group Affiliations",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    response.write "<OPTION CLASS=Small VALUE=""1"""
    if SC5 = 1 then response.write " SELECTED"
    response.write ">" & Translate("EMail",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Small VALUE=""2"""
    if SC5 = 2 then response.write " SELECTED"
    response.write ">" & Translate("Postal Code",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Small VALUE=""3"""
    if SC5 = 3 then response.write " SELECTED"
    response.write ">" & Translate("Customer Number",Login_Language,conn) & "</OPTION>" & vbCrLf

    if Admin_Access >=8 then
      response.write "<OPTION CLASS=Small VALUE=""9"""
      if SC5 = 9 then response.write " SELECTED"
      response.write ">" & Translate("User Name",Login_Language,conn) & " &amp; " & Translate("Password",Login_language,conn) & "</OPTION>" & vbCrLf
    end if

    if Admin_Access >=6 then
      response.write "<OPTION CLASS=Small VALUE=""9"""
      if SC5 = Export_Delimited then response.write " SELECTED"
      response.write ">" & Translate("Export to Excel",Login_Language,conn) & "</OPTION>" & vbCrLf
    end if

    response.write "</SELECT>" & vbCrLf
    response.write "</TD>" & vbCrLf
  end if
  
  response.write "</TR>" & vbCrLf
  response.write "</TABLE>" & vbCrLf
  response.write "</FORM>" & vbCrLf
  
end sub

' --------------------------------------------------------------------------------------

sub Main_Menu()

  response.write "<A HREF=""/SW-Administrator/Default.asp?Site_ID=" & Site_ID & """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Home",Login_Language,conn) & "&nbsp;&nbsp;</A>"

end sub

' --------------------------------------------------------------------------------------

sub Group_Code_Table()

  response.write "&nbsp;&nbsp;&nbsp;<A HREF=""/SW-Administrator/SubGroup_Codes.asp?Site_ID=" & Site_ID & """ onclick=""openit('/SW-Administrator/SubGroup_Codes.asp?Site_ID=" & Site_ID & "','Vertical');return false;"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("View Group Affiliation Codes",Login_Language,conn) & "&nbsp;&nbsp;</A>"
  
end sub

' --------------------------------------------------------------------------------------

sub Write_SubGroups()

  if not isblank(rsUser("SubGroups")) then
    SubGroups = rsUser("SubGroups")
    SubGroups = Replace(SubGroups,"administrator","<FONT COLOR=""Red""><B>Site</B></FONT>")    
    SubGroups = Replace(SubGroups,"submitter","<FONT COLOR=""Red""><B>Submitter</B></FONT>")
    SubGroups = Replace(SubGroups,"content","<FONT COLOR=""Red""><B>Content</B></FONT>")
    SubGroups = Replace(SubGroups,"account","<FONT COLOR=""Red""><B>Account</B></FONT>")
    SubGroups = Replace(SubGroups,"PIK, Users,","")
    SubGroups = Replace(SubGroups,"PIK, ","")    
  else
    SubGroups = "&nbsp;"
  end if
      
  response.write SubGroups

end sub

' --------------------------------------------------------------------------------------

sub Write_Account_Status()

  if rsUser("NewFlag") = True then                
    response.write "<FONT COLOR=""Red"">" & Translate("Pending Approval",Login_Language,conn) & "</FONT>"
  else
    response.write "Active"
  end if    

end sub

' --------------------------------------------------------------------------------------

sub Write_Last_Logon()

    if isblank(rsUser("Logon")) or instr(1,rsUser("Logon"),"9999") > 0 or not isDate(rsUser("Logon")) then
      response.write "<FONT COLOR=""#FF9900"">" & Translate("Never",Login_Language,conn) & "</FONT>"
    else  
      response.write FormatDate(1,rsUser("Logon"))
    end if  

end sub

' --------------------------------------------------------------------------------------

sub Write_Logon_History()

    SQL = "SELECT Logon.Account_ID, Logon.Logon, Logon.Logoff FROM Logon WHERE Logon.Site_ID=" & Site_ID & " Logon.Account_ID=" & rsUser("ID") & " ORDER BY Logon.Logon DESC"
    Set rsLogon = Server.CreateObject("ADODB.Recordset")
    rsLogon.Open SQL, conn, 3, 3                    
    
    if not rsLogon.EOF then
       
      do while not rsLogon.EOF
        response.write FormatDate(1,rsLogon("Logon")) & "<BR>"
        rsLogon.MoveNext
      loop  
    else
      response.write Translate("No History",Login_Language,conn)
    end if  

    rsLogon.close
    set rsLogon = nothing                       

end sub

' --------------------------------------------------------------------------------------
' Record Set Page Navigation
' --------------------------------------------------------------------------------------

sub Page_Navigation

  Page_QS = "Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 &  "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&SC5=" & SC5

  if Record_Pages > 1 then
  
    response.write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>" & vbCrLf
    response.write "<TR>" & vbCrLf
    response.write "<TD CLASS=SmallBold VALIGN=TOP>"
    response.write "&nbsp;&nbsp;" & Translate("Page", Login_Language, conn) & ":&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    response.write "</TD>" & vbCrLf    

  	if Page = 1 then
        response.write "<TD CLASS=SmallBold VALIGN=TOP>"
    		response.write "<A HREF=""/SW-Administrator/Account_List.asp?" & Page_QS & "&Page=" & Page + 1 & """ TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;"
        response.write "&nbsp;&nbsp;" & vbCrLf
        response.write "</TD>" & vbCrLf
    		Call Page_Numbers        
        response.write "</TR>" & vbCrLf
  	else
  		if Page = Record_Pages then
        response.write "<TD CLASS=SmallBold VALIGN=TOP>"
  			response.write "<A HREF=""/SW-Administrator/Account_List.asp?" & Page_QS & "&Page=" & Page - 1 & """ TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;" & vbCrLf
        response.write "</TD>" & vbCrLf
    		Call Page_Numbers
        response.write "</TR>" & vbCrLf
  		else
        response.write "<TD CLASS=SmallBold VALIGN=Top>"
  			response.write "<A HREF=""/SW-Administrator/Account_List.asp?" & Page_QS & "&Page=" & Page - 1 &  """ TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;" & vbCrLf
        response.write "</TD>" & vbCrLf
        response.write "<TD CLASS=SmallBold VALIGN=Top>"
  			response.write "<A HREF=""/SW-Administrator/Account_List.asp?" & Page_QS & "&Page=" & Page + 1 &  """ TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;" & vbCrLf
        response.write "</TD>" & vbCrLf
    		Call Page_Numbers
        response.write "</TR>" & vbCrLf
  		end if
  		
  	end if

    response.write "</TABLE>" & vbCrLf & vbCrLf
    
  end if

end Sub

' --------------------------------------------------------------------------------------
' Record Set Page Numbers
' --------------------------------------------------------------------------------------

sub Page_Numbers

  Box_Count = 0
  
  response.write "<TD CLASS=SmallBold VALIGN=TOP>"  

  for i = 1 to Record_Pages
    if Box_Count > 24 then
      Box_Count = 1
      response.write "<BR>"
    else
      Box_Count = Box_Count + 1  
    end if  

   	response.write "<A HREF=""/SW-Administrator/Account_List.asp?" & Page_QS & "&Page=" & i & """>"

  	if i = Page then
      response.write "<FONT CLASS=NavLeftHighLight1>&nbsp;"
  	else
      response.write "<FONT CLASS=NavTopHighlight>&nbsp;"      
  	end if

    if i < 10 then response.write "&nbsp;"
    response.write CStr(i)
    if i < 10 then response.write "&nbsp;"    
    response.write "&nbsp;</FONT></A>&nbsp;&nbsp;"

  next
  
  response.write "<BR>&nbsp;</TD>"  & vbCrLf

end sub

' --------------------------------------------------------------------------------------

sub Count_Records

    response.write "<FONT CLASS=SmallBold>&nbsp;&nbsp;" & Translate("Records",Login_Language,conn) & ": "
    
    if Page = 1 then
      response.write "1"
    else
      response.write (((Page * Record_Limit) - Record_Limit) + 1)
    end if

    if Page * Record_Limit > Record_Count then
      response.write " - " & Record_Count
    else
      response.write " - " & Page * Record_Limit
    end if
        
    response.write "&nbsp;&nbsp;&nbsp;" & Translate("Total",Login_Language,conn) & ": " & Record_Count & "</FONT>" & vbCrLf & vbCrLf
    
end sub

' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/Pop-Up.asp"-->
<%

Call Disconnect_SiteWide

%>
  