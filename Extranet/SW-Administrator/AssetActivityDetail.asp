<%
if isblank(Session("LOGON_USER")) then
	 response.redirect "/register/login.asp?Site_ID=100"
end if


Site_ID=request("Site_ID")

%>

<% 
Assetid=request("ID")
title=request("title")
Response.Clear()
Response.Buffer=True
Response.ContentType = "xls"
Response.AddHeader "Content-Disposition", "attachment;filename=" & Assetid & ".xls"
Response.Charset = "utf-16" 
'Response.Codepage = "936" 
%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->

<%





   Call Connect_SiteWide



Session.timeout = 60            ' Set to 1 Hour
Server.ScriptTimeout = 6 * 60   ' Set to 6 Minutes



dim Begin_date, Interval,Site_ID,Region, Item_Number,AssetID
dim SQL,SQL_WHERE


if isblank(request("Begin_Date")) then
      Begin_Date = Date()
    elseif isdate(request("Begin_Date")) then
      Begin_Date = request("Begin_Date")
    else
      response.write Translate("Invalid Date - Reseting to Today's Date",Login_Language,conn) & "<P>" & vbCrLf
      Begin_Date = Date()      
    end if

Site_ID_Change=Site_ID
if Site_ID = 82 then
      if isblank(request("Interval")) then
      Interval = 1
    elseif isnumeric(request("Interval")) then
      if FormatDateTime(Begin_Date) = FormatDateTime(Date()) and request("Interval") >= 0 then
        Interval = 1
      elseif request("Interval") > 365 then
        Interval = 365
      elseif request("Interval") < -365 then
        Interval = -365
      else  
        Interval = request("Interval")
      end if  
    end if
    
  else
  if isblank(request("Interval")) then
      Interval = 1
    elseif isnumeric(request("Interval")) then
      if FormatDateTime(Begin_Date) = FormatDateTime(Date()) and request("Interval") >= 0 then
        Interval = 1
      elseif request("Interval") > 90 then
        Interval = 90
      elseif request("Interval") < -90 then
        Interval = -90
      else  
        Interval = request("Interval")
      end if  
    end if
    
    
  end if
dim SQLWhere, SQLWhereLOS

   Item_Number=request("Item_Number")

if Interval >= 0 and isblank(Item_Number) then
	
      SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "
      SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "      
    elseif not isblank(Item_Number) then

      SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
      SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "              
    else

      SQLWhere    = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
      SQLWhereLOS = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "      
    end if




if Site_ID_Change > 0 then
      SQLWhere = SQLWhere & "AND dbo.Activity.Site_ID=" & Site_ID_Change & " "
    else
      SQLWhere = SQLWhere & " "
    end if  

Region=request("Region")

 select case Region
      case 1, 2, 3
        SQLWhere = SQLWhere & "AND dbo.UserData.Region=" & Region & " "
    end select


 SQL = "SELECT     dbo.Calendar.Title as AssetTitle, dbo.Calendar_Category.Title as AssetCategory, Convert(varchar, View_Time, 101) as AssetDate, Convert(varchar, View_Time, 108) as AssetTime, dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone  " &_
          "FROM      dbo.UserData RIGHT OUTER JOIN " &_
          "          dbo.Activity ON dbo.UserData.ID = dbo.Activity.Account_ID and dbo.UserData.Site_ID=dbo.Activity.Site_ID LEFT OUTER JOIN " &_
          "          dbo.Content_Sub_Category RIGHT OUTER JOIN " &_
          "          dbo.Calendar ON dbo.Content_Sub_Category.Sub_Category = dbo.Calendar.Sub_Category AND " &_
          "          dbo.Content_Sub_Category.Site_ID = dbo.Calendar.Site_ID ON dbo.Activity.Calendar_ID = dbo.Calendar.ID LEFT OUTER JOIN " &_
          "          dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code AND dbo.Calendar.Site_ID = dbo.Calendar_Category.Site_ID "
     if not isblank(item_number) then
	    SQL = SQL & SQLWhere & "AND (Activity.Calendar_ID=" & AssetID & " )"
	else
	    SQL = SQL & SQLWhere & "AND (Activity.Calendar_ID=" & AssetID & " ) AND (Calendar_Category.Title= " & title & ")"
	end if
      





          
        
	if not isblank(item_number) then
SQL = SQL & "AND ( "  
        if len(item_number) < 6 then
          SQL = SQL & "dbo.Activity.Calendar_ID='" & item_number & "' "
        else  
          SQL = SQL & "dbo.Calendar.Item_Number='" & item_number & "' "
        end if  
     
      SQL = SQL & ") "
	end if
'SQL = SQL & "ORDER BY dbo.Calendar_Category.Title, dbo.Content_Sub_Category.Sub_Category, dbo.Activity.Calendar_ID"
''SQL = SQL & "group BY dbo.Activity.Account_ID, dbo.UserData.FirstName,dbo.UserData.LastName, dbo.UserData.Job_Title,dbo.UserData.Company,Business_Address,Business_Address_2,dbo.UserData.Business_City,dbo.UserData.Business_State,dbo.UserData.Business_Postal_Code,dbo.UserData.Business_Country,Email,dbo.UserData.Business_Fax,Business_Phone"

'response.write SQL
'response.end
Set rsActivity = Server.CreateObject("ADODB.Recordset")
 rsActivity.Open SQL, conn, 3, 3


%>
<HTML>
<TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>
<tr><th colspan=7>Asset Activity Detail Report for Asset ID: <%=AssetID%></th></tr>
<TR>
   <TD><font size=2 face="Arial"><b>Asset ID</b></font></TD>
   <TD><font size=2 face="Arial"><b>Asset Title</b></font></TD>
   <TD><font size=2 face="Arial"><b>Asset Category</b></font></TD>
   <TD><font size=2 face="Arial"><b>Date</b></font></TD>
   <TD><font size=2 face="Arial"><b>Time</b></font></TD>
   <TD><font size=2 face="Arial"><b>Name</b></font></TD>
   <TD><font size=2 face="Arial"><b>Title</b></font></TD>
   <TD><font size=2 face="Arial"><b>Company</b></font></TD>
   <TD><font size=2 face="Arial"><b>Address</b></font></TD>
   <TD><font size=2 face="Arial"><b>Telephone</b></font></TD>
   <TD><font size=2 face="Arial"><b>Email</b></font></TD>
   <TD><font size=2 face="Arial"><b>Fax</b></font></TD>
   
</TR>
<%
dim Address, Country
do while not rsActivity.eof
	if not rsActivity("Name")="" then

Address=""
Country=""
Address=rsActivity("Address")

if not rsActivity("Country")="" then
	Country=DisplayCountry(rsActivity("Country"))
end if

Address= Address+ ", " & Country

 %>
 <TR>
   <TD><font size=2 face="Arial"><%=AssetID%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("AssetTitle")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("AssetCategory")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("AssetDate")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("AssetTime")%></font></TD>

   <TD><font size=2 face="Arial"><%=rsActivity("Name")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("Title")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("Company")%></font></TD>
   <TD><font size=2 face="Arial"><%=Address%></font></TD>
   <TD><font size=2 face="Arial"><%=CStr(rsActivity("Phone"))%></font></TD>
   <TD><font size=2 face="Arial"><%=rsActivity("Email")%></font></TD>
   <TD><font size=2 face="Arial"><%=Cstr(rsActivity("Fax"))%></font></TD>
</TR>
<% 
	end if
   rsActivity.MoveNext
   loop
   ' Clean up
   rsActivity.Close
   set rsActivity = Nothing
   
%>
</HTML>
<%
	Function DisplayCountry(byval countryCode)

	dim code
	code=countryCode
	SQL = "SELECT * FROM Country WHERE Enable=" & CInt(True) & " and Abbrev='" & code & "'"
  	Set rsCountries = Server.CreateObject("ADODB.Recordset")
 	 rsCountries.Open SQL, conn, 3, 3
	if not rsCountries.eof then
		Name=rsCountries("Name")
	end if
	rsCountries.close()
	DisplayCountry = Name
	end function
%>

