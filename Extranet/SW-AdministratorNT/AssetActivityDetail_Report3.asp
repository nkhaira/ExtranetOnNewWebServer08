<%
if isblank(Session("LOGON_USER")) then
	 response.redirect "/register/login.asp?Site_ID=100"
end if
'Response.Clear()
'Response.Buffer=True
'Response.ContentType = "xls"
'Response.AddHeader "Content-Disposition", "attachment;filename=" & Assetid & ".xls"
'Response.Charset = "utf-16" 
'Response.Codepage = "936" 
wsite=request("wsite")

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

if wsite = "" then

Site_ID=request("Site_ID")

%>

<% 
Assetid=request("ID")
title=request("title")

%>


<%





  


'dim Begin_date, Interval,Site_ID,Region, Item_Number,AssetID
'dim SQL,SQL_WHERE


if isblank(request("Begin_Date")) then
      Begin_Date = Date()
    elseif isdate(request("Begin_Date")) then
      Begin_Date = request("Begin_Date")
    else
'      response.write Translate("Invalid Date - Reseting to Today's Date",Login_Language,conn) & "<P>" & vbCrLf
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
'dim SQLWhere, SQLWhereLOS

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


 SQL = "SELECT    dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone  " &_
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

<%
	else

	Site_ID=request("Site_ID")
	Assetid=request("ID")
	Item_Number=request("Item_Number")
if (request("Site_Path") <> "") then
	Site_Path=request("Site_Path")
end if
	
'	dim Begin_date, Interval,Site_ID,Region, Item_Number,AssetID
'	dim SQL,SQL_WHERE



	if isblank(request("Begin_Date")) then
      Begin_Date = Date()
    elseif isdate(request("Begin_Date")) then
      Begin_Date = request("Begin_Date")
    else
      Begin_Date = Date()      
    end if


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

	if not isblank(request("Category_Code")) then
     	 Category_Code = request("Category_Code")
    else
      Category_Code = "all"
    end if  
	if not isblank(request("Country_Code")) then
     	 Country_Code = request("Country_Code")
    else
      Country_Code = "all"
    end if  
 	if not isblank(request("Local_Code")) then
      Local_Code = request("Local_Code")
    else
      Local_Code = "all"
    end if  

SQLWhere=""
SQL=""
if Interval >= 0  then
      SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "
    else
      SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Activity.View_Time,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Activity.View_Time,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
    end if

SQL = "SELECT    dbo.Activity.Account_ID, dbo.Activity.Site_ID, dbo.Activity.Calendar_ID, dbo.Calendar.Item_Number, dbo.Calendar.Revision_Code, dbo.Calendar.Product, dbo.Calendar.Title, dbo.Calendar.Sub_Category, dbo.Calendar.Language, dbo.Activity.View_Time, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone," &_
          " dbo.Activity.CMS_Site, dbo.CMS_XReference.CMS_Path AS CMS_Path " &_
          "FROM  dbo.UserData RIGHT OUTER JOIN dbo.Activity ON dbo.UserData.ID = dbo.Activity.Account_ID and dbo.UserData.Site_ID=dbo.Activity.Site_ID LEFT OUTER JOIN " &_
          " dbo.Calendar ON dbo.Activity.Calendar_ID = dbo.Calendar.ID LEFT OUTER JOIN " &_
          " dbo.CMS_XReference ON dbo.Activity.CMS_ID = dbo.CMS_XReference.ID "

    SQL = SQL & SQLWhere & "AND (dbo.Activity.CMS_Site IS NOT NULL) "
    
    if Category_Code <> "all" then
      SQL = SQL & "AND (dbo.Calendar.Sub_Category='" & Category_Code & "') "
    end if
    
    if Country_Code <> "all" then
      SQL = SQL & "AND (SUBSTRING(dbo.Activity.CMS_Site,1,2)='" & Country_Code & "') "
    end if  
      
    if Local_Code <> "all" then
      SQL = SQL & "AND (SUBSTRING(dbo.Activity.CMS_Site,3,2)='" & Local_Code & "') "
    end if  

 	if isblank(Site_Path) then
	    SQL = SQL & "AND (Activity.Calendar_ID=" & AssetID & " ) "
	else
	    SQL = SQL & "AND (Activity.Calendar_ID=" & AssetID & " ) AND (CMS_Path= '" & Site_Path& "')"
	end if

 if Category_Code <> "all" then
      SQL = SQL & "AND (dbo.Calendar.Sub_Category='" & Category_Code & "') "
    end if
    
    if Country_Code <> "all" then
      SQL = SQL & "AND (SUBSTRING(dbo.Activity.CMS_Site,1,2)='" & Country_Code & "') "
    end if  
      
    if Local_Code <> "all" then
      SQL = SQL & "AND (SUBSTRING(dbo.Activity.CMS_Site,3,2)='" & Local_Code & "') "
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

	
Set rsActivity = Server.CreateObject("ADODB.Recordset")
 rsActivity.Open SQL, conn, 3, 3


%>
<HTML>
<TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>
<tr><th colspan=7>Asset Activity Detail Report for Asset ID: <%=AssetID%></th></tr>
<TR>
   <TD><font size=2 face="Arial"><b>Name</b></font></TD>
   <TD><font size=2 face="Arial"><b>Title</b></font></TD>
   <TD><font size=2 face="Arial"><b>Company</b></font></TD>
   <TD><font size=2 face="Arial"><b>Address</b></font></TD>
   <TD><font size=2 face="Arial"><b>Telephone</b></font></TD>
   <TD><font size=2 face="Arial"><b>Email</b></font></TD>
   <TD><font size=2 face="Arial"><b>Fax</b></font></TD>
   
</TR>
<%
'dim Address, Country
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
<% end if %>