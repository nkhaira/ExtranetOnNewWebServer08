<% 
Response.Buffer=True
Response.Clear()
Response.ContentType = "xls"
Response.AddHeader "Content-Disposition", "attachment;filename=SummaryReport.xls"
Response.Charset = "utf-16" 
Response.Codepage = "936" 
%>

<%


On Error Resume Next
Session.timeout = 60            ' Set to 1 Hour


if isblank(Session("LOGON_USER")) then
	 response.redirect "/register/login.asp?Site_ID=100"
end if

z=request("z")
y=request("y")
'response.write Z
Site_ID=request("Site_ID")
Region=request("region")

Summary_Year=request("Summary_Year")
%>

<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->

<%
Server.ScriptTimeout = 60 * 60   ' Set to 1 hr
   Call Connect_SiteWide

dim SQL
SQL=""

'for z=0 to 7

Select case z

	case 0
		SQL= "Select dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.Calendar_ID=101 and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 

		select case UCase(request("Region"))     ' Filter Results by Region

              case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select
			
case 1
		SQL= "Select dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.Calendar_ID=102 and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 

		select case UCase(request("Region"))     ' Filter Results by Region

		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select

case 2
		'SQL= "Select dbo.Calendar.Title as AssetTitle, Calendar_Category.Title as AssetCategory, Convert(varchar, View_Time, 101) as AssetDate, Convert(varchar, View_Time, 108) as AssetTime, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		 '    " From dbo.Activity, UserData, Calendar where Activity.Site_ID= " & Site_ID & " and Activity.Calendar_ID > 200 and Activity.Calendar_ID is not null and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID and Activity.Calendar_ID=Calendar.ID and Activity.Site_ID=Calendar.Site_ID " 


		if y="" then

		SQL=  "SELECT dbo.Calendar.Title as AssetTitle, dbo.Activity.Calendar_ID as AssetID,dbo.Calendar_Category.Title as AssetCategory, " &_
			"Convert(varchar, View_Time, 101) as AssetDate, Convert(varchar, View_Time, 108) as "&_
			"AssetTime, dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, "&_
			"(isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,"&_
			"(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', "&_
			"' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') "&_
			"+ ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') "&_
			"as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, "&_
			"isnull(dbo.UserData.Business_Phone,'') as Phone, Activity.Country as ActCountry  "&_
			"FROM dbo.UserData RIGHT OUTER JOIN dbo.Activity "&_
			"ON dbo.UserData.ID = dbo.Activity.Account_ID and dbo.UserData.Site_ID=dbo.Activity.Site_ID "&_
			"RIGHT OUTER JOIN dbo.Calendar "&_
			"ON dbo.Activity.Calendar_ID = dbo.Calendar.ID "&_
			"Left Outer JOIN dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code AND dbo.Calendar.Site_ID = "&_
			"dbo.Calendar_Category.Site_ID WHERE DATEPART(yyyy, View_Time) = " & Summary_Year & ""&_
			" AND dbo.Activity.Site_ID=" & Site_ID & "  and Activity.Calendar_ID > 200 and Activity.Calendar_ID is not null"&_
			" and UserData.FirstName is not null and Calendar.Category_ID = Calendar_Category.ID"

		else
			SQL=  "SELECT dbo.Calendar.Title as AssetTitle, dbo.Activity.Calendar_ID as AssetID,dbo.Calendar_Category.Title as AssetCategory, " &_
			"Convert(varchar, View_Time, 101) as AssetDate, Convert(varchar, View_Time, 108) as "&_
			"AssetTime, dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, "&_
			"(isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,"&_
			"(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', "&_
			"' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') "&_
			"+ ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') "&_
			"as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, "&_
			"isnull(dbo.UserData.Business_Phone,'') as Phone, Activity.Country as ActCountry "&_
			"FROM dbo.UserData RIGHT OUTER JOIN dbo.Activity "&_
			"ON dbo.UserData.ID = dbo.Activity.Account_ID and dbo.UserData.Site_ID=dbo.Activity.Site_ID "&_
			"RIGHT OUTER JOIN dbo.Calendar "&_
			"ON dbo.Activity.Calendar_ID = dbo.Calendar.ID "&_
			"Left Outer JOIN dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code AND dbo.Calendar.Site_ID = "&_
			"dbo.Calendar_Category.Site_ID WHERE (DATEPART(m, dbo.Activity.View_Time) = " & y & ") and DATEPART(yyyy, View_Time) = " & Summary_Year & ""&_
			" AND dbo.Activity.Site_ID=" & Site_ID & "  and Activity.Calendar_ID > 200 and Activity.Calendar_ID is not null"&_
			" and UserData.FirstName is not null and Calendar.Category_ID = Calendar_Category.ID"
		end if


		select case UCase(request("Region"))     ' Filter Results by Region
		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " AND (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select

		'SQL=SQL & "group by dbo.Calendar.Title,dbo.Activity.Calendar_ID,dbo.Activity.Account_ID,dbo.Calendar.Code,View_Time,FirstName,LastName,Job_Title,Company,Business_Address,Business_Address_2,Business_City,Business_State,Business_Postal_Code,Business_Country,Email,Business_Fax,Business_Phone"
		
		GenerateExcel()
'''' Call Excel. For printing Category, loop through Calendar Code and find title from Calendar_Category table'''

case 3
		SQL= "Select dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.Account_ID=1 and Activity.CMS_Site is null and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID and Activity.Calendar_ID is not null" 
select case UCase(request("Region"))     ' Filter Results by Region
		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select

GenerateExcel()

'''' Call Excel. '''

case 4
		SQL= "Select dbo.Activity.Account_ID,(dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.CMS_Site is not null  and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID and Activity.Calendar_ID is not null" 
select case UCase(request("Region"))     ' Filter Results by Region
		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select

'''' Call Excel. '''
GenerateExcel()

case 5
		SQL= "Select dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.CID=9004 and Activity.Calendar_ID is not null and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 
select case UCase(request("Region"))     ' Filter Results by Region
		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select
GenerateExcel()
'' Call Excel'''
case 6
		SQL= "Select dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.Calendar_ID=0 and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 
select case UCase(request("Region"))     ' Filter Results by Region
		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select
'SQL=SQL & "group by dbo.Calendar.Title,dbo.Activity.Calendar_ID,dbo.Activity.Account_ID,dbo.Calendar.Code,View_Time,FirstName,LastName,Job_Title,Company,Business_Address,Business_Address_2,Business_City,Business_State,Business_Postal_Code,Business_Country,Email,Business_Fax,Business_Phone"
		
'Response.write SQL
		GenerateExcel()
case 7
		SQL= "Select dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & " and Activity.Calendar_ID is not null and DATEPART(yyyy, View_Time) = " & Summary_Year & " and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 

		select case UCase(request("Region"))     ' Filter Results by Region

		 case "0", ""  ' Worldwide no filter
              case "1"
                SQL = SQL & " AND (Activity.Region=0 OR Activity.Region=1) "
              case "2"
                SQL = SQL & " AND (Activity.Region=2) "
              case "3"
                SQL = SQL & " AND (Activity.Region=3) "
              case else
                if not isblank(request("Region")) then              
                  SQL = SQL & " (Activity.Country='" & UCase(request("Region")) & "') "
                end if  
            end select

GenerateExcel()

case 8
	
		SQL="SELECT *FROM UserData WHERE Site_ID=" & Site_ID &"  AND (DATEPART(yyyy, UserData.Reg_Request_Date) = " & Summary_Year & ")"

			select case UCase(request("Region"))     ' Filter Results by Region
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "')"
                    end if  
                end select
GenerateExcel()
case 9

		SQL="SELECT *FROM UserData WHERE NewFlag=-1 AND Site_ID=" & Site_ID & " AND (DATEPART(yyyy, UserData.Reg_Request_Date) = " & Summary_Year & ")" 

                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') "
                    end if  
                end select
GenerateExcel()
case 10

		SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=" & Site_ID & " AND (DATEPART(yyyy, UserData.Reg_Approval_Date) = " & Summary_Year & ")"
 select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') "
                    end if  
                end select
GenerateExcel()
case 11

	dim Curr_Year
		Curr_Year=Year(Date)
 		Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
		Last_Day_Pre_Month=DateAdd("m",-1,"1/1/" & (Summary_Year + 1))
		First_Day_Curr_Year=DateAdd("d",0,"1/1/" & Summary_Year)
		
		if CInt(Curr_Year)=Cint(Summary_Year) then

                 

			SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=" & Site_ID & " AND (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Month & "')AND (ExpirationDate >= '" & First_Day_Curr_Year & "' AND ExpirationDate < '" & Last_Day_Pre_Month & "')"
		else
			First_Day_Pre_Year=DateAdd("d",-1,"1/1/" & (Summary_Year))
			SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=" & Site_ID & " AND (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Month & "')AND (ExpirationDate > '" & First_Day_Pre_Year & "' AND ExpirationDate <= '" & Last_Day_Month & "')"

		end if
 		 select case UCase(request("Region"))     ' Filter Results by Region
                  
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') "
                    end if  
                end select
GenerateExcel()
case 12

		Year_Start_Date=DateAdd("d",0,"1/1/" & Summary_Year)
		Year_End_Date=DateAdd("d",-1,"1/1/" & (Summary_Year+1))

		SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=" & Site_ID & " AND (Reg_Approval_Date >= '" & Year_Start_Date & "' AND Reg_Approval_Date <= '" & Year_End_Date & "') AND Logon IS NULL or Logon=''"
 		
		select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then                  
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "')"
                    end if  
                end select

GenerateExcel()
case 13
		Curr_Year=Year(Date)
		if Cint(Curr_Year)=Cint(Summary_Year) then
			Curr_Month=Month(date)
			Last_Day_Curr_Month=DateAdd("d",-1,(Curr_Month + 1) & "/1/" & Summary_Year)
		else
			Last_Day_Curr_Month=DateAdd("d",-1,"1/1/" & (Summary_Year + 1))

		end if
		
		SQL="SELECT *FROM UserData WHERE NewFlag=0 AND Site_ID=" & Site_ID & " AND (Reg_Approval_Date >= '1/1/2001' AND Reg_Approval_Date <= '" & Last_Day_Curr_Month & "') AND ExpirationDate >'" & Last_Day_Curr_Month & "' "

                select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0"
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') "
                    end if  
                end select
GenerateExcel()
case 14

		Curr_Year=Year(Date)
		if Cint(Curr_Year)=Cint(Summary_Year) then
			Curr_Month=Month(date)
			Last_Day_Month =DateAdd("d",-1,(Curr_Month + 1) & "/1/" & Summary_Year)
		else
			Curr_Month=12
			Last_Day_Month =DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
		end if

		'SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=11 AND (Logon >= '1/1/" & Summary_Year & "' AND Logon <= '" & Last_Day_Month  & "')"


'''S
Curr_Year=Year(Date)
		if Cint(Curr_Year)=Cint(Summary_Year) then
			Curr_Month=Month(date)

		for j=1 to Curr_Month

			if j+1 < 13 then
				Last_Day_Month =DateAdd("d",-1,(j + 1) & "/1/" & (Summary_Year ))
			else
				Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			end if


			SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=11 AND (Logon >= '" & j & "/1/" & Summary_Year & "' AND Logon <= '" & Last_Day_Month  & "')"

			
				''' Call Excel '''	
		next
	
		else
	
		for j=1 to 12
			if j+1 < 13 then
				Last_Day_Month =DateAdd("d",-1,(j + 1) & "/1/" & (Summary_Year ))
			else
				Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			end if
			
			SQL="SELECT * FROM UserData WHERE NewFlag=0 AND Site_ID=11 AND (Logon >= '" & j & "/1/" & Summary_Year & "' AND Logon <= '" & Last_Day_Month  & "')"


		
			
				''' Call Excel '''
		next
		
		end if

'''E

 select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Business_Country='" & UCase(request("Region")) & "') "
                    end if  
                end select
GenerateExcel()
case 15
		Curr_Year=Year(Date)
		if Cint(Curr_Year)=Cint(Summary_Year) then
			Curr_Month=Month(date)

		for j=1 to Curr_Month

			if j+1 < 13 then
				Last_Day_Month =DateAdd("d",-1,(j + 1) & "/1/" & (Summary_Year ))
			else
				Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			end if


			SQL= "Select DISTINCT dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & "  and (View_Time >= '" & j & "/1/" & Summary_Year & "' and View_Time <= '" & Last_Day_Month & "') and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 
		
				''' Call Excel '''	
		next
	
		else
	
		for j=1 to 12
			if j+1 < 13 then
				Last_Day_Month =DateAdd("d",-1,(j + 1) & "/1/" & (Summary_Year ))
			else
				Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			end if
			response.write Last_Day_Month & "<br>"
			'Last_Day_Month =DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			SQL= "Select DISTINCT dbo.Activity.Account_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & "  and (View_Time >= '" & j & "/1/" & Summary_Year & "' and View_Time <= '" & Last_Day_Month & "') and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID" 
			
		'response.write SQL & "<br>"
			
				''' Call Excel '''
		next
		
		end if

		'SQL="SELECT *FROM Activity WHERE Site_ID=11 AND (View_Time >= '1/1/2009' AND View_Time <= '12/31/2009')"
		 select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Country='" & UCase(request("Region")) & "') AND "
                    end if  
                end select

GenerateExcel()
case 16

	
Curr_Year=Year(Date)
		if Cint(Curr_Year)=Cint(Summary_Year) then
			Curr_Month=Month(date)

		for j=1 to Curr_Month

			if j+1 < 13 then
				Last_Day_Month =DateAdd("d",-1,(j + 1) & "/1/" & (Summary_Year ))
			else
				Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			end if


			SQL= "Select DISTINCT Session_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & "  and (View_Time >= '" & j & "/1/" & Summary_Year & "' AND View_Time <= '" & Last_Day_Month & "') and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID and Activity.Session_ID <> 0" 
		'response.write SQL & "<br>"

				''' Call Excel '''	
		next
	
		else
	
		for j=1 to 12
			if j+1 < 13 then
				Last_Day_Month =DateAdd("d",-1,(j + 1) & "/1/" & (Summary_Year ))
			else
				Last_Day_Month = DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			end if
			'response.write Last_Day_Month & "<br>"
			'Last_Day_Month =DateAdd("d",-1,"1/1/" & (Summary_Year + 1))
			SQL= "Select DISTINCT Session_ID, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     " From dbo.Activity, UserData where Activity.Site_ID= " & Site_ID & "  and (View_Time >= '" & j & "/1/" & Summary_Year & "' AND View_Time <= '" & Last_Day_Month & "') and Activity.Account_ID=UserData.ID and Activity.Site_ID=UserData.Site_ID and Activity.Session_ID <> 0" 
		

		'response.write SQL & "<br>"
			
				''' Call Excel '''
		next
		
		end if

		'SQL="SELECT *FROM Activity WHERE Site_ID=11 AND (View_Time >= '1/1/2009' AND View_Time <= '12/31/2009')"
		 select case UCase(request("Region"))     ' Filter Results by Region
    
                  case "0", ""
                  case "1"
                    SQL = SQL & " (Region=0 OR Region=1) AND "
                  case "2"
                    SQL = SQL & " (Region=2) AND "
                  case "3"
                    SQL = SQL & " (Region=3) AND "
                  case else
                    if not isblank(request("Region")) then
                      SQL = SQL & " (Country='" & UCase(request("Region")) & "') "
                    end if                      
                end select
GenerateExcel()
case 17

		SQL="SELECT Distinct Order_Number, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     "FROM dbo.Shopping_Cart_Lit S, dbo.UserData  WHERE (S.Site_ID = " & Site_ID & ")  AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & ")" &_
		     "and S.Account_ID=UserData.ID " &_
		     "and S.Site_ID=UserData.Site_ID"
 
		select case UCase(request("Region"))     ' Filter Results by Region

                case "0", ""                 
                case "1"
                  SQL = SQL & " AND(Region=0 OR Region=1)) "
                case "2"
                  SQL = SQL & " AND(Region=2)) "
                case "3"
                  SQL = SQL & " AND(Region=3)) "
                case else
                  if not isblank(request("Region")) then                
                    SQL = SQL & " AND(Country='" & UCase(request("Region")) & "')) "
                  end if  
              end select
GenerateExcel()
Case 18

		SQL="SELECT Distinct Item_Number, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     "FROM dbo.Shopping_Cart_Lit S, dbo.UserData  WHERE (S.Site_ID = " & Site_ID & ")  AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & ")" &_
		     "and S.Account_ID=UserData.ID " &_
		     "and S.Site_ID=UserData.Site_ID"

 select case UCase(request("Region"))     ' Filter Results by Region
  
                case "0", ""
                 
                case "1"
                  SQL = SQL & " AND(Region=0 OR Region=1)) "
                case "2"
                  SQL = SQL & " AND(Region=2)) "
                case "3"
                  SQL = SQL & " AND(Region=3)) "
                case else
                  if not isblank(request("Region")) then                
                    SQL = SQL & " AND(Country='" & UCase(request("Region")) & "')) "
                  end if  
              end select
GenerateExcel()

Case 19

		SQL="SELECT Quantity, (dbo.UserData.FirstName + ' ' + dbo.UserData.LastName) as Name, (isnull(dbo.UserData.Job_Title,'')) as Title, (isnull(dbo.UserData.Company,'')) as Company,(dbo.UserData.Business_Address + ' ' + isnull(dbo.UserData.Business_Address_2,'') + ', ' + isnull(dbo.UserData.Business_City,'') + ', ' + isnull(dbo.UserData.Business_State,'') + ' ' + isnull(dbo.UserData.Business_Postal_Code,'')) as Address, isnull(dbo.UserData.Business_Country,'') as Country, isnull(dbo.UserData.Email,'') as Email, isnull(dbo.UserData.Business_Fax,'') as Fax, isnull(dbo.UserData.Business_Phone,'') as Phone " &_
		     "FROM dbo.Shopping_Cart_Lit S, dbo.UserData  WHERE (S.Site_ID = " & Site_ID & ")  AND (DATEPART(yyyy, Submit_Date) = " & Summary_Year & ")" &_
		     "and S.Account_ID=UserData.ID " &_
		     "and S.Site_ID=UserData.Site_ID"

 select case UCase(request("Region"))     ' Filter Results by Region
  
                case "0", ""
                  
                case "1"
                  SQL = SQL & " AND(Region=0 OR Region=1)) "
                case "2"
                  SQL = SQL & " AND(Region=2)) "
                case "3"
                  SQL = SQL & " AND(Region=3)) "
                case else
                  if not isblank(request("Region")) then                
                    SQL = SQL & " AND(Country='" & UCase(request("Region")) & "')) "
                  end if  
              end select
GenerateExcel()
End Select

		
'next 


%>
<%

Sub GenerateExcel()
'response.write SQL
'response.end
'conn.CommandTimeout = 7200
set rsReport=Server.CreateObject("Adodb.Recordset")
rsReport.open SQL,conn,3, 3
if not rsReport.eof then
	ActCountry=rsReport("ActCountry")
end if
set rsCountry=Server.CreateObject("Adodb.recordset")
rsCountry.open "select Country.name from Country where Country.Abbrev='" & ActCountry & "'",conn,3,3
if not rsCountry.eof then
	ActCountryName=rsCountry("Name")
end if
rsCountry.close()
'response.write y

If region="0" or region="" then
	regionName=Translate("Worldwide",Login_Language,conn)
elseif region="1" then
	regionName=Translate("United States",Login_Language,conn)
elseif region="2" then
	regionName=Translate("Europe",Login_Language,conn)
elseif region="3" then
	regionName=Translate("Intercon",Login_Language,conn) 
else
	regionName=ActCountryName

end if
%>

<TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>
<% if y="" then %>
<tr><th colspan=9>Site Activity Summary Report for Year <%=Summary_Year%> &nbsp; &nbsp; &nbsp; &nbsp; Region: <%=regionName%></th></tr>
<%
else
mname=Monthname(Cint(y))
dtDisplay=mname& ", " & Summary_Year
%>
<tr><th colspan=9>Site Activity Summary Report for <%=dtDisplay%> &nbsp; &nbsp; &nbsp; &nbsp; Region: <%=regionName%></th></tr>
<%
end if
%>
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
dim cnt
cnt=0
'response.write "Timeout is " & Server.ScriptTimeout

do while not rsReport.eof
	if not rsReport("Name")="" then

Address=""
Country=""
Address=rsReport("Address")

if not rsReport("Country")="" then
	Country=DisplayCountry(rsReport("Country"))
end if

Address= Address+ ", " & Country

 %>
 <TR>
   <TD><font size=2 face="Arial"><%=rsReport("AssetID")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("AssetTitle")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsREport("AssetCategory")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("AssetDate")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("AssetTime")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("Name")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("Title")%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("Company")%></font></TD>
   <TD><font size=2 face="Arial"><%=Address%></font></TD>
   <TD><font size=2 face="Arial"><%=CStr(rsReport("Phone"))%></font></TD>
   <TD><font size=2 face="Arial"><%=rsReport("Email")%></font></TD>
   <TD><font size=2 face="Arial"><%=Cstr(rsReport("Fax"))%></font></TD>
</TR>
<% 

	end if
	cnt=cnt+1
	if cnt=100 then
		Response.Flush
		cnt=0
	end if
   rsReport.MoveNext
   loop
   rsReport.Close
   set rsReport= Nothing
  End Sub
%>

<%
	Function DisplayCountry(byval countryCode)

	dim code
	code=countryCode
	SQL1 = "SELECT * FROM Country WHERE Enable=" & CInt(True) & " and Abbrev='" & code & "'"
  	Set rsCountries = Server.CreateObject("ADODB.Recordset")
 	 rsCountries.Open SQL1, conn, 3, 3
	if not rsCountries.eof then
		Name=rsCountries("Name")
	end if
	rsCountries.close()
	DisplayCountry = Name
	end function
%>
<%
	Function DisplayAssetCode(byval CatCode)

	dim Catcode1
	Catcode1=CatCode
	SQL2 = "SELECT * FROM Calendar_Category WHERE Code='" & CatCode1 & "' and Site_ID=" & Site_ID & ""
  	Set rsCatCode = Server.CreateObject("ADODB.Recordset")
 	rsCatCode.Open SQL2, conn, 3, 3
	if not rsCatCode.eof then
		CategoryCode=rsCatCode("Title")
	end if
	rsCatCode.close()
	DisplayAssetCode=CategoryCode
	end function
%>

