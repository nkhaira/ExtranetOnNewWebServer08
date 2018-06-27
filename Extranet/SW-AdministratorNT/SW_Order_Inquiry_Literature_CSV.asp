<%
' --------------------------------------------------------------------------------------
' Title:  Literature Order CSV File Extract
' Author: Kelly Whitlock
' Date:   12/17/2003
' --------------------------------------------------------------------------------------

Dim Site_ID, Site_Code

if not isblank(request("Site_ID")) then
  Site_ID = request("Site_ID")
else
  Site_ID = 3 
end if

if not isblank(request("Site_Code")) then
  Site_Code = request("Site_Code")
else
  Site_Code = "Find-Sales" 
end if

if not isblank(request("Language")) then
  Login_Language = request("Login_Language")
else
  Login_Language = "eng"
end if



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

Set fstemp = server.CreateObject("Scripting.FileSystemObject")
Extract_Path = "/" & Site_Code & "/download/LOS_Extract.csv"
File_path = server.mappath(Extract_Path)
Set FileTemp = fstemp.CreateTextFile(File_Path, True)

' Begin Date


' Search for Item Numbers (Not Null will exclude certain other filters)

if not isblank(request("Item_Numbers")) then
  item_numbers = replace(request("Item_Numbers")," ","")
else
  item_numbers = ""
end if

' Begin Date

if isblank(request("Begin_Date")) then
  Begin_Date = Date()
elseif isdate(request("Begin_Date")) then
  Begin_Date = request("Begin_Date")
end if

' Interval

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

' Country

if not isblank(request("Country_Code")) then
  Country_Code = request("Country_Code")
else
  Country_Code = "all"
end if  

' Sort by

if not isblank(request("Sort_By")) then
  Sort_By = request("Sort_By")
else
  Sort_By = 0
end if

Dim BackURL
BackURL = "/sw-administrator/Site_Utility.asp?ID=site_utility&Utility_ID=73&Begin_Date=" & Begin_Date & "&Interval=" & Interval & "&Country_Code=" & Country_Code & "&Sort_By=" & Sort_By & "&Item_Numbers=" & Item_Numbers
   
' --------------------------------------------------------------------------------------
' Individual Item Numbers
' --------------------------------------------------------------------------------------

if Interval >= 0  then
  SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= CONVERT(DATETIME, '" & Begin_Date & "', 102) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date, 102), 102) <= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102))) "
else
  SQLWhere = "WHERE (CONVERT(DATETIME,CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) >= DATEADD(d, " & Interval & ", CONVERT(DATETIME, '" & Begin_Date & "', 102)) AND CONVERT(DATETIME, CONVERT(Char(10),dbo.Shopping_Cart_Lit.Submit_Date,102), 102) <= CONVERT(DATETIME, '" & Begin_Date & "', 102)) "
end if

SQL = "SELECT dbo.UserData.*, dbo.Shopping_Cart_Lit.*, dbo.Literature_Items_US.COST_CENTER AS Cost_Center, " &_
      "       dbo.Shopping_Cart_Ship_Tracking.Name AS Ship_Carrier, dbo.Shopping_Cart_Lit.Site_ID AS Site_ID " &_
      "FROM   dbo.Shopping_Cart_Lit LEFT OUTER JOIN " &_
      "       dbo.Literature_Items_US ON dbo.Shopping_Cart_Lit.Item_Number = dbo.Literature_Items_US.ITEM LEFT OUTER JOIN " &_
      "       dbo.Shopping_Cart_Ship_Tracking ON dbo.Shopping_Cart_Lit.Ship_Carrier = dbo.Shopping_Cart_Ship_Tracking.ID LEFT OUTER JOIN " &_
      "       dbo.UserData ON dbo.Shopping_Cart_Lit.Account_ID = dbo.UserData.ID " 

      SQL = SQL & SQLWhere & " AND (dbo.Shopping_Cart_Lit.Submit_Date IS NOT NULL) AND (dbo.Shopping_Cart_Lit.Cart_Type = 'lit_us') "
              
item_numbers_max = 0
if not isblank(item_numbers) then
  if instr(item_numbers,",") > 0 then
    item_number = Split(item_numbers,",")
    item_number_max = Ubound(item_number)
  else
    reDim item_number(0)
    item_number(0) = item_numbers
  end if

  for x = 0 to item_number_max
    if x = 0 then
      SQL = SQL & "AND ("
    else
      SQL = SQL & "OR "  
    end if
    if len(item_number(x)) = 8 and instr(1,trim(item_number(x)),"0") = 1 then
      SQL = SQL & "Order_Number LIKE '%" & Trim(item_number(x)) & "' "
    else  
      SQL = SQL & "Item_Number='" & Trim(item_number(x)) & "' "
    end if  
  next
  SQL = SQL & ") "
end if    
      
if LCase(Country_Code) <> "all" then
    SQL = SQL & " AND (dbo.UserData.Business_Country='" & Country_Code & "') "
end if  

select case Sort_By
  case 0
    SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.Item_Number, dbo.Shopping_Cart_Lit.Submit_Date Desc"
  case 1
    SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.Order_Number DESC, dbo.Shopping_Cart_Lit.Item_Number"
  case 2
    SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.LastName, dbo.Shopping_Cart_Lit.FirstName, dbo.Shopping_Cart.Item_Number"
  case 3
    SQL = SQL & "ORDER BY dbo.Shopping_Cart_Lit.Submit_Date Desc, dbo.Shopping_Cart_Lit.LastName, dbo.Shopping_Cart_Lit.FirstName, dbo.Shopping_Cart.Item_Number"
  case 4
    SQL = SQL & "ORDER BY dbo.UserData.Company, dbo.Shopping_Cart_Lit.Submit_Date Desc, dbo.Shopping_Cart_Lit.LastName, dbo.Shopping_Cart_Lit.FirstName, dbo.Shopping_Cart.Item_Number"
end select

'response.write SQL & "<P>"

Set rsActivity = Server.CreateObject("ADODB.Recordset")
rsActivity.Open SQL, conn, 3, 3

' --------------------------------------------------------------------------------------

Column_Titles = """Item Number""," &_
                """Quantity""," &_
                """Cost Center""," &_
                """Asset ID""," &_
                """Order Number""," &_
                """Ship Status""," &_
                """Order Date""," &_
                """Ship Date""," &_
                """Ship Delta Days""," &_
                """Ship Carrier""," &_
                """Ship Tracking No""," &_
                """Ordr Company""," &_
                """Ordr First Name""," &_
                """Ordr Last Name""," &_
                """Ordr Job Title""," &_
                """Ordr Email""," &_
                """Ordr Phone""," &_
                """Ordr Address""," &_
                """Ordr Address 2""," &_
                """Ordr City""," &_
                """Ordr State""," &_
                """Ordr State Other""," &_
                """Ordr Postal Code""," &_
                """Ordr Country""," &_
                """Ordr Region""," &_
                """Ordr Fluke ID""," &_
                """Ship Company""," &_
                """Ship First Name""," &_
                """Ship Last Name""," &_
                """Ship Job Title""," &_
                """Ship Address""," &_
                """Ship Address 2""," &_
                """Ship City""," &_
                """Ship State""," &_
                """Ship State Other""," &_
                """Ship Country""," &_
                """Ship Postal Code""," &_
                """Ship Phone""," &_
                """Ship Email"""

filetemp.writeLine(Column_Titles)

Old_Ship_ID = -1

do while not rsActivity.EOF

  Order_Info =  """" & rsActivity("Item_Number") & """," &_
                """" & rsActivity("Quantity") & """," &_
                """" & rsActivity("Cost_Center") & """," &_
                """" & rsActivity("Asset_ID") & """," &_
                """" & Replace(rsActivity("Order_Number"),"FLUKECO","") & ""","
                
                select case rsActivity("Order_Status")
                  case 10, 15
                    Order_Status =  """Shipped"","
                  case 16
                    Order_Status =  """See Note 1"","
                  case 80
                    Order_Status =  """Back-Ordered"","
                  case 99
                    Order_Status =  """Cancelled"","
                  case 0
                    Order_Status =  """In Process"","
                  case else
                    Order_Status =  """Unknown"","
                 end select
                
  Order_Info =  Order_Info & Order_Status &_                        
                """" & rsActivity("Submit_Date") & """," &_
                """" & rsActivity("Order_Ship_Date") & ""","
                
                Delta = rsActivity("Order_Ship_Date")
                if isblank(Delta) then
                  Delta = Date()
                end if  

  Order_Info =  Order_Info & """" & DateDiff("D",rsActivity("Submit_Date"),Delta) & """," &_
                """" & rsActivity("Ship_Carrier") & """," &_
                """" & rsActivity("Ship_Tracking_No") & """," &_
                """" & rsActivity("Company") & """," &_
                """" & rsActivity("FirstName") & """," &_
                """" & rsActivity("LastName") & """," &_
                """" & rsActivity("Job_Title") & """," &_
                """" & rsActivity("Email") & """," &_
                """" & rsActivity("Business_Phone") & """," &_
                """" & rsActivity("Business_Address") & """," &_
                """" & rsActivity("Business_Address_2") & """," &_
                """" & rsActivity("Business_City") & """," &_
                """" & rsActivity("Business_State") & """," &_
                """" & rsActivity("Business_State_Other") & """," &_
                """" & rsActivity("Business_Postal_Code") & """," &_
                """" & rsActivity("Business_Country") & """," &_
                """" & rsActivity("Region") & """," &_
                """" & rsActivity("Business_System") & " " & rsActivity("Fluke_ID") & ""","

  if CInt(Old_Ship_ID) <> CInt(rsActivity("Shipping_Address_ID")) and CInt(rsActivity("Shipping_Address_ID")) > 0 then

    SQLShip = "SELECT * FROM Shopping_Cart_Ship_To WHERE ID=" & rsActivity("Shipping_Address_ID")
    Set rsShip = Server.CreateObject("ADODB.Recordset")
    rsShip.Open SQLShip, conn, 3, 3
    
    if not rsShip.EOF then
      Shipping_Info = """" & rsShip("Company") & """," &_
                      """" & rsShip("FirstName") & """," &_
                      """" & rsShip("LastName") & """," &_
                      """" & rsShip("Job_Title") & """," &_
                      """" & rsShip("Shipping_Address") & """," &_
                      """" & rsShip("Shipping_Address_2") & """," &_
                      """" & rsShip("Shipping_City") & """," &_
                      """" & rsShip("Shipping_State") & """," &_
                      """" & rsShip("Shipping_State_Other") & """," &_
                      """" & rsShip("Shipping_Country") & """," &_
                      """" & rsShip("Shipping_Postal_Code") & """," &_
                      """" & rsShip("Business_Phone") & """," &_
                      """" & rsShip("Email") & """"
    end if
    
    rsShip.Close
    set rsShip = nothing
    
  else  ' If Shiping_Address_ID = 0 then Shipping Address is the Account

      Shipping_Info = """" & rsActivity("Company") & """," &_
                      """" & rsActivity("FirstName") & """," &_
                      """" & rsActivity("LastName") & """," &_
                      """" & rsActivity("Job_Title") & """," &_
                      """" & rsActivity("Shipping_Address") & """," &_
                      """" & rsActivity("Shipping_Address_2") & """," &_
                      """" & rsActivity("Shipping_City") & """," &_
                      """" & rsActivity("Shipping_State") & """," &_
                      """" & rsActivity("Shipping_State_Other") & """," &_
                      """" & rsActivity("Shipping_Country") & """," &_
                      """" & rsActivity("Shipping_Postal_Code") & """," &_
                      """" & rsActivity("Business_Phone") & """," &_
                      """" & rsActivity("Email") & """"
  end if  
  
  Old_Ship_ID = rsActivity("Shipping_Address_ID")

  filetemp.writeLine(Order_Info & Shipping_Info)

  rsActivity.MoveNext

loop
  
rsActivity.close
set rsActivity = nothing

filetemp.close
set filetemp=nothing
set fstemp=nothing

' --------------------------------------------------------------------------------------
' Download Screen
' --------------------------------------------------------------------------------------

Site_Description = "Literature Order System"
    
Screen_Title     = Site_Description & " - " & Translate("LOS Administrator",Alt_Language,conn)
Bar_Title        = Screen_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Literature Order Extract - CSV File",Login_Language,conn) & "</FONT>"

Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%
response.write "<FONT CLASS=SMALL>"
response.write "<OL>"
response.write "<LI><A HREF=""" & Extract_Path & """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Click Here to View the Extract",Login_Language,conn) & "&nbsp;&nbsp;</A></LI><BR><BR>"
response.write "<LI>" & Translate("After the File Dowload Dialog appears (see example below), answer the question, &quot;What do you want to do with this file?&quot;",Login_Language,conn) & "<BR><BR>"
response.write Translate("Select either &quot;Open this file from its current location&quot; to view in Excel, or &quot;Save this file to Disk&quot; to save this file to your local drive to be opened at a later time.",Login_Language,conn) & "</LI><BR><BR>"
response.write "<LI><IMG SRC=""/images/File_Download_PopUp.jpg"" BORDER=0><BR><BR>"
response.write Translate("Click on [OK] to begin.",Login_Language,conn) & "</LI><BR><BR>"
response.write "<LI>" & Translate("After the Extract file loads, if you are viewing this file on-line, click on the [Back] button of your browser to return to this screen.",Login_Language,conn) & "</LI><BR><BR>"
response.write "<LI>"
Call Main_Menu
response.write "</LI><BR><BR>"
response.write "</OL>"
response.write "</FONT>"
%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

call Disconnect_Sitewide

sub Main_Menu

  response.write "<A HREF=""javascript:void(0);"" TITLE=""Literature Order Query"
  response.write """ LANGUAGE=""JavaScript"" onclick=""location.href='" & BackURL & "'; return false;"">"
  response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;" & Translate("Click Here to Return to Your Previous View",Login_Language,conn) & "&nbsp;</SPAN></A>"

end sub 

%>
