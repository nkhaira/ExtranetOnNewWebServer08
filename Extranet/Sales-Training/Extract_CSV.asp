<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   4/1/2003
' --------------------------------------------------------------------------------------
%>

<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/Pop-Up.asp"-->  
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_formdata.asp" -->

<%
'Dim Site_ID
'Site_ID=3
Call Connect_SiteWide
  
Site_Description = "Fluke Sales Training Extranet Site"
Screen_Title    = Translate(Site_Description,Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>CSV Data File Extract</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

if not isblank(request("Course")) and isnumeric(request("Course")) then
  Course = request("Course")
else
  Course = -1
end if

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
<%
 
with response
  
SQL = "SELECT dbo.Country.Name AS Country, dbo.UserData_Sales_Training.* " &_
      "FROM dbo.Country RIGHT OUTER JOIN " &_
      "dbo.UserData_Sales_Training ON dbo.Country.Abbrev = dbo.UserData_Sales_Training.Business_Country " &_
      "WHERE "
      
      if not isblank(request("Begin_Date")) and isDate(request("Begin_Date")) then
        SQL = SQL & " (CAST(dbo.UserData_Sales_Training.Training_Dates AS DateTime) >='" & request("Begin_Date") & "' "
      else  
        SQL = SQL & " (CAST(dbo.UserData_Sales_Training.Training_Dates AS DateTime) >='" & Date() & "' "
      end if  
      
      if not isblank(request("End_Date")) and isDate(request("End_Date")) then
        SQL = SQL & " AND CAST(dbo.UserData_Sales_Training.Training_Dates AS DateTime) <='" & request("End_Date") & "') "
      else  
        SQL = SQL & " AND CAST(dbo.UserData_Sales_Training.Training_Dates AS DateTime) <='" & Date() & "') "
      end if  
      
      if not isblank(request("Country")) then
        if instr(1,request("Country"),",") > 0 then
          Country = Split(request("Country"),", ")
        else
          Redim Country(0)
          Country(0) = request("Country")
        end if

        Country_Max = UBound(Country)          
                  
        for x = 0 to Country_Max
          if x = 0 then
            SQL = SQL & " AND"
          else
            SQL = SQL & " OR"
          end if    
          SQL = SQL & " dbo.UserData_Sales_Training.Business_Country='" & Country(x) & "'"
        next
        
      end if
      
      if not isblank(Course) then
        if Course > 0 then
          SQL = SQL & " AND dbo.UserData_Sales_Training.Training_Course=" & Course
        end if
      end if
          
          
      SQL = SQL & " ORDER BY dbo.Country.Name, dbo.UserData_Sales_Training.Company, dbo.UserData_Sales_Training.LastName, dbo.UserData_Sales_Training.FirstName "

'response.write SQL & "<P>"

Set rsProfile = Server.CreateObject("ADODB.Recordset")
rsProfile.Open SQL, conn, 3, 3

SQL_Extract_Total = rsProfile.RecordCount

.write "<HR NOSHADE COLOR=""Black"">"

.write "<SPAN CLASS=SMALL>"
.write "<FORM ACTION=""Extract_CSV.asp"" METHOD=""GET"">"
.write "Beginning Date:&nbsp;"
.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""Begin_Date"" SIZE=""10"" MAXLENGTH=""10"" VALUE="""

if not isblank(request("Begin_Date")) and isDate(request("Begin_Date")) then
  .write request("Begin_Date")
else
  .write Date()
end if

.write """>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf

.write "Ending Date:&nbsp;"
.write "<INPUT CLASS=Small TYPE=""TEXT"" NAME=""End_Date"" SIZE=""10"" MAXLENGTH=""10"" VALUE="""

if not isblank(request("End_Date")) and isDate(request("End_Date")) then
  .write request("End_Date")
else
  .write Date()
end if
.write """><P>" & vbCrLf

.write "Course Filter:&nbsp;&nbsp;"

SQLCourse = "SELECT ID, Course_Title FROM Sales_Training WHERE Enabled=-1 ORDER BY ID"
Set rsCourse = Server.CreateObject("ADODB.Recordset")
rsCourse.Open SQLCourse, conn, 3, 3

.write "<SELECT NAME=""Course"" CLASS=SMALL>" & vbCrLf
.write "<OPTION VALUE=""-1"">Select course from list</OPTION>" & vbCrLf
.write "<OPTION VALUE=""0"""
if Course = 0 then .write " SELECTED"
.write ">All Courses</OPTION>" & vbCrLf

do while not rsCourse.EOF
  .write "<OPTION VALUE=""" & rsCourse("ID") & """"
  if CInt(Course) = CInt(rsCourse("ID")) then .write " SELECTED"
  .write ">" & rsCourse("Course_Title") & "</OPTION>" & vbCrLf
  
  rsCourse.MoveNext
loop

.write "</SELECT><P>" & vbCrLf

rsCourse.close
set rsCourse = nothing
  
.write "Country Filter:&nbsp;&nbsp;"

Users_Country = request("Country")
Call Connect_FormDatabase
Call DisplayCountryList("Country",Users_Country,"","Small")
Call Disconnect_FormDatabase

.write "&nbsp;&nbsp;(To Select Multiple Countries, use [CTRL] + Left Mouse Click.)<P>"

.write "<INPUT TYPE=""RESET""  NAME=""RESET"" VALUE=""Clear"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;"
.write "<INPUT TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""Go"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;"
.write "</FORM>"

.write "<P>"
.write "<LI>" & "This extract utility will create a &quot;comma separated value&quot; (CSV) file of records from the Sales-Training database." & "</LI>"
.write "<LI>" & "This extract will allow you to automatically view the data with Microsoft&reg; Excel, or any other database or spreadsheet that accepts CSV formated data." & "</LI>"
.write "<LI>" & "<B>Total records in current listing:</B><SPAN CLASS=SmallBoldRed>&nbsp;&nbsp;"  & SQL_Extract_Total & "</SPAN></LI>"
.write "</SPAN>"
.write "<BR><HR NOSHADE COLOR=""Black"">"
  
if SQL_Extract_Total > 0 then  

  Set fstemp = server.CreateObject("Scripting.FileSystemObject")
  Extract_Path = "/Find-Sales/Download/Sales_Training_Extract.csv"
  File_path = server.mappath(Extract_Path)
  ' True  = File can be over-written if it exists
  ' False = File CANNOT be over-written if it exists
  Set FileTemp = fstemp.CreateTextFile(File_Path, True)

  Header = Split("Course,First Name,Last Name,Email,Phone,Company,Address 1,Address 2,City,State/Province,Postal Code,Country,Language,Training Date,Training Modules,More_Info,Fulfillment",",")
  Header_Max = UBound(Header)

  for x = 0 to Header_Max
    Header_Row = Header_Row & """" & Header(x) & ""","
  next

  filetemp.writeLine(Mid(Header_Row,1, Len(Header_Row)-1))

  do while not rsProfile.EOF

    Data_Row = """"  & rsProfile("Training_Course") & """," &_
               """"  & rsProfile("FirstName") & """," &_  
               """"  & rsProfile("LastName") & """," &_  
               """"  & rsProfile("Email") & """," &_
               """"  & rsProfile("Business_Phone") & """," &_
               """"  & rsProfile("Company") & """," &_
               """"  & rsProfile("Business_Address") & """," &_
               """"  & rsProfile("Business_Address_2") & """," &_
               """"  & rsProfile("Business_City") & """," &_
               """"  & rsProfile("Business_State") & """," &_
               """" & rsProfile("Business_Postal_Code") & """," &_
               """"  & rsProfile("Country") & """," &_
               """"  & rsProfile("Language") & """," &_             
               """" & rsProfile("Training_Dates") & """," &_
               """"  & rsProfile("Training_Modules") & ""","
               if CInt(rsProfile("More_Info")) = CInt(True) then
                 Data_Row = Data_Row & """YES"","
               else
                 Data_Row = Data_Row & """NO"","             
               end if  
               if CInt(rsProfile("Fulfillment")) = CInt(True) then
                 Data_Row = Data_Row & """YES"""
               else
                 Data_Row = Data_Row & """NO"""             
               end if  
               
    filetemp.writeLine(Data_Row)
    rsProfile.MoveNext
    
  loop  
  
  rsProfile.close
  set rsProfile = nothing
  
  filetemp.close
  set filetemp=nothing
  set fstemp=nothing
  
  .write "<FONT CLASS=SMALL>"
  .write "<OL>"
  .write "<LI><A HREF=""/Find-Sales/Download/Sales_Training_Extract.csv"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;Click Here to View the CSV Extract File&nbsp;&nbsp;</A></LI><BR><BR>"
  .write "<LI>After the File Dowload Dialog appears (see example below), answer the question, &quot;What do you want to do with this file?&quot;<BR><BR>"
  .write "Select either &quot;Open this file from its current location&quot; to view in Excel, or &quot;Save this file to Disk&quot; to save this file to your local drive to be opened at a later time.</LI><BR><BR>"
  .write "<LI><IMG SRC=""/images/File_Download_PopUp.jpg"" BORDER=0><BR><BR>"
  .write "Click on [OK] to begin.</LI><BR><BR>"
  .write "<LI>After the Extract file loads, if you are viewing this file on-line, click on the [Back] button of your browser to return to this screen.</LI><BR><BR>"
  .write "</OL>"
  .write "</FONT>"

end if

end with

%>  

<!--End Content -->
<BR><BR>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<!--#include virtual="/include/core_countries_multi.inc"-->

<%
Call Disconnect_SiteWide
%>
