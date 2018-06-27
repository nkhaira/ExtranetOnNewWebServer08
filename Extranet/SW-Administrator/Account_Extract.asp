<%
' --------------------------------------------------------------------------------------
' Name:   Account File Extract to CSV File
' Author: K. D. Whitlock
' Date:   03/14/2001
' --------------------------------------------------------------------------------------

Script_Debug = false

if Script_Debug then
  response.write "<HTML><BODY><TABLE Border=1>"
  for each item in request.querystring
    response.write "<TR><TD>" & item & "</TD><TD>" & request.querystring(item) & "</TD></TR>"
  next
  response.write "</TABLE></BODY></HTML>"
  response.end
end if

Dim Site_ID

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
<!--#include virtual="/connections/connection_Login_Admin.asp"-->
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim File_Path
Dim Extract_Path

Dim Site_Code
Dim Site_Description
Dim SC0
Dim SC1
Dim SC2
Dim SC3
Dim SC4
Dim SC5
Dim Page
Dim Z

Site_Code = Request("Site_Code")
Site_Description = Request("Site_Description")
SC0  = Request("SC0")
SC1  = Request("SC1")
SC2  = Request("SC2")
SC3  = Request("SC3")
SC4  = Request("SC4")
SC5  = ""
Page = Request("Page")
Z    = Request("Z")

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

' Use only the WHERE clause from request("sql") since all columns need to be used but based on WHERE clause.
SQL = "SELECT UserData.* FROM UserData" & Mid(request("SQL"),Instr(1,UCase(Request("SQL"))," WHERE "))
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open SQL, conn, 3, 3

if not rsUser.EOF then

  Export_Field = Split(Request("F"),",")
  Export_Field_Max = Ubound(Export_Field)

  SQL = "SELECT * FROM Field_Names WHERE Field_Names.Table_Name='" & Request("Table") & "'"
  Set rsField_Names = Server.CreateObject("ADODB.Recordset")
  rsField_Names.Open SQL, conn, 3, 3
   
  Dim Column_Name(100)
  Column_Name_Counter = -1
  
  do while not rsField_Names.EOF
    for i = 0 to Export_Field_Max
      if CInt(rsField_Names("ID")) = CInt(Export_Field(i)) then
        Column_Name_Counter = Column_Name_Counter + 1
        Header_Row = Header_Row & """" & rsField_Names("Description") & ""","
        Column_Name(Column_Name_Counter) = Trim(rsField_Names("Field_Name"))
      end if  
    next
    rsField_Names.MoveNext
  loop
  
  rsField_Names.Close
  set rsField_Names = nothing
  
  Header_Row = Mid(Header_Row,1,Len(Header_Row)-1)
  
  ' Attempt to write file as ASCII, if fails repeat using UTF-8

  Set fstemp = server.CreateObject("Scripting.FileSystemObject")
  Extract_Path = "/" & Site_Code & "/download/Account_Extract.csv"
  File_path = server.mappath(Extract_Path)
  ' True  = File can be over-written if it exists
  ' False = File CANNOT be over-written if it exists
  Set FileTemp = fstemp.CreateTextFile(File_Path, True, False)

  filetemp.WriteLine(Header_Row)
  
  rsUser.MoveFirst

  UTF8 = false
  
  do while not rsUser.EOF
    Data_Row = ""
    for i = 0 to Column_Name_Counter
      Data_Row = Data_Row & """" & rsUser(Column_Name(i)) & ""","
    next
    if not isblank(Data_Row) then
      Data_Row = Mid(Data_Row,1,Len(Data_Row)-1)
      on error resume next
      filetemp.writeLine(Data_Row)
      if err.number <> 0 then
        UTF8 = true
        on error goto 0
        exit do
      end if
    end if
    rsUser.MoveNext  
  loop  
 
  filetemp.close
  set filetemp=nothing
  set fstemp=nothing
  
  if UTF8 = true then   ' Ascii Failed so re-write file in UTF-8
  
    Set fstemp = server.CreateObject("Scripting.FileSystemObject")
    Extract_Path = "/" & Site_Code & "/download/Account_Extract.csv"
    File_path = server.mappath(Extract_Path)
    ' True  = File can be over-written if it exists
    ' False = File CANNOT be over-written if it exists
    Set FileTemp = fstemp.CreateTextFile(File_Path, True, True)
  
    filetemp.WriteLine(Header_Row)
  
    rsUser.MoveFirst
    
    do while not rsUser.EOF
      Data_Row = ""
      for i = 0 to Column_Name_Counter
        Data_Row = Data_Row & """" & rsUser(Column_Name(i)) & ""","
      next
      if not isblank(Data_Row) then
        Data_Row = Mid(Data_Row,1,Len(Data_Row)-1)
        on error resume next
        filetemp.writeLine(Data_Row)
        if err.number <> 0 then
          UTF8 = true
          on error goto 0
          exit do
        end if
      end if
      rsUser.MoveNext  
    loop  
   
    filetemp.close
    set filetemp=nothing
    set fstemp=nothing

  end if  
  
end if

rsUser.close
set rsUser = nothing

if err.number = 0 then

  Screen_Title     = Site_Description & " - " & Translate("Account Administrator",Alt_Language,conn)
  Bar_Title        = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Administrator",Login_Language,conn) & "</FONT>"
  Bar_Title        = Bar_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Data Extract",Login_Language,conn) & "</FONT>"

  Navigation       = false
  Top_Navigation   = false
  Content_Width    = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%
  response.write "<FONT CLASS=SMALL>"
  response.write "<OL>"
  response.write "<LI><A HREF=""/" & Site_Code & "/Download/Account_Extract.csv"" CLASS=NavLeftHighlight1>&nbsp;&nbsp;"
  if UTF8 = false then
    response.write Translate("Click Here to View the Extract",Login_Language,conn)
    response.write "&nbsp;&nbsp;</A>"    
  else  
    response.write Translate("Right click here then select: ""Save Target As..."" to save the extract to your computer",Login_Language,conn)
    response.write "&nbsp;&nbsp;</A>&nbsp;&nbsp;&nbsp;&nbsp;"
    Call Main_Menu
  end if
  response.write "</LI><P>"

  if UTF8 = true then
    response.write "<LI>"
    response.write Translate("After saving the file to your computer, open Microsoft Excel or an equivalent spreadsheet software application capable of importing a .CSV file",Login_Language,conn) & "<P>"
    response.write Translate("From the Excel menu bar, select ""File | Open...""  click on the file named, ""Account_Extract.csv"" and then click on ""Open"" button.",Login_Language,conn) & "<P>"
    response.write Translate("<B>Note</B>: You may have to select from ""Files of type:"", ""All Files (*.*)"" to see this file appear in the listing.",Login_Language,conn) & "</LI><P>"
    
    response.write "<LI><B>" & Translate("Excel Text Import Wizard",Login_Language,conn) & "</B><P>"
    response.write Translate("Using the Text Import Wizard screens, configure each screen as shown in the examples below. Click [Next] to advance to the next Wizard Screen.",Login_Language,conn) & "<P>"
    response.write "<IMG SRC=""/images/Excel-TIW-1.jpg"" BORDER=0><P>"
    response.write "<IMG SRC=""/images/Excel-TIW-2.jpg"" BORDER=0><P>"
    response.write "<IMG SRC=""/images/Excel-TIW-3.jpg"" BORDER=0><P>"    
    response.write Translate("Click on [Finish] to import the Account_Extract.CSV file.",Login_Language,conn) & "</LI><BR><BR>"
  end if
  if UTF8 = false then
    response.write "<LI>" & Translate("After the Extract file loads, if you are viewing this file on-line, click on the [Back] button of your browser to return to this screen.",Login_Language,conn) & "</LI><BR><BR>"
  end if  
  response.write "<LI>"
  Call Main_Menu
  response.write "</LI><BR><BR>"
  response.write "</OL>"
  response.write "</FONT>"
  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%
  
else

  Screen_Title     = Site_Description & " - " & Translate("Account Administrator",Alt_Language,conn)
  Bar_Title        = Site_Description & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Account Administrator",Login_Language,conn) & "</FONT>"
  Bar_Title        = Bar_Title & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Error",Login_Language,conn) & "</FONT>"

  Navigation       = false
  Top_Navigation   = false
  Content_Width    = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <%

  response.write "<BR><BR>"
  Call Main_Menu
  response.write "<BR><BR>"  

  response.write "Sorry, the file you requested is not available at this time.<BR>"
  response.write "Click on the [Back] button of your browser and try again.<BR>"
  response.write "If after repeated attempts, the file is still not available, please copy the URL of this page and the complete error message below and send to David.Whitlock@Fluke.com<BR><BR>"

  response.write "<U>Error Information</U><br>"
  response.write "Error Number=#<B>" & err.number & "</B><BR>"
  response.write "Error Desc. =<B>" & err.description & "</B><BR>"
  response.write "Help Path =<B>" & err.helppath & "</B><BR>"
  response.write "Native Error=<B>" & err.nativeerror & "</B><BR>"
  response.write "Error Source =<B>" & err.source & "</B><BR>"
  response.write "SQL State=#<B>" & err.sqlstate & "</B><BR>"

  %>
  <!--#include virtual="/SW-Common/SW-Footer.asp"-->
  <%
  
end if

Call Disconnect_SiteWide

sub Main_Menu

  response.write "<A HREF=""/SW-Administrator/Account_List.asp"
  response.write "?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 & "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&SC5=" & SC5 & "&Page=" & Page & "&Z=" & Z & "&Encoding=UTF-8"
  if isnumeric(Z) then
    if Z > 0 then
      response.write "#" & Z
    end if
  end if    
  response.write """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Click Here to Return to Your Previous View",Login_Language,conn) & "&nbsp;&nbsp;</A>"

end sub 
%>