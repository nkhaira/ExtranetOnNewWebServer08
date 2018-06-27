<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   4/1/2003
'         10/18/2005 added TI30 Course that does not follow standard STC configuration...
'         so I only had 1 day to implement and did a semi-hack job.  When I have time,
'         I will rebuild the code to eliminate the funky code. (Kelly)
' --------------------------------------------------------------------------------------

response.buffer = true    ' Set in case there is a redirect

%>
  <!--#include virtual="/include/functions_string.asp"-->
  <!--#include virtual="/include/functions_file.asp"-->
  <!--#include virtual="/include/functions_date_formatting.asp"-->
  <!--#include virtual="/include/functions_translate.asp"-->
  <!--#include virtual="/SW-Common/Preferred_Language.asp"-->
  <!--#include virtual="/connections/connection_SiteWide.asp"-->
  <!--#include virtual="/connections/connection_FormData.asp"-->
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Site_ID          = 96 ' Hardwired see SiteWide.Site

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Debug_Flag
Dim SValue

if request("Debug") = "on" then
  Session("Debug_Flag") = True
elseif request("Debug") = "off" then
  Session("Debug_Flag") = False
elseif isblank(Session("Debug_Flag")) then
  Session("Debug_Flag") = False
end if    
'Session("Debug_Flag") = true
Debug_Flag = Session("Debug_Flag")

Dim Sequence, seqRegister, seqDownload, seqFlash, seqShipping, seqFulfillment, seqPortal

seqRegister    = 1
seqDownload    = 2
seqFlash       = 3
seqShipping    = 4
seqFulfillment = 5
seqDownload_2  = 6
seqPortalReg   = 7
seqPortalLogon = 8

Dim Caption(8)
Caption(0) = "Register"
Caption(1) = "Register"
Caption(2) = "Download Training Materials"
Caption(3) = "On-Line Training"
Caption(4) = "Shipping Address Information"
Caption(5) = "Fulfillment Request"
Caption(6) = "Download Resource Center"
Caption(7) = "Register for the Partner Portal"
Caption(8) = "Logon to the Partner Portal"

Dim Account_ID

Dim Form_Name
Form_Name = "Not_Set"

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title
Contrast = "#FFCC00"

if not isblank(request("Sequence")) then
  Sequence = CInt(request("Sequence"))
else
  Sequence = seqRegister
end if

Session("Sequence") = Sequence

if not isblank(request("Account_ID")) then
  Account_ID = CInt(request("Account_ID"))
  Session("Account_ID_Course") = Account_ID
elseif not isblank(Session("Account_ID_Course")) then
  Account_ID = CInt(request("Account_ID"))
else
  Account_ID = 0
end if

Dim Course
if not isblank(request("Course")) then
  Course = CInt(request("Course"))
else
  Course = 0
end if

Dim Course_Code
if not isblank(request("Course_Code")) then
  Course_Code = UCase(request("Course_Code"))
else
  Course_Code = ""
end if

select case UCase(Course_Code)

  case "TI20"
    Course = 49
    Login_Language = "eng"
    if sequence = seqDownload then
      Caption(2) = "Take the Course"
    end if

  case "TI30"
    Course = 50
    Login_Language = "eng"
    if sequence = seqDownload then
      Caption(2) = "Take the Course"
    end if
    
  case "FOODPRO"
    Course = 10
    Login_Language = "eng"
    if sequence = seqDownload then
      Caption(2) = "Take the Course"
    end if

end select

SQL = "SELECT Site_Description FROM Sales_Training WHERE ID=" & Course
Set rsCourse = Server.CreateObject("ADODB.Recordset")
rsCourse.Open SQL, conn, 3, 3

if not rsCourse.EOF then
  Site_Description = rsCourse("Site_Description")
end if

rsCourse.close
set rsCourse = nothing

Screen_Title    = Translate(Site_Description,Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><SPAN CLASS=SmallBoldGold>" & Translate(Caption(Sequence),Login_Language,conn) & "</SPAN>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

Session("Language") = Login_Language

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
<%

if Sequence = seqFulfillment then
  'Set Mailer = Server.CreateObject("SMTPsvg.Mailer") 
  'adding new email method
  %>
  <!--#include virtual="/connections/connection_email_new.asp"-->
  <%
end if  

with response

' --------------------------------------------------------------------------------------

select case Sequence

' --------------------------------------------------------------------------------------

  case seqRegister  ' Registration

    .write "<SPAN CLASS=Heading3>" & Translate(Caption(Sequence),Login_Language,conn) & "</SPAN>" & vbCrLf
    .write "<BR><BR>" & vbCrLf
    .write "<SPAN CLASS=Medium>" & vbCrLf

    Form_Name = "Register"
    %>
    <FORM ACTION="Default.asp" NAME="<%=Form_Name%>" METHOD="POST" onsubmit="return(CheckRequiredFields(this.form));" onKeyUp="highlight(event)" onClick="highlight(event)">
    <INPUT TYPE=HIDDEN NAME="Sequence" VALUE="<%=Sequence + 1%>">
    <%
    Call Table_Begin
  %>
      <TABLE CELLPADDING=4 CELLSPACING=0 BORDER=0 BGCOLOR="<%=Contrast%>" WIDTH="100%">
      
        <!-- Preferred Language -->

        <% if Course_Code <> "TI30" and Course_Code <> "TI20" and Course_Code <> "FOODPRO" then %>    
        <TR>
        	<TD CLASS=SMALLBOLD>
            <%=Translate("Preferred Language",Login_Language,conn)%> :
          </TD>
        	<TD CLASS=SMALL>
            <%
            .write "<SELECT CLASS=Small NAME=""Language"" CLASS=Medium LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='default.asp" & "?Language='+this.options[this.selectedIndex].value"">" & vbCrLf

            SQL = "SELECT Language.* FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
            Set rsLanguage = Server.CreateObject("ADODB.Recordset")
            rsLanguage.Open SQL, conn, 3, 3
                                  
            Do while not rsLanguage.EOF
              select case rsLanguage("Code")
                case "eng", "por", "spa", "fre"
                  if LCase(rsLanguage("Code")) = LCase(Login_Language) then
                 	  .write "<OPTION CLASS=Small SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
                  else
                 	  .write "<OPTION Class=Small VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
                  end if
              end select  
          	  rsLanguage.MoveNext 
            loop
            
            rsLanguage.close
            set rsLanguage=nothing

            .write "</SELECT>" & vbCrLf
            %>
          </TD>
        </TR>
        <%
        else
          .write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""eng"">"
        end if
        
        if isblank(Course_Code) or Course_Code = "MANUALLY" then
          %>
          <TR>
            <TD WIDTH="40%" CLASS=SMALLBOLD><%=Translate("Course",Login_Language,conn)%> :</TD>
            <TD WIDTH="60%" CLASS=SMALL>
              <%
              if Course_Code <> "MANUALLY" then
                SQL = "SELECT ID, Course_Code, Course_Title FROM Sales_Training WHERE DropDown_Show=" & CInt(True) & " ORDER BY Course_Title"
                  
                Set rsCourse = Server.CreateObject("ADODB.Recordset")
                rsCourse.Open SQL, conn, 3, 3
              
                .write "<SELECT CLASS=Small NAME=""Course_Code"" ONCHANGE=""CkManual();"">" & vbCrLf
                .write "<OPTION Class=Small VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
                
                do while not rsCourse.EOF
                  .write "<OPTION Class=Small VALUE=""" & UCase(rsCourse("Course_Code")) & """"
                  if Course_Code = UCase(rsCourse("Course_Code")) then response.write " SELECTED"                    
'                    if CInt(Course) = CInt(rsCourse("ID")) then response.write " SELECTED"
                  .write ">" & Translate(rsCourse("Course_Title"),Login_Language,conn) & "</OPTION>" & vbCrLf
                  rsCourse.MoveNext
                loop
                .write "</OPTION>"
                .write "<OPTION VALUE=""""></OPTION>" & vbCrLf
                .write "<OPTION Class=ZeroValue VALUE=""manually"">" & Translate("Enter Course Code Manually",Login_Language,conn) & "" & "</OPTION>" & vbCrLf
              
                rsCourse.close
                set rsCourse = nothing
                
                .write "</SELECT>" & vbCrLf
              else
                .write "<INPUT TYPE=""Text"" NAME=""Course_Code"" SIZE=""15"" MAXLENGTH=""50"" VALUE="""" CLASS=Small>" & vbCrLf
                .write "&nbsp;&nbsp;"
                .write "<A HREF=""/sales-training/default.asp?Language=" & Login_Language & """ CLASS=NavLeftHighlight1><SPAN CLASS=NAVLEFT1>&nbsp;&nbsp;" & Translate("Show List",Login_Language,conn) & "&nbsp;&nbsp;</SPAN></A>" & vbCrLf
              end if  
              %>
            </TD>
          </TR>
        
          <TR><TD COLSPAN=2 BGCOLOR="#777777" CLASS=SMALLWHITE><%=Translate("Select Preferred Language and Course before continuing below.",Login_language,conn)%></TR>

        <%
        else
          .write "<INPUT TYPE=""HIDDEN"" NAME=""Course_Code"" VALUE=""" & Course_Code & """>" & vbCrLf
        end if  
        %>
          
        <TR>
          <TD WIDTH="40%" CLASS=SMALLBOLD><%=Translate("First Name",Login_Language,conn)%> :</TD>
          <TD WIDTH="60%" CLASS=SMALL>
          <INPUT CLASS=SMALL NAME="FirstName" VALUE="<%=request("FIRSTNAME")%>" SIZE="20" MAXLENGTH="50">
          </TD>
        </TR>
      
        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("Last Name",Login_Language,conn)%> :</TD>
          <TD CLASS=SMALL>
          <INPUT CLASS=SMALL NAME="LastName" VALUE="<%=request("LASTNAME")%>" SIZE="20" MAXLENGTH="50">
          </TD>
        </TR>
        
        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("Company",Login_Language,conn)%> :</TD>
          <TD CLASS=SMALL>
          <INPUT CLASS=SMALL NAME="Company" VALUE="<%=request("COMPANY")%>" SIZE="20" MAXLENGTH="100">
          </TD>
        </TR>
        
        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("Email",Login_Language,conn)%> :</TD>
          <TD CLASS=SMALL>
          <INPUT CLASS=SMALL NAME="Email" VALUE="<%=request("EMAIL")%>" SIZE="20" MAXLENGTH="50">
          </TD>
        </TR>
      
        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("City",Login_Language,conn)%> :</TD>
          <TD CLASS=SMALL>
          <INPUT CLASS=SMALL NAME="Business_City" VALUE="<%=request("BUSINESS_CITY")%>" SIZE="20" MAXLENGTH="50">
          </TD>
        </TR>

        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("Country",Login_Language,conn)%> :</TD>
          <TD CLASS=SMALL>
            <%
            Call Connect_FormDatabase
            Call displayCountryList("Business_Country",request("Business_Country"),Translate("Select from list",Login_Language,conn),"Small")
            Call Disconnect_FormDatabase
            %>
          </TD>
        </TR>

        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("State or Province",Login_Language,conn)%> :<BR>
          <SPAN CLASS=Smallest>(<%=Translate("US and Canada Only",Login_Language,conn)%>)</TD>

          <TD CLASS=SMALL>
             <SELECT NAME="Business_State" CLASS=SMALL>
             <%
             .write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"
             SValue = request("Business_State")
             %>  
             <!--#include virtual="/include/core_states.inc"-->
            </SELECT>
          </TD>
        </TR>

        <% if Course_Code <> "TI30" and Course_Code <> "TI20" then %>
        <TR>
          <TD CLASS=SMALLBOLD><%=Translate("Audio",Login_Language,conn)%> :</TD>
          <TD CLASS=SMALL>
          <INPUT CLASS=SMALL NAME="Audio" TYPE="Radio" VALUE="on" <%IF ISBLANK(REQUEST("Audio")) OR REQUEST("Audio") = "on" THEN RESPONSE.WRITE " CHECKED"%>>&nbsp;<%=Translate("On",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;<INPUT CLASS=SMALL NAME="Audio" TYPE="Radio" VALUE="off" <%IF REQUEST("Audio") = "off" THEN RESPONSE.WRITE " CHECKED"%>>&nbsp;<%=Translate("Off",Login_Language,conn)%>
          </TD>
        </TR>
        <% else
          .write "<INPUT TYPE=""HIDDEN"" NAME=""Audio"" VALUE=""off"">"
        end if
        %>  

        <TR>
          <TD CLASS=SMALL BGCOLOR="#777777">&nbsp;</TD>
          <TD CLASS=SMALL BGCOLOR="#777777"><INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="SUBMIT" VALUE="<%=Translate("Continue",Login_Language,conn)%>"></TD>
        </TR>
      </TABLE>
      <%
      Call Table_End
      %>
    </FORM>
  <%

  if Course_Code <> "TI30" and Course_Code <> "TI20" then
    .write "(" & Translate("Note",Login_Language,conn) & ": " & Translate("The audio requires a high speed Internet connection. If you have a slow, dial-up connection, please select the no-audio version.",Login_Language,conn) & ")"
  end if
  
' --------------------------------------------------------------------------------------

  case seqDownload  ' Download Materials
  
    SQL = "SELECT ID, Course_Code, Download_Sequence, Flash_File, Fulfillment, Fulfillment_Countries FROM Sales_Training WHERE Course_Code='" & UCase(request("Course_Code")) & "'"
    Set rsCourse = Server.CreateObject("ADODB.Recordset")
    rsCourse.Open SQL, conn, 3, 3
    
    if not rsCourse.EOF then
      Course                               = CInt(rsCourse("ID"))
      if (CInt(rsCourse("Fulfillment"))    = CInt(True) and instr(1,LCase(rsCourse("Fulfillment_Countries")),LCase(request("Business_Country"))) > 0) _
         or (CInt(rsCourse("Fulfillment")) = CInt(True) and isblank(rsCourse("Fulfillment_Countries"))) then
        Course_Fulfillment = CInt(True)
      else
        Course_Fulfillment = CInt(False)
      end if
      Download_Sequence = UCase(rsCourse("Download_Sequence"))  
    else
      Course = 0
    end if
    
    rsCourse.close
    set rsCourse = nothing
    
    if Course = 0 then

      .write "<SPAN CLASS=Heading3>" & Translate("Course Code Error",Login_Language,conn) & "</SPAN>" & vbCrLf
      .write "<BR><BR>" & vbCrLf
      .write "<SPAN CLASS=Medium>" & vbCrLf
      .write "<P>" & vbCrLf   
      .write "<SPAN CLASS=MediumBoldRed>"
      .write Translate("You have entered an invalid Course Code.",Login_Language,conn) & "&nbsp;&nbsp;" & Translate("Please click on the [Back] button below to select a course from the dropdown list.",Login_Langugae,conn) & vbCrLf
      .write "</SPAN>" & vbCrLf
      .write "<P>" & vbCrLf
      .write "<FORM ACTION=""/sales-training/Default.asp"" METHOD=""POST"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""FirstName"" VALUE=""" & request("FirstName") & """>" & vbCrLf      
      .write "<INPUT TYPE=""HIDDEN"" NAME=""LastName"" VALUE=""" & request("LastName") & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Company"" VALUE=""" & request("Company") & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Business_City"" VALUE=""" & request("Business_City") & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Business_State"" VALUE=""" & request("Business_State") & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Business_Country"" VALUE=""" & request("Business_Country") & """>" & vbCrLf      
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Audio"" VALUE=""" & request("Audio") & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Email"" VALUE=""" & request("Email") & """>" & vbCrLf      
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrLf
      .write "<INPUT CLASS=NavLeftHighlight1 TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""" & Translate("Back",Login_Language,conn) & """>" & vbCrLf
      .write "</FORM>" & vbCrLf

    else
  
      ' Check for exsisting Account
      
      SQL = "SELECT * FROM UserData_Sales_Training WHERE LastName='" & Replace(request("LastName"),"'","''") & "' AND Email='" & request("Email") & "' AND Training_Course=" & Course
      Set rsProfile = Server.CreateObject("ADODB.Recordset")
      rsProfile.Open SQL, conn, 3, 3
      
      if not rsProfile.EOF then
        Account_ID  = rsProfile("ID")
        if isblank(rsProfile("Fulfillment")) then
          Fulfillment = CInt(False)
        else
          Fulfillment = CInt(rsProfile("Fulfillment"))
        end if  
        Modules     = rsProfile("Training_Modules")
        ' Update Old Account with Limited Registration Data
        SQL = "UPDATE UserData_Sales_Training SET " &_
              "FirstName='" & Replace(request("FirstName"),"'","''")  & "', " &_
              "LastName='"  & Replace(request("LastName"),"'","''")   & "', " &_
              "Company='"   & Replace(request("Company"),"'","''")    & "', " &_
              "Business_City='"   & Replace(request("Business_City"),"'","''")    & "', " &_          
              "Business_State='"   & Replace(request("Business_State"),"'","''")    & "', " &_
              "Business_Country='"   & Replace(request("Business_Country"),"'","''")    & "', " &_
              "Language='" & Login_Language & "' " &_
              "WHERE ID=" & Account_ID
      else
  
        ' Create New Account using Add/Update Method to get New_Account_ID
        Account_ID = CInt(Get_New_Record_ID ("UserData_Sales_Training", "Training_Dates", Date(), conn))
      
        ' Update New Account with Registration Data
  
        SQL = "UPDATE UserData_Sales_Training SET " &_
              "FirstName='" & Replace(request("FirstName"),"'","''")  & "', " &_
              "LastName='"  & Replace(request("LastName"),"'","''")   & "', " &_
              "Company='"   & Replace(request("Company"),"'","''")    & "', " &_
              "Email='"   & Replace(request("Email"),"'","''")    & "', " &_          
              "Business_City='"   & Replace(request("Business_City"),"'","''")    & "', " &_          
              "Business_State='"   & Replace(request("Business_State"),"'","''")    & "', " &_
              "Business_Country='"   & Replace(request("Business_Country"),"'","''")    & "', " &_
              "Language='" & Login_Language & "', " &_
              "Training_Course=" & Course & ", " &_            
              "Fulfillment=" & CInt(False) & _
              "WHERE ID=" & Account_ID
        Fulfillment = CInt(False)       ' Set Fulfillment Flag to False for new user
      end if
      
      rsProfile.close
      set rsProfile = nothing  
  
      conn.execute (SQL)
  
      if Download_Sequence = "POST" then
      
        select case Course_Code
        
          case "TI20"
          
            if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then
              select case UCase(request("Business_Country"))
                case "US", "CA", "SG", "MY", "IN", "ID"
                  .write "<SPAN CLASS=Heading3>" & Translate("Fluke Ti20 Thermal Imager Training",Login_Language,conn) & "</SPAN>" & vbCrLf
                  .write "<BR><BR>" & vbCrLf
                  .write "<SPAN CLASS=Medium>" & vbCrLf
                  .write Translate("Thank you for your interest in learning about the Fluke Ti20 Thermal Imager.<P>In appreciation for your time, we'll send you a free Fluke hat. Click the button below to begin the training.<P>Note: Your contact information and interest in thermography will be retained by Fluke.",Login_Language,conn) & "</SPAN><P>" & vbCrLf
                  .write Translate("Thank you for your interest in learning about the Fluke Ti20 Thermal Imager.<P>Click the button below to begin the training.<P>Note: Your contact information and interest in thermography will be retained by Fluke.",Login_Language,conn) & "</SPAN><P>" & vbCrLf                  
                  .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
                  .write "<IMG SRC=""/Sales-Training/Images/FLC_102-120.jpg"">"
                case else
                  .write "<SPAN CLASS=Heading3>" & Translate("Fluke Ti20 Thermal Imager Training",Login_Language,conn) & "</SPAN>" & vbCrLf
                  .write "<BR><BR>" & vbCrLf
                  .write "<SPAN CLASS=Medium>" & vbCrLf
                  .write Translate("Thank you for your interest in learning about the Fluke Ti20 Thermal Imager. Click the button below to begin the training.<P>Note: Your contact information and interest in thermography will be retained by Fluke.",Login_Language,conn) & "<P>"
                  '.write "<B>" & Translate("Fluke Hats are only available for the first 300 FlukePlus members living in the United States, Canada, Singapore, Malaysia, India, and Indonesia only, who complete the training.",Login_Language,conn) & "</B>"
                  .write "</SPAN><P>" & vbCrLf
                  .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
                  .write "<IMG SRC=""/Sales-Training/Images/56042.jpg"">"
              end select
              .write "</TD>" & vbCrLf
              .write "</TR></TABLE>" & vbCrLf

            else

              .write "<SPAN CLASS=Heading3>" & Translate("Fluke Ti20 Sales Training",Login_Language,conn) & "</SPAN>" & vbCrLf
              .write "<BR><BR>" & vbCrLf
              .write "<SPAN CLASS=Medium>" & vbCrLf
              .write Translate("Thank you for returning to Fluke Ti20 Thermal Imager Training. Just click on the button below to begin the training course.",Login_Language,conn) & "<P>" & vbCrLf
              
            end if    

          case "TI30"
          
            if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then
              select case UCase(request("Business_Country"))
                case "US"
                  .write "<SPAN CLASS=Heading3>" & Translate("Fluke Ti30 Distributor Sales Training",Login_Language,conn) & "</SPAN>" & vbCrLf
                  .write "<BR><BR>" & vbCrLf
                  .write "<SPAN CLASS=Medium>" & vbCrLf
                  .write Translate("Thank you for visiting this site.  As a token of our appreciation for completing this training, we will send you a free Fluke hat.  Just click on the button below to begin the training course.",Login_Language,conn) & "</SPAN><P>" & vbCrLf
                  .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
                  .write "<IMG SRC=""/Sales-Training/Images/FLC_102-120.jpg"">"
                case else
                  .write "<SPAN CLASS=Heading3>" & Translate("Fluke Distributor Sales Training - Ti30",Login_Language,conn) & "</SPAN>" & vbCrLf
                  .write "<BR><BR>" & vbCrLf
                  .write "<SPAN CLASS=Medium>" & vbCrLf
                  .write Translate("Thank you for visiting this site.  As a token of our appreciation for completing this training, we will send you a free gift.  Just click on the button below to begin the training course.",Login_Language,conn) & "</SPAN><P>" & vbCrLf
                  .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
                  .write "<IMG SRC=""/Sales-Training/Images/56042.jpg"">"
              end select
              .write "</TD>" & vbCrLf
              .write "</TR></TABLE>" & vbCrLf
              
            else

              .write "<SPAN CLASS=Heading3>" & Translate("Fluke Distributor Sales Training - Ti30",Login_Language,conn) & "</SPAN>" & vbCrLf
              .write "<BR><BR>" & vbCrLf
              .write "<SPAN CLASS=Medium>" & vbCrLf
              .write Translate("Thank you for returning to this site. Just click on the button below to begin the training course.",Login_Language,conn) & "<P>" & vbCrLf
              
            end if    

          case "FOODPRO"
          
            if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then
              select case UCase(request("Business_Country"))
                case "US"
                  .write "<SPAN CLASS=Heading3>" & Translate("Fluke FoodPro Sales Training",Login_Language,conn) & "</SPAN>" & vbCrLf
                  .write "<BR><BR>" & vbCrLf
                  .write "<SPAN CLASS=Medium>" & vbCrLf
                  .write Translate("Thank you for visiting this site.  As a token of our appreciation for completing this training, we will send you a free Fluke hat.  Just click on the button below to begin the training course.",Login_Language,conn) & "</SPAN><P>" & vbCrLf
                  .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
                  .write "<IMG SRC=""/Sales-Training/Images/FLC_102-120.jpg"">"
                case else
                  .write "<SPAN CLASS=Heading3>" & Translate("Fluke FoodPro Sales Training",Login_Language,conn) & "</SPAN>" & vbCrLf
                  .write "<BR><BR>" & vbCrLf
                  .write "<SPAN CLASS=Medium>" & vbCrLf
                  .write Translate("Thank you for visiting this site.  As a token of our appreciation for completing this training, we will send you a free gift.  Just click on the button below to begin the training course.",Login_Language,conn) & "</SPAN><P>" & vbCrLf
                  .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
                  .write "<IMG SRC=""/Sales-Training/Images/56042.jpg"">"
              end select
              .write "</TD>" & vbCrLf
              .write "</TR></TABLE>" & vbCrLf

            else

              .write "<SPAN CLASS=Heading3>" & Translate("Fluke FoodPro Sales Training",Login_Language,conn) & "</SPAN>" & vbCrLf
              .write "<BR><BR>" & vbCrLf
              .write "<SPAN CLASS=Medium>" & vbCrLf
              .write Translate("Thank you for returning to this site. Just click on the button below to begin the training course.",Login_Language,conn) & "<P>" & vbCrLf
              
            end if    
      
          case else
          
            .write "<SPAN CLASS=Heading3>" & Translate("Fluke Distributor Sales Training - Ti30",Login_Language,conn) & "</SPAN>" & vbCrLf
            .write "<BR><BR>" & vbCrLf
            .write "<SPAN CLASS=Medium>" & vbCrLf
            .write Translate("Thank you for returning to this site. Just click on the button below to begin the training course.",Login_Language,conn) & "<P>" & vbCrLf

        end select
        
        Form_Name = "Material_Download"
        %>

        <SCRIPT LANGUAGE=VBScript>
        Private i, x, MM_FlashControlVersion
        On Error Resume Next
        x = null
        MM_FlashControlVersion = 0
        var Flashmode
        FlashMode = False
  
        For i = 9 To 1 Step -1
        
        Set x = CreateObject("ShockwaveFlash.ShockwaveFlash." & i)
        	
        	MM_FlashControlInstalled = IsObject(x)
        	
        	If MM_FlashControlInstalled Then
        		MM_FlashControlVersion = CStr(i)
        		Exit For
        	End If
        Next
        -->
        FlashMode = (MM_FlashControlVersion >= 6)
        If FlashMode = False Then
          document.write "<DIV ALIGN=CENTER><SPAN CLASS=MediumBoldRed><%=Translate("To participate in the training course, your browser must have Macromedia's Flash Version 6 enabled.",Login_Language,conn) & "<BR>" & Translate("Click on the icon below to update your Macromedia Flash Player, before beginning the training course.",Login_Language,conn)%></SPAN></DIV><P>"
        End If
  
        </SCRIPT>
        
        <FORM ACTION="Default.asp" NAME="<%=Form_Name%>" METHOD="POST">
        <INPUT TYPE=HIDDEN NAME="Sequence" VALUE="<%=Sequence + 1%>">
        <INPUT TYPE=HIDDEN NAME="Account_ID" VALUE="<%=Account_ID%>">
        <INPUT TYPE=HIDDEN NAME="Language" VALUE="<%=Login_Language%>">
        <INPUT TYPE=HIDDEN NAME="Audio" VALUE="<%=request("Audio")%>">    
        <INPUT TYPE=HIDDEN NAME="Modules" VALUE="<%=Modules%>">
        <INPUT TYPE=HIDDEN NAME="Fulfillment" VALUE="<%if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then response.write "YES" else response.write "NO"%>">              
        <INPUT TYPE=HIDDEN NAME="Country" VALUE="<%=request("Business_Country")%>">        
        <INPUT TYPE=HIDDEN NAME="Course" VALUE="<%=Course%>">
        &nbsp;&nbsp;&nbsp;&nbsp;
        <INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="SUBMIT" NAME="SUBMIT" VALUE="<%=Translate("Click Here to Begin the Training Course",Login_Language,conn)%>">
        </FORM>            
        <%
        
      else
        
        .write "<SPAN CLASS=Heading3>" & Translate(Caption(Sequence),Login_Language,conn) & "</SPAN>" & vbCrLf
        .write "<BR><BR>" & vbCrLf
        .write "<SPAN CLASS=Medium>" & vbCrLf
     
        if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then
          select case UCase(request("Business_Country"))
            case "US"
              .write Translate("Get a free Fluke hat when you successfully complete this training program.",Login_Language,conn) & "<P>" & vbCrLf
              .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
              .write "<IMG SRC=""/Sales-Training/Images/FLC_102-120.jpg"">"
            case else
              .write Translate("Get a free Fluke gift when you successfully complete this training program.",Login_Language,conn) & "<P>" & vbCrLf
              .write "<TABLE BORDER=""0""><TR><TD>" & vbCrLf
              .write "<IMG SRC=""/Sales-Training/Images/56042.jpg"">"
          end select
          .write "</TD>" & vbCrLf
          .write "<TD VALIGN=""MIDDLE"">" & vbCrLf
        end if
        
        select case UCase(Course_Code)  
          case "GRAYBAR"
            .write Translate("You may have recently received a packet of sales tools from Fluke test instruments in the mail.",Login_Language,conn) & "&nbsp;&nbsp;" & vbCrLf
          case else  
            .write Translate("You may have recently received a packet of sales tools from Fluke and Meterman test instruments in the mail.",Login_Language,conn) & "&nbsp;&nbsp;" & vbCrLf
        end select
            
        .write Translate("You will need to refer to those materials during the course of this training, so please have them handy.",Login_Language,conn)  & "<P>" & vbCrLf
        .write Translate("If you do not have them, please download the files below and print them out <U>before</U> you start the training program.",Login_Language,conn)  & "<P>" & vbCrLf
        if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then
          .write "</TD></TR></TABLE>" & vbCrLf
        end if    
  
        %>
        <SCRIPT LANGUAGE=VBScript>
        Private i, x, MM_FlashControlVersion
        On Error Resume Next
        x = null
        MM_FlashControlVersion = 0
        var Flashmode
        FlashMode = False
  
        For i = 9 To 1 Step -1
        
        Set x = CreateObject("ShockwaveFlash.ShockwaveFlash." & i)
        	
        	MM_FlashControlInstalled = IsObject(x)
        	
        	If MM_FlashControlInstalled Then
        		MM_FlashControlVersion = CStr(i)
        		Exit For
        	End If
        Next
        -->
        FlashMode = (MM_FlashControlVersion >= 6)
        If FlashMode = False Then
          document.write "<DIV ALIGN=CENTER><SPAN CLASS=MediumBoldRed><%=Translate("To participate in the training course, your browser must have Macromedia's Flash Version 6 enabled.",Login_Language,conn) & "<BR>" & Translate("Click on the icon below to update your Macromedia Flash Player, before beginning the training course.",Login_Language,conn)%></SPAN></DIV><P>"
        End If
  
        </SCRIPT>
        
        <%        
        Form_Name = "Material_Download"
        %>
        
        <FORM ACTION="Default.asp" NAME="<%=Form_Name%>" METHOD="POST">
        <INPUT TYPE=HIDDEN NAME="Sequence" VALUE="<%=Sequence + 1%>">
        <INPUT TYPE=HIDDEN NAME="Account_ID" VALUE="<%=Account_ID%>">
        <INPUT TYPE=HIDDEN NAME="Language" VALUE="<%=Login_Language%>">
        <INPUT TYPE=HIDDEN NAME="Audio" VALUE="<%=request("AUDIO")%>">    
        <INPUT TYPE=HIDDEN NAME="Modules" VALUE="<%=Modules%>">
        <INPUT TYPE=HIDDEN NAME="Fulfillment" VALUE="<%if CInt(Fulfillment) = CInt(False) and CInt(Course_Fulfillment) = CInt(True) then response.write "YES" else response.write "NO"%>">              
        <INPUT TYPE=HIDDEN NAME="Country" VALUE="<%=request("BUSINESS_COUNTRY")%>">        
        <INPUT TYPE=HIDDEN NAME="Course" VALUE="<%=Course%>">
        <%
        
        SQL = "SELECT Doc_Name, Doc_Number, Doc_Break FROM Sales_Training WHERE ID=" & Course
        Set rsCourse = Server.CreateObject("ADODB.Recordset")
        rsCourse.Open SQL, conn, 3, 3
  
        if UCase(request("Business_Country")) <> "US" and UCase(Login_Language) = "ENG" then
          if instr(1,LCase(rsCourse("Doc_Number")),"[row]") > 0 then
            Course_Language = "row"
          else
            Course_Language = Login_Language
          end if   
        else
          Course_Language = Login_Language
        end if
  
        Doc_Name   = Split(replace(replace(rsCourse("Doc_Name"),chr(10),""),chr(13),""),"|")
        Doc_Number = replace(replace(rsCourse("Doc_Number"),chr(10),""),chr(13),"")
        Doc_Break  = Split(replace(replace(rsCourse("Doc_Break"),chr(10),""),chr(13),""),"|")
  
        Doc_Number_Group = Split(Doc_Number,"|")
        for x = 0 to UBound(Doc_Number_Group)
          if "[" & Course_Language & "]" = Mid(Doc_Number_Group(x),1,5) then
            Doc_File = Split(Mid(Doc_Number_Group(x),7),",")
            Doc_Count = UBound(Doc_File)
            exit for
          end if
        next
        
        rsCourse.close
        set rsCourse = nothing  
  
        .write "<DIV ALIGN=CENTER CLASS=Medium>" & vbCrLf
    
        Call Table_Begin
        %>
          <TABLE CELLPADDING=4 CELLSPACING=0 CELLPADDING=0 BORDER=0 BGCOLOR="<%=Contrast%>">
            <TR>
              <TD COLSPAN=2 BGCOLOR="#777777" CLASS=SMALLWHITE>
              <%=Translate("Please download and print these files before you proceed with the training course, unless you already have printed copies of them.",Login_Language,conn)%></TD>
              </TD>
            </TR>
              <%
              break = 0
              for x = 0 to Doc_Count
                if Doc_File(x) = "[break]" then
                  .write "<TR>" & vbCrLf
                  .write "<TD COLSPAN=2 BGCOLOR=""#777777"" CLASS=SMALLWHITE>" & vbCrLf
                  .write Translate(Doc_Break(break),Login_Language,conn)
                  .write "</TD>" & vbCrLf
                  .write "</TR>" & vbCrLf
                  break = break + 1
                else  
                  .write "<TR>" & vbCrLf
                  .write  "<TD CLASS=Small WIDTH=""2%"" ALIGN=CENTER>" & vbCrLf
                  .write  "<A HREF=""javascript:void(0);"" Language=""javascript"" TITLE=""Click to Download File"" onclick=""openit_mini('/Find_It.asp?Document=" & Doc_File(x) & "','Vertical');return false;"">"
                  .write  "<IMG SRC=""/images/Button-PDF.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
                  .write  "</A>"
                  .write  "</TD>" & vbCrLf
                  .write  "<TD Class=SmallBold>" & vbCrLf
                  .write  Translate(Doc_Name(x),Login_Language,conn)
                  .write  "</TD>" & vbCrLf
                  .write  "</TR>" & vbCrLf
                end if  
              next
              %>
              <TR>
                <TD COLSPAN=2 ALIGN=CENTER BGCOLOR="#777777">
                <%
                .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.adobe.com/support/downloads/main.html','Vertical');return false;""><IMG SRC=""images/acrobat.gif"" BORDER=0 ALT=""Click to Download Acrobat Reader"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
                .write "&nbsp;&nbsp;&nbsp;"
    
                select case Login_Language
                  case "eng"
                    .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&P5_Language=English','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
                  case "fre"
                    .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&Lang=French&P5_Language=French','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
                  case "por"
                    .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&Lang=BrazilianPortuguese&P5_Language=BrazilianPortuguese','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
                  case "spa"
                    .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&Lang=Spanish&P5_Language=Spanish','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
                  end select                
                %>
                &nbsp;&nbsp;&nbsp;&nbsp;
                <INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="SUBMIT" NAME="SUBMIT" VALUE="<%=Translate("Click Here to Begin the Training Course",Login_Language,conn)%>">
              </TR>
          </TABLE>
          <%
          Call Table_End
          
          .write "</DIV>" & vbCrLf
          %>
        </FORM>

        <%
        end if
        
      end if

  ' --------------------------------------------------------------------------------------

  case seqFlash   ' Embedded Flash Modules
  
    SQL = "SELECT Flash_File FROM Sales_Training WHERE ID=" & course
    Set rsCourse = Server.CreateObject("ADODB.Recordset")
    rsCourse.Open SQL, conn, 3, 3
      
    Flash_SRC = "Course_" & course & "/" & rsCourse("Flash_File")
      

    select case course
    
      case 10

        %>
        <script type="text/javascript" src="flashobject.js"></script>
      
        <div id="flashcontent" style="position:relative; top:-18px"></div>
      
    	  <script type="text/javascript">
    		  var fo = new FlashObject("<%=Flash_SRC%>", "FoodPro", "775", "510", "7.0.20", "#ffffff", true);
        	fo.addVariable("BackURL", "http://<%=request.ServerVariables("Server_Name")%>/sales-training/default.asp");
        	fo.addVariable("Language", "<%=request("Language")%>");
    		  fo.addVariable("Account_ID", "<%=request("ACCOUNT_ID")%>");
      		fo.addVariable("Country", "<%=request("Country")%>");
      		fo.addVariable("Audio", "<%=request("Audio")%>");
      		fo.addVariable("Modules", "1");
      		fo.addVariable("Sequence", "<%=request("Sequence")+1%>");
      		fo.addVariable("Course", "<%=Course%>");
       		fo.addVariable("Course_Code", "FOODPRO");
     	  	fo.addVariable("Score", "100");
        	fo.addVariable("Fulfillment", "YES");
       		fo.addParam("base", "/sales-training/<%="Course_" & course%>");        
      		fo.write("flashcontent");
  		    // ]]>
      	</script>
        <%      
    
      case 11
        %>
        <SCRIPT LANGUAGE="JavaScript">
        window.location.href="<%=Flash_SRC & "?BackURL=http://" & request.ServerVariables("Server_Name") & "/sales-training/default.asp&Course_Code=FOODPRO&Course=10&Account_ID=" & Account_ID & "&Sequence=" & Sequence + 1 & "&Language=" & Login_Language & "&Modules=1" & "&Fulfillment=-1" & "&Audio=" & request("Audio") & "&Country=" & request("Country") & "&Score=100"%>";
        </SCRIPT>
        <%

      case 49
        %>
        <SCRIPT LANGUAGE="JavaScript">
        window.location.href="<%=Flash_SRC & "?BackURL=http://" & request.ServerVariables("Server_Name") & "/sales-training/default.asp&Course_Code=TI20&Course=49&Account_ID=" & Account_ID & "&Sequence=" & Sequence + 1 & "&Language=" & Login_Language & "&Modules=1&Score=100"%>";
        </SCRIPT>
        <%
      case 50
        %>
        <SCRIPT LANGUAGE="JavaScript">
        window.location.href="<%=Flash_SRC%>";
        </SCRIPT>
        <%
      case else
      
        .write "<SPAN CLASS=Medium>"
        .write "<DIV ALIGN=""center"">" & vbCrLf
  
        Call Table_Begin
        
        Flash_Language = request("Language")

        'response.write"<PARAM NAME=FLASHVARS VALUE=""Account_ID=" & request("ACCOUNT_ID") & "&Language=" & Flash_Language & "&Country=" & request("Country") & "&Audio=" & request("Audio") & "&Modules=" & request("Modules") & "&Sequence=" & request("Sequence") & "&Course=" & request("Course") & "&Fulfillment=" & request("Fulfillment") & ">"        
        'response.end
        %>
        
        <OBJECT  CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
                 CODEBASE="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"
                 WIDTH="750"
                 HEIGHT="450"
                 ID="flashTest"
                 ALIGN="">
          <PARAM NAME=MOVIE VALUE="<%=Flash_SRC%>">
          <PARAM NAME=MENU VALUE=FALSE>
          <PARAM NAME=QUALITY VALUE=HIGH>
          <PARAM NAME=SCALE VALUE=NOSCALE>
          <PARAM NAME=BGCOLOR VALUE=#FFFFFF>
          <PARAM NAME=FLASHVARS VALUE="Account_ID=<%=request("ACCOUNT_ID")%>&Language=<%=Flash_Language%>&Country=<%=request("Country")%>&Audio=<%=request("Audio")%>&Modules=<%=request("Modules")%>&Sequence=<%=request("Sequence")%>&Course=<%=request("Course")%>&Fulfillment=0">
        
          <EMBED  SRC="<%=Flash_SRC%>"
                  menu=false
                  quality=high
                  scale=noscale
                  bgcolor=#FFFFFF
                  WIDTH="750"
                  HEIGHT="450"
                  NAME="flashTest"
                  ALIGN=""
                  TYPE="application/x-shockwave-flash"
                  PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
                  FlashVars="Account_ID=<%=request("Account_ID")%>&Language=<%=request("Language")%>&Country=<%=request("Country")%>&Audio=<%=request("Audio")%>&Modules=<%=request("Modules")%>&Sequence=<%=request("Sequence")%>&Course=<%=request("Course")%>&Fulfillment=0"
                  >
          </EMBED>
        </OBJECT-->
        <%
    
        Call Table_End
        .write "</DIV>" & vbCrLf
        
    end select  

  ' --------------------------------------------------------------------------------------

  case seqShipping

    SQL = "SELECT ID, Fulfillment, Fulfillment_Countries FROM Sales_Training WHERE ID=" & Course
    Set rsCourse = Server.CreateObject("ADODB.Recordset")
    rsCourse.Open SQL, conn, 3, 3
    
    if (CInt(rsCourse("Fulfillment"))    = CInt(True) and instr(1,LCase(rsCourse("Fulfillment_Countries")),LCase(request("Business_Country"))) > 0) _
        or (CInt(rsCourse("Fulfillment")) = CInt(True) and isblank(rsCourse("Fulfillment_Countries"))) then
      Course_Fulfillment = CInt(True)
    else
      Course_Fulfillment = CInt(False)
    end if  
    
    rsCourse.close
    set rsCourse = nothing

    if not isblank(session("Account_ID_Course")) then
      Account_ID = session("Account_ID_Course")
    else
      Account_ID = request("Account_ID")
    end if
    
    SQL = "SELECT * FROM UserData_Sales_Training WHERE ID=" & Account_ID & " AND FulFillment=" & CInt(False) & " AND Training_Course=" & Course
    Set rsProfile = Server.CreateObject("ADODB.Recordset")
    rsProfile.Open SQL, conn, 3, 3

    .write "<SPAN CLASS=Heading3>"
    if not rsProfile.EOF and CInt(Course_Fulfillment) = CInt(True) then
      select case UCase(Course_Code)
        case "TI20"
          select case UCase(rsProfile("Business_Country"))
            case "US", "CA", "SG", "MY", "IN", "ID"
              .write Translate(Caption(Sequence),Login_Language,conn)            
            case else
              .write Translate("Contact Information",Login_Language,conn)
          end select            
        case else
          .write Translate(Caption(Sequence),Login_Language,conn)
      end select    

    else
      .write Translate("Thank You!",Login_Language,conn)
    end if
    .write "</SPAN>" & vbCrLf        
    
    .write "<BR><BR>" & vbCrLf
    
    .write "<SPAN CLASS=Medium>" & vbCrLf

    select case UCase(Course_Code)
      case "GRAYBAR"
        .write Translate("Thank you for taking the time to participate in this training program to learn to sell more Fluke test instruments.",Login_Language,conn) & "<P>" & vbCrLf
        .write Translate("Your overall score on the assessment questions is:",Login_Language,conn) & "&nbsp;&nbsp;<B>" & request("Score") & " %</B>" & "<P>" & vbCrLf        
      case "TI20"
        .write Translate("Thank you for watching the Ti20 Thermal Imager training.",Login_Language,conn) & "<P>" & vbCrLf
      case "TI30"
        .write Translate("Thank you for taking the time to participate in this training program to learn to sell Fluke Thermal Imagers.",Login_Language,conn) & "<P>" & vbCrLf      
      case "FOODPRO"
        .write Translate("Thank you for taking the time to participate in this training program to learn to sell Fluke FoodPro thermometers.",Login_Language,conn) & "<P>" & vbCrLf
      case else
        .write Translate("Thank you for taking the time to participate in this training program to learn to sell more Fluke and Meterman test instruments.",Login_Language,conn) & "<P>" & vbCrLf
        .write Translate("Your overall score on the assessment questions is:",Login_Language,conn) & "&nbsp;&nbsp;<B>" & request("Score") & " %</B>" & "<P>" & vbCrLf
    end select

    if not rsProfile.EOF and CInt(Course_Fulfillment) = CInt(True) then

      select case UCase(Course_Code)
      
        case "TI20"
          select case UCase(rsProfile("Business_Country"))
            case "US", "CA", "SG", "MY", "IN", "ID"
              '.write "<TABLE BORDER=""0"">" & vbCrLf
              '.write "<TR>" & vbCrLf
              '.write "<TD>" & vbCrLf
              '.write "<IMG SRC=""/Sales-Training/Images/FLC_102-120.jpg"">"
              '.write "</TD>" & vbCrLf
              '.write "<TD VALIGN=""MIDDLE"">" & vbCrLf     
              '.write "<P>" & Translate("In appreciation, we'll send a Fluke hat to you within 3 weeks.",Login_Language,conn)
              '.write "</TD>" & vbCrLf
              '.write "</TR>" & vbCrLf
              '.write "</TABLE>" & vbCrLf
              '.write Translate("Please confirm your contact information and the address to which you would like the hat to be shipped.",Login_Language,conn) & "<BR>" & vbCrLf
              '.write "(" & Translate("Hats are shipped via UPS. No P.O. Boxes, please.",Login_Language,conn)  & ")" & vbCrLf
              
              .write "If you'd like to learn more about Fluke thermal imagers, please complete the form below to request product information, or visit <A HREF=""http://www.fluke.com/thermography"">www.fluke.com/thermography</A>."

 


            case else
              .write Translate("Please confirm your contact and address information.",Login_Language,conn) & vbCrLf              
          end select  
      
        case else     

          select case UCase(rsProfile("Business_Country"))
            case "US"
              .write "<TABLE BORDER=""0"">" & vbCrLf
              .write "<TR>" & vbCrLf
              .write "<TD>" & vbCrLf
              .write "<IMG SRC=""/Sales-Training/Images/FLC_102-120.jpg"">"
              .write "</TD>" & vbCrLf
              .write "<TD VALIGN=""MIDDLE"">" & vbCrLf     
              .write Translate("As a token of our appreciation, Fluke will send your Fluke hat to you within 3 weeks.",Login_Language,conn)
              .write "</TD>" & vbCrLf
              .write "</TR>" & vbCrLf
              .write "</TABLE>" & vbCrLf
              .write Translate("Please confirm your contact information and the address to which you would like the hat to be shipped.",Login_Language,conn) & "<BR>" & vbCrLf
              .write "(" & Translate("Hats are shipped via UPS. No P.O. Boxes, please.",Login_Language,conn)  & ")" & vbCrLf
            case else
              .write "<TABLE BORDER=""0"">" & vbCrLf
              .write "<TR>" & vbCrLf
              .write "<TD>" & vbCrLf
              .write "<IMG SRC=""/Sales-Training/Images/56042.jpg"">"
              .write "</TD>" & vbCrLf
              .write "<TD VALIGN=""MIDDLE"">" & vbCrLf     
              .write Translate("As a token of our appreciation, Fluke will send your Fluke gift to you within 3 weeks.",Login_Language,conn)
              .write "</TD>" & vbCrLf
              .write "</TR>" & vbCrLf
              .write "</TABLE>" & vbCrLf
              .write Translate("Please confirm your contact information and the address to which you would like the gift to be shipped.",Login_Language,conn) & "<BR>" & vbCrLf
              .write "(" & Translate("Gifts are shipped via UPS. No P.O. Boxes, please.",Login_Language,conn)  & ")" & vbCrLf
          end select

      end select    
      
      Form_Name = "Shipping"

      SQLInfo = "SELECT More_Info FROM Sales_Training WHERE Course_Code='" & Course_Code & "'"
      Set rsInfo = Server.CreateObject("ADODB.Recordset")
      rsInfo.Open SQLInfo, conn, 3, 3
      
      More_Info = false
      if not rsInfo.EOF then
        More_Info = CInt(rsInfo("More_Info"))
      end if
      rsInfo.close
      set rsInfo = nothing

      %>
      <FORM ACTION="Default.asp" NAME="<%=Form_Name%>" METHOD="POST" onsubmit="return(CheckRequiredFields(this.form));" onKeyUp="highlight(event)" onClick="highlight(event)">
      <INPUT TYPE=HIDDEN NAME="Sequence" VALUE="<%=Sequence + 1%>">
      <INPUT TYPE=HIDDEN NAME="Account_ID" VALUE="<%=Account_ID%>">
      <INPUT TYPE=HIDDEN NAME="Score" VALUE="<%=request("Score")%>">
      <INPUT TYPE=HIDDEN NAME="Modules" VALUE="<%=request("Modules")%>">      
      <INPUT TYPE=HIDDEN NAME="Language" VALUE="<%=request("Language")%>">                
      <INPUT TYPE=HIDDEN NAME="Course" VALUE="<%=request("Course")%>">    
      <%
      Call Table_Begin
      %>
            <TABLE CELLPADDING=4 CELLSPACING=0 BORDER=0 BGCOLOR="<%=Contrast%>" WIDTH="100%">
            <%
            if More_Info = CInt(true) then
              %>
              <TR>
                <TD WIDTH="40%" CLASS=SMALLBOLD><%=Translate("I would like more information about this product",Login_Language,conn)%> :</TD>
                <TD WIDTH="60%" CLASS=SMALL><INPUT TYPE=CHECKBOX CLASS=SMALL NAME="More_Info">  <%=Translate("Yes",Login_Language,conn)%>
                </TD>
              </TR>
              <%
            else
              %>
                <INPUT TYPE=HIDDEN NAME="More_Info" VALUE="0">
              <%  
            end if
              %>
              <TR>
                <TD WIDTH="40%" CLASS=SMALLBOLD><%=Translate("First Name",Login_Language,conn)%> :</TD>
                <TD WIDTH="60%" CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="FirstName" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("FIRSTNAME")%>">
                </TD>
              </TR>
            
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Last Name",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="LastName" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("LASTNAME")%>">
                </TD>
              </TR>
              
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Email",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Email" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("EMAIL")%>">
                </TD>
              </TR>

              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Phone",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Business_Phone" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("BUSINESS_PHONE")%>">
                </TD>
              </TR>

              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Company",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Company" SIZE="20" MAXLENGTH="100" VALUE="<%=rsProfile("COMPANY")%>">
                </TD>
              </TR>
                      
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Address",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Business_Address" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("BUSINESS_ADDRESS")%>">
                </TD>
              </TR>
  
              <TR>
                <TD CLASS=SMALLBOLD>&nbsp;</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Business_Address_2" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("BUSINESS_ADDRESS_2")%>">
                </TD>
              </TR>
  
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("City",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Business_City" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("BUSINESS_CITY")%>">
                </TD>
              </TR>
      
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Postal Code",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <INPUT CLASS=SMALL NAME="Business_Postal_Code" SIZE="20" MAXLENGTH="50" VALUE="<%=rsProfile("BUSINESS_POSTAL_CODE")%>">
                </TD>
              </TR>
  
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Country",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                  <%
                  Call Connect_FormDatabase
                  Call DisplayCountryList("Business_Country",rsProfile("Business_Country"),Translate("Select from list",Login_Language,conn),"Small")
                  Call Disconnect_FormDatabase
                  %>
                </TD>
              </TR>
              
              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("State or Province",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                   <SELECT NAME="Business_State" CLASS=SMALL>
                   <%
                   .write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
                   sValue = rsProfile("Business_State")
                   %>  
                   <!--#include virtual="/include/core_states.inc"-->
                  </SELECT>
                </TD>
              </TR>

              <TR><TD COLSPAN=2 BGCOLOR="#777777" CLASS=SMALLWHITE><%=Translate("Help us improve.",Login_Language,conn) & "&nbsp;&nbsp;" & Translate("We would appreciate your comments.",Login_language,conn)%></TR>              

              <TR>
                <TD CLASS=SMALLBOLD><%=Translate("Comments",Login_Language,conn)%> :</TD>
                <TD CLASS=SMALL>
                <TEXTAREA NAME="Comment" COLS=50 ROWS=6 MAXLENGTH="500" CLASS=Small></TEXTAREA>
                </TD>
              </TR>
              
              <TR>
                <TD CLASS=SMALL BGCOLOR="#777777">&nbsp;</TD>
                <TD CLASS=SMALL BGCOLOR="#777777"><INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="SUBMIT" VALUE="<%=Translate("Continue",Login_Language,conn)%>"></TD>
              </TR>
            </TABLE>
            <%
            Call Table_End
            %>
        </FORM>
        <%
        rsProfile.close
        set rsProfile = nothing
       
      else
        rsProfile.close
        set rsProfile = nothing
        
        %>      
        <FORM ACTION="Default.asp" NAME="<%=Form_Name%>" METHOD="POST">        
        <INPUT TYPE=HIDDEN NAME="Sequence" VALUE="<%=Sequence + 2%>">
        <INPUT TYPE=HIDDEN NAME="Account_ID" VALUE="<%=Account_ID%>">
        <INPUT TYPE=HIDDEN NAME="Score" VALUE="<%=request("Score")%>">
        <INPUT TYPE=HIDDEN NAME="Modules" VALUE="<%=request("Modules")%>">      
        <INPUT TYPE=HIDDEN NAME="Language" VALUE="<%=request("Language")%>">                
        <INPUT TYPE=HIDDEN NAME="Course" VALUE="<%=request("Course")%>">    
        <INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="SUBMIT" VALUE="<%=Translate("Continue",Login_Language,conn)%>">
        </FORM>
        <%

      end if
      
  ' --------------------------------------------------------------------------------------

   case seqFulfillment, seqDownload_2, seqPortalReg, seqPortalLogon   ' Hat/Gift Fulfillment, Register for Partner Portal Access
   
    Account_ID = request("Account_ID")

    if Sequence = seqFulfillment then
    
      ' Update New Account with Registration Data

      SQL = "UPDATE UserData_Sales_Training SET " &_
            "FirstName='" & Replace(request("FirstName"),"'","''")  & "', " &_
            "LastName='"  & Replace(request("LastName"),"'","''")   & "', " &_
            "Company='"   & Replace(request("Company"),"'","''")    & "', " &_
            "Email='"   & Replace(request("Email"),"'","''")    & "', " &_          
            "Business_Phone='"   & Replace(request("Business_Phone"),"'","''")   & "', " &_          
            "Business_Address='"   & Replace(request("Business_Address"),"'","''")    & "', " &_          
            "Business_Address_2='"   & Replace(request("Business_Address_2"),"'","''")    & "', " &_          
            "Business_City='"   & Replace(request("Business_City"),"'","''")    & "', " &_          
            "Business_State='"   & Replace(request("Business_State"),"'","''")    & "', " &_
            "Business_Postal_Code='"   & Replace(request("Business_Postal_Code"),"'","''")    & "', " &_
            "Business_Country='"   & Replace(request("Business_Country"),"'","''")    & "', " & _
            "Training_Dates='" & Date() & "', " &_
            "Training_Modules='" & request("Modules") & "', " &_
            "More_Info="
            if request("More_Info") = "on" then
              SQL = SQL & "-1 "
            else
              SQL = SQL & "0 "
            end if                
            SQL = SQL & "WHERE ID=" & Account_ID & " AND Training_Course=" & request("Course")

      conn.execute (SQL)

      ' Send Fulfillment Email
        
      SQL = "SELECT Fulfillment FROM UserData_Sales_Training WHERE ID=" & Account_ID
      Set rsSite = Server.CreateObject("ADODB.Recordset")
      rsSite.Open SQL, conn, 3, 3
      Fulfillment = CInt(rsSite("Fulfillment"))
      rsSite.close
      set rsSite = nothing
      
      if CInt(Fulfillment) = CInt(False) then
      
        SQL = "SELECT FromName,FromAddress,ReplyToName,ReplyTo FROM Sales_Training Where ID=" & Course
        Set rsSite = Server.CreateObject("ADODB.Recordset")
        rsSite.Open SQL, conn, 3, 3
  
        MailFromName    = rsSite("FromName")
        MailFromAddress = rsSite("FromAddress")
        MailReplyToName = rsSite("ReplyToName")
        MailReplyTo     = rsSite("ReplyTo")
        'MailBCCName     = "Whitlock Kelly"
        MailBCCName     = "Extranet Group"
        'MailBCC         = "Kelly.Whitlock@Fluke.com"
        MailBCC         = "extranetalerts@Fluke.com"
        
        rsSite.close
        set rsSite = nothing
        
        'Mailer.ReturnReceipt = False  
        'Mailer.Priority      = 3
        'Mailer.QMessage      = False
    
        ' Use Primary Site Administrator's Info
        'Mailer.FromName      = MailFromName
        'Mailer.FromAddress   = MailFromAddress 
  
        msg.From = """" & MailFromName & """" & MailFromAddress

        'Mailer.ReplyTo       = MailReplyTo
        msg.ReplyTo = MailReplyTo
  
        if not isblank(MailBCC) then
          'Mailer.AddBCC        MailBCCName, MailBCC
          msg.Bcc = """" & MailBCCName & """" & MailBCC
        end if
  
        'Mailer.Subject         = "Fluke Fulfillment Request"
        msg.Subject = "Fluke Fulfillment Request"
        
        MailMessage = "This is an automated notification message from the " & MailFromName & " " & "Extranet Support Server" & "." & vbCrLf & vbCrLf
  
        MailMessage = MailMessage & "Please send (1) Fluke gift to our Sales Associate listed below for the succesfull completion of course number " & Course & " :" & vbCrLf & vbCrLf
  
        MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf
        MailMessage = MailMessage & "Sales Associate Shipping Information" & vbCrLf
        MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Name: " & request("FirstName") & " " & request("LastName") & vbCrLf
        MailMessage = MailMessage & "Company: " & request("Company") & vbCrLf
        MailMessage = MailMessage & "Phone: " & request("Business_Phone") & vbCrLf
        MailMessage = MailMessage & "Email: " & request("Email") & vbCrLf            
        MailMessage = MailMessage & "Address 1: " & request("Business_Address") & vbCrLf
        if not isblank(request("Business_Address_2")) then
          MailMessage = MailMessage & "Address 2: " & request("Business_Address_2") & vbCrLf
        end if  
        MailMessage = MailMessage & "City: " & request("Business_City") & vbCrLf
        if not isblank(request("Business_State")) and request("Business_State") <> "ZZ" then
          MailMessage = MailMessage & "State / Province: " & request("Business_State") & vbCrLf
        end if  
        MailMessage = MailMessage & "Postal Code: " & request("Business_Postal_Code") & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Country: "
        
        SQL = "SELECT Country.Name AS Name, Country_Sales_Training_Fulfillment.Email AS Email, Country_Sales_Training_Fulfillment.Email_Comments AS Email_Comments " &_
              "FROM Country LEFT OUTER JOIN " &_
              "Country_Sales_Training_Fulfillment ON Country.Abbrev = Country_Sales_Training_Fulfillment.Abbrev " &_
              "WHERE Country.Abbrev='" & request("Business_Country") & "'"
              
        Set rsCountry = Server.CreateObject("ADODB.Recordset")
        rsCountry.Open SQL, conn, 3, 3
        MailMessage = MailMessage & rsCountry("Name")
  
        if CInt(Debug_Flag) = CInt(False) then         
          if not isblank(rsCountry("Email")) then
            'Mailer.AddRecipient    "Fulfillment Administrator", rsCountry("Email")
            msg.To = msg.To & ";" & """Fulfillment Administrator""" & rsCountry("Email")
          else
            'Mailer.AddRecipient    "Fulfillment Administrator", "FK01@csepromo.com"
            msg.To = msg.To & ";" & """Fulfillment Administrator""" & "FK01@csepromo.com"
          end if  
        end if  
        
        Email_Comments = rsCountry("Email_Comments")
        
        rsCountry.close
        set rsCountry = nothing
        
        MailMessage = MailMessage & vbCrLf & vbCrLf
        MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf & vbCrLf
        MailMessage = MailMessage & "Sincerely," & vbCrLf
        MailMessage = MailMessage & "The " & MailFromName & " Support Team" & vbCrLf
        
        'Mailer.BodyText = MailMessage
        msg.TextBody = MailMessage

        strError = ""
        'Mailer.SendMail
        msg.Configuration = conf
        On Error Resume Next
        msg.Send
        If Err.Number = 0 then
          'Success
        Else
          ' Email Unsuccessful, Send Advisory to Webmaster
          strError = Err.Description         
          msg.From = """Support Server""" & "Webmaster@Fluke.com"
          msg.To = """Support Webmaster""" & "Webmaster@Fluke.com"
          msg.Subject = "Send Email Failure"
          msg.TextBody = strError
          msg.send

          if Debug_Flag and not isblank(strError) then
            ErrorMessage = ErrorMessage & vbCrLf & "<LI>" & Translate("Send email failure",Login_Language,conn) & ".<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & strError & ".</LI>"
            .write ErrorMessage
          end if
        End If
        
        ' Fulfillment Email Error Handler
        
        'strError = ""
        'if Mailer.Response <> "" then
    '
        '  ' Email Unsuccessful, Send Advisory to Webmaster
        '  strError = Mailer.Response
        '  Mailer.ClearAllRecipients
        '  Mailer.FromName    = "Support Server"
        '  Mailer.FromAddress = "Webmaster@Fluke.com"
        '  Mailer.AddRecipient  "Support Webmaster", "Webmaster@Fluke.com"
        '  Mailer.Subject     = "Send Email Failure"
        '  Mailer.BodyText    = strError
        '  Mailer.SendMail
        '  if Debug_Flag and not isblank(strError) then
        '    ErrorMessage = ErrorMessage & vbCrLf & "<LI>" & Translate("Send email failure",Login_Language,conn) & ".<BR><BR>" & Translate("Error Description",Login_Language,conn) & ": " & strError & ".</LI>"
        '    .write ErrorMessage
        '  end if
        'end if     

        ' Email Successful, Set Fulfilment Flag
        SQL = "UPDATE UserData_Sales_Training SET Fulfillment=" & CInt(True) & " WHERE ID=" & Account_ID
        conn.execute (SQL)
        
        ' User Comments
        
        if not isblank(request("Comment")) then
        
          'Mailer.ClearRecipients
          'Mailer.ClearBodyText
          if UCase(request("Business_Country")) = "US" then
            'Mailer.AddRecipient  MailReplyToName, MailReplyTo
            msg.To = """" & MailReplyToName & """" & MailReplyTo
          else  
            'Mailer.AddRecipient  "Dana Banning", "Dana.Banning@Fluke.com"
            msg.To = """" & "Dana Banning" & """" & "Dana.Banning@Fluke.com"
          end if
          if not isblank(Email_Comments) then
            'Mailer.AddBCC        Email_Comments, Email_Comments
            msg.Bcc = Email_Comments
          end if            
          if not isblank(MailBCC) then
            'Mailer.AddBCC        MailBCCName, MailBCC
            msg.Bcc = """" & MailBCCName & """" & MailBCC
          end if

          'Mailer.Subject = "Extranet Training - Feedback"
          msg.Subject = "Extranet Training - Feedback"

          MailMessage = "This is an automated notification message from the " & MailFromName & " " & "Extranet Support Server" & "." & vbCrLf & vbCrLf
          
          MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf
          MailMessage = MailMessage & "Comment" & vbCrLf
          MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf & vbCrLf

          MailMessage = MailMessage & request("Comment") & vbCrLf & vbCrLf
          
          MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf
          MailMessage = MailMessage & "Sales Associate Information" & vbCrLf
          MailMessage = MailMessage & "--------------------------------------------"   & vbCrLf & vbCrLf
          MailMessage = MailMessage & "Name: " & request("FirstName") & " " & request("LastName") & vbCrLf
          MailMessage = MailMessage & "Company: " & request("Company") & vbCrLf
          MailMessage = MailMessage & "Phone: " & request("Business_Phone") & vbCrLf
          MailMessage = MailMessage & "Email: " & request("Email") & vbCrLf
          MailMessage = MailMessage & "Language: " & request("Language") & vbCrLf & vbCrLf          
          MailMessage = MailMessage & "Address 1: " & request("Business_Address") & vbCrLf
          if not isblank(request("Business_Address_2")) then
            MailMessage = MailMessage & "Address 2: " & request("Business_Address_2") & vbCrLf
          end if  
          MailMessage = MailMessage & "City: " & request("Business_City") & vbCrLf
          if not isblank(request("Business_State")) and request("Business_State") <> "ZZ" then
            MailMessage = MailMessage & "State / Province: " & request("Business_State") & vbCrLf
          end if  
          MailMessage = MailMessage & "Postal Code: " & request("Business_Postal_Code") & vbCrLf
          MailMessage = MailMessage & "Country: "
          
          SQL = "SELECT Country.Name AS Name, Country_Sales_Training_Fulfillment.Email AS Email " &_
                "FROM Country LEFT OUTER JOIN " &_
                "Country_Sales_Training_Fulfillment ON Country.Abbrev = Country_Sales_Training_Fulfillment.Abbrev " &_
                "WHERE Country.Abbrev='" & request("Business_Country") & "'"
                
          Set rsCountry = Server.CreateObject("ADODB.Recordset")
          rsCountry.Open SQL, conn, 3, 3
          MailMessage = MailMessage & rsCountry("Name")
          rsCountry.close
          set rsCountry = nothing
         
          'Mailer.BodyText = MailMessage
          msg.TextBody = MailMessage

          'Mailer.SendMail
          msg.Configuration = conf
          On Error Resume Next
          msg.Send
          If Err.Number = 0 then
            'Success
          Else
            'Fail
          End If

        end if
        
        ' Visitor Confirmation and Invite
        
        'Mailer.ClearRecipients
        'Mailer.ClearBodyText
        'Mailer.AddRecipient  request("FirstName") & " " & request("LastName"), request("Email")
        'Mailer.Subject = Translate("Fluke Acknowledgement",Login_Language,conn)

        msg.To = """" & request("FirstName") & " " & request("LastName") & """" & request("Email")
        msg.Subject = Translate("Fluke Acknowledgement",Login_Language,conn)
        
        SQL = "SELECT ID, Course_Title, Course_Description FROM Sales_Training WHERE ID=" & request("Course")
        Set rsCourse = Server.CreateObject("ADODB.Recordset")
        rsCourse.Open SQL, conn, 3, 3

        MailMessage = Translate("Dear Fluke Reseller",Login_Language,conn) & "," & vbCrLf & vbCrLf
        MailMessage = MailMessage & Translate("Thank you for participating in Fluke�s Distributor Sales Training Program.",Login_Language,conn) & "  "
        MailMessage = MailMessage & Translate("This email confirms that you have passed the training course and you will be receiving your Fluke gift within 3-5 weeks.",Login_Language,conn) & vbCrLf & vbCrLf
        MailMessage = MailMessage & Translate("Did you know that other sales associates in your organization can participate in Fluke's Distributor Sales Training Program and if they successfully complete the course can receive a Fluke gift too?",Login_Language,conn) & "  "
        MailMessage = MailMessage & Translate("Just pass along this invitation to your fellow associates.",Login_Language,conn) & vbCrLf & vbCrLf
        MailMessage = MailMessage & Translate("Fluke On-Line Distributor Sales Training",Login_Language,conn) & vbCrLf
        MailMessage = MailMessage & "http://Support.Fluke.com/Sales-Training" & vbCrLf & vbCrLf
        MailMessage = MailMessage & Translate("Then select course",Login_Language,conn) & ": "

        MailMessage = MailMessage & Translate(rsCourse("Course_Title"),Login_Language,conn) & vbCrLf & vbCrLf
        MailMessage = MailMessage & Translate("Sincerely",Login_Language,conn) & "," & vbCrLf & vbCrLf & vbCrLf
        MailMessage = MailMessage & Translate(rsCourse("Course_Description"),Login_Language,conn) & " - " & Translate("Support Team",Login_Language,conn) & vbCrLf

        rsCourse.close
        set rsCourse = nothing
        
        'Mailer.BodyText = MailMessage
        msg.TextBody = MailMessage

        if UCase(request("Course")) <> "TI20" then
          'Mailer.SendMail
          msg.send
        end if  
        
      end if
      
      Sequence = Sequence + 1
    
    end if

    ' Download_2 Post
    
    SQL = "SELECT Download_Sequence, Doc_Name, Doc_Number, Doc_Break FROM Sales_Training WHERE ID=" & Course
    Set rsCourse = Server.CreateObject("ADODB.Recordset")
    rsCourse.Open SQL, conn, 3, 3
        
    if rsCourse("Download_Sequence") = "POST" then
  
      if UCase(request("Business_Country")) <> "US" and UCase(Login_Language) = "ENG" then
        if instr(1,LCase(rsCourse("Doc_Number")),"[row]") > 0 then
          Course_Language = "row"
        else
          Course_Language = Login_Language
        end if   
      else
        Course_Language = Login_Language
      end if
  
      Doc_Name   = Split(replace(replace(rsCourse("Doc_Name"),chr(10),""),chr(13),""),"|")
      Doc_Number = replace(replace(rsCourse("Doc_Number"),chr(10),""),chr(13),"")
      Doc_Break  = Split(replace(replace(rsCourse("Doc_Break"),chr(10),""),chr(13),""),"|")
  
      Doc_Number_Group = Split(Doc_Number,"|")
      for x = 0 to UBound(Doc_Number_Group)
        if "[" & Course_Language & "]" = Mid(Doc_Number_Group(x),1,5) then
          Doc_File = Split(Mid(Doc_Number_Group(x),7),",")
          Doc_Count = UBound(Doc_File)
          exit for
        end if
      next

      .write "<SPAN CLASS=Heading3>" & Translate(Caption(Sequence),Login_Language,conn) & "</SPAN>" & vbCrLf
      .write "<BR><BR>" & vbCrLf
    
      Call Table_Begin
    
      %>
      <TABLE CELLPADDING=4 CELLSPACING=0 CELLPADDING=0 BORDER=0 BGCOLOR="<%=Contrast%>">
        <TR>
          <TD COLSPAN=2 BGCOLOR="#777777" CLASS=SMALLWHITE>
          <%
          if Course = 50 or Course = 49 or Course = 10 then
            .write "<B>" & Translate("Click on the icons below to view or download these informative sales resource materials.",Login_Language,conn) & "</B>"          
          else
            .write "<B>" & Translate("Please download and print these files before you proceed with the training course, unless you already have printed copies of them.",Login_Language,conn) & "</B>"
          end if  

          if Course = 10 then
            .write "<P>"
            .write Translate("End-user Customer Information",Login_Language,conn)          
          end if  
          %>
          </TD>
        </TR>
          <%
          break = 0
          for x = 0 to Doc_Count
            if Doc_File(x) = "[break]" then
              .write "<TR>" & vbCrLf
              .write "<TD COLSPAN=2 BGCOLOR=""#777777"" CLASS=SMALLWHITE>" & vbCrLf
              .write Translate(Doc_Break(break),Login_Language,conn)
              .write "</TD>" & vbCrLf
              .write "</TR>" & vbCrLf
              break = break + 1
            else  
              .write "<TR>" & vbCrLf
              .write  "<TD CLASS=Small WIDTH=""2%"" ALIGN=CENTER>" & vbCrLf
              .write  "<A HREF=""javascript:void(0);"" Language=""javascript"" TITLE=""Click to Download File"" onclick=""openit_mini('/Find_It.asp?Document=" & Doc_File(x) & "','Vertical');return false;"">"
              
              .write  "<IMG SRC=""/images/Button_Down.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
              .write  "</A>"
              .write  "</TD>" & vbCrLf
              .write  "<TD Class=SmallBold>" & vbCrLf
              .write  Translate(Doc_Name(x),Login_Language,conn)
              .write  "</TD>" & vbCrLf
              .write  "</TR>" & vbCrLf
            end if  
          next
          %>
          <TR>
            <TD COLSPAN=2 ALIGN=CENTER BGCOLOR="#777777">
            <%
            .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.adobe.com/support/downloads/main.html','Vertical');return false;""><IMG SRC=""images/acrobat.gif"" BORDER=0 ALT=""Click to Download Acrobat Reader"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
            .write "&nbsp;&nbsp;&nbsp;"

            select case Login_Language
              case "eng"
                .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&P5_Language=English','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
              case "fre"
                .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&Lang=French&P5_Language=French','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
              case "por"
                .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&Lang=BrazilianPortuguese&P5_Language=BrazilianPortuguese','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
              case "spa"
                .write "<A HREF=""javascript:void(0);"" Language=""javascript"" ONCLICK=""openit_mini('http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash&Lang=Spanish&P5_Language=Spanish','Vertical');return false;""><IMG SRC=""images/flash.jpg"" BORDER=0 ALT=""Click to Download Macromedia Flash"" VSPACE=0 HSPACE=0 ALIGN=""absbottom""></A>"
              end select                
            %>
          </TR>
      </TABLE>
      <%
      Call Table_End
      
      .write "<P>&nbsp;<P>" & vbCrLf

    end if
    
    rsCourse.close
    set rsCourse = nothing  
    
    Sequence = Sequence + 1
    
    SQL = "SELECT PP_Register FROM Sales_Training WHERE ID=" & Course
    Set rsCourse = Server.CreateObject("ADODB.Recordset")
    rsCourse.Open SQL, conn, 3, 3
    
    PP_Register = rsCourse("PP_Register")
    
    rsCourse.close
    set rsCourse = nothing
    
    ' PP Registration and access
    
    if CInt(PP_Register) = CInt(true) then
    
      SQL = "SELECT * FROM UserData_Sales_Training WHERE ID=" & Account_ID
      Set rsProfile = Server.CreateObject("ADODB.Recordset")
      rsProfile.Open SQL, conn, 3, 3
  
      ppSQL = "SELECT * FROM UserData WHERE LastName='" & Replace(rsProfile("LastName"),"'","''") & "' AND Email='" & rsProfile("Email") & "'"
  
      Set rsPortal = Server.CreateObject("ADODB.Recordset")
      rsPortal.Open ppSQL, conn, 3, 3
      
      if not rsPortal.EOF then Sequence = Sequence + 1
      
      .write "<SPAN CLASS=Heading3>" & Translate(Caption(Sequence),Login_Language,conn) & " - " & Translate("Fluke Industrial Tools",Login_Language,conn) & "</SPAN>" & vbCrLf
      .write "<BR><BR>" & vbCrLf
      .write "<SPAN CLASS=Medium>" & vbCrLf
      .write "<IMG SRC=""images/Find-Home.jpg"" WIDTH=""400""><P>" & vbCrLf
        
      if rsPortal.EOF then
         
         select case UCase(Course_Code)
          case "GRAYBAR"
            .write Translate("The Fluke Partner Portal is a web site designed specifically for Fluke distributors, and contains a wealth of information to help you sell more Fluke test instruments, such as order status, product data sheets, application notes, sales guides and images.",Login_Language,conn) & "<P>" & vbCrLf
          case else
            .write Translate("The Fluke Partner Portal is a web site designed specifically for Fluke distributors, and contains a wealth of information to help you sell more Fluke and Meterman test instruments, such as order status, product data sheets, application notes, sales guides and images.",Login_Language,conn) & "<P>" & vbCrLf
        end select
        
        .write Translate("If you would like to register for access to the Fluke Partner Portal, click on the [ Partner Portal Registration ] button below.",Login_Language,conn) & "<P>" & vbCrLf
        
      ' Update Profile
  
        .write "<FORM NAME=""Portal_Register"" ACTION=""/register/register.asp"" METHOD=""POST"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""3"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Account_ID"" VALUE=""new"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" VALUE=""Search"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""FirstName_MI"" VALUE=""" & rsProfile("FirstName") & """>" & vbCrLf      
        .write "<INPUT TYPE=""HIDDEN"" NAME=""LastName"" VALUE=""" & rsProfile("LastName") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Company"" VALUE=""" & rsProfile("Company") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Address1"" VALUE=""" & rsProfile("Business_Address") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Address2"" VALUE=""" & rsProfile("Business_Address_2") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""City"" VALUE=""" & rsProfile("Business_City") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""State_Province"" VALUE=""" & rsProfile("Business_State") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Country"" VALUE=""" & rsProfile("Business_Country") & """>" & vbCrLf      
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Email"" VALUE=""" & rsProfile("Email") & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Phone"" VALUE=""" & rsProfile("Business_Phone") & """>" & vbCrLf      
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Type_Code"" VALUE=""1"">" & vbCrlf      
        .write "<INPUT TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""" & Translate("Partner Portal Registration",Login_Language,conn) & """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        .write "<INPUT TYPE=""BUTTON"" NAME=""CLOSE"" VALUE=""" & Translate("Close Window",Login_Language,conn) & """ ONCLICK=""window.close();"" CLASS=NavLeftHighlight1>" & vbCrLf
        .write "</FORM>" & vbCrLf
          
      else
  
        .write Translate("If you would like to logon to your Fluke Partner Portal account, click on the [ Logon ] button below.",Login_Languagne,conn) & "<P>" & vbCrLf
         
        .write "<FORM NAME=""Portal_Logon"" ACTION=""/register/login.asp"" METHOD=""POST"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""3"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & Login_Language & """>" & vbCrlf
        .write "<INPUT TYPE=""SUBMIT"" NAME=""SUBMIT"" VALUE=""" & Translate("Logon",Login_Language,conn) & """ CLASS=NavLeftHighlight1>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        .write "<INPUT TYPE=""BUTTON"" NAME=""CLOSE"" VALUE=""" & Translate("Close Window",Login_Language,conn) & """ ONCLICK=""window.close();"" CLASS=NavLeftHighlight1>" & vbCrLf
        .write "</FORM>" & vbCrLf
  
      end if
        
      rsPortal.close
      set rsPortal = nothing
      
      rsProfile.close
      set rsProfile = nothing
  
    else
    
      .write "<FORM NAME=""Close"">" & vbCrLf
      .write "<INPUT TYPE=""BUTTON"" NAME=""CLOSE"" VALUE=""" & Translate("Close Window",Login_Language,conn) & """ ONCLICK=""window.close();"" CLASS=NavLeftHighlight1>" & vbCrLf
      .write "</FORM>" & vbCrLf

    end if
    
  case else            
  
end select

' --------------------------------------------------------------------------------------
' Debug
' --------------------------------------------------------------------------------------

if CInt(Session("Debug_Flag")) = CInt(True) then

  .write "<P>------------------------------------------------<BR>" & vbCrLf
  .write "Debug - Key=Value Pairs<BR>" & vbCrLf
  .write "------------------------------------------------<P>" & vbCrLf
  .write "<B>Sequence: " & Sequence & "<P>" & vbCrLf
  .write "<B><FONT COLOR=RED>QueryString Object</FONT></B><BR>" & vbCrLf
  for each item in request.querystring
    .write item & "=" & request.querystring(item) & "<BR>" & vbCrLf
  next  
  .write "<P><B><FONT COLOR=RED>Form Object</FONT></B><BR>" & vbCrLf
  for each item in request.form
    .write item & "=" & request.form(item) & "<BR>" & vbCrLf
  next  
  .write "<BR>------------------------------------------------<P>" & vbCrLf
end if  

end with

' End Content

response.write "<P>" & vbCrLf

' --------------------------------------------------------------------------------------

%>  
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<!--#include virtual="/include/core_countries.inc"-->
<%

' --------------------------------------------------------------------------------------
' Functions and Subroutines
' --------------------------------------------------------------------------------------

function Get_New_Record_ID (myTable, myField, myValue, connection)

  ' Creates new Record, then Returns ID Number for UPDATE
 
  set objRS = Server.CreateObject("ADODB.Recordset")

  with objRS
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
  	.ActiveConnection = connection
  end with

  objRS.Open myTable,,,,adCmdTable
  objRS.AddNew
  objRS(myField) = myValue
  objRS.Update
  objRS.Close
  objRS.open "select @@identity", conn, 3,3
  my_ID = objRS(0)
  set objRS = Nothing

  Get_New_Record_ID = my_ID
  
end function

'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" BGCOLOR=""#666666"">" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/SideNav_TL_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/SideNav_TR_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "        </TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/SideNav_BL_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""images/SideNav_BR_corner.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "    </TABLE>" & vbCrLf
end sub  

'--------------------------------------------------------------------------------------

select case Sequence
  case seqRegister, seqShipping
    %>

    <SCRIPT LANGUAGE=JAVASCRIPT>
    <!--
    // --------------------------------------------------------------------------------------
    
    function CheckRequiredFields() {
    
      var df = document.<%=Form_Name%>;
      var ErrorMsg;
      var strVal;
      var strChk;
      ErrorMsg = "";
      strVal = "";
    
      var TestRadio;
      var ctr;
      var RadioChecked;
      var valid = "0123456789-";
      var temp;
      var LastField = "";
          
      var badchars = /[\[\]:;|=,+*?<>"\\\/]/;
      var D_BadChars = '\\ / [ ] : ; | = , + * ? < > \' "';              //"'
    
      if (df.FirstName.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("First Name",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "FirstName";}
        df.FirstName.style.backgroundColor = "#FFB9B9";
      }
    
      if (df.LastName.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Last Name",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "LastName";}
        df.LastName.style.backgroundColor = "#FFB9B9";
      }
    
      if (df.Email.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email",Alt_Language,conn))%>\r\n";  
        if (LastField.length == 0) {LastField = "Email";}
        df.Email.style.backgroundColor = "#FFB9B9";
      }
      
      if (!df.Email.value.match(/^[\w]{1}[\w\.\-_]*@[\w]{1}[\w\-_\.]*\.[\w]{2,6}$/i)) { 
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Invalid Email Address",Alt_Language,conn))%>\r\n";  
        if (LastField.length == 0) {LastField = "Email";}
        df.Email.style.backgroundColor = "#FFB9B9";
      } 
    
      if (df.Company.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Company Name",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Company";}
        df.Company.style.backgroundColor = "#FFB9B9";
      }
      
      if ("<%=Form_Name%>" != "Register") {
        if (df.Business_Phone.value == "") {
          ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Phone",Alt_Language,conn))%>\r\n";
          if (LastField.length == 0) {LastField = "Business_Phone";}
          df.Business_Phone.style.backgroundColor = "#FFB9B9";
        }
      }

      if ("<%=Form_Name%>" == "Register") {
        if (df.Course_Code.value == "") {
          ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Select Course Name from list or scroll to the bottom of the list and select 'Enter Course Code Manually' to enter Course Code",Alt_Language,conn))%>\r\n";
          if (LastField.length == 0) {LastField = "Course_Code";}
          df.Course_Code.style.backgroundColor = "#FFB9B9";
        }
      }

      
      if ("<%=Form_Name%>" != "Register") {
        if (df.Business_Address.value == "") {
          ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Address",Alt_Language,conn))%>\r\n";
          if (LastField.length == 0) {LastField = "Business_Address";}
          df.Business_Address.style.backgroundColor = "#FFB9B9";
        }
      }

      if (df.Business_City.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("City",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Business_City";}
        df.Business_City.style.backgroundColor = "#FFB9B9";
      }
      
      if (df.Business_Country.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Country",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Business_Country";}
        df.Business_Country.style.backgroundColor = "#FFB9B9";
      }
      
      if ((df.Business_State.value == "" && df.Business_Country.value == "US") || (df.Business_State.value == "" && df.Business_Country.value == "CA") || (df.Business_State.value == "" && df.Business_Country.value == "")) {
        ErrorMsg = ErrorMsg + "<%=Translate("USA State or Canadian Province",Alt_Language,conn) & " " & Translate("or",Alt_Language,conn) & " " & Translate("N/A",Alt_Language,conn)%>\r\n";
        if (LastField.length == 0) {LastField = "Business_State";}
        df.Business_State.style.backgroundColor = "#FFB9B9";
      }    

      if ("<%=Form_Name%>" != "Register") {
        if (df.Business_Postal_Code.value == "") {
          ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code",Alt_Language,conn))%>\r\n";
          if (LastField.length == 0) {LastField = "Business_Postal_Code";}
          df.Business_Postal_Code.style.backgroundColor = "#FFB9B9";
        }  
        
        if (df.Business_Country.value == "US" && df.Business_Postal_Code.value != "") {
          strVal = df.Business_Postal_Code.value;
          if (strVal.length != 5 && strVal.length != 10) {
            ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code - 5 digit or 5 digit + 4",Alt_Language,conn))%>\r\n";
            if (LastField.length == 0) {LastField = "Business_Country";}
            df.Business_Postal_Code.style.backgroundColor = "#FFB9B9";
          }
          else {
            for (var i=0; i < strVal.length; i++) {
              strChk = "" + strVal.substring(i, i+1);
              if (valid.indexOf(strChk) == "-1") {
                ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code - Invalid Characters",Alt_Language,conn))%>\r\n";
                if (LastField.length == 0) {LastField = "Business_Postal_Code";}
                df.Business_Postal_Code.style.backgroundColor = "#FFB9B9";
              }
            }
          }
        }  
      }
      
      if (ErrorMsg != "") {
        ErrorMsg = "<%=KillQuote(Translate("Please enter the missing information for following REQUIRED fields (or use N/A (Not Applicable))",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
        if (LastField.length > 0) {
          df[LastField].focus();
        }  
        alert (ErrorMsg);
        return false;
      }
      else {
      	return true;
      } 
    }
    
    // --------------------------------------------------------------------------------------
    
    var highlightcolor="lightyellow"
    var ns6=document.getElementById&&!document.all
    var previous=''
    var eventobj
    
    //Regular expression to highlight only form elements
    var intended=/INPUT|TEXTAREA/
    
    //Function to check whether element clicked is form element
    
    function checkel(which){
      if (which.style&&intended.test(which.tagName)){
        if (ns6&&eventobj.nodeType==3)
          eventobj=eventobj.parentNode.parentNode
        return true;
      }
      else
      return false;
    }
    
    //Function to highlight form element
    
    function highlight(e){
      eventobj=ns6? e.target : event.srcElement
      if (previous!=''){
      if (checkel(previous))
        previous.style.backgroundColor=''
        previous=eventobj
        if (checkel(eventobj))
        eventobj.style.backgroundColor=highlightcolor
      }
      else{
      if (checkel(eventobj))
        eventobj.style.backgroundColor=highlightcolor
        previous=eventobj
      }
    }
    // --------------------------------------------------------------------------------------
    //-->
    </SCRIPT>
    
    <SCRIPT LANGUAGE="JavaScript">
      function CkManual() {
        if (document.<%=Form_Name%>.Course_Code.value == "MANUALLY") {
        window.location.href="/sales-training/default.asp?Course_Code=MANUALLY&Language=<%=Login_Language%>";
        }
        else {
        window.location.href="/sales-training/default.asp?Course_Code=" + document.<%=Form_Name%>.Course_Code.value + "&Language=<%=Login_Language%>";
        }
      }
    </SCRIPT>  
    
  <%  
  case else
end select  

' --------------------------------------------------------------------------------------
Set conf = Nothing
Set msg = Nothing
Call Disconnect_SiteWide
%>
<!--#include virtual="/include/Pop-Up.asp"-->  