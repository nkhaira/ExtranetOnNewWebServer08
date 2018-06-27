<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author: K. D. Whitlock
' Date:   2/1/2000
' 06/19/2002 - Added Fields and Re-Ordered Form to work with Euro DCM
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_formdata.asp" -->
<%
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

Dim BackURL
Dim HomeURL
Dim Site_ID
Dim Site_Code
Dim Site_Description
Dim Account_ID
Dim strUser
Dim DebugFlag
Dim ErrorMessage
Dim SendPassword
Dim HideCancel
Dim Border_Toggle

Border_Toggle = 0

HideCancel    = false

DebugFlag     = False
SendPassword  = False

BackURL       = "/register/register.asp"
HomeURL       = "/register/default.asp"
ErrorMessage  = ""

Dim CoreMax
CoreMax = 22

Dim Core(22)
for i = 0 to CoreMax
  Core(i) = ""
next

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

Site_ID          = request("Site_ID")
Account_ID       = request("Account_ID")

if Site_ID = 100 then
  response.redirect HomeURL
end if

if request("ELO") = "on" then
  Session("ELO") = "on"
  Session("Language") = "eng"
end if

' --------------------------------------------------------------------------------------  
' New Record
' --------------------------------------------------------------------------------------

if lcase(Account_ID) = "new" or lcase(Account_ID) = "password" then
 
  SQL = "SELECT * FROM Site WHERE ID=" & CInt(Site_ID)
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3
  
  Site_Code         = rsSite("Site_Code")
  Site_Description  = rsSite("Site_Description")
  Logo              = rsSite("Logo")
  Logo_Left         = rsSite("Logo_Left")
  Footer_Disabled   = rsSite("Footer_Disabled")
  Business          = CInt(rsSite("Business"))
  Contrast          = rsSite("Contrast")
'  if CInt(rsSite("Business")) = -1 then Business = True else Business = False
  Privacy_Statement = rsSite("Privacy_Statement_Link")
  
  if lcase(Account_ID) = "new" then
    Screen_Title      = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Registration Request",Alt_Language,conn)
    Bar_Title         = Translate(Site_Description,Login_Language,conn) &  "<BR><SPAN CLASS=SmallBoldBar>" & Translate("Registration Request",Login_Language,conn) & "</SPAN>"
  else
    Screen_Title      = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Send Password Request",Alt_Language,conn)
    Bar_Title         = Translate(Site_Description,Login_Language,conn) &  "<BR><SPAN CLASS=SmallBoldBar>" & Translate("Send Password Request",Login_Language,conn) & "</SPAN>"
  end if
  
  Navigation        = False
  Top_Navigation    = False
  Content_Width     = 95  ' Percent
  
  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <%

  response.write "<TABLE BORDER=0 WIDTH=""100%"" CELLPADDING=0 CELLSPACING=0>"
  response.write "<TR>"
  response.write "<TD WIDTH=""100%"" HEIGHT=6 CLASS=TopColorBar><IMG SRC=""/images/1x1trans.gif"" HEIGHT=6 BORDER=0 VSPACE=0></TD>" & vbCrLf
  response.write "</TR>"

  response.write "<TR>"
  response.write "<TD CLASS=SMALL WIDTH=""100%"" BGCOLOR=WHITE>"
  
  response.write "<DIV ALIGN=CENTER>"
  response.write "<TABLE BORDER=0 WIDTH=""" & Content_Width & "%"">"
  response.write "<TR>"
  response.write "<TD WIDTH=""100%"" CLASS=Medium>"
  
  response.write "<BR><BR>"
  response.write "<SPAN CLASS=Normal><IMG SRC=""/images/lock.gif"" BORDER=0 WIDTH=37 ALIGN=""ABSMIDDLE"">&nbsp;&nbsp;&nbsp;" & Translate("This is a secure site connection to protect your personal information.",Login_Language,conn) & "</SPAN><BR><BR>"    
  rsSite.close
  set rsSite=nothing

  if not isblank(trim(request("Core_Email"))) then

    ' First check SiteWide DB to determine if email already exists, if True then deny new registration.
    
    SQL = "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & " AND UserData.EMail='" & trim(request("Core_Email")) & "'"
    if not isblank(trim(request("Core_LastName"))) then
      SQL = SQL & " AND LastName='" & trim(request("Core_LastName")) & "'"
    end if  
    Set rsUser = Server.CreateObject("ADODB.Recordset")
    rsUser.Open SQL, conn, 3, 3
    
    if NOT rsUser.EOF then      ' Account Found - Do not allow new registration

      if rsUser("NewFlag") = True then
        ErrorMessage = Translate("Your New Registration has been found and currently in the review process.  When approved, you will be contacted by email with information on how to access this site.",Login_Language,conn)
      elseif LCase(Account_ID) <> "password" then
        ErrorMessage = Translate("Your account information has been found, however, you cannot request an addition account at this site using the same email address.  If you have forgotten your Logon Name and Password, this information can be sent to your registered email address.",Login_Language,conn)
        SendPassword = True
      else
        ErrorMessage = Translate("Your account information has been found.",Login_Language,conn)
        SendPassword = True
        HideCancel   = True
      end if
      
    end if
    
    rsUser.close
    set rsUser = nothing
    
    if isblank(ErrorMessage) then    ' Check Core DB to Prepopulate form for New Account.

      ' Form is being populated from an auxillary form post.

        Call Connect_FormDatabase
              
        SQL = "SELECT * FROM tbl_Core WHERE tbl_Core.Email='" & trim(request("Core_Email")) & "'"
        if not isblank(trim(request("Core_LastName"))) then
          SQL = SQL & " AND LastName='" & trim(request("Core_LastName")) & "'"
        end if  
  
        Set rsCore = dbConnFormData.Execute(SQL)	
        
        ' Get core table data
        
        if not rsCore.EOF then
          Core(0) = trim(rsCore.Fields("CoreID").value)
          Core(1) = trim(rsCore.Fields("Prefix").value)
          Core(2) = trim(rsCore.Fields("FirstName_MI").value)
          if instr(1,Core(2),",") > 0 then
            Core(3) = trim(mid(Core(2),instr(1,Core(2)," ") + 1))
            Core(2) = trim(mid(Core(2),1,instr(1,Core(2)," ") - 1))
          else
            Core(3) = ""                  
          end if  
          Core(4)  = trim(rsCore.Fields("LastName").value)
          Core(5)  = trim(rsCore.Fields("Suffix").value)
          Core(6)  = trim(rsCore.Fields("Title").value)
          Core(7)  = trim(rsCore.Fields("MailStop").value)
          Core(8)  = trim(rsCore.Fields("Company").value)
          Core(9)  = trim(rsCore.Fields("Address1").value)
          Core(10) = trim(rsCore.Fields("Address2").value)
          Core(11) = trim(rsCore.Fields("City").value)
          Core(12) = trim(rsCore.Fields("State_Province").value)
          Core(13) = trim(rsCore.Fields("State_Other").value)
          Core(14) = trim(rsCore.Fields("Zip").value)
          Core(15) = trim(rsCore.Fields("Country").value)
          if trim(rsCore.Fields("Email").value) = "" then
            Core(16) = trim(request("Core_EMail"))
          else  
            Core(16) = trim(rsCore.Fields("Email").value)
          end if  
          Core(17) = trim(rsCore.Fields("Phone").value)
          Core(18) = trim(rsCore.Fields("Extension").value)
          Core(19) = trim(rsCore.Fields("Fax").value)
          Core(20) = trim(rsCore.Fields("Title").value)
          Core(21) = trim(rsCore.Fields("NativeLanguage"))
          
          if LCase(Session("elo")) = "on" then
            Core(21) = "elo"
          elseif not isblank(trim(request("Language"))) then
            Core(21) = Login_Language
          else
            Core(21) = Login_Language  
          end if
  
          ' Check with the eStore for ID if available.
          if not isblank(Core(0)) then        
          
            %>  
        	  <!-- #include virtual="/connections/connections_parts.asp" -->
            <%
          
            Call Connect_Parts
          
            SQLeStore = "SELECT * FROM vcturbo_shopper WHERE vcturbo_shopper.coreid='" & Core(0) & "'"
            Set rseStore = Server.CreateObject("ADODB.Recordset")
            rseStore.Open SQLeStore, dbconn, 3, 3
            
            if not rseStore.EOF then
              if not isblank(rseStore("Shopper_ID")) then
                Core(22) = rseStore("Shopper_ID")
              else
                Core(22) = ""  
              end if
            else
              core(22) = ""    
            end if
            
            rseStore.Close
            set rseStore = nothing
            
            Call Disconnect_Parts
              
          end if
          if LCase(Account_ID) <> "password" then
            response.write "<SPAN CLASS=Medium><UL><LI>"
            response.write Translate("Your profile information has been found in our database by searching for the email address that you have provided. Please ensure that this information is accurate and complete on the Registration Request form below",Login_Language,conn) & "."
          end if  
          response.write "</LI></UL></SPAN>"        
          
        else
    
          response.write "<SPAN CLSS=MediumRed><UL><LI>"
          response.write Translate("Your profile information was not found in our database by searching for the email address that you have provided. Please complete the following Registration Request form below",Login_Language,conn) & "."        
          response.write "</LI></UL></SPAN>"
  
          Core(16) = trim(request("Core_EMail"))
          Core(21) = Login_Language
          
          Account_ID = "new"
  
        end if

        Call Disconnect_FormDatabase
           
    end if
    
  else  ' Pre-Population with QueryString Values
  
    Core(0)  = trim(request("CoreID"))  ' www.fluke.com Core ID
    Core(1)  = trim(request("Prefix"))
    Core(2)  = (request("FirstName_MI"))
    if instr(1,Core(2),",") > 0 then
      Core(3) = trim(mid(Core(2),instr(1,Core(2)," ") + 1))
      Core(2) = trim(mid(Core(2),1,instr(1,Core(2)," ") - 1))
    else
      Core(3) = ""                  
    end if  

    Core(4)  = trim(request("LastName"))
    Core(5)  = trim(request("Suffix"))
    Core(6)  = trim(request("Title"))
    Core(7)  = trim(request("MailStop"))
    Core(8)  = trim(request("Company"))
    Core(9)  = trim(request("Address1"))
    Core(10) = trim(request("Address2"))
    Core(11) = trim(request("City"))
    Core(12) = trim(request("State_Province"))
    Core(13) = trim(request("State_Other"))
    Core(14) = trim(request("Zip"))
    Core(15) = trim(request("Country"))
    Core(16) = trim(request("Email"))
    Core(17) = trim(request("Phone"))
    Core(18) = trim(request("Extension"))
    Core(19) = trim(request("Fax"))
    Core(20) = trim(request("Title"))

    if LCase(Session("ELO")) = "on" then
      Core(21) = "elo"
    elseif not isblank(trim(request("Language"))) then
      Core(21) = trim(request("Language"))
    else
      Core(21) = Login_Language  
    end if
    
    Core(22) = trim(request("msscid"))  ' eStore Shopper ID
         
  end if      
  %>              
                                
  <!-- With Email Check Core-->
  
  <%
  if isblank(trim(request("Core_EMail"))) and isblank(trim(request("Email"))) and isblank(ErrorMessage) and Business then %>
  
    <FORM NAME="Check_Email" ACTION="register.asp" METHOD="GET">
    <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
    <INPUT TYPE="Hidden" NAME="Site_Description" VALUE="<%=Site_Description%>">
    <%
    if lcase(Account_ID) = "new" then
      response.write "<INPUT TYPE=""Hidden"" NAME=""Account_ID"" VALUE=""new"">"
    else
      response.write "<INPUT TYPE=""Hidden"" NAME=""Account_ID"" VALUE=""password"">"
    end if
    %>
    <INPUT TYPE="Hidden" NAME="Action" VALUE="Search">
    <INPUT TYPE="Hidden" NAME="Language" VALUE="<%=Login_Language%>">    

    <%Call Table_Begin%>
		<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>                   
	  	<TR>
      	<TD BGCOLOR="<%=Contrast%>" VALIGN=TOP WIDTH="50%" CLASS=MediumBold>
          <%
          response.write Translate("Type your Email Address",Login_Language,conn) & ":<BR>"
          response.write "<SPAN CLASS=Small>"
          if lcase(Account_ID) = "new" then
            response.write "<LI>" & Translate("to attempt to pre-populate this form for a new registration, or",Login_Language,conn) & "</LI>"
          end if
          response.write "<LI>" & Translate("to email your password if you already have an account,",Login_Language,conn) & "</LI><BR>"
          response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("then click on the [Search] button.",Login_Language,conn)
          response.write "</SPAN>"
          %>
        </TD>
       <TD BGCOLOR="White" ALIGN=LEFT WIDTH="50%" VALIGN=MIDDLE CLASS=MEDIUM>
          <INPUT TYPE="Text" NAME="Core_Email" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=MEDIUM>&nbsp;&nbsp;<INPUT TYPE="Submit" VALUE=" <%=Translate("Search",Login_Language,conn)%> " CLASS=NavLeftHighlight1>&nbsp;&nbsp;
          <INPUT TYPE="Button" NAME=CANCEL VALUE=" <%=Translate("Cancel",Login_Language,conn)%> " CLASS=NavLeftHighlight1 LANGUAGE="JavaScript" ONCLICK="window.location.href='/'">

        </TD>
      </TR>
    </TABLE>
    <%Call Table_End%>
    </FORM>
                
    <%
    if lcase(Account_ID) = "new" then
      response.write "<SPAN CLASS=MediumBOLD>" & Translate("New Account Registration Request Form",Login_Language,conn) & "<P>"
    end if
    
  end if %>
  
  <% if isblank(ErrorMessage) and lcase(Account_ID) = "new" then %>

    <FORM NAME="NTAccount" ACTION="register_admin.asp" METHOD="POST" onsubmit="return(CheckRequiredFields(this.form));" onKeyUp="highlight(event)" onClick="highlight(event)">
    <INPUT TYPE="Hidden" NAME="ID" VALUE="new">
    <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
    <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
    <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
    <INPUT TYPE="Hidden" NAME="Action" VALUE="registration">
    <INPUT TYPE="Hidden" NAME="send_email_admin" value="on">
    <INPUT TYPE="Hidden" NAME="send_email_fcm" value="on">
    <INPUT TYPE="Hidden" NAME="Subscription" value="on">
    <INPUT TYPE="Hidden" NAME="ChangeDate" VALUE="<%response.write FormatDate(1, Date())%>">
    <INPUT TYPE="Hidden" NAME="Core_ID" VALUE="<%=core(0)%>">
    <INPUT TYPE="Hidden" NAME="Mscssid" VALUE="<%=core(22)%>">    
    <% response.write "<INPUT TYPE=""Hidden"" NAME=""Business"" VALUE=""" & Business & """>" %>   
    
    <%
    if Login_Language <> Alt_Language then
      response.write "<SPAN CLASS=MediumBoldRed>" & Translate("Please enter your registration information in English",Login_Language,conn) & "</SPAN><P>"
    end if
    %>  

    <%Call Table_Begin%>
    <TABLE WIDTH="100%" BGCOLOR="#666666" BORDER=0 CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
    	<TR>
    		<TD WIDTH="100%" BGCOLOR="#EEEEEE"-->
    			<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>
          
            <!-- Header -->
    				<TR>
            	<TD WIDTH="40%" BGCOLOR="Black" COLSPAN=2 CLASS=MediumBoldGold>
                <%=Translate("Description",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>(<%=Translate("Note",Login_Language,conn)%>:&nbsp;<IMG SRC="images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> = <%=Translate("Required Information or use N/A",Login_Language,conn)%>)</SPAN>
              </TD>
    	        <TD WIDTH="60%" BGCOLOR="Black" ALIGN=LEFT CLASS=MediumBoldGold>
                <%=Translate("Your Information",Login_Language,conn)%>
              </TD>
            </TR>
                                                           
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Account Information",Login_Language,conn)%></TD></TR>

            <!-- Preferred Language -->
            
    				<%
            if LCase(Session("ELO")) <> "on" then
            %>
              <TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                  <%=Translate("Preferred Language",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  response.write "<SELECT CLASS=Medium NAME=""Language"" CLASS=Medium LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='register.asp?Site_ID=" & Site_ID & "&Core_Email=" & request("Core_EMail") & "&Account_ID=new&Type_Code=" & Request("Type_Code") & "&Language='+this.options[this.selectedIndex].value"">" & vbCrLf                      

                  SQL = "SELECT Language.* FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
                  Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                  rsLanguage.Open SQL, conn, 3, 3

                  Do while not rsLanguage.EOF
                    if LCase(rsLanguage("Code")) = LCase(Login_Language) then
                   	  response.write "<OPTION CLASS=Medium SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
                    else
                   	  response.write "<OPTION Class=Medium VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
                    end if
                	  rsLanguage.MoveNext 
                  loop
                  
                  rsLanguage.close
                  set rsLanguage=nothing
                  %>
                  </SELECT>              
                </TD>
              </TR>
            <%
            else
              response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""elo"">" & vbCrLf
            end if 

            ' Account Type
            
            SQL = "SELECT UserType.* FROM UserType WHERE UserType.Site_ID=" & CInt(Site_ID) & " ORDER BY UserType.Order_Num"
            Set rsUserType = Server.CreateObject("ADODB.Recordset")
            rsUserType.Open SQL, conn, 3, 3                    
            Account_Type = False
            if not rsUserType.EOF then
              Account_Type = True
		
  		        with response
      		      .write "<TR>"
            	  .write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
	  	          .write Translate("Your relationship to",Login_Language,conn) & " " & Translate(Site_Description,Login_Language,conn) & ":</TD>" & vbcrlf
            	  .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER WIDTH=""2%"" CLASS=Medium>"
                .write "    <IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                .write "</TD>"
    	          .write "<TD BGCOLOR=""White"" WIDTH=""50%"">"
                .write "  <INPUT TYPE=""HIDDEN"" NAME=""Type_Code_Required"" VALUE=""on"">"                                
                .write "  <SELECT NAME=""Type_Code"" CLASS=MEDIUM LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='register.asp?Site_ID=" & Site_ID & "&Core_Email=" & request("Core_EMail") & "&Account_ID=new&Language=" & Request("Language") & "&Type_Code='+this.options[this.selectedIndex].value"">" & vbCrLf
                .write "    <OPTION CLASS=Region5NavMedium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
	          	end with
           
              do while not rsUserType.EOF
                if CInt(rsUserType("Type_Code")) < 99 then 
                  response.write "    <OPTION CLASS=Medium VALUE=""" & rsUserType("Type_Code") & """"
                  if not isblank(request("Type_Code")) then
                    if CInt(rsUserType("Type_Code")) = CInt(request("Type_Code")) then response.write " SELECTED"
                  end if  
                  response.write ">" & Translate(rsUserType("Type_Description"),Login_Language,conn) & "</OPTION>"
                end if
                rsUserType.MoveNext
              loop
              response.write "</SELECT>"  
              response.write "&nbsp;"
              response.write "</TD>"
              response.write "</TR>"
            end if
            %> 
        
            <!-- NT Login -->
   
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Enter a New Logon User Name",Login_Language,conn)%>:<BR>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="NTLogin" SIZE="50" MAXLENGTH="20" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;<SPAN CLASS=SMALLest>(<%response.write Translate("7 or more characters",Login_Language,conn) & ", " & Translate("maximum",Login_Language,conn) & " 20"%>)</SPAN>
              </TD>
            </TR>
    
            <!-- NT Password -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Enter a New Logon Password",Login_Language,conn)%>:                                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Password" NAME="Password" SIZE="50" MAXLENGTH="14" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;<SPAN CLASS=SMALLest>(<%response.write Translate("7 or more characters",Login_Language,conn) & ", " & Translate("maximum",Login_Language,conn) & " 14"%>)</SPAN>
              </TD>
            </TR>
            
            <!-- NT Password Confirm -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Confirm New Logon Password",Login_Language,conn)%>:                                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Password" NAME="Password_Confirm" SIZE="50" MAXLENGTH="14" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;<SPAN CLASS=SMALLest>(<%response.write Translate("7 or more characters",Login_Language,conn) & ", " & Translate("maximum",Login_Language,conn) & " 14"%>)</SPAN>
              </TD>
            </TR>

            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Contact Information",Login_Language,conn)%></TD></TR>

            <!-- Name -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Name",Login_Language,conn)%>:&nbsp;&nbsp;<SPAN CLASS=Smallest><IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("First",Login_Language,conn)%>]&nbsp;&nbsp;[<%=Translate("Middle or Surname Prefix",Login_Language,conn)%>]&nbsp;<IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("Surname",Login_Language,conn)%>],&nbsp;&nbsp;[<%=Translate("Suffix",Login_Language,conn)%>]</SPAN>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                <INPUT TYPE="Text" NAME="FirstName" SIZE="10" MAXLENGTH="50" VALUE="<%=Core(2)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" NAME="MiddleName" SIZE="6" MAXLENGTH="50" VALUE="<%=Core(3)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;&nbsp;<IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><INPUT TYPE="Text" NAME="LastName" SIZE="11" MAXLENGTH="50" VALUE="<%=Core(4)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());"> <B>,</B>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" NAME="Suffix" SIZE="2" MAXLENGTH="50" VALUE="<%=Core(5)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>

            <!-- Initials -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Initials",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                <INPUT TYPE="Text" NAME="Initials" SIZE="10" MAXLENGTH="10" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>

    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Gender",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                <% 
                sValue = Core(1)
                %>
                <SELECT CLASS=Medium NAME="Gender" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
                  <OPTION CLASS=Medium VALUE=""><%=Translate("Select",Login_Language,conn)%></OPTION>
                  <OPTION CLASS=Region2 VALUE="0"<% If sValue = "Mr" Then Response.Write " SELECTED" %>><%=Translate("Male",Login_Language,conn)%></OPTION>
                  <OPTION CLASS=Region3 VALUE="1"<% If sValue = "Ms" or sValue = "Miss" or sValue="Mrs" Then Response.Write " SELECTED" %>><%=Translate("Female",Login_Language,conn)%></OPTION>
                </SELECT>
              </TD>
            </TR>

            <!-- Job Title -->

            <% if Business then %>    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Job Title",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Job_Title" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(20)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
            <% else %>
            <INPUT TYPE="Hidden" NAME="Job_Title" VALUE="N/A">
            <% end if %>
            
            <!-- Business Phone -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% 
                'if business then
                '  response.write Translate("Office",Login_Language,conn) & "&nbsp;"
                'end if  
                response.write Translate("Phone",Login_Language,conn)
                response.write " (" & Translate("Direct",Login_Languge,conn) & "):"
                %>                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=Core(17)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;<%=Translate("Extension",Login_Language,conn)%>: <INPUT TYPE="Text" NAME="Business_Phone_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=Core(18)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>

            <!-- Mobile Phone -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Mobile Phone",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Mobile_Phone" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Pager -->
    
            <% if business then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Pager",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Pager" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
                
            <!-- Email -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("EMail",Login_Language,conn)%> (<%=Translate("Direct",Login_Language,conn)%>):
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Email" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(16)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
            
            <!-- EMail Method -->
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("EMail Format",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
  
                <SELECT Name="EMail_Method" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
                  <OPTION CLASS=Medium VALUE=""><%=Translate("Select from List",Login_Language,conn)%></OPTION>
                  <OPTION Class=Region3 VALUE="0"><%=Translate("Plain Text without Graphics",Login_Language,conn)%></OPTION>
                  <OPTION Class=Region2 VALUE="1"><%=Translate("Rich Text with Graphics",Login_Language,conn)%></OPTION>
                </SELECT>                                  
  
              </TD>
            </TR>
                           
            <!-- Connection Speed -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("What is your typical connection speed to the internet?",Login_Language,conn)%>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                SQL = "SELECT Download_Time.* FROM Download_Time WHERE Download_Time.Enabled=" & CInt(True)
                SQL = SQL & " ORDER BY DownLoad_Time.bps, DownLoad_Time.Description"
                Set rsDownload = Server.CreateObject("ADODB.Recordset")
                rsDownload.Open SQL, conn, 3, 3
                response.write "<SELECT NAME=""Connection_Speed"" CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""return(Ck_Type_Code());"">" & vbCrLf
                response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                response.write "<OPTION CLASS=Region3 Value=""6"">" & Translate("Slow",Login_Language,conn) & "</OPTION>" & vbCrLf
                response.write "<OPTION CLASS=Region2 Value=""33"">" & Translate("Medium",Login_Language,conn) & "</OPTION>" & vbCrLf                
                response.write "<OPTION CLASS=Region1 Value=""13"">" & Translate("High",Login_Language,conn) & "</OPTION>" & vbCrLf
                response.write "<OPTION CLASS=Medium Value="""">" & "</OPTION>" & vbCrLf                                                                                                
                response.write "<OPTION CLASS=NavLeftHighlight1 Value="""">" & Translate("or Select Exact Speed from List Below",Login_Language,conn) & "</OPTION>" & vbCrLf
                response.write "<OPTION CLASS=Medium Value="""">" & "</OPTION>" & vbCrLf                                                                                
                do while not rsDownload.EOF
                  response.write "<OPTION CLASS=Medium VALUE=""" & rsDownload("ID") & """>" & rsDownload("Description") & "</OPTION>" & vbCrLf
                  rsDownload.MoveNext
                loop
                response.write "</SELECT>" & vbCrLf
                
                rsDownload.close
                set rsDownload = nothing
                %>
              </TD>
            </TR>            
           
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Company Information",Login_Language,conn)%></TD></TR>
                                      
            <!-- Company -->
    
            <% if Business then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Company Name",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Company" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(8)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
            <% else %>
            <INPUT TYPE="Hidden" NAME="Company" VALUE="N/A">
            <% end if %>
            
            <!-- Company Website -->
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Company Website Address",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Company_Website" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
              
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Office Information",Login_Language,conn)%></TD></TR>

            <!-- Business Mail Stop -->
                  
            <% if Business then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Mail Stop",Login_Language,conn)%> / <%=Translate("Building Number",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(7)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
            <% end if %>            

            <!-- Business Address -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Top CLASS=Medium>
                <% 
                response.write Translate("Address",Login_Language,conn)                  
                response.write ":"
                %>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(9)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());"><BR>
                <INPUT TYPE="Text" NAME="Business_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(10)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">                
              </TD>
            </TR>
       
            <!-- Business City -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% 
                response.write Translate("City",Login_Language,conn)
                response.write ":"
                %>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_City" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(11)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Business State -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% 
                response.write Translate("USA State or Canadian Province",Login_Language,conn)                
                response.write ":"
                %>                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <% SValue = Core(12)%>
                <SELECT CLASS=Medium NAME="Business_State" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
    
                 <%
                   response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List or Enter Below",Login_Language,conn) & "</OPTION>"
                 %>  
                  
                  <!--#include virtual="/include/core_states.inc"-->
    
                </SELECT>
    
              </TD>
            </TR>
    
            <!-- Business State Other-->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% response.write "<B>" & Translate("or",Login_Language,conn) & "</B>&nbsp;" & Translate("Other State, Province or Local",Login_Language,conn) & ":"%>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(13)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Business Postal Code -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% 
                response.write Translate("Postal Code",Login_Language,conn)                  
                response.write ":"
                %>                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(14)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Business Country -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% 
                response.write Translate("Country",Login_Language,conn)                  
                response.write ":"
                %>                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                Call Connect_FormDatabase
                Call displayCountryList("Business_Country",Core(15),Translate("Select from List",Login_Language,conn),"Medium")
                Call Disconnect_FormDatabase
                %>
              </TD>
            </TR>
    
            <!-- Email 2 -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("EMail",Login_Language,conn) & "&nbsp;&nbsp;&nbsp;(" & Translate("General Office",Login_Language,conn) & "):"%>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Email_2" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>

            <!-- Business Phone 2-->

    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%
                response.write Translate("Phone",Login_Language,conn)                
                response.write " (" & Translate("General Office",Login_Language,conn) & "):"
                %>                                
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_Phone_2" SIZE="28" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">&nbsp;&nbsp;<%=Translate("Extension",Login_Language,conn)%>: <INPUT TYPE="Text" NAME="Business_Phone_2_Extension" SIZE="10" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>

            <!-- Fax -->
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%
                response.write Translate("Fax",Login_Language,conn) & ":"
                %>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Business_Fax" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(19)%>" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
                          
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Postal Information",Login_Language,conn)%></TD></TR>

            <!-- Postal Address -->
       
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Postal Box Number",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Postal_Address" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Postal Same -->
       	
      			<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=SmallRed>
                <%=Translate("If the remainder of the Postal Address is the same as Office Address",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=SmallRed>
                <INPUT CLASS=NavLeftHighlight1 TYPE="BUTTON" VALUE="<%=Translate("Click Here",Login_Language,conn)%>" ONCLICK="Postal_Same();" ONFOCUS="return(Ck_Type_Code());">
                <!--INPUT TYPE="Checkbox" NAME="PostalSame" CLASS=Medium>&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
                -->
              </TD>
            </TR>

            <!-- Postal City -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("City",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Postal_City" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Postal State -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
    
                <% SValue = "" %>              
                <SELECT CLASS=Medium NAME="Postal_State" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
    
                 <%
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List or Enter Below",Login_Language,conn) & "</OPTION>"
                 %>  
    
                  <!--#include virtual="/include/core_states.inc"-->
    
                </SELECT>
              </TD>
            </TR>
    
            <!-- Postal State Other-->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <B><%=Translate("or",Login_Language,conn)%></B> <%=Translate("Other State, Province or Local",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Postal_State_Other" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Postal Code -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Postal Code",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Postal_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Postal Country -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Country",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                Call Connect_FormDatabase
                Call displayCountryList("Postal_Country","",Translate("Select from List",Login_Language,conn),"Medium")
                Call Disconnect_FormDatabase
                %>
              </TD>
            </TR>

            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Shipping Information",Login_Language,conn)%></TD></TR>

            <!-- Shipping Same -->
       	
      			<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=SmallRed>
                <%=Translate("If Shipping Address is the same as Office Address",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=SmallRed>
                <INPUT CLASS=NavLeftHighlight1 TYPE="BUTTON" VALUE="<%=Translate("Click Here",Login_Language,conn)%>" ONCLICK="Shipping_Same();" ONFOCUS="return(Ck_Type_Code());">
                <!--INPUT TYPE="Checkbox" NAME="ShippingSame" CLASS=Medium>&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
                -->
              </TD>
            </TR>
    
            <!-- Shipping Mail Stop -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Mail Stop",Login_Language,conn)%> / <%=Translate("Building Number",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Shipping_MailStop" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>              
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Top CLASS=Medium>
                <%=Translate("Address",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Shipping_Address" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());"><BR>
                <INPUT TYPE="Text" NAME="Shipping_Address_2" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Shipping City -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("City",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Shipping_City" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Shipping State -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
    
                <% SValue = "" %>              
                <SELECT CLASS=Medium NAME="Shipping_State" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
    
                 <%
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List or Enter Below",Login_Language,conn) & "</OPTION>"
                 %>  
    
                  <!--#include virtual="/include/core_states.inc"-->
    
                </SELECT>
              </TD>
            </TR>
    
            <!-- Shipping State Other-->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <B><%=Translate("or",Login_Language,conn)%></B> <%=Translate("Other State, Province or Local",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Shipping_State_Other" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Shipping Postal Code -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Postal Code",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Shipping_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="return(Ck_Type_Code());">
              </TD>
            </TR>
    
            <!-- Shipping Country -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Country",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                Call Connect_FormDatabase
                Call displayCountryList("Shipping_Country","",Translate("Select from List",Login_Language,conn),"Medium")
                Call Disconnect_FormDatabase
                %>
              </TD>
            </TR>
<% end if %>
                   
            <!-- Auxilliary Fields -->
    
            <%
            
            Dim Aux_Required(9)
            Dim Aux_Method(9)
            for i = 0 to 9
              Aux_Required(i) = False
              Aux_Method(i)   = 0
            next  
            
            SQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.Site_ID=" & CInt(Site_ID) & " AND Auxiliary.Enabled=" & CInt(True) & " AND Auxiliary.Registration=" & CInt(True) & " ORDER BY Auxiliary.Order_Num"
            Set rsAuxiliary = Server.CreateObject("ADODB.Recordset")
            rsAuxiliary.Open SQL, conn, 3, 3
            
            if not rsAuxiliary.EOF then

              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Silver"" CLASS=MediumBold>" & Translate("Other Information",Login_Language,conn) & "</TD></TR>"
            
              do while not rsAuxiliary.EOF

                if rsAuxiliary("Enabled") = CInt(True) and rsAuxiliary("Registration") = CInt(True) then              

            			response.write "<TR>"
                	response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=MIDDLE CLASS=Medium>"
                  response.write Translate(rsAuxiliary("Description"),Login_Language,conn) & ":"
                  response.write "</TD>"
                  response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                  if CInt(rsAuxiliary("Required")) = True then				  
                    Aux_Required(rsAuxiliary("Order_Num")) = True
                    Aux_Method(rsAuxiliary("Order_Num"))   = rsAuxiliary("Input_Method")
                    response.write "<IMG SRC=""images/required.gif"" Border=0 WIDTH=10 HEIGHT=10 ALIGN=ABSMIDDLE>"
                  else
                    response.write "&nbsp;"
                  end if  
                  response.write "</TD>"

   	              response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
    
                  Dim Aux_Selection
                  Dim Aux_Selection_Max

                  Aux_Selection     = Split(rsAuxiliary("Radio_Text"),",")
                  Aux_Selection_Max = Ubound(Aux_Selection)
                  
                  Select Case rsAuxiliary("Input_Method")
                    Case 0      ' Text
                      response.write "<INPUT TYPE=""Text"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ SIZE=""50"" MAXLENGTH=""50"" VALUE="""" CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""return(Ck_Type_Code());"">" & vbCrLf
                    Case 1      ' Drop-Down
                      response.write "<SELECT NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""return(Ck_Type_Code());"">" & vbCrLf
                      response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                      for i = 0 to Aux_Selection_Max
                        response.write "<OPTION CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "</OPTION>" & vbCrLf
                      next
                      response.write "</SELECT>" & vbCrLf
                    Case 2      ' Radio
                      for i = 0 to Aux_Selection_Max
                        response.write "<INPUT TYPE=RADIO NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium VALUE=""" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & """ LANGUAGE=""JavaScript"" ONFOCUS=""return(Ck_Type_Code());"">&nbsp;" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "&nbsp;&nbsp;" & vbCrLf
                      next
                  end select
  	        			response.write "<input type=""hidden"" Name=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Description"" Value=""" & Translate(rsAuxiliary("Description"),Login_Language,conn) & """>&nbsp;" & vbCrLf 
                  response.write "</TD>" & vbCrLf
                  response.write "</TR>" & vbCrLf

                end if                  
                
                rsAuxiliary.MoveNext
                
              loop
                           
            end if
            
            rsAuxiliary.Close
            set rsAuxiliary = nothing
            
            %>

            <!-- Navigation Buttons -->
    
            <TR>
              <TD COLSPAN=3>
                <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666" BORDER=0>
                  <TR>
                    <TD ALIGN=CENTER WIDTH="40%" CLASS=Medium>&nbsp;
                      <INPUT TYPE="BUTTON" CLASS=NavLeftHighLight1 NAME="Cancel" VALUE=" <%=Translate("Cancel Registration",Login_Language,conn)%> " onclick="Redirect('<%=HomeURL%>');">                      
                    </TD>             
                    <TD ALIGN=LEFT WIDTH="60%" CLASS=Medium>
                      <INPUT TYPE="Submit" CLASS=NavLeftHighLight1 NAME="Submit" VALUE=" <%=Translate("Submit Registration",Login_Language,conn)%> ">
                    </TD>
                  </TR>        
                </TABLE>
              </TD>
            </TR>        
          </TABLE>
        </TD>
      </TR>
    </TABLE>
    <%Call Table_End%>    
    
    </FORM>
    
    <!-- End of Content Container -->
      
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%   
  elseif ErrorMessage <> "" then %>
  
    <SPAN CLASS=MediumRed>
    <UL>
      <LI><%=ErrorMessage%>
    
      <% if SendPassword = True then %>
      
        <BR><BR>
        <%=Translate("Click on the [Send Password] to send Logon and Password information for this account to",Login_Language,conn)%>: <FONT COLOR="Black"><B><%=request("Core_Email")%></B></FONT>
        <BR><BR>      
        <FORM NAME="NTAccount-Send Password" ACTION="register_admin.asp" METHOD="POST">        
        <INPUT TYPE="Hidden" NAME="ID" VALUE="SendPassword">   
        <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
        <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
        <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
        <INPUT TYPE="Hidden" NAME="EMail" VALUE="<%=request("Core_Email")%>">
        <INPUT TYPE="Hidden" NAME="Language" VALUE="<%=Login_Language%>">
        <INPUT TYPE="Hidden" NAME="Action" VALUE=" Send Password ">
        <%
        Call Nav_Border_Begin

        if HideCancel <> true then
          response.write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" CLASS=NavLeftHighlight1 Value=""" & Translate("Cancel Registration",Login_Language,conn) & """ LANGUAGE=""JavaScript"" ONCLICK=""window.location.href='" & HomeURL & "';"">&nbsp;&nbsp;&nbsp;"
        end if
        response.write "<INPUT TYPE=""Submit"" VALUE=""" & Translate("Send Password",Login_Language,conn) & """ CLASS=NavLeftHighlight1>"

        Call Nav_Border_End
        %>
        </FORM>
        
      <% else %>
      
        <BR><BR>
        <FORM NAME="Dummy0">
        <%Call Nav_Border_Begin%>
        <INPUT TYPE="BUTTON" NAME="Cancel" CLASS=NavLeftHighlight1 Value=" <%=Translate("Cancel Registration",Login_Language,conn)%> " onclick="Redirect('<%=HomeURL%>');">                      
        <%Call Nav_Border_End%>
        </FORM>

      <% end if %>         
      
      </LI>
    </UL>
    </SPAN>
    
        <!-- End of Content Container -->
                                
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
    
  <%  
  end if  

else

	Call Disconnect_SiteWide  
	response.redirect HomeURL

end if

' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------

%>
<!-- #include virtual="/include/core_countries.inc"-->

<SCRIPT LANGUAGE=JAVASCRIPT>

<!--
// --------------------------------------------------------------------------------------

function Redirect(MyURL){ 
  window.location = MyURL;
}

// --------------------------------------------------------------------------------------
<%
if lcase(Account_ID) = "new" then
%>

function CheckRequiredFields() {

  var df = document.NTAccount;
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

  var badchars = /[\[\]:;|=,+*?<>"\\\/]/
  var D_BadChars = '\\ / [ ] : ; | = , + * ? < > "';              //"'

  if (df.Type_Code.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Relationship to ",Alt_Language,conn)) & " " & Translate(Site_Description,Login_Language,conn)%>\r\n";  
    if (LastField.length == 0) {LastField = "Type_Code";}
    df.Type_Code.style.backgroundColor = "#FFB9B9";
  }

  strVal = df.NTLogin.value;     
  if (strVal.length < 7 || strVal.length > 20) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Logon User Name must be at least 7 characters",Alt_Language,conn)) & ", " & Translate("maximum",Alt_Language,conn) & " 20"%>\r\n";
    if (LastField.length == 0) {LastField = "NTLogin";}
    df.NTLogin.style.backgroundColor = "#FFB9B9";
  }
  if (badchars.test(strVal)) {
	  ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Logon User Name contains illegal characters",Alt_Language,conn)) & ": "%>" + D_BadChars + "\r\n";
    if (LastField.length == 0) {LastField = "NTLogin";}
    df.NTLogin.style.backgroundColor = "#FFB9B9";
  }
  
  strVal = df.Password.value;
  if (strVal.length < 7 || strVal.length > 14) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Logon Password must be at least 7 characters",Alt_Language,conn)) & ", " & Translate("maximum",Alt_Language,conn) & " 14"%>\r\n";
    if (LastField.length == 0) {LastField = "Password";}
    df.Password.style.backgroundColor = "#FFB9B9";
  }

  if (badchars.test(strVal)) {
	  ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Logon Password contains illegal characters",Alt_Language,conn)) & ": "%>" + D_BadChars + "\r\n";
    if (LastField.length == 0) {LastField = "Password";}
    df.Password.style.backgroundColor = "#FFB9B9";
  }
  
  strVal = df.Password_Confirm.value;
  if (strVal.length < 7 || strVal.length > 14) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Confirm Logon Password must be at least 7 characters",Alt_Language,conn)) & ", " & Translate("maximum",Alt_Language,conn) & " 14"%>\r\n";
    if (LastField.length == 0) {LastField = "Password_Confirm";}
    df.Password_Confirm.style.backgroundColor = "#FFB9B9";
  }
  
  if (badchars.test(strVal)) {
	  ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Confirm Logon Password contains illegal characters",Alt_Language,conn)) & ": "%>" + D_BadChars + "\r\n";
    if (LastField.length == 0) {LastField = "Password_Confirm";}
    df.Password_Confirm.style.backgroundColor = "#FFB9B9";
  }

  if (df.Password.value != df.Password_Confirm.value) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Logon Password and Confirm Logon Password must match",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Password";}
    df.Password.style.backgroundColor = "#FFB9B9";
  } 

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

  if (df.Gender.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Gender",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Gender";}
    df.Gender.style.backgroundColor = "#FFB9B9";
  }

  if (df.Job_Title.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Job Title",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Job_Title";}
    df.Job_Title.style.backgroundColor = "#FFB9B9";
  }
  
  if (df.Business_Phone.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Office",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Phone",Alt_Language,conn))%> (<%=ReplaceRSQuote(Translate("Direct",Alt_Language,conn))%>)\r\n";
    if (LastField.length == 0) {LastField = "Business_Phone";}
    df.Business_Phone.style.backgroundColor = "#FFB9B9";
  }

  if (df.Email.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email",Alt_Language,conn))%> (<%=ReplaceRSQuote(Translate("Direct",Alt_Language,conn))%>)\r\n";  
    if (LastField.length == 0) {LastField = "Email";}
    df.Email.style.backgroundColor = "#FFB9B9";
  }
  
  if (!df.Email.value.match(/^[\w]{1}[\w\.\-_]*@[\w]{1}[\w\-_\.]*\.[\w]{2,6}$/i)) { 
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Invalid Email Address",Alt_Language,conn))%>\r\n";  
    if (LastField.length == 0) {LastField = "Email";}
    df.Email.style.backgroundColor = "#FFB9B9";
  } 

  if (df.EMail_Method.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email Format",Alt_Language,conn))%>\r\n";  
    if (LastField.length == 0) {LastField = "EMail_Method";}
    df.EMail_Method.style.backgroundColor = "#FFB9B9";
  }
  
  if (df.Connection_Speed.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Internet Connection Speed",Alt_Language,conn))%>\r\n";  
    if (LastField.length == 0) {LastField = "Connection_Speed";}
    df.Connection_Speed.style.backgroundColor = "#FFB9B9";
  }

  if (df.Company.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Company Name",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Company";}
    df.Company.style.backgroundColor = "#FFB9B9";
  }
  
  if (df.Business_Address.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Office",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Address",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Business_Address";}
    df.Business_Address.style.backgroundColor = "#FFB9B9";
  }

  if (df.Business_City.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Office",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("City",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Business_City";}
    df.Business_City.style.backgroundColor = "#FFB9B9";
  }

  if (df.Business_State.value == "" && df.Business_State_Other.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Office",Login_Language,conn) & "\t" & Translate("USA State or Canadian Province",Alt_Language,conn)) & " " & Translate("or",Alt_Language,conn) & " " & Translate("Other State, Province or County",Alt_Language,conn)%>\r\n";
    if (LastField.length == 0) {LastField = "Business_State";}
    df.Business_State.style.backgroundColor = "#FFB9B9";
  }
  else if (df.Business_State.value == "" && df.Business_State_Other.value != "") {
    df.Business_State.value = "ZZ";
  }

  if (df.Business_Country.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Office",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Country",Alt_Language,conn))%>\r\n";
    if (LastField.length == 0) {LastField = "Business_Country";}
    df.Business_Country.style.backgroundColor = "#FFB9B9";
  }

  if (df.Business_Postal_Code.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Office",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Postal Code",Alt_Language,conn))%>\r\n";
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
  
  if (df.Postal_State.value == "" && df.Postal_State_Other.value != "") {
    df.Postal_State.value = "ZZ";
  }

  if (df.Shipping_State.value == "" && df.Shipping_State_Other.value != "") {
    df.Shipping_State.value = "ZZ";
  }

  <%  
    for i = 0 to 9
      if Aux_Required(i) = true then 
        select case Aux_Method(i)
    			case 0  ' Text
    				response.write "if (df.Aux_" + Trim(CStr(i)) + ".value == """") {" & vbCrLf
    				response.write "  ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + df.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
            response.write "  if (LastField.length == 0) {LastField = ""Aux_" + Trim(CStr(i)) +""";}" & vbCrLf
            response.write "  df.Aux_" + Trim(CStr(i)) + ".style.backgroundColor = ""#FFB9B9"";" & vbCrLf
    				response.write "}" & vbCrLF & vbCrLF
    			case 1  ' Drop-Down
    				response.write "strVal = df.Aux_" + Trim(CStr(i)) + ".value; " & vbCrLF				
    				response.write "strChk = strVal.substring(0, 7); " & vbCrLF 
    				response.write "  if ((strVal == """") || (strChk.toLowerCase() == ""select"")) {" & vbCrLf
    				response.write "     ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + df.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
            response.write "  if (LastField.length == 0) {LastField = ""Aux_" + Trim(CStr(i)) +""";}" & vbCrLf
            response.write "  df.Aux_" + Trim(CStr(i)) + ".style.backgroundColor = ""#FFB9B9"";" & vbCrLf
    				response.write "  }" & vbCrLF & vbCrLF				 
          case 2  ' Radio or Checkbox
            response.write "RadioChecked = 0;" & vbCrLf
            response.write "TestRadio = df.Aux_" + Trim(Cstr(i)) & ";" & vbCrLf
            response.write "for (ctr=0; ctr < TestRadio.length; ctr++) {" & vbCrLf
            response.write "  if (TestRadio[ctr].checked) {" & vbCrLf
            response.write "   RadioChecked = 1;" & vbCrLf
            response.write "   break;" & vbCrLf
            response.write "   }" & vbCrLf
            response.write "}" & vbCrLf            
    				response.write "if (RadioChecked == 0) {" & vbCrLf 
    				response.write "  ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + df.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
    				response.write "}" & vbCrLF & vbCrLF
        end select
      end if
    next  
  %>
  
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
<% end if %>
function Postal_Same() {

  var df = document.NTAccount;
  var ErrorMsg = "";
  
  //df.Postal_Address.value     = df.Business_Address.value;
  df.Postal_City.value        = df.Business_City.value;            
  df.Postal_State.value       = df.Business_State.value;
  df.Postal_State_Other.value = df.Business_State_Other.value;
  df.Postal_Postal_Code.value = df.Business_Postal_Code.value;      
  df.Postal_Country.value     = df.Business_Country.value;            

  if (df.Postal_Address.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Box Number",Login_Language,conn))%>\r\n";
  }

  if (ErrorMsg != "") {
    ErrorMsg = "<%=KillQuote(Translate("Please enter the missing information for following REQUIRED fields (or use N/A (Not Applicable))",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    df.Postal_Address.focus();
    return false;
  }
  else {
    ErrorMsg = "<%=KillQuote(Translate("Your Office Address Information has been copied to Postal Address section of this form.",Alt_Language,conn) & "\n\n" & Translate("Please verify that your Postal Address information is correct.",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    df.Postal_Address.focus();
  	return true;
  } 
}

function Shipping_Same() {

  var df = document.NTAccount;
  var ErrorMsg = "";
    
  df.Shipping_MailStop.value    = df.Business_MailStop.value;
  df.Shipping_Address.value     = df.Business_Address.value;        
  df.Shipping_Address_2.value   = df.Business_Address_2.value;
  df.Shipping_City.value        = df.Business_City.value;            
  df.Shipping_State.value       = df.Business_State.value;
  df.Shipping_State_Other.value = df.Business_State_Other.value;
  df.Shipping_Postal_Code.value = df.Business_Postal_Code.value;      
  df.Shipping_Country.value     = df.Business_Country.value;            

  ErrorMsg = "<%=KillQuote(Translate("Your Office Address Information has been copied to Shipping Address section of this form.",Alt_Language,conn) & "\n\n" & Translate("Please verify that your Shipping Address information is correct.",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
  alert (ErrorMsg);

  df.Shipping_MailStop.focus();
}

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

function Ck_Type_Code() {

  var df = document.NTAccount;
  var ErrorMsg;
  var strVal;
  var strChk;
  ErrorMsg = "";
  strVal = "";

  strVal = df.Type_Code.value;
  strChk = strVal.substring(0, 7);
  if ((strVal == "") || (strChk.toLowerCase() == "select")) {
    ErrorMsg = "\n" + "<%=Translate("Before completing the rest of this registration form, please select your relationship to:",Alt_Language,conn)%>" + " " + "<%=Translate(Site_Description,Alt_Language,conn)%>" + "\r\n";
    alert(ErrorMsg);
    df.Type_Code.focus();
    return true;
  }
  return false;
}

// --------------------------------------------------------------------------------------
//-->
</SCRIPT>

<%
Call Disconnect_SiteWide
%>