<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   2/1/2001
' 07/04/2001 - Added Account Recprocity
' 01/10/2001 - Added Order Inquiry
' 06/19/2002 - Added Fields and Re-Ordered Form to work with Euro DCM
' --------------------------------------------------------------------------------------

Dim DebugFlag
DebugFlag = false

if Request("Language") = "XON" then
  Session("ShowTranslation") = True
elseif Request("Language")="XOF" then
  Session("ShowTranslation") = False
end if  

Dim Site_ID

Post_Method = "POST"

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_formdata.asp" -->
<%
Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim Site_Code

Account_ID    = request("Account_ID")

Dim strUser

Dim Aux_Selection
Dim Aux_Selection_Max
Dim Aux_Required(9)
Dim Aux_Method(9)
for i = 0 to 9
  Aux_Required(i) = False
  Aux_Method(i)   = 0
next  

Dim Region
Region         = 0
Dim RegionValue
RegionValue    = ""
Dim RegionColor(4)
RegionColor(0) = "#0000CC"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"
RegionColor(4) = "#FFCC99"

Dim SC0
Dim SC1
Dim SC2
Dim SC3
Dim SC4
Dim SC5
Dim SC9
Dim Page

SC0  = Request("SC0")
SC1  = Request("SC1")
SC2  = Request("SC2")
SC3  = Request("SC3")
SC4  = Request("SC4")
SC5  = Request("SC5")
SC9  = Request("SC9")
Page = Request("Page")

Dim BackURL
Dim HomeURL
Dim FormName

if SC0 > 0 then
  BackURL = "/SW-Administrator/Account_List.asp?Site_ID=" & Site_ID & "&SC0=" & SC0 & "&SC1=" & SC1 & "&SC2=" & SC2 & "&SC3=" & SC3 & "&SC4=" & SC4 & "&SC5=" & SC5 & "&SC9=" & SC9 & "&Page=" & Page & "&Z=" & Account_ID & "#" & Account_ID
  HomeURL = "/SW-Administrator/Default.asp?Site_ID=" & Site_ID
else
  BackURL = "/SW-Administrator/Default.asp?Site_ID=" & Site_ID
  HomeURL = "/SW-Administrator/Default.asp?Site_ID=" & Site_ID  
end if

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Screen_Title = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Users Account Administration Screen",Alt_Language,conn)
Bar_Title = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=NormalBoldGold>" & Translate("Users Account Administration Screen",Login_Language,conn) & "</FONT>"
Navigation = False
Top_Navigation = False
Content_Width = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<IFRAME STYLE="display:none;position:absolute;width:148;height:194;z-index=100" ID="CalFrame" MARGINHEIGHT=0 MARGINWIDTH=0 NORESIZE FRAMEBORDER=0 SCROLLING=NO SRC="/SW-Common/SW-Calendar_PopUp.asp"></IFRAME>
<%

response.write "<FONT CLASS=NormalBoldRed>"
select case Admin_Access
  case 2
    response.write Translate("Content Submitter",Login_Language,conn)
  case 4
    response.write Translate("Content Administrator",Login_Language,conn)
  case 6
    response.write Translate("Account Administrator",Login_Language,conn)
  case 8
    response.write Translate("Site Administrator",Login_Language,conn)
  case 9
    response.write Translate("Domain Administrator",Login_Language,conn)
end select
response.write "</FONT><BR>" & Admin_FirstName & " " & Admin_LastName & "<BR>" & Admin_Company & "<BR><BR>"

' if the page user is an admin then...
if (cint(Site_ID) = cint(Admin_Site_ID) and Admin_Access = 6) or (cint(Site_ID) = cint(Admin_Site_ID) and Admin_Access = 8) or Admin_Access = 9 then

' --------------------------------------------------------------------------------------
' Edit
' --------------------------------------------------------------------------------------  

  if IsNumeric(Account_ID) then  ' Account_ID came from the request object
  
      SQL = "SELECT UserData.*, UserData.ID "
      SQL = SQL & "FROM UserData "
      SQL = SQL & "WHERE (((UserData.ID)=" & CLng(Account_ID) & "))"
   
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sql, conn, 3, 3
  
      if not rs.EOF then
        FormName = "NTAccount"
      %>
        <FORM NAME="<%=FormName%>" ACTION="account_admin.asp" METHOD="<%=Post_Method%>" onsubmit="return(CheckRequiredFields(this.form));" onKeyUp="highlight(event)" onClick="highlight(event)">
        <INPUT TYPE="Hidden" NAME="ID" VALUE="<%=Account_ID%>">
        <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
        <INPUT TYPE="Hidden" NAME="Site_Code" VALUE="<%=Site_Code%>">
        <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
        <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
        <INPUT TYPE="Hidden" NAME="ChangeID" VALUE="<%=Admin_ID%>">
        <INPUT TYPE="Hidden" NAME="ChangeDate" Value="<%=Date%>">
        <INPUT TYPE="Hidden" NAME="CM_ID" Value="<%=rs("CM_ID")%>">        

        <%
        if CInt(rs("NewFlag")) = -2 then
          response.write "<FONT CLASS=MediumBoldRed>" & Translate("Note: Reciprocal Site Access Account.  See comment section below.",Login_Language,conn) & "</FONT><BR><BR>"
        end if
        
        Call Table_Begin
        %>

        <TABLE WIDTH="100%" BORDER=0 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
      	<TR>
      		<TD WIDTH="100%" BGCOLOR="#EEEEEE">
      			<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>
            
              <!-- Header -->
    		  		<TR>
              	<TD WIDTH="40%" COLSPAN=2 CLASS=NavLeftSelected1>
                  <%=Translate("Description",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallBold>(<%=Translate("Note",Login_Language,conn)%>:&nbsp;<IMG SRC="/images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> = <%=Translate("Required Information or use N/A",Login_Language,conn)%>)</SPAN>
                </TD>
      	        <TD WIDTH="60%" ALIGN=LEFT CLASS=NavLeftSelected1>
                  <%=Translate("User&acute;s Account Profile Information",Login_Language,conn)%>
                </TD>
              </TR>

    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=2 CLASS=MediumBold>
                  <%=Translate("Account Information",Login_Language,conn)%>
                </TD>
                <TD BGCOLOR="Silver" VALIGN=TOP CLASS=Medium NOWRAP>
                  <A HREF="/SW-Help/AA-UG.pdf"><IMG SRC="/images/help_button.gif" BORDER=0 ALIGN=RIGHT VALIGN=TOP ALT="Account Managers - User Guide"></A>
                    <%
                      response.write "<INPUT TYPE=""BUTTON"" Value=""" & Translate("Main Menu",Login_Language,conn) & """ Language=""JavaScript"" onclick=""location.href='" & BackURL & "'"" CLASS=NavLeftHighlight1 onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
                    %>                                     
                </TD>
              </TR>        
                  
              <!-- Account ID -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" WIDTH="38%" CLASS=Medium>
                  <%=Translate("Account ID Number",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER WIDTH="2%" CLASS=Medium>
                  &nbsp;
                </TD>                                
      	        <TD BGCOLOR="White" WIDTH="60%" CLASS=Medium>
                   <FONT COLOR="Gray"><%=rs("ID")%></FONT>
                </TD>
              </TR>
                          
              <%
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
				          .write Translate("Fluke Customer Type",Login_Language,conn) & ":</TD>" & vbcrlf
              	  .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                  .write "    <IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                  .write "</TD>"
      	          .write "<TD BGCOLOR=""White"">"
                  .write "  <INPUT TYPE=""HIDDEN"" NAME=""Type_Code_Required"" VALUE=""on"">"                                
                  .write "  <SELECT NAME=""Type_Code"" CLASS=MEDIUM>"
                  .write "    <OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
			        	end with
                
                do while not rsUserType.EOF
                  if (CInt(rsUserType("Type_Code")) < 99 and Admin_Access < 9) or Admin_Access = 9 then 
                    response.write "    <OPTION CLASS=Medium VALUE=""" & rsUserType("Type_Code") & """"
                    if CInt(rs("Type_Code")) = CInt(rsUserType("Type_Code")) then
                      response.write " SELECTED"
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

              rsUserType.close
              set rsUserType = nothing
              
              ' Account Manager or Account Manager Status
              
              response.write "<TR>"
              response.write "  <TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Account Manager",Login_Language,conn) & ":</TD>"
              response.write "	<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "    <IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
              response.write "  </TD>"
      	      response.write "  <TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "    <TABLE WIDTH=""100%"" CELLPADDING=0 CELLSPACING=0>"
              response.write "      <TR>"
              response.write "        <TD WIDTH=""50%"" CLASS=Medium>"
                
              response.write "          <SELECT NAME=""Fcm_ID"" CLASS=Medium>"
              response.write "          <OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
              
              SQL = "SELECT UserData.* "
              SQL = SQL & "FROM UserData "
              SQL = SQL & "WHERE (((UserData.Site_ID)=0 Or (UserData.Site_ID)=" & CInt(Site_ID) & ") AND ((UserData.Fcm)=" & CInt(True) & "))"
              SQL = SQL & "ORDER BY UserData.LastName"
              
              Set rsManager = Server.CreateObject("ADODB.Recordset")
              rsManager.Open SQL, conn, 3, 3                    
                                                              
              Do while not rsManager.EOF
                response.write "<OPTION"
                if isnumeric(rs("Fcm_ID")) then
                  if CLng(rs("Fcm_ID")) = CLng(rsManager("ID")) then
                    response.write " SELECTED"
                  end if
                end if                  
                response.write " CLASS=Region" & Trim(CStr(rsManager("Region")))
             	  response.write " VALUE=""" & rsManager("ID") & """>" & rsManager("LastName") & ", " & rsManager("FirstName") & "</OPTION>"
            	  rsManager.MoveNext 
              loop
                 
              rsManager.close
              set rsManager=nothing
              
              if rs("Fcm_ID") = "0" then
                response.write "<OPTION VALUE=""0"" SELECTED>" & Translate("N/A",Login_Language,conn) & "</OPTION>"
              else
                response.write "<OPTION VALUE=""0"">" & Translate("N/A",Login_Language,conn) & "</OPTION>"
              end if  
              
              response.write "          </SELECT>"
              response.write "        </TD>"
              response.write "        <TD WIDTH=""50%"" BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium><B>" & Translate("or",Login_Language,conn) & "</B> " & Translate("Account Manager",Login_Language,conn) & " ?&nbsp;&nbsp;"
              response.write "          <INPUT TYPE=""Checkbox"" NAME=""Fcm"""
              if rs("Fcm") = CInt(True) then response.write " CHECKED"
              response.write " CLASS=Medium>"
              response.write "        </TD>"
              response.write "      </TR>"
              response.write "    </TABLE>"
                  
              response.write "  </TD>"
              response.write "</TR>"

              ' Fluke Customer Number & Business System
			        ' when adding another business system add another:
      			  '		telltale variable
      			  '		case clause
      			  '		option clause  (this last step is also required in the New section of code)
      			  ora_chk = ""
              mfg_chk = ""
              dgu_chk = ""
      			  select case rs("Business_System")
      			    case "ORA"
        				  ora_chk = "SELECTED"
      			    case "MFG"
        				  mfg_chk = "SELECTED"
      			    case "DGU"
        				  dgu_chk = "SELECTED"
      			  end select
              
      		    with response
                .write "<TR>"
              	.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
                .write Translate("Fluke Customer Number",Login_Language,conn) & " and "
  			      	.write Translate("Business System",Login_Language,conn)
	        			.write ":</TD>" & vbcrlf
                .write "</TD>"
              	.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                .write "&nbsp;"
                .write "</TD>"
                .write "<TD BGCOLOR=""White"" CLASS=Medium>"
                .write "<INPUT TYPE=""Text"" NAME=""Fluke_ID"" SIZE=""15"" MAXLENGTH=""50"" CLASS=MEDIUM VALUE=""" & rs("Fluke_ID") & """>"
        				.write "&nbsp;&nbsp;&nbsp;<SELECT Name=""Business_system"" CLASS=""MEDIUM"">" & vbcrlf
        				.write "    <OPTION CLASS=Medium VALUE="""">Select from List</OPTION>" & vbcrlf
        				.write "    <OPTION CLASS=MEDIUM VALUE=""ORA"" " & ora_chk & ">Oracle</OPTION>" & vbcrlf
        				.write "    <OPTION CLASS=MEDIUM VALUE=""MFG"" " & mfg_chk & ">Mfg/Pro</OPTION>" & vbcrlf
        				.write "    <OPTION CLASS=MEDIUM VALUE=""DGU"" " & dgu_chk & ">Canada</OPTION>" & vbcrlf
        				.write " </select>" & vbcrlf
        				.write "</TD>"
                .write "</TR>"
              end with
              %>
              
    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Contact Information",Login_Language,conn)%>
                </TD>
              </TR>        
                          
               <!-- Name -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                  <%=Translate("Name",Login_Language,conn)%>:&nbsp;&nbsp;<SPAN CLASS=Smallest><IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("First",Login_Language,conn)%>]&nbsp;&nbsp;[<%=Translate("Middle",Login_Language,conn)%>]&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("Surname",Login_Language,conn)%>],&nbsp;&nbsp;[<%=Translate("Suffix",Login_Language,conn)%>]</SPAN>
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                  <INPUT TYPE="Text" NAME="FirstName" SIZE="10" MAXLENGTH="50" VALUE="<%=rs("FirstName")%>" CLASS=Medium>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" NAME="MiddleName" SIZE="6" MAXLENGTH="50" VALUE="<%=rs("MiddleName")%>" CLASS=Medium>&nbsp;&nbsp;&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><INPUT TYPE="Text" NAME="LastName" SIZE="11" MAXLENGTH="50" VALUE="<%=rs("LastName")%>" CLASS=Medium> <B>,</B>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" NAME="Suffix" SIZE="2" MAXLENGTH="50" VALUE="<%=rs("Suffix")%>" CLASS=Medium>
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
                  <INPUT TYPE="Text" NAME="Initials" SIZE="10" MAXLENGTH="10" VALUE="<%=rs("Initials")%>" CLASS=Medium>
                </TD>
              </TR>
              
              <!-- Gender -->
              
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                  <%=Translate("Gender",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                  <% 
                  sValue = rs("Gender")
                  %>
                  <SELECT CLASS=Medium NAME="Gender" CLASS=Medium>
                    <OPTION CLASS=Medium VALUE=""><%=Translate("Select",Login_Language,conn)%></OPTION>
                    <OPTION CLASS=Region2 VALUE="0"<% If sValue = "0" Then Response.Write " SELECTED" %>><%=Translate("Male",Login_Language,conn)%></OPTION>
                    <OPTION CLASS=Region3 VALUE="1"<% If sValue = "1" Then Response.Write " SELECTED" %>><%=Translate("Female",Login_Language,conn)%></OPTION>
                  </SELECT>
                </TD>
              </TR>

               <!-- Job Title -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Job Title",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Job_Title" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Job_Title")%>" CLASS=Medium>
                </TD>
              </TR>

               <!-- Office Phone (Direct) -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Phone",Login_Language,conn) & " (" & Translate("Direct",Login_Language,conn)%>):
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=rs("Business_Phone")%>" CLASS=MEDIUM>&nbsp;&nbsp;<%=Translate("Extension",Login_Language,conn)%>: <INPUT TYPE="Text" NAME="Business_Phone_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=rs("Business_Phone_Extension")%>" CLASS=Medium>
                </TD>
              </TR>

               <!-- Mobile Phone -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Mobile Phone",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Mobile_Phone" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Mobile_Phone")%>" CLASS=Medium>
                </TD>
              </TR>
  
               <!-- Pager -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Pager",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Pager" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Pager")%>" CLASS=Medium>
                </TD>
              </TR>
  
               <!-- Email -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("EMail",Login_Language,conn)%> (<%=Translate("Direct",Login_Language,conn)%>):
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Email" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Email")%>" CLASS=Medium>
                </TD>
              </TR>

              <!-- EMail Method -->
            
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("EMail Format",Login_Language,conn)%>:</TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <SELECT Name="EMail_Method" CLASS=Medium>
                    <OPTION CLASS=Medium VALUE=""><%=Translate("Select from List",Login_Language,conn)%></OPTION>
                    <OPTION CLASS=Medium Value=""></OPTION>
                    <OPTION Class=Region3 VALUE="0"<%if rs("Email_Method") = 0 then response.write(" SELECTED")%>><%=Translate("Plain Text without Graphics",Login_Language,conn)%></OPTION>
                    <OPTION Class=Region2 VALUE="1"<%if rs("Email_Method") = 1 then response.write(" SELECTED")%>><%=Translate("Rich Text with Graphics",Login_Language,conn)%></OPTION>
                  </SELECT>
                </TD>
              </TR>

             <!-- Connection Speed -->
    
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                  <% response.write Translate("Internet Connection Speed",Login_Language,conn) & ":"%>
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  SQL = "SELECT Download_Time.* FROM Download_Time WHERE Download_Time.Enabled=" & CInt(True)
                  SQL = SQL & " ORDER BY DownLoad_Time.bps, DownLoad_Time.Description"
                  Set rsDownload = Server.CreateObject("ADODB.Recordset")
                  rsDownload.Open SQL, conn, 3, 3
                  response.write "<SELECT Class=Medium NAME=""Connection_Speed"">" & vbCrLf
                  response.write "<OPTION CLASS=Medium Value="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                  response.write "<OPTION CLASS=Medium Value="""">" & "</OPTION>" & vbCrLf                                                                                                                
                  response.write "<OPTION CLASS=Region3 Value=""6"">" & Translate("Slow",Login_Language,conn) & "</OPTION>" & vbCrLf
                  response.write "<OPTION CLASS=Region2 Value=""33"">" & Translate("Medium",Login_Language,conn) & "</OPTION>" & vbCrLf                
                  response.write "<OPTION CLASS=Region1 Value=""13"">" & Translate("High",Login_Language,conn) & "</OPTION>" & vbCrLf
                  response.write "<OPTION CLASS=Medium Value="""">" & "</OPTION>" & vbCrLf                                                                                                
                  response.write "<OPTION CLASS=NavLeftHighlight1 Value="""">" & Translate("or Select Exact Speed from List Below",Login_Language,conn) & "</OPTION>" & vbCrLf
                  response.write "<OPTION CLASS=Medium Value="""">" & "</OPTION>" & vbCrLf                                                                                
                  do while not rsDownload.EOF
                    response.write "<OPTION CLASS=Medium VALUE=""" & rsDownload("ID") & """"
                    if not isblank(rs("Connection_Speed")) then
                      if rs("Connection_Speed") = rsDownLoad("ID") then response.write " SELECTED"
                    elseif rsDownload("ID") = 6 and (isblank(rs("Connection_Speed")) or rs("Connection_Speed") = 0) then
                      response.write "SELECTED"
                    end if  
                    response.write ">" & rsDownload("Description") & "</OPTION>" & vbCrLf
                    rsDownload.MoveNext
                  loop
                  response.write "</SELECT>" & vbCrLf
                  
                  rsDownload.close
                  set rsDownload = nothing
                  %> 
                </TD>
              </TR>            
                            
              <!-- Subscription Service -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Subscription Service",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                  
                  <INPUT TYPE="Checkbox" NAME="Subscription" <%if rs("Subscription") = True then response.write " CHECKED"%> CLASS=Medium>
                </TD>
              </TR>

              <!-- Preferred Language -->
              
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Preferred Language",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                                
                  <SELECT Name="Account_Language" CLASS=Medium>    
                  <%
                  SQL = "SELECT * FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
                  Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                  rsLanguage.Open SQL, conn, 3, 3
                                        
                  Do while not rsLanguage.EOF
                    if rs("Language") = rsLanguage("Code") then
                   	  response.write "<OPTION SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
                    else
                   	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"
                    end if
                	  rsLanguage.MoveNext 
                  loop
                  
                  rsLanguage.close
                  set rsLanguage=nothing
                  %>
                  </SELECT>              
                </TD>
              </TR>

    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Account Credentials",Login_Language,conn)%>
                </TD>
              </TR>        

               <!-- Expiration Date -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium>
                  <% if CInt(rs("NewFlag")) = CInt(True) then
                       response.write "<B>" & Translate("Verify",Login_Language,conn) & "</B> "
                     elseif CInt(rs("NewFlag")) = -2 then
                       response.write "<B>" & Translate("Pre-Set",Login_Language,conn) & "</B> "
                     end if
                  %>                      
                  <%=Translate("Expiration Date",Login_Language,conn)%> <SPAN CLASS=Smallest> (<%=Translate("mm/dd/yyyy or Blank for Never",Login_Language,conn)%>)</SPAN>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                    <TABLE WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
                      <TR>
                        <TD WIDTH="50%" CLASS=Medium>
                          <%
                          if FormatDate(1,rs("ExpirationDate")) = "09/09/9999" then
                            response.write "<INPUT TYPE=""Text"" NAME=""ExpirationDate"" SIZE=""10"" MAXLENGTH=""50"" VALUE="""" CLASS=Medium>&nbsp;&nbsp;"
                          elseif isdate(rs("ExpirationDate")) then
                            response.write "<INPUT TYPE=""Text"" NAME=""ExpirationDate"" SIZE=""10"" MAXLENGTH=""50"" VALUE=""" & FormatDate(1,rs("ExpirationDate")) & """ CLASS=Medium>&nbsp;&nbsp;"
                          else
                            response.write "<INPUT TYPE=""Text"" NAME=""ExpirationDate"" SIZE=""10"" MAXLENGTH=""50"" VALUE="""" CLASS=Medium>&nbsp;&nbsp;"
                          end if
                          %>
                          <A HREF="javascript:void()" LANGUAGE="JavaScript" onClick="window.dateField = document.<%=FormName%>.ExpirationDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false"><IMG SRC="/images/calendar/calendar_icon.gif" BORDER=0 HEIGHT="21"ALIGN=TOP></A>
                          <%
                                              
                          if CInt(rs("NewFlag")) = CInt(True) or CInt(rs("NewFlag")) = -2 then
                            response.write "<INPUT TYPE=""HIDDEN"" NAME=""NewFlag"" VALUE=""on"">"
                          end if 
                          %>      
                        </TD>
                        <TD WIDTH="25%" BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium><%=Translate("Last Logon",Login_Language,conn)%>:</TD>
                        <TD WIDTH="25%" BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                          
                          <%
                            SQL = "SELECT Logon.Account_ID, Logon.Logon, Logon.Logoff FROM Logon WHERE Logon.Account_ID=" & rs("ID") & " ORDER BY Logon.Logon DESC"
                            Set rsLogon = Server.CreateObject("ADODB.Recordset")
                            rsLogon.Open SQL, conn, 3, 3                    
                            
                            if not rsLogon.EOF then
                              if instr(1,rsLogon("Logon"),"9999") = 0 then
                                if isdate(rs("ExpirationDate")) and CDate(rs("ExpirationDate")) < CDate(Date) then
                                  response.write "<FONT COLOR=""Red""><B>" & Translate("Expired",Login_Language,conn) & "</B></FONT>"
                                else  
                                  response.write FormatDate(1,rsLogon("Logon"))
                                end if  
                              else
                                response.write "Never"
                              end if
                            elseif isdate(rs("ExpirationDate")) and CDate(rs("ExpirationDate")) < CDate(Date) then
                              response.write "<FONT COLOR=""Red""><B>" & Translate("Expired",Login_Language,conn) & "</B></FONT>"
                            else
                              response.write Translate("Never",Login_Language,conn)
                            end if
                            
                            rsLogon.close
                            set rsLogon=nothing                       
                          %>                                                                                                                                                     
                      </TR>
                    </TABLE>
                </TD>
              </TR>
      
              <!-- NT Login -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <% if CInt(rs("NewFlag")) = CInt(True) then
                       response.write "<B>" & Translate("Requested",Login_Language,conn) & "</B>"
                     elseif CInt(rs("NewFlag")) = -2 then
                       response.write "<B>" & Translate("Pre-Set",Login_Language,conn) & "</B> "
                     else
                       response.write Translate("Current",Login_Language,conn)
                     end if
                      
                     response.write " " & Translate("Logon User Name",Login_Language,conn) & ":<BR>"
                  %>  
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <% if isblank(rs("NTLogin")) or CInt(rs("NewFlag")) = CInt(true) then %>
                    <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                  <% else %>
                    &nbsp;
                  <% end if %>    
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                    if isblank(rs("NTLogin")) and CInt(rs("NewFlag")) = CInt(True) then
                      response.write "<INPUT TYPE=""Text"" NAME=""NTLogin"" SIZE=""50"" MAXLENGTH=""20"" VALUE="""" CLASS=Medium>"
                    elseif isblank(rs("NTLogin")) then
                      response.write "<INPUT TYPE=""Text"" NAME=""NTLogin"" SIZE=""50"" MAXLENGTH=""20"" VALUE="""" CLASS=Medium>"
                    else
                      response.write "<INPUT TYPE=""Hidden"" NAME=""NTLogin"" VALUE=""" & rs("NTLogin") & """ CLASS=Medium>" & rs("NTLogin")
                    end if
                  %>    
                </TD>
              </TR>
      
               <!-- NT Password -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <% if CInt(rs("NewFlag")) = CInt(True) then
                       response.write "<B>" & Translate("Requested",Login_Language,conn) & "</B>"
                     elseif CInt(rs("NewFlag")) = -2 then
                       response.write "<B>" & Translate("Pre-Set",Login_Language,conn) & "</B> "
                     else
                       response.write Translate("Current",Login_Language,conn)
                     end if
                     response.write " " & Translate("Password",Login_Language,conn) & ":"
                   %>  
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <% if isblank(rs("Password")) or CInt(rs("NewFlag")) = CInt(true) then %>
                    <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                  <% else %>
                    &nbsp;
                  <% end if %>    
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                    if not isblank(rs("Password")) and CInt(rs("NewFlag")) = -2 then
                      response.write "<INPUT TYPE=""Hidden"" NAME=""Password"" VALUE=""" & rs("Password") & """ CLASS=Medium>" & rs("Password")
                    else
                      response.write "<INPUT TYPE=""Text"" NAME=""Password"" SIZE=""50"" MAXLENGTH=""14"" VALUE=""" & rs("Password") & """ CLASS=Medium>"
                    end if
                  %>    
                </TD>
              </TR>
              
               <!-- Change NT Password -->
      
              <% if rs("Password") <> "(hidden)" and not isblank("Password") and CInt(rs("NewFlag")) = CInt(False) then %>
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Change Password",Login_Language,conn)%>:                                
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                    <INPUT TYPE="Text" NAME="Password_Change" SIZE="50" MAXLENGTH="14" VALUE="" CLASS=Medium>
                </TD>
              </TR>
              <% else %>
                <INPUT TYPE="HIDDEN" NAME="Password_Change" Value="">
              <% end if %>
             
    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Company Information",Login_Language,conn)%>
                </TD>
              </TR>        

               <!-- Company -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Company",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Company" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Company")%>" CLASS=Medium>
                </TD>
              </TR>
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Company Website Address",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Company_Website" SIZE="50" MAXLENGTH="255" VALUE="<%=rs("Company_Website")%>" CLASS=Medium>
                </TD>
              </TR>

    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Office Information",Login_Language,conn)%>
                </TD>
              </TR>        
              
              <!-- Business Address -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Address",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                               
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Business_Address")%>" CLASS=Medium><BR>
                  <INPUT TYPE="Text" NAME="Business_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Business_Address_2")%>" CLASS=Medium>                
                </TD>
              </TR>
  
               <!-- Mail Stop -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Mail Stop",Login_Language,conn) & " / " & Translate("Building Number",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Business_MailStop")%>" CLASS=Medium>
                </TD>
              </TR>

              <!-- Business City -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("City",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Business_City")%>" CLASS=Medium>
                </TD>
              </TR>
  
              <!-- Business State -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <% SValue = rs("Business_State") %>
                  <SELECT NAME="Business_State" CLASS=Medium>  
                   <%
                   if isblank(rs("Business_State")) then
                    response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_language,conn) & "</OPTION>"
                   end if
                   %>  
                    
                  <!--#include virtual="/include/core_states.inc"-->
  
                  </SELECT>
                </TD>
              </TR>

              <!-- Business State Other-->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%="<B>" & Translate("or",Login_Language,conn) & "</B> " & Translate("Other State, Province or Local",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Business_State_Other")%>" CLASS=Medium>
                </TD>
              </TR>
  
              <!-- Business Postal Code -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Postal Code",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Business_Postal_Code")%>" CLASS=Medium>
                </TD>
              </TR>
  
              <!-- Business Country -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Country",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  Call Connect_FormDatabase
                  Call DisplayCountryList("Business_Country",rs("Business_Country"),Translate("Select from List",Login_Language,conn),"Medium")
                  Call Disconnect_FormDatabase                  
                  %>
                </TD>
              </TR>

              <!-- Email 2 -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("EMail",Login_Language,conn)%> (<%=Translate("General Office",Login_Language,conn)%>):
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Email_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Email_2")%>" CLASS=Medium>
                </TD>
              </TR>              

              <!-- Business Phone 2 -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Phone",Login_Language,conn)%> (<%=Translate("General Office",Login_Language,conn)%>):
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_Phone_2" SIZE="28" MAXLENGTH="50" VALUE="<%=rs("Business_Phone_2")%>">&nbsp;&nbsp;<%=Translate("Extension",Login_Language,conn)%>: <INPUT TYPE="Text" NAME="Business_Phone_2_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=rs("Business_Phone_2_Extension")%>" CLASS=Medium>
                </TD>
              </TR>
            
              <!-- Business Fax -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Fax",Login_Language,conn)%> (<%=Translate("General Office",Login_Language,conn)%>):
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Business_Fax" SIZE="28" MAXLENGTH="50" VALUE="<%=rs("Business_Fax")%>">
                </TD>
              </TR>

    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Postal Information",Login_Language,conn)%>
                </TD>
              </TR>        

              <!-- Postal Address -->
                    
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Postal Box Number",Login_Langugage,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Postal_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Postal_Address")%>" CLASS=Medium><BR>
                </TD>
              </TR>
  
              <!-- Postal Same -->

      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP COLSPAN=2 CLASS=SmallRed>
                  <%=Translate("If the remainder of the Postal Address is the same as Office Address",Login_Language,conn)%>:
                </TD>
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=SmallRed>
                  <INPUT TYPE="Checkbox" NAME="PostalSame" CLASS=Medium>&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
                </TD>
              </TR>

               <!-- Postal City -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("City",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Postal_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Postal_City")%>" CLASS=Medium>
                </TD>
              </TR>
  
               <!-- Postal State -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
  
                  <% SValue = rs("Postal_State") %>              
                  <SELECT NAME="Postal_State" CLASS=Medium>
  
                   <%
                   if isblank(rs("Postal_State")) then
                    response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
                   end if
                   %>  
  
                  <!--#include virtual="/include/core_states.inc"-->
  
                  </SELECT>
                </TD>
              </TR>
  
              <!-- Postal State Other-->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%="<B>" & Translate("or",Login_Langugage,conn) & "</B> " & Translate("Other State, Province or Local",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Postal_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Postal_State_Other")%>" CLASS=Medium>
                </TD>
              </TR>

              <!-- Postal Postal Code -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Postal Code",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Postal_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Postal_Postal_Code")%>" CLASS=Medium>
                </TD>
              </TR>
  
              <!-- Postal Country -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Country",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  Call Connect_FormDatabase
                  Call DisplayCountryList("Postal_Country",rs("Postal_Country"),Translate("Select from List",Login_Language,conn),"Medium")
                  Call Disconnect_FormDatabase
                  %>
                </TD>
              </TR>
  
    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Shipping Information",Login_Language,conn)%>
                </TD>
              </TR>        

              <!-- Ship Same -->

      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP COLSPAN=2 CLASS=SmallRed>
                  <%=Translate("If Shipping Address is the same as Office Address",Login_Language,conn)%>:
                </TD>
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=SmallRed>
                  <INPUT TYPE="Checkbox" NAME="ShippingSame" CLASS=Medium>&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
                </TD>
              </TR>

              <!-- Shipping Address -->
                    
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Address",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Shipping_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Shipping_Address")%>" CLASS=Medium><BR>
                  <INPUT TYPE="Text" NAME="Shipping_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Shipping_Address_2")%>" CLASS=Medium>                
                </TD>
              </TR>
  
              <!-- Shipping Mail Stop -->

      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Mail Stop",Login_Language,conn) & " / " & Translate("Building Number",Login_Langugage,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Shipping_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Shipping_MailStop")%>" CLASS=Medium>
                </TD>
              </TR>

               <!-- Shipping City -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("City",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Shipping_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Shipping_City")%>" CLASS=Medium>
                </TD>
              </TR>
  
               <!-- Shipping State -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
  
                  <% SValue = rs("Shipping_State") %>              
                  <SELECT NAME="Shipping_State" CLASS=Medium>
  
                   <%
                   if isblank(rs("Shipping_State")) then
                    response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
                   end if
                   %>  
  
                  <!--#include virtual="/include/core_states.inc"-->
  
                  </SELECT>
                </TD>
              </TR>
  
              <!-- Shipping State Other-->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%="<B>" & Translate("or",Login_Langugage,conn) & "</B> " & Translate("Other State, Province or Local",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Shipping_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Shipping_State_Other")%>" CLASS=Medium>
                </TD>
              </TR>

              <!-- Shipping Postal Code -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Postal Code",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Shipping_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rs("Shipping_Postal_Code")%>" CLASS=Medium>
                </TD>
              </TR>
  
              <!-- Shipping Country -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Country",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  Call Connect_FormDatabase
                  Call DisplayCountryList("Shipping_Country",rs("Shipping_Country"),Translate("Select from List",Login_Language,conn),"Medium")
                  Call Disconnect_FormDatabase
                  %>
                </TD>
              </TR>
  
    				  <TR>
            	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
                  <%=Translate("Site Administrative Permissions",Login_Language,conn)%>
                </TD>
              </TR>        
                       
              <!-- Groups Selection -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  &nbsp;
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" CLASS=Medium>               
            
                <%  
                  SQL = "SELECT SubGroups.*, SubGroups.Order_Num, SubGroups.Site_ID, SubGroups.Enabled "
                  SQL = SQL & "FROM SubGroups "
                  SQL = SQL & "WHERE ((SubGroups.Site_ID=" & Site_ID & ") AND (SubGroups.Enabled=" & CInt(True) & ")) "
                  SQL = SQL & "ORDER BY SubGroups.Order_Num"
  
                  Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
                  rsSubGroups.Open SQL, conn, 3, 3
                  
                  if not rsSubGroups.EOF then
                    
                    response.write "<TABLE WIDTH=""100%"">" & vbCrLf
        			
                    if Admin_Access = 9 then
					  
                      if instr(1,lcase(rs("SubGroups")), "domain") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""domain"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Domain Administrator",Login_Language,conn) & "</FONT><BR></TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""domain"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Domain Administrator",Login_Language,conn) & "</FONT></TD></TR>" & vbCrLf
                      end if
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
					  
                      if instr(1,lcase(rs("SubGroups")), "administrator") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""administrator"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Site Administrator",Login_Language,conn) & "</FONT><BR></TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""administrator"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Site Administrator",Login_Language,conn) & "</FONT></TD></TR>" & vbCrLf
                      end if
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
  
  			              ' we'll access "CSubGroups" in the javascript so I need it...                      
                      response.write "<INPUT TYPE=""HIDDEN"" NAME=""SubGroups"">"
					  
					          else
        					    ' Only a Domain Admin can demote a Site Admin, this maintains Site Admin status
				        	    ' if less than a Domain Admin is editing an administrator or a domain
                      if instr(1,lcase(rs("SubGroups")), "domain") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""SubGroups"" VALUE=""domain"">"
                      elseif instr(1,lcase(rs("SubGroups")), "administrator") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""SubGroups"" VALUE=""administrator"">"
	        				    else
					              ' we'll access "CSubGroups" in the javascript so I need it...
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""SubGroups"">"
                      end if
                    end if
					
                    if Admin_Access >= 8 then                    
                      response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Select <U>only one</U> Administrative Permission from this group, if applicable.",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
                      
                      if instr(1,lcase(rs("SubGroups")), "account") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""account"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>"
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""account"" CLASS=Medium></TD><TD CLASS=MediumRed>"
                      end if
'                     if instr(1,lcase(rs("SubGroups")), "account") > 0 then                      
                      response.write Translate("Account Administrator",Login_Language,conn) & "&nbsp;&nbsp;&nbsp;&nbsp;<FONT CLASS=Medium>"
                      
                      response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=4"
                      if rs("Account_Region") = 4 then response.write " CHECKED"
                      response.write "> " & Translate("ALL",Login_Language,conn) & "&nbsp;&nbsp;"

                      response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=1"
                      if rs("Account_Region") = 1 then response.write " CHECKED"
                      response.write "> " & Translate("USA",Login_Language,conn) & "&nbsp;&nbsp;"
                      
                      response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=2"
                      if rs("Account_Region") = 2 then response.write " CHECKED"
                      response.write "> " & Translate("EUR",Login_Language,conn) & "&nbsp;&nbsp;"
                      
                      response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=3"
                      if rs("Account_Region") = 3 then response.write " CHECKED"
                      response.write "> " & Translate("INT",Login_Language,conn) & "&nbsp;&nbsp;"
                      
                      response.write "</TD></TR>" & vbCrLf
                     
                      response.write "<TR><TD COLSPAN=2></TD></TR>" & vbCrLf
                      if instr(1,lcase(rs("SubGroups")), "content") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""content"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Content Administrator",Login_Language,conn)
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""content"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Content Administrator",Login_Language,conn)
                      end if

                      response.write "&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RTE_ENABLED"""
                      if not isblank(rs("RTE_Enabled")) then
                        if CInt(rs("RTE_Enabled")) = CInt(True) then response.write " CHECKED"
                      end if  
                      response.write ">&nbsp;<SPAN CLASS=Medium>" & Translate("Rich Text Editor",Login_Language,conn)
                      response.write "</SPAN></TD></TR>" & vbCrLf
                    end if
                    
                    ' Submitter
                    if Admin_Access >=6 then  
                      if instr(1,lcase(rs("SubGroups")), "submitter") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""submitter"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Content Submitter",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""submitter"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Content Submitter",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      end if
                    end if

                    if Admin_Access >=8 and 1=2 then
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
                      response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("ImageStore",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
                      if instr(1,lcase(rs("SubGroups")), "imstadm") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""imstadm"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("ImageStore Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""imstadm"" CLASS=Medium LANGUAGE=""JavaScript"" ONCLICK=""alert('Reminder: To complete the registration procedure for ImageStore Administration access for this account, you must email Gerda Meijer (Eindhoven) this user's Login Username and Password to set up an ImageStore Administration account.');""></TD><TD CLASS=MediumRed>" & Translate("ImageStore Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      end if
                    end if

                    if Admin_Access >=6 then
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
                    end if  
                                                                                                      
                    response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Forums",Login_Language,conn) & "</TD></TR>" & vbCrLf
                    response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf

                    ' Forum Moderator
                    
                    if Admin_Access >=6 then  
                      if instr(1,lcase(rs("SubGroups")), "forum") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""forum"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Forum or Discussion Group Moderator",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""forum"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Forum or Discussion Group Moderator",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      end if
                    end if

                    ' Inquiry

                    if Admin_Access >= 6 then

                      response.write "</TABLE>"
              				response.write "<TR>"
                      response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
                      response.write Translate("Inquiry Permissions",Login_Language,conn)
                      response.write "  </TD>"
            	        response.write "</TR>"
  
              				response.write "<TR>"
                      response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
                      response.write "    &nbsp;"
                      response.write "  </TD>"
                      response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
                      response.write "    &nbsp;"
                      response.write "  </TD>"
            	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
                      response.write "<TABLE WIDTH=""100%"">"
  
                      response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Order Inquiry",Login_Language,conn) & " (" & Translate("Select only if applicable",Login_Language,conn) & ")</TD></TR>" & vbCrLf

                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
                    end if  
                    
                    if Admin_Access >= 6 then
                    
                      if instr(1,lcase(rs("SubGroups")), "ordad") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""ordad"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Order Inquiry Administrator",Login_Language,conn) & " (" & Translate("Fluke Internal Use Only",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""ordad"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Order Inquiry Administrator",Login_Language,conn) & " (" & Translate("Fluke Internal Use Only",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      end if
                      
                      if instr(1,lcase(rs("SubGroups")), "order") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""order"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Order Inquiry with Search Capability",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""order"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Order Inquiry with Search Capability",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      end if

                      if instr(1,lcase(rs("SubGroups")), "literature") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""literature"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Literature Order Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""literature"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Literature Order Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
                      end if
                    end if

                    if Admin_Access >=6 then
                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf
                    end if  

'                    response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Branch Location Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
'                    response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf

                    ' Branch Location Administrator
'                    if Admin_Access >= 6 and Account_Type = True then  
'                      if instr(1,lcase(rs("SubGroups")), "branch") > 0 then
'                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""branch"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Branch Location Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
'                      else
'                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""branch"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Branch Location Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
'                     end if
'                    end if

                    ' WTB Address Administrator

                    if webadm = true or Admin_Access = 9 then
                                                                                                      
                      response.write "</TABLE>"
              				response.write "<TR>"
                      response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
                      response.write Translate("WTB Administrator Permission",Login_Language,conn)
                      response.write "  </TD>"
            	        response.write "</TR>"
  
              				response.write "<TR>"
                      response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
                      response.write "    &nbsp;"
                      response.write "  </TD>"
                      response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
                      response.write "    &nbsp;"
                      response.write "  </TD>"
            	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
                      
                      response.write "<TABLE WIDTH=""100%"">"
  
                      response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("WTB Editor",Login_Language,conn) & " (" & Translate("Select only one",Login_Language,conn) & ")</TD></TR>" & vbCrLf

                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf

                      if Admin_Access = 9 then  
                        if instr(1,lcase(rs("SubGroups")), "wtbadm") > 0 then
                          response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbadm"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Super Administrator",Login_Language,conn) & " (" & Translate("Fluke Internal Use Only",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                        else
                          response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbadm"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Super Administrator",Login_Language,conn) & " (" & Translate("Fluke Internal Use Only",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                        end if
                      else
                        if instr(1,lcase(rs("SubGroups")), "wtbadm") > 0 then
                          response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""wtbadm"">" & vbCrLf
                        end if  
                      end if  
                      
                      if instr(1,lcase(rs("SubGroups")), "wtbtsm, wtbaap") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbtsm, wtbaap"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Fluke Territory Sales Manager",Login_Language,conn) & " (" & Translate("Auto-Approve",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbtsm, wtbaap"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Fluke Territory Sales Manager",Login_Language,conn) & " (" & Translate("Auto-Approve",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      end if

                      if instr(1,lcase(rs("SubGroups")), "wtbtsm") > 0 and instr(1,lcase(rs("SubGroups")), "wtbaap") = 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbtsm"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Fluke Territory Sales Manager",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbtsm"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Fluke Territory Sales Manager",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      end if

                      if instr(1,lcase(rs("SubGroups")), "wtbdis, wtbaap") > 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbdis, wtbaap"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Reseller",Login_Language,conn) & ") (" & Translate("Auto-Approve",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbdis, wtbaap"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Reseller",Login_Language,conn) & ") (" & Translate("Auto-Approve",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      end if

                      if instr(1,lcase(rs("SubGroups")), "wtbdis") > 0 and instr(1,lcase(rs("SubGroups")), "wtbaap") = 0 then
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbdis"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Reseller",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      else
                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""wtbdis"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Store Front Editor",Login_Language,conn) & " (" & Translate("Reseller",Login_Language,conn) & ")</TD></TR>" & vbCrLf
                      end if

                    else
                    
                      if instr(1,lcase(rs("SubGroups")), "wtbadm") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""wtbadm"">" & vbCrLf
                      end if  
                      if instr(1,lcase(rs("SubGroups")), "wtbtsm, wtbaap") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""wtbtsm, wtbaap"">" & vbCrLf
                      elseif instr(1,lcase(rs("SubGroups")), "wtbtsm") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""wtbtsm"">" & vbCrLf
                      end if  
                      if instr(1,lcase(rs("SubGroups")), "wtbdis, wtbaap") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""wtbdis, wtbaap"">" & vbCrLf
                      elseif instr(1,lcase(rs("SubGroups")), "wtbdis") > 0 then
                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""wtbdis"">" & vbCrLf
                      end if  
                      
                    end if
                    
                    response.write "</TABLE>"
                    
                    ' Gateway Application Administrators

                    SQLGateway = "SELECT * FROM Gateway_Applications ORDER BY Gateway_Title"
                    Set rsGateway = Server.CreateObject("ADODB.Recordset")
                    rsGateway.Open SQLGateway, conn, 3, 3

                    if Admin_Access = 9 then
                                                                                                      
              				response.write "<TR>"
                      response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
                      response.write Translate("Gateway Administrator Permission",Login_Language,conn)
                      response.write "  </TD>"
            	        response.write "</TR>"
  
              				response.write "<TR>"
                      response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
                      response.write "    &nbsp;"
                      response.write "  </TD>"
                      response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
                      response.write "    &nbsp;"
                      response.write "  </TD>"
            	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
                      response.write "<TABLE WIDTH=""100%"">"
  
                      response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Gateway Applications",Login_Language,conn) & " (" & Translate("Select one or more Group Affiliations.",Login_Language,conn) & ")</TD></TR>" & vbCrLf

                      response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>" & vbCrLf

                      do while not rsGateway.EOF
                        if instr(1,lcase(rs("SubGroups")), lcase(rsGateway("Gateway_Code"))) > 0 then
                            response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""" & Trim(rsGateway("Gateway_Code")) & """ CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate(Trim(rsGateway("Gateway_Title")),Login_Language,conn) & "</TD></TR>" & vbCrLf
                        else
                            response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""" & Trim(rsGateway("Gateway_Code")) & """         CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate(Trim(rsGateway("Gateway_Title")),Login_Language,conn) & "</TD></TR>" & vbCrLf
                        end if
                        rsGateway.MoveNext
                      loop
                      
                      response.write "</TABLE>"
                      
                    else
                      do while not rsGateway.EOF
                        if instr(1,lcase(rs("SubGroups")), lcase(rsGateway("Gateway_Code"))) > 0 then
                          response.write "<INPUT TYPE=""HIDDEN"" NAME=""CSubGroups"" VALUE=""" & Trim(rsGateway("Gateway_Code")) & """>" & vbCrLf
                        end if
                        rsGateway.MoveNext
                      loop
                    end if
                    
                    rsGateway.close
                    set rsGateway = nothing
                    
                  end if
                  
                  ' Group Affiliations
          				response.write "<TR>"
                  response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
                  response.write Translate("Group Affiliations",Login_Language,conn)
                  response.write "  </TD>"
        	        response.write "</TR>"

          				response.write "<TR>"
                  response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
                  response.write "    &nbsp;"
                  response.write "  </TD>"
                  response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
                  response.write "    <IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" VALIGN=TOP>"
                  response.write "  </TD>"
        	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
                  response.write "<TABLE WIDTH=""100%"">"

                  response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Select one or more Group Affiliations.",Login_Language,conn) & "</TD></TR>" & vbCrLf
                  
                  Do while not rsSubGroups.EOF
                                
                    if RegionValue <> Mid(rsSubGroups("Code"),1,1) then
                      RegionValue = Mid(rsSubGroups("Code"),1,1)
                      Region = Region + 1
                      
                      if Region > 4 then Region = 4
                      
                      response.write "<TR><TD HEIGHT=8 WIDTH=20></TD><TD HEIGHT=8></TD></TR>" & vbCrLf
                    end if

                    if rsSubGroups("Enabled") = True and instr(1,lcase(rs("SubGroups")), lcase(rsSubGroups("Code"))) > 0 then
                      response.write "<TR><TD CLASS=Normal><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED CLASS=Normal></TD><TD CLASS=Normal BGCOLOR=""" & RegionColor(Region) & """>" & rsSubGroups("X_Description") & "</TD></TR>" & vbCrLf
                    elseif rsSubGroups("Enabled") <> True then
                      'response.write "<TR><TD CLASS=Normal>&nbsp;</TD><TD CLASS=Normal><FONT COLOR=""Gray"">" & rsSubGroups("X_Description") & "</FONT></TD></TR>" & vbCrLf
                    elseif rsSubGroups("Enabled") = True then' and rsSubGroups("Default_Select") <> True then
                      response.write "<TR><TD CLASS=Normal><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""" & rsSubGroups("Code") & """ CLASS=Normal></TD><TD CLASS=Normal BGCOLOR=""" & RegionColor(Region) & """>" & rsSubGroups("X_Description") & "</TD></TR>" & vbCrLf
                    end if
                
                	  rsSubGroups.MoveNext 

                  loop
                
                  response.write "</TABLE>" & vbCrLf
                     
                  rsSubGroups.close
                  set rsSubGroups=nothing
  
                  %>
                </TD>
              </TR>             
              
              <!-- Groups Affiliation Codes Listing-->
               
              <% if Admin_Access >=8 then %>      

      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Site / Group Affiliation Codes",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium><FONT COLOR="Gray">
                <% if not isblank(rs("Groups")) then
                     response.write "<FONT COLOR=""Red"">" & rs("Groups") & ",</FONT> "
                   end if                  
                   response.write rs("SubGroups")
                %>
                </FONT>
                </TD>
              </TR>
              
              <% end if
              
              ' Additional Site Access
              
              SQL =       "SELECT Site_Aux.Site_ID, Site.Site_Code, Site.Enabled, Site.Site_Description, Site.ID "
              SQL = SQL & "FROM Site_Aux LEFT JOIN Site ON Site_Aux.Site_ID_Aux = Site.ID "
              SQL = SQL & "WHERE Site_Aux.Site_ID=" & Site_ID & "AND (dbo.Site.Site_Code IS NOT NULL) ORDER BY Site.Site_Description"

              Set rsSite_Aux = Server.CreateObject("ADODB.Recordset")
              rsSite_Aux.Open SQL, conn, 3, 3
                  
              if not rsSite_Aux.EOF then

        				response.write "<TR>"
                response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
                response.write Translate("Reciprocal Site Permissions",Login_Language,conn)
                response.write "  </TD>"
      	        response.write "</TR>"
                
        				response.write "<TR>"
                response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
                response.write "    &nbsp;"
                response.write "  </TD>"
                response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
                response.write "    &nbsp;"
                response.write "  </TD>"
      	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
                response.write "<TABLE WIDTH=""100%"">"
                
                do while not rsSite_Aux.EOF
                  response.write "<TR><TD WIDTH=20 CLASS=Normal><INPUT TYPE=""Checkbox"" NAME=""Groups_Aux"" VALUE=""" & LCase(rsSite_Aux("Site_Code")) & """ CLASS=Normal"
                  if instr(1,LCase(rs("Groups_Aux")),LCase(rsSite_aux("Site_Code"))) > 0 then
                    response.write " CHECKED"
                  end if
                  response.write "></TD>"
                  response.write "<TD CLASS=Normal BGCOLOR=""#FFCC00"">" & rsSite_Aux("Site_Description") & "</TD>"
                  
                  ' Check Active Account Status

                  response.write "<TD WIDTH=""20%"" ALIGN=""Center"" CLASS=Small BGCOLOR="

                  SQL = "SELECT UserData.Site_ID, UserData.NTLogin, UserData.NewFlag, UserData.ExpirationDate FROM UserData WHERE UserData.Site_ID=" & rsSite_aux("ID") & " AND UserData.NTLogin='" & rs("NTLogin") & "'"
                  Set rsSite_NewFlag = Server.CreateObject("ADODB.Recordset")

                  rsSite_NewFlag.Open SQL, conn, 3, 3

                  if not rsSite_NewFlag.EOF then
                    if rsSite_NewFlag("NewFlag") = CInt(False) then
                      if CDate(rsSite_NewFlag("ExpirationDate")) < Date then
                        response.write """Red"">"
                        response.write Translate("Expired",Login_Language,conn)
                      else
                        response.write """#FFCC00"">"
                        response.write Translate("Active",Login_Language,conn)
                      end if
                    elseif rsSite_NewFlag("NewFlag") = CInt(True) or rsSite_NewFlag("NewFlag") = -2 then
                      response.write """#FFFF00"">"
                      response.write Translate("Pending",Login_Language,conn)
                    else
                      response.write """#FFCC00"">"
                      response.write "&nbsp;" 
                    end if
                  else
                    response.write """#FFCC00"">"
                    response.write "&nbsp;" 
                  end if 
     
                  rsSite_NewFlag.close
                  set rsSite_NewFlag = nothing
     
                  response.write "</TD>"

                  response.write "</TR>"
            
                  rsSite_Aux.MoveNext

                loop
    
                response.write "</TABLE>"
                response.write "</TD>"
                response.write "</TR>"
    
              end if
    
              rsSite_Aux.close
              set rsSite_Aux = nothing

              ' Auxiliary Fields
  
              SQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.Site_ID=" & CInt(Site_ID) & " AND Auxiliary.Enabled=" & CInt(True) & " ORDER BY Auxiliary.Order_Num"
              Set rsAuxiliary = Server.CreateObject("ADODB.Recordset")
              rsAuxiliary.Open SQL, conn, 3, 3
              
              if not rsAuxiliary.EOF then
              
         				response.write "<TR>"
                response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
                response.write "<INPUT TYPE=""Hidden"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Required"" VALUE=""" & rsAuxiliary("Required") & """>"
                response.write Translate("Other Information",Login_Language,conn)
                response.write "  </TD>"
      	        response.write "</TR>"
                
                do while not rsAuxiliary.EOF

                  if rsAuxiliary("Enabled") = CInt(True) then

          				  response.write "<TR>" & vbCrLf
                  	response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=MIDDLE CLASS=Medium>"
                    response.write Translate(rsAuxiliary("Description"),Login_Language,conn) & ":"
                    response.write "</TD>" & vbCrLf
                    response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                    if CInt(rsAuxiliary("Required")) = True then
                      Aux_Required(rsAuxiliary("Order_Num")) = True
                      Aux_Method(rsAuxiliary("Order_Num"))   = rsAuxiliary("Input_Method")
                      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Description"" VALUE=""" & rsAuxiliary("Description") & """>"
                      response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10 ALIGN=ABSMIDDLE>"
                    else
                      response.write "&nbsp;"
                    end if  
                    response.write "</TD>" & vbCrLf

       	            response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf

                    Aux_Selection     = Split(rsAuxiliary("Radio_Text"),",")
                    Aux_Selection_Max = Ubound(Aux_Selection)

                    Select Case rsAuxiliary("Input_Method")
                      Case 0      ' Text
                        response.write "<INPUT TYPE=""Text"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ SIZE=""50"" MAXLENGTH=""50"" VALUE=""" & rs("Aux_" & Trim(rsAuxiliary("Order_Num"))) & """ CLASS=Medium>" & vbCrLf
                      Case 1      ' Drop-Down
                        response.write "<SELECT NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium>" & vbCrLf
                        response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                        for i = 0 to Aux_Selection_Max
                          response.write "<OPTION" 
                          if rs("Aux_" & Trim(rsAuxiliary("Order_Num"))) = Trim(Aux_Selection(i)) then
                            response.write " SELECTED"
                          end if                            
                          response.write " CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "</OPTION>" & vbCrLf
                        next
                        response.write "</SELECT>" & vbCrLf
                      Case 2      ' Radio
                        for i = 0 to Aux_Selection_Max
                          response.write "<INPUT"
                          if rs("Aux_" & Trim(rsAuxiliary("Order_Num"))) = Trim(Aux_Selection(i)) then
                            response.write " CHECKED"
                          end if                                                      
                          response.write " TYPE=RADIO NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>&nbsp;" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "&nbsp;&nbsp;" & vbCrLf
                        next
                    end select
                        
                    response.write "</TD>" & vbCrLf
                    response.write "</TR>" & vbCrLf
                  
                  end if
                  
                  rsAuxiliary.MoveNext
                  
                loop
                
                response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>" & vbCrLf
                
              end if
              
              rsAuxiliary.Close
              set rsAuxiliary = nothing
                
              ' Comment
  
      				response.write "<TR>" & vbCrLf
       				response.write "	<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & Translate("Comment",Login_Language,conn) & ":</TD>" & vbCrLf
       				response.write "	<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>" & vbCrLf
       				response.write "  <TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "    <TEXTAREA NAME=""Comment"" COLS=52 ROWS=6  MAXLENGTH=""255"" CLASS=Medium>" & rs("Comment") & "</TEXTAREA>"
              response.write "  </TD>" & vbCrLf
              response.write "</TR>" & vbCrLf
  
              ' Instant Message
  
      				response.write "<TR>" & vbCrLf
       				response.write "	<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & Translate("Instant Message",Login_Language,conn) & ":<BR>"
  
              SQLMessage = "SELECT COUNT(*) AS Count FROM Messages WHERE NTLogin='" & rs("NTLogin") & "'"
              Set rsMessage = Server.CreateObject("ADODB.Recordset")
              rsMessage.Open SQLMessage, conn, 3, 3
              
              if not rsMessage.EOF then
                if rsMessage("Count") > 0 then
                  response.write "<SPAN CLASS=Small>" & Translate("View Message Queue",Login_Language,conn) & "</SPAN>&nbsp;&nbsp"
                  response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""Button"" NAME=""Message_Queue"" VALUE=""" & ": [ " & Trim(CStr(rsMessage("Count"))) & " ] "" "
                  response.write "LANGUAGE='JavaScript' "
                  response.write "onclick=""Message_Window = window.open('/sw-administrator/Instant_Message_Queue.asp?NTLogin=" & rs("NTLogin") & "&Language=" & Login_Language &  "','Message_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no'); Message_Window.moveTo(400,100); Message_Window.window.focus(); return false;"">"
                end if  
              end if
              rsMessage.close
              set rsMessage  = nothing
              set sqlMessage = nothing  
                  
              response.write "</TD>" & vbCrLf
       				response.write "	<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>" & vbCrLf
       				response.write "  <TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "    <TEXTAREA NAME=""Message"" COLS=52 ROWS=6  MAXLENGTH=""255"" CLASS=Medium>" & "</TEXTAREA>"
              response.write "  </TD>" & vbCrLf
              response.write "</TR>" & vbCrLf

              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>" & vbCrLf
              %>             

               <!-- EMail Update -->
      
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Send Update EMail to",Login_Language,conn)%>:
                </TD>
              	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                                
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                  
                   <INPUT TYPE="Checkbox" NAME="send_email_user"<%if CInt(rs("NewFlag")) = CInt(True) or CInt(rs("NewFlag")) = -2 then response.write " CHECKED"%> CLASS=Medium><%=Translate("User",Login_Language,conn)%><BR>
                   <INPUT TYPE="Checkbox" NAME="send_email_fcm" CLASS=Medium><%=Translate("Account Manager",Login_Language,conn)%><BR>
                   <INPUT TYPE="Checkbox" NAME="send_email_admin" Class CLASS=Medium><%=Translate("Site Administrator",Login_Language,conn)%>
                </TD>
              </TR>
              
              <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                    
              
              <!-- Navigation Buttons -->
      
              <TR>
              <TD COLSPAN=3 CLASS=Medium>
                <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666">
                  <TR>
                    <TD ALIGN=CENTER WIDTH="38%" CLASS=Medium>
                    <%
                      response.write "<INPUT TYPE=""BUTTON"" Value=""" & Translate("Main Menu",Login_Language,conn) & """ Language=""JavaScript"" onclick=""location.href='" & BackURL & "'"" CLASS=NavLeftHighlight1 onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
                    %>  
                    </TD>
                    <TD ALIGN=CENTER WIDTH="2%" CLASS=Medium>
                    &nbsp;
                    </TD>             
                    <TD ALIGN=LEFT WIDTH="30%" CLASS=Medium>&nbsp;
                      <INPUT TYPE="Submit" NAME="Update" VALUE=" <%=Translate("Update",Login_Language,conn)%> " CLASS=NavLeftHighlight1 onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'">
                    </TD>
                    <TD ALIGN=CENTER WIDTH="30%" CLASS=Medium>
                      <INPUT TYPE="Submit" NAME="Delete" VALUE=" <%=Translate("Delete",Login_Language,conn)%> " CLASS=NavLeftHighlight1 onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'">
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
      <BR><BR>      
      
  <%  else 
      response.write "No Records Found<BR>"
  
    end if
         
    rs.close
    set rs=nothing
    
  end if
  

' --------------------------------------------------------------------------------------  
' Add Record
' --------------------------------------------------------------------------------------

  if lcase(Account_ID) = "add" then
  
    ' Check Core Table on WWW to determine if user information already exists to pre-populate form

    ErrorMessage      = ""

    CoreMax = 23
    DIM Core(23)

    for i = 0 to CoreMax
      Core(i) = ""
    next
  
    if not isblank(request("Core_Email")) then
    
    ' First check SiteWide DB to determine if email already exists, if true then deny new registration for same Site_ID code.
    
      SQL = "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & CInt(Site_ID) & "AND UserData.EMail='" & request("Core_Email") & "'"
      Set rsUser = Server.CreateObject("ADODB.Recordset")
      rsUser.Open SQL, conn, 3, 3
    
      if NOT rsUser.EOF then      ' Account Found - Do not allow new registration

        if (CInt(rsUser("NewFlag")) = CInt(True) or CInt(rsUser("NewFlag")) = -2) and rsUser("Site_ID") = CInt(Site_ID) then
          ErrorMessage = "<FONT COLOR=""Red""><UL><LI>" & Translate("This New Account Request is currently in your queue waiting review / approval / deletion, under the name of",Login_Language,conn) & ": <B>" & rsUser("FirstName") & " " & rsUser("LastName") & ", " & rsUser("Company") & "</B>, " & Translate("Record ID",Login_Language,conn) & ": " & rsUser("ID") & ".</LI></UL></FONT>"
        elseif rsUser("Site_ID") = CInt(Site_ID) then
          ErrorMessage = "<FONT COLOR=""Red""><UL><LI>" & Translate("This Account is currently active. You will not be allowed to add a second occurrence of this account.  You can review / update / delete, this account under the name of",Login_Language,conn) & ": <B><FONT COLOR=""Black"">" & rsUser("FirstName") & " " & rsUser("LastName") & ", " & rsUser("Company") & "</B>, " & Translate("Record ID",Login_Language,conn) & ": " & rsUser("ID") & "</FONT>.</LI></UL></FONT>"
        end if
        
      end if

      rsUser.close
      set rsUser = nothing
     
      Core(16) = Request("Core_Email")

      if isblank(ErrorMessage) then
        
        Call Connect_FormDatabase   

        SQL = "SELECT * FROM tbl_Core WHERE tbl_Core.Email='" & request("Core_Email") & "'"
        Set rsCore = dbConnFormData.Execute(SQL)	
        
        ' Get core table data
        
        if not rsCore.EOF then
          Core(0) = rsCore.Fields("CoreID").Value
          Core(1) = rsCore.Fields("Prefix").Value
          Core(2) = rsCore.Fields("FirstName_MI").Value
          if instr(1,Core(2),",") > 0 then
            Core(3) = trim(mid(Core(2),instr(1,Core(2)," ") + 1))
            Core(2) = trim(mid(Core(2),1,instr(1,Core(2)," ") - 1))
          else
            Core(3) = ""
          end if  
          Core(4) = rsCore.Fields("LastName").Value
          Core(5) = rsCore.Fields("Suffix").Value
          Core(6) = rsCore.Fields("Title").Value
          Core(7) = rsCore.Fields("MailStop").Value
          Core(8) = rsCore.Fields("Company").Value
          Core(9) = rsCore.Fields("Address1").Value
          Core(10) = rsCore.Fields("Address2").Value
          Core(11) = rsCore.Fields("City").Value
          Core(12) = rsCore.Fields("State_Province").Value
          Core(13) = rsCore.Fields("State_Other").Value
          Core(14) = rsCore.Fields("Zip").Value
          Core(15) = rsCore.Fields("Country").Value
          if rsCore.Fields("Email").value = "" then
            Core(16) = request("Core_EMail")
          else  
            Core(16) = rsCore.Fields("Email").Value
          end if  
          Core(17) = rsCore.Fields("Phone").Value
          Core(18) = rsCore.Fields("Extension").Value
          Core(19) = rsCore.Fields("Fax").Value
          Core(20) = rsCore.Fields("JobFunction").Value
          Core(21) = rsCore.Fields("NativeLanguage")

          ErrorMessage = "<UL><LI>" & Translate("The Requestor&acute;s profile information has been found in the database by searching by the email address that you have provided. Please ensure that this information is accurate and complete on the account form below.",Login_Language,conn) & "</LI></UL>"
          
        else

          Core(16) = request("Core_Email")
          ErrorMessage = "<FONT COLOR=""Red""><UL><LI>" & Translate("The Requestor&acute;s profile information was not found in the database by searching by the email address",Login_Language,conn) & ": <FONT COLOR=""Black""><B>" & core(16) & "</B></FONT> " & Translate("that you have provided. Please complete the following account form below or ",Login_Language,conn) & " <A HREF=""account_edit.asp?Site_ID=" & Site_ID & "&ID=account_edit&Account_ID=add"">" & Translate("click here",Login_Language,conn) & "</A> " & Translate("to try another search.",Login_Language,conn) & "</LI></UL></FONT>"

        end if

        rsCore.close
        set rsCore = nothing
          
        Call Disconnect_FormDatabase

      end if
      
    end if
         
   ' With Email Check Core

   if isblank(request("Core_EMail")) then %>

    <FORM NAME="Check_Email" ACTION="account_edit.asp" METHOD="<%=Post_Method%>" onKeyUp="highlight(event)" onClick="highlight(event)">
    <INPUT TYPE="Hidden" NAME="ID" VALUE="edit_account">
    <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
    <INPUT TYPE="Hidden" NAME="Account_ID" VALUE="add">
    <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
    <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
    <INPUT TYPE="Hidden" NAME="Reg_Request_Date" VALUE="<%=Now%>">

    <%Call Table_Begin%>                  
    <TABLE WIDTH="100%" BORDER=0 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
    	<TR>
     		<TD WIDTH="100%" BGCOLOR="#EEEEEE">
   	  		<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>                   
   		 	  	<TR>
            	<TD BGCOLOR="<%=Contrast%>" VALIGN=TOP WIDTH="50%" CLASS=Medium>
                <%=Translate("<B>Search by User&acute;s EMail Address</B><BR>to attempt to pre-populate this form",Login_Language,conn)%>:
              </TD>
    	        <TD BGCOLOR="White" ALIGN=LEFT WIDTH="50%" VALIGN=MIDDLE CLASS=Medium>                
                <INPUT TYPE="Text" NAME="Core_Email" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>&nbsp;&nbsp;<INPUT TYPE="Submit" VALUE=" Search " CLASS=NavLeftHighlight1>
              </TD>
            </TR>
          </TABLE>
        </TD>
      </TR>
    </TABLE>
    </FORM>            

    <%
    end if
    
    if not isblank("ErrorMessage") then
      response.write ErrorMessage
      ErrorMessage = ""
    end if

    FormName = "NTAccount"
    %>
 
 
    <FORM NAME="<%=FormName%>" ACTION="account_admin.asp" METHOD="<%=Post_Method%>" onsubmit="return CheckRequiredFields(this.form);" onKeyUp="highlight(event)"  Language="JavaScript" onClick="highlight(event)">
    <INPUT TYPE="Hidden" NAME="ID" VALUE="add">
    <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
    <INPUT TYPE="Hidden" NAME="Site_Code" VALUE="<%=Site_Code%>">
    <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
    <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
    <INPUT Type="Hidden" NAME="NewFlag" Value="on">
    <INPUT TYPE="Hidden" NAME="ChangeID" VALUE="<%=Admin_ID%>">                                  
    <INPUT TYPE="Hidden" NAME="ChangeDate" Value="<%=Date%>">
	  <INPUT TYPE="Hidden" NAME="SubGroups">

    <TABLE WIDTH="100%" BORDER=1 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
  	<TR>
  		<TD WIDTH="100%" BGCOLOR="#EEEEEE" CLASS=Medium>
  			<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>
        
          <!-- Header -->
		  		<TR>
          	<TD WIDTH="40%" COLSPAN=2 CLASS=NavLeftSelected1>
              <%=Translate("Description",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallBoldGold>(<%=Translate("Note",Login_Language,conn)%>:&nbsp;<IMG SRC="/images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> = <%=Translate("Required Information or use N/A",Login_Language,conn)%>)</SPAN>
            </TD>
  	        <TD WIDTH="60%" ALIGN=LEFT CLASS=NavLeftSelected1>
              <%=Translate("User&acute;s Account Profile Information",Login_Language,conn)%>
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=2 CLASS=MediumBold>
              <%=Translate("Account Information",Login_Language,conn)%>
            </TD>
            <TD BGCOLOR="Silver" VALIGN=TOP CLASS=Medium NOWRAP>
              <A HREF="/SW-Help/AA-UG.pdf"><IMG SRC="/images/help_button.gif" BORDER=0 ALIGN=RIGHT VALIGN=TOP ALT="Account Managers - User Guide"></A>
                <%
                  response.write "<INPUT TYPE=""BUTTON"" Value=""" & Translate("Main Menu",Login_Language,conn) & """ Language=""JavaScript"" onclick=""location.href='" & BackURL & "'"" CLASS=NavLeftHighlight1>"
                %>                                     
            </TD>
          </TR>        
                            
          <!-- Account ID and Account Manager Status-->
          
  				<TR>
          	<TD BGCOLOR="#EEEEEE" WIDTH="38%" CLASS=Medium>
              <%=Translate("Account ID Number",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER WIDTH="2%" CLASS=Medium>
              &nbsp;
            </TD>                                                
            <TD BGCOLOR="White" WIDTH="60%" CLASS=MediumRed><B><%=Translate("Add",Login_Language,conn)%></B></TD>
          </TR>
        
          <%
          ' Account Type
          
          SQL = "SELECT UserType.* FROM UserType WHERE UserType.Site_ID=" & CInt(Site_ID) & " ORDER BY UserType.Order_Num"
          Set rsUserType = Server.CreateObject("ADODB.Recordset")
          rsUserType.Open SQL, conn, 3, 3                    
          
          Account_Type = False
          if not rsUserType.EOF then
            Account_Type = True
    				response.write "<TR>"
          	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Fluke Customer Type",Login_Language,conn) & ":</TD>"
          	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
            response.write "    <IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
            response.write "</TD>"
  	        response.write "<TD BGCOLOR=""White"">"
            response.write "  <INPUT TYPE=""HIDDEN"" NAME=""Type_Code_Required"" VALUE=""on"">"            
            response.write "  <SELECT NAME=""Type_Code"" CLASS=MEDIUM>"
            response.write "    <OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
            
            do while not rsUserType.EOF
              if (CInt(rsUserType("Type_Code")) < 99 and Admin_Access < 9) or Admin_Access = 9 then
                response.write "    <OPTION CLASS=Medium VALUE=""" & rsUserType("Type_Code") & """>" & Translate(rsUserType("Type_Description"),Login_Language,conn) & "</OPTION>"
              end if
              rsUserType.MoveNext
            loop
            response.write "</SELECT>"  
            response.write "&nbsp;"
            response.write "</TD>"
            response.write "</TR>"
            
          end if
          
          rsUserType.close
          set rsUserType = nothing
          
          ' Account Manager or Account Manager Status
          
          response.write "<TR>"
          response.write "  <TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Account Manager",Login_Language,conn) & ":</TD>"
          response.write "	<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
          response.write "    <IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
          response.write "  </TD>"
  	      response.write "  <TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
          response.write "    <TABLE WIDTH=""100%"">"
          response.write "      <TR>"
          response.write "        <TD WIDTH=""50%"" CLASS=Medium>"
            
          response.write "          <SELECT NAME=""Fcm_ID"" CLASS=Medium>"
          response.write "          <OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"

          SQL = "SELECT UserData.* "
          SQL = SQL & "FROM UserData "
          SQL = SQL & "WHERE (((UserData.Site_ID)=0 Or (UserData.Site_ID)=" & CInt(Site_ID) & ") AND ((UserData.Fcm)=" & CInt(True) & "))"
          SQL = SQL & "ORDER BY UserData.LastName"

          Set rsManager = Server.CreateObject("ADODB.Recordset")
          rsManager.Open SQL, conn, 3, 3                    

          Do while not rsManager.EOF
            response.write "<OPTION"
            response.write " CLASS=Region" & Trim(CStr(rsManager("Region")))
         	  response.write " VALUE=""" & rsManager("ID") & """>" & rsManager("LastName") & ", " & rsManager("FirstName") & "</OPTION>"
        	  rsManager.MoveNext 
          loop
             
          rsManager.close
          set rsManager=nothing
          
          response.write "            <OPTION VALUE=""0"">" & Translate("N / A",Login_Language,conn) & "</OPTION>"
          response.write "          </SELECT>"
          response.write "        </TD>"
          response.write "        <TD WIDTH=""50%"" BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium><B>" & Translate("or",Login_Language,conn) & "</B> " & Translate("Account Manager",Login_Language,conn) & " ?&nbsp;&nbsp;"
          response.write "          <INPUT TYPE=""Checkbox"" NAME=""Fcm"" CLASS=Medium>"
          response.write "        </TD>"
          response.write "      </TR>"
          response.write "    </TABLE>"
              
          response.write "  </TD>"
          response.write "</TR>"

          ' Fluke Customer Number
          
    		  with response
            .write "<TR>"
          	.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
            .write Translate("Fluke Customer Number",Login_Language,conn) & " / "
	      		.write Translate("Business System",Login_Language,conn)
      			.write ":</TD>" & vbcrlf
          	.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
            .write "&nbsp;"
            .write "</TD>"
            .write "<TD BGCOLOR=""White"" CLASS=Medium>"
            .write "<INPUT TYPE=""Text"" NAME=""Fluke_ID"" SIZE=""15"" MAXLENGTH=""50"" CLASS=""MEDIUM"">"
      			.write "&nbsp;&nbsp;&nbsp;<SELECT Name=""Business_system"" CLASS=""MEDIUM"">" & vbcrlf
      			.write "    <OPTION CLASS=Medium VALUE="""">Select from List</OPTION>" & vbcrlf
      			.write "    <OPTION CLASS=MEDIUM VALUE=""ORA"">Oracle</OPTION>" & vbcrlf
    				.write "    <OPTION CLASS=MEDIUM VALUE=""MFG"">Mfg/Pro</OPTION>" & vbcrlf
    				.write "    <OPTION CLASS=MEDIUM VALUE=""DGU"">DGUX</OPTION>" & vbcrlf
      			.write " </select>" & vbcrlf
      			.write "</TD>"
            .write "</TR>"
          end with
          %>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Contact Information",Login_Language,conn)%>
            </TD>
          </TR>        

           <!-- Name -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
              <%=Translate("Name",Login_Language,conn)%>:&nbsp;&nbsp;<SPAN CLASS=Smallest><IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("First",Login_Language,conn)%>]&nbsp;&nbsp;[<%=Translate("Middle",Login_Language,conn)%>]&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("Surname",Login_Language,conn)%>],&nbsp;&nbsp;[<%=Translate("Suffix",Login_Language,conn)%>]</SPAN>
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
              <INPUT TYPE="Text" NAME="FirstName" SIZE="10" MAXLENGTH="50" VALUE="<%=Core(2)%>" CLASS=Medium>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" NAME="MiddleName" SIZE="6" MAXLENGTH="50" VALUE="<%=Core(3)%>" CLASS=Medium>&nbsp;&nbsp;&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><INPUT TYPE="Text" NAME="LastName" SIZE="11" MAXLENGTH="50" VALUE="<%=Core(4)%>" CLASS=Medium> <B>,</B>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" NAME="Suffix" SIZE="2" MAXLENGTH="50" VALUE="<%=Core(5)%>" CLASS=Medium>
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
              <INPUT TYPE="Text" NAME="Initials" SIZE="10" MAXLENGTH="10" VALUE="" CLASS=Medium>
            </TD>
          </TR>

          <!-- Gender -->
          
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
              <%=Translate("Gender",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
              <% 
              sValue = Core(1)
              %>
              <SELECT CLASS=Medium NAME="Gender" CLASS=Medium>
                <OPTION CLASS=Medium VALUE=""><%=Translate("Select",Login_Language,conn)%></OPTION>
                <OPTION CLASS=Region2 VALUE="0"<% If sValue = "Mr" Then Response.Write " SELECTED" %>><%=Translate("Male",Login_Language,conn)%></OPTION>
                <OPTION CLASS=Region3 VALUE="1"<% If sValue = "Ms" or sValue = "Miss" or sValue="Mrs" Then Response.Write " SELECTED" %>><%=Translate("Female",Login_Language,conn)%></OPTION>
              </SELECT>
            </TD>
          </TR>

           <!-- Job Title -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Job Title",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Job_Title" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(20)%>" CLASS=Medium>
            </TD>
          </TR>

           <!-- Business Phone -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Phone",Login_Language,conn)%> (<%=Translate("Direct",Login_Language,conn)%>):
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=Core(17)%>">&nbsp;&nbsp;<%=Translate("Extension",Login_Language,conn)%>: <INPUT TYPE="Text" NAME="Business_Phone_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=Core(18)%>" CLASS=Medium>
            </TD>
          </TR>

           <!-- Mobile Phone -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Mobile Phone",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Mobile_Phone" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

           <!-- Pager -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Pager",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Pager" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

           <!-- Email -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("EMail",Login_Language,conn)%> <FONT CLASS=Small>(<%=Translate("Direct",Login_Language,conn)%>)</FONT>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Email" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(16)%>" CLASS=Medium>
            </TD>
          </TR>
          
          <!-- EMail Method -->
            
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("EMail Format",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <SELECT Name="EMail_Method" CLASS=Medium>
                <OPTION CLASS=Medium VALUE=""><%=Translate("Select from List",Login_Language,conn)%></OPTION>
                <OPTION CLASS=Medium Value=""></OPTION>
                <OPTION Class=Region3 VALUE="0"><%=Translate("Plain Text without Graphics",Login_Language,conn)%></OPTION>
                <OPTION Class=Region2 VALUE="1"><%=Translate("Rich Text with Graphics",Login_Language,conn)%></OPTION>
              </SELECT>
            </TD>
          </TR>

           <!-- Connection Speed -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
              <% response.write Translate("Internet Connection Speed",Login_Language,conn) & ":"%>
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <%
              SQL = "SELECT Download_Time.* FROM Download_Time WHERE Download_Time.Enabled=" & CInt(True)
              SQL = SQL & " ORDER BY DownLoad_Time.bps, DownLoad_Time.Description"
              Set rsDownload = Server.CreateObject("ADODB.Recordset")
              rsDownload.Open SQL, conn, 3, 3
              response.write "<SELECT CLASS=Medium NAME=""Connection_Speed"">" & vbCrLf
              response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
              response.write "<OPTION CLASS=Medium Value="""">" & "</OPTION>" & vbCrLf
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

           <!-- Subscription Service -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Subscription Service",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Checkbox" NAME="Subscription" CHECKED CLASS=Medium>
            </TD>
          </TR>
          
          <!-- Preferred Language -->
          
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Preferred Language",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                            
              <SELECT Name="Account_Language" CLASS=Medium>    
              <%
              SQL = "SELECT * FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
              Set rsLanguage = Server.CreateObject("ADODB.Recordset")
              rsLanguage.Open SQL, conn, 3, 3
                                    
              Do while not rsLanguage.EOF
                if not isblank(Core(21)) then
                  if rsLanguage("Code") = Core(21) then
                    response.write "<OPTION SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"
                  else
                 	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"
                  end if
                elseif rsLanguage("Code") = "eng" then
               	  response.write "<OPTION SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
                else
               	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
                end if
            	  rsLanguage.MoveNext 
              loop
              
              rsLanguage.close
              set rsLanguage=nothing
              %>
              </SELECT>              
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Account Credentials",Login_Language,conn)%>
            </TD>
          </TR>        

          <!-- Expiration Date -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <B><%=Translate("Verify",Login_Language,conn)%></B> <%=Translate("Expiration Date",Login_Language,conn)%><FONT CLASS=Smallest> (<%=Translate("mm/dd/yyyy or blank for Never",Login_Language,conn)%>):</FONT>
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="ExpirationDate" SIZE="10" MAXLENGTH="50" VALUE="12/31/<%=DatePart("yyyy",Date)%>" CLASS=Medium>&nbsp;&nbsp;
              <A HREF="javascript:void()" LANGUAGE="JavaScript" onClick="window.dateField = document.<%=FormName%>.ExpirationDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false"><IMG SRC="/images/calendar/calendar_icon.gif" BORDER=0 HEIGHT="21"ALIGN=TOP></A>
            </TD>
          </TR>
           <!-- NT Login -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>              
              <%=Translate("Logon User Name",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>              
              <INPUT TYPE="Text" NAME="NTLogin" SIZE="50" MAXLENGTH="30" VALUE="" CLASS=Medium>
            </TD>
          </TR>
  
           <!-- NT Password -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>              
              <%=Translate("Logon Password",Login_Language,conn)%>:                                
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Password" SIZE="50" MAXLENGTH="30" VALUE="" CLASS=Medium>
              <INPUT TYPE="HIDDEN" NAME="Password_Change" Value="">              
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Company Information",Login_Language,conn)%>
            </TD>
          </TR>        

           <!-- Company -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Company Name",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Company" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(8)%>" CLASS=Medium>
            </TD>
          </TR>

           <!-- Company Website -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Company Website Address",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Company_Website" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium>
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Office Information",Login_Language,conn)%>
            </TD>
          </TR>        

          <!-- Business Address -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Address",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" VALIGN=TOP>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(9)%>" CLASS=Medium><BR>
              <INPUT TYPE="Text" NAME="Business_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(10)%>" CLASS=Medium>                
            </TD>
          </TR>

          <!-- Business Mail Stop -->
                
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Mail Stop",Login_Language,conn) & " / " & Translate("Building Number",Login_Langugage,conn)%> :
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(7)%>" CLASS=Medium>
            </TD>
          </TR>

          <!-- Business City -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("City",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_City" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(11)%>" CLASS=Medium>
            </TD>
          </TR>

          <!-- Business State -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <% SValue = Core(12)%>
              <SELECT NAME="Business_State" CLASS=Medium>

               <%
               response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
               %>  
                
                <!--#include virtual="/include/core_states.inc"-->

              </SELECT>

            </TD>
          </TR>

           <!-- Business State Other-->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%="<B>" & Translate("or",Login_Language,conn) & "</B> " & Translate("Other State, Province or Local",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(13)%>" CLASS=Medium>
            </TD>
          </TR>

           <!-- Business Postal Code -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Postal Code",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=Core(14)%>" CLASS=Medium>
            </TD>
          </TR>

           <!-- Business Country -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Country",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <%
              Call Connect_FormDatabase              
              Call DisplayCountryList("Business_Country",Core(15),Translate("Select from List",Login_Language,conn),"Medium")
              Call Disconnect_FormDatabase
              %>
            </TD>
          </TR>

          <!-- Business Phone 2-->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Phone",Login_Language,conn)%> (<%=Translate("General Office",Login_Language,conn)%>):
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_Phone_2" SIZE="28" MAXLENGTH="50" VALUE="">&nbsp;&nbsp;<%=Translate("Extension",Login_Language,conn)%>: <INPUT TYPE="Text" NAME="Business_Phone_2_Extension" SIZE="10" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

           <!-- Email 2 -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("EMail",Login_Language,conn)%> (<%=Translate("General Office",Login_Language,conn)%>):
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Email_2" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

          <!-- Business Fax -->
      
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Fax",Login_Language,conn)%> (<%=Translate("General Office",Login_Language,conn)%>):
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Business_Fax" SIZE="28" MAXLENGTH="50" VALUE="">
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Postal Information",Login_Language,conn)%>
            </TD>
          </TR>        

  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Postal Box Number",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Postal_Address" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium><BR>
            </TD>
          </TR>

          <!-- Postal Same -->

  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=SmallRed>
              <%=Translate("If the remainder of the Postal Address is the same as Office Address",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=SmallRed>
              <INPUT TYPE="Checkbox" NAME="PostalSame" CLASS=Medium>&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
            </TD>
          </TR>
  
           <!-- Postal City -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("City",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Postal_City" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

           <!-- Postal State -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>

              <% SValue = "" %>              
              <SELECT NAME="Postal_State" CLASS=Medium>

               <%
                response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
               %>  

                <!--#include virtual="/include/core_states.inc"-->

              </SELECT>
            </TD>
          </TR>

           <!-- Postal State Other-->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%="<B>" & Translate("or",Login_Language,conn) & "</B> " & Translate("Other State, Province or Local",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Postal_State_Other" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

           <!-- Postal Postal Code -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Postal Code",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Postal_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

           <!-- Postal Country -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Country",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <%
              Call Connect_FormDatabase              
              Call DisplayCountryList("Postal_Country",Core(15),Translate("Select from List",Login_Language,conn),"Medium")
              Call Disconnect_FormDatabase              
              %>
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Shipping Information",Login_Language,conn)%>
            </TD>
          </TR>        
                        
          <!-- Shipping Same -->

  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=SmallRed>
              <%=Translate("If Shipping Address is the same as Office Address",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=SmallRed>
              <INPUT TYPE="Checkbox" NAME="ShippingSame" CLASS=Medium>&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
            </TD>
          </TR>
  
          <!-- Shipping Address -->
           
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Address",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Shipping_Address" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium><BR>
              <INPUT TYPE="Text" NAME="Shipping_Address_2" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>                
            </TD>
          </TR>

           <!-- Shipping Mail Stop -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Mail Stop",Login_Language,conn) & " / " & Translate("Building Number",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Shipping_MailStop" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>              

          <!-- Shipping City -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("City",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Shipping_City" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

          <!-- Shipping State -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("USA State or Canadian Province",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>

              <% SValue = "" %>              
              <SELECT NAME="Shipping_State" CLASS=Medium>

               <%
                response.write "<OPTION VALUE="""" CLASS=Medium>" & Translate("Select from List",Login_Language,conn) & "</OPTION>"
               %>  

                <!--#include virtual="/include/core_states.inc"-->

              </SELECT>
            </TD>
          </TR>

          <!-- Shipping State Other-->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%="<B>" & Translate("or",Login_Language,conn) & "</B> " & Translate("Other State, Province or Local",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Shipping_State_Other" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

          <!-- Shipping Postal Code -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Postal Code",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT TYPE="Text" NAME="Shipping_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="" CLASS=Medium>
            </TD>
          </TR>

          <!-- Shipping Country -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Country",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <%
              Call Connect_FormDatabase              
              Call DisplayCountryList("Shipping_Country",Core(15),Translate("Select from List",Login_Language,conn),"Medium")
              Call Disconnect_FormDatabase              
              %>
            </TD>
          </TR>

				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold>
              <%=Translate("Site Administrative Permissions",Login_Language,conn)%>
            </TD>
          </TR>        
                    
          <!-- Groups Selection -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              &nbsp;
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" CLASS=Medium>               
        
          <%
          SQL = "SELECT SubGroups.* "
          SQL = SQL & "FROM SubGroups "
          SQL = SQL & "WHERE (((SubGroups.Site_ID)=" & Site_ID & ") AND ((SubGroups.Enabled)=" & CInt(True) & ")) "
          SQL = SQL & "ORDER BY SubGroups.Order_Num"

          Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
          rsSubGroups.Open SQL, conn, 3, 3
          
          if not rsSubGroups.EOF then
          
            response.write "<TABLE WIDTH=""100%"">"

            if Admin_Access = 9 then
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""domain"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Domain Administrator",Login_Language,conn) & "</TD></TR>"
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""administrator"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Site Administrator",Login_Language,conn) & "</TD></TR>"
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"                  
            end if
            
            if Admin_Access >= 8 then
              response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Select <U>only one</U> Administrative Permission from this group, if applicable.",Login_Language,conn) & "</TD></TR>"
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"                  
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""account"" CLASS=Medium></TD><TD CLASS=MediumRed>"
              response.write Translate("Account Administrator",Login_Language,conn) & "&nbsp;&nbsp;&nbsp;&nbsp;<FONT CLASS=Medium>"
              response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=4> " & Translate("ALL",Login_Language,conn) & "&nbsp;&nbsp;"                  
              response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=1> " & Translate("USA",Login_Language,conn) & "&nbsp;&nbsp;"
              response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=2> " & Translate("EUR",Login_Language,conn) & "&nbsp;&nbsp;"
              response.write "<INPUT TYPE=""RADIO"" NAME=""Account_Region"" VALUE=3> " & Translate("INT",Login_Language,conn) & "</FONT>"
              
              response.write "</TD></TR>"
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""content"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Content Administrator",Login_Language,conn)
              
              response.write "&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RTE_ENABLED"">"
              response.write "&nbsp;<SPAN CLASS=Medium>" & Translate("Rich Text Editor",Login_Language,conn)
              response.write "</SPAN></TD></TR>" & vbCrLf
            end if

            ' Submitter
            if Admin_Access >= 6 then                
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""submitter"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Content Submitter",Login_Language,conn) & "</TD></TR>"
            end if

            ' ImageStore Administrator
            if Admin_Access >=8 and 1=2 then
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"
              response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("ImageStore",Login_Language,conn) & "</TD></TR>"
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""imstadm"" CLASS=Medium LANGUAGE=""JavaScript"" ONCLICK=""alert('You must email User Account Information plus Username and Password to Gerda Meijer to set up an ImageStore Administration account.');""></TD><TD CLASS=MediumRed>" & Translate("ImageStore Administrator",Login_Language,conn) & "</TD></TR>" & vbCrLf
            end if

            if Admin_Access >=6 then
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"
            end if  
                                                                                                  
            ' Forums
            response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Forums",Login_Language,conn) & "</TD></TR>"
            response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"
            if Admin_Access >=6 then  
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""forum"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Forum or Discussion Group Moderator",Login_Language,conn) & "</TD></TR>"
            end if
     
            if Admin_Access >=6 then  
              response.write "</TABLE>"
      				response.write "<TR>"
              response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
              response.write Translate("Order Inquiry Permissions",Login_Language,conn)
              response.write "  </TD>"
    	        response.write "</TR>"

      				response.write "<TR>"
              response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              response.write "    &nbsp;"
              response.write "  </TD>"
              response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
              response.write "    &nbsp;"
              response.write "  </TD>"
    	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
              response.write "<TABLE WIDTH=""100%"">"
              response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Order Inquiry",Login_Language,conn) & " (" & Translate("Select only one",Login_Language,conn) & ")</TD></TR>"
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"

              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""ordad"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Order Inquiry Administrator",Login_Language,conn) & " (" & Translate("Fluke Internal Use Only",Login_Language,conn) & ")</TD></TR>"
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""order"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Order Inquiry with Search Capability",Login_Language,conn) & "</TD></TR>"
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""literature"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Literature Order Administrator",Login_Language,conn) & "</TD></TR>"              
            end if

            if Admin_Access >=6 then
              response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"
            end if  

            response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Branch Location Administrator",Login_Language,conn) & "</TD></TR>"
            response.write "<TR><TD COLSPAN=2 CLASS=Medium HEIGHT=8></TD></TR>"

            ' Branch Location Administrator
            if Admin_Access >=6 and Account_Type = True then
              response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""branch"" CLASS=Medium></TD><TD CLASS=MediumRed>" & Translate("Branch Location Administrator",Login_Language,conn) & "</TD></TR>"
            end if

            response.write "</TABLE>"
    				response.write "<TR>"
            response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
            response.write Translate("Group Affiliations",Login_Language,conn)
            response.write "  </TD>"
  	        response.write "</TR>"

    				response.write "<TR>"
            response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
            response.write "    &nbsp;"
            response.write "  </TD>"
            response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
            response.write "    &nbsp;"
            response.write "  </TD>"
  	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
            response.write "<TABLE WIDTH=""100%"">"

            response.write "<TR><TD WIDTH=20 BGCOLOR=""Gray"" CLASS=Medium></TD><TD BGCOLOR=""#EEEEEE"" CLASS=Medium>" & Translate("Select one or more Group Affiliations.",Login_Language,conn) & "</TD></TR>"
                            
            Do while not rsSubGroups.EOF

              if RegionValue <> Mid(rsSubGroups("Code"),1,1) then
                RegionValue = Mid(rsSubGroups("Code"),1,1)
                Region = Region + 1
                
                if Region > 4 then Region = 4
                
                response.write "<TR><TD HEIGHT=8 WIDTH=20></TD><TD HEIGHT=8></TD></TR>"
              end if

              if rsSubGroups("Enabled") = True then
                response.write "<TR><TD"
                    if rsSubGroups("Default_Select") = True then
                     response.write " BGCOLOR=""#FF0000"""
                    end if  
                response.write " CLASS=Normal>"
                response.write "<INPUT TYPE=""Checkbox"" NAME=""CSubGroups"" VALUE=""" & rsSubGroups("Code") & """"
                if rsSubGroups("Default_Select") = True then
                   response.write   " CHECKED"
                end if
                
                response.write "  CLASS=Normal></TD><TD CLASS=Normal BGCOLOR=""" & RegionColor(Region) & """>" & rsSubGroups("X_Description") & "</TD></TR>"
              end if  
                                       
          	  rsSubGroups.MoveNext 

            loop
          
            response.write "</TABLE>"
          
          end if  
             
          rsSubGroups.close
          set rsSubGroups=nothing

          response.write "</TD>" & vbCrLf
          response.write "</TR>" & vbCrLf
          
          ' Additional Site Access
          
          SQL =       "SELECT Site_Aux.Site_ID, Site.Site_Code, Site.Enabled, Site.Site_Description "
          SQL = SQL & "FROM Site_Aux LEFT JOIN Site ON Site_Aux.Site_ID_Aux = Site.ID "
          SQL = SQL & "WHERE Site_Aux.Site_ID=" & Site_ID & " ORDER BY Site.Site_Description"

          Set rsSite_Aux = Server.CreateObject("ADODB.Recordset")
          rsSite_Aux.Open SQL, conn, 3, 3
              
          if not rsSite_Aux.EOF then
            
    				response.write "<TR>"
            response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
            response.write Translate("Reciprocal Site Permissions",Login_Language,conn)
            response.write "  </TD>"
  	        response.write "</TR>"
            
    				response.write "<TR>"
            response.write "  <TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
            response.write "    &nbsp;"
            response.write "  </TD>"
            response.write "  <TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER VALIGN=TOP CLASS=Medium>"
            response.write "    &nbsp;"
            response.write "  </TD>"
  	        response.write "  <TD BGCOLOR=""White"" CLASS=Medium>"
            response.write "<TABLE WIDTH=""100%"">"

            do while not rsSite_Aux.EOF
              response.write "<TR><TD CLASS=Normal WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Groups_Aux"" VALUE=""" & rsSite_Aux("Site_Code") & """ CLASS=Normal></TD>"
              response.write "<TD CLASS=Normal BGCOLOR=""#FFCC00"">" & rsSite_Aux("Site_Description") & "</TD></TR>"
              rsSite_Aux.MoveNext
            loop
            
            response.write "</TABLE>"
            response.write "</TD>"
            response.write "</TR>"
            
          end if
            
          rsSite_Aux.close
          set rsSite_Aux = nothing

      
  				response.write "<TR>"
          response.write "  <TD BGCOLOR=""SILVER"" COLSPAN=3 VALIGN=TOP CLASS=MediumBold>"
          response.write Translate("Other Information",Login_Language,conn)
          response.write "  </TD>"
	        response.write "</TR>"

          ' Auxiliary Fields

          SQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.Site_ID=" & CInt(Site_ID) & " ORDER BY Auxiliary.Order_Num"
          Set rsAuxiliary = Server.CreateObject("ADODB.Recordset")
          rsAuxiliary.Open SQL, conn, 3, 3
          
          if not rsAuxiliary.EOF then

            do while not rsAuxiliary.EOF
            
              response.write "<INPUT TYPE=""Hidden"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Required"" VALUE=""" & rsAuxiliary("Required") & """>"            

              if rsAuxiliary("Enabled") = CInt(True) then

      				  response.write "<TR>" & vbCrLf
              	response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=MIDDLE CLASS=Medium>"
                response.write Translate(rsAuxiliary("Description"),Login_Language,conn) & ":"
                response.write "</TD>" & vbCrLf
                response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                if CInt(rsAuxiliary("Required")) = True then
                  Aux_Required(rsAuxiliary("Order_Num")) = True
                  Aux_Method(rsAuxiliary("Order_Num"))   = rsAuxiliary("Input_Method")
                  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Description"" VALUE=""" & rsAuxiliary("Description") &  """>"
                  response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10 ALIGN=ABSMIDDLE>"
                else
                  response.write "&nbsp;"
                end if  
                response.write "</TD>" & vbCrLf
  
   	            response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
  
                Aux_Selection     = Split(rsAuxiliary("Radio_Text"),",")
                Aux_Selection_Max = Ubound(Aux_Selection)               
  
                Select Case rsAuxiliary("Input_Method")
                  Case 0      ' Text
                    response.write "<INPUT TYPE=""Text"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ SIZE=""50"" MAXLENGTH=""50"" VALUE="""" CLASS=Medium>" & vbCrLf
                  Case 1      ' Drop-Down
                    response.write "<SELECT NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium>" & vbCrLf
                    response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                    for i = 0 to Aux_Selection_Max
                      response.write "<OPTION CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "</OPTION>" & vbCrLf
                    next
                    response.write "</SELECT>" & vbCrLf
                  Case 2      ' Radio
                    for i = 0 to Aux_Selection_Max
                      response.write "<INPUT TYPE=RADIO NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>&nbsp;" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "&nbsp;&nbsp;" & vbCrLf
                    next
                    response.write "</SELECT>" & vbCrLf
                end select
                    
                response.write "</TD>" & vbCrLf
                response.write "</TR>" & vbCrLf
              
              end if
              
              rsAuxiliary.MoveNext
              
            loop
            
            response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>" & vbCrLf
            
          end if
          
          rsAuxiliary.Close
          set rsAuxiliary = nothing

          ' Comment

  				response.write "<TR>" & vbCrLf
          response.write "	<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & Translate("Comment",Login_Language,conn) & ":</TD>" & vbCrLf
          response.write "	<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>" & vbCrLf
  	      response.write "  <TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
          response.write "    <TEXTAREA NAME=""Comment"" COLS=52 ROWS=6  MAXLENGTH=""255"" CLASS=Medium></TEXTAREA>"
          response.write "  </TD>" & vbCrLf
          response.write "</TR>" & vbCrLf

          ' Instant Message

  				response.write "<TR>" & vbCrLf
          response.write "	<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & Translate("Instant Message",Login_Language,conn) & ":</TD>" & vbCrLf
          response.write "	<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>" & vbCrLf
  	      response.write "  <TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
          response.write "    <TEXTAREA NAME=""Message"" COLS=52 ROWS=6  MAXLENGTH=""255"" CLASS=Medium></TEXTAREA>"
          response.write "  </TD>" & vbCrLf
          response.write "</TR>" & vbCrLf

          response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>" & vbCrLf  
          %>

          <!-- EMail Update -->

  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
              <%=Translate("Send Update EMail to",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <% if Account_ID <> false then %>                    
                <INPUT TYPE="Checkbox" NAME="send_email_user" CHECKED CLASS=Medium><%=Translate("User",Login_Language,conn)%><BR>
                <INPUT TYPE="Checkbox" NAME="send_email_fcm" CLASS=Medium><%=Translate("Account Manager",Login_Language,conn)%><BR>
                <INPUT TYPE="Checkbox" NAME="send_email_admin" CLASS=Medium><%=Translate("Site Administrator",Login_Language,conn)%>
              <% else %>
                 &nbsp;  
              <% end if %>
            </TD>
          </TR>
          
          <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>

          
          <!-- Navigation Buttons -->
  
          <TR>
          <TD COLSPAN=3 CLASS=Medium>
            <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666">
              <TR>
                <TD ALIGN=CENTER WIDTH="48%" CLASS=Medium>
                  <%
                    response.write "<INPUT TYPE=""BUTTON"" Value=""" & Translate("Main Menu",Login_Language,conn) & """ Language=""JavaScript"" onclick=""location.href='" & BackURL & "'"" CLASS=NavLeftHighlight1>"
                  %>  
                </TD>
                <TD ALIGN=CENTER WIDTH="2%" CLASS=Medium>
                </TD>             
                <TD ALIGN=CENTER WIDTH="25%" CLASS=Medium>                 
                  <%
                  if Account_ID <> false then
                    response.write "<INPUT TYPE=""Submit"" NAME=""Update"" VALUE="" " & Translate("Save",Login_Language,conn) & " / " & Translate("Update",Login_Language,conn) & " "" CLASS=NavLeftHighlight1>"
                  else
                    response.write "&nbsp;"
                  end if
                  %>
                </TD>
                <TD ALIGN=CENTER WIDTH="25%" CLASS=Medium>
                  &nbsp;
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
        
  <% end if %>

<%
else

  %>
  <B><%=Translate("You are not authorized to edit this account.",Login_Language,conn)%></B>
  <BR><BR>
  <%Call Nav_Border_Begin%>
  <INPUT TYPE="BUTTON" Value=" <%=Translate("Main Menu",Login_Language,conn)%> " Language="JavaScript" onclick="location.href='default.asp?Site_ID=<%=Site_ID%>'" CLASS=NavLeftHighlight1>
  <%Call Nav_Border_End%>
  <%

end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------

%>
<!-- #include virtual="/include/core_countries.inc"-->

<SCRIPT LANGUAGE=JAVASCRIPT>

<!--

function IsDate(datein){
	
	var indate=datein;

	if (indate.indexOf("-") != -1) {
		var sdate = indate.split("-");
	}
	else {
		var sdate = indate.split("/");
	}

	var chkDate=new Date(Date.parse(indate));
	
	var cmpDate=(chkDate.getMonth()+1) + "/" + (chkDate.getDate())+"/" + (chkDate.getYear());
	var indate2=(Math.abs(sdate[0])) + "/" + (Math.abs(sdate[1]))+"/" + (Math.abs(sdate[2]));

	if (indate2 != cmpDate) {  
    return 0;
	}
	else {
		if (cmpDate == "NaN/NaN/NaN") {
      return 0;
		}
		else {
      return 1;
		}	
	}
}

// --------------------------------------------------------------------------------------

function CheckRequiredFields() {

  var ErrorMsg = "";
  var str;
  var strVal;
  var strChk;
  var badchars;
  var D_BadChars;
  var TestRadio;
  var ctr;
  var ctr0;
  var RadioChecked = 0;
  var dNT = document.NTAccount;
   
  <%
  if Account_Type = True then
    response.write "if (dNT.Type_Code.value == """") {" & vbCrLf
    response.write "ErrorMsg = ErrorMsg + ""Customer Type\r\n"";" & vbCrLf
    response.write "}" & vbCrLf
  end if
  %>

  if (! dNT.Fcm.checked) {
    if (dNT.Fcm_ID.selectedIndex < 1) {
      ErrorMsg = ErrorMsg + "<%=Translate("Account Managers Name",Alt_Language,conn)%>\r\n";
    }
  }

  // check that there is a business_system if there is a Fluke_ID
  if (dNT.Fluke_ID.value.length) {
    if (! dNT.Business_system.options[dNT.Business_system.selectedIndex].value.length) {
      ErrorMsg = ErrorMsg + "<%=Translate("Business System is required with Fluke Customer Number",Alt_Language,conn)%>\r\n";
	  }
  }
      
  if (dNT.ExpirationDate.value.length) {
    if (IsDate(dNT.ExpirationDate.value) == 0) {
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Expiration Date or Date Format.  Use: mm/dd/yyyy or leave blank.",Alt_Language,conn)%>\r\n";
    }
  }

  strlen = dNT.NTLogin.value.length;
  if (strlen < 7 || strlen > 20) {   
    if (dNT.ID.value == "add" ) {
      ErrorMsg = ErrorMsg + "<%=Translate("Logon User Name must be at least 7 characters",Alt_Language,conn)%>, <%=Translate("maximum",Alt_Language,conn)%> 20\r\n";
    }
  }

  str = dNT.NTLogin.value;  
  badchars = /[\[\]:;|=,+*?<>"\\\/]/
  D_BadChars = '\\ / [ ] : ; | = , + * ? < > "';              //"'
  if (badchars.test(str)) {
	  ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Logon User Name contains illegal characters",Alt_Language,conn)) & ": "%>" + D_BadChars + "\r\n";
  }

  strlen = dNT.Password.value.length;
  if (strlen < 7 || strlen > 14) {
      ErrorMsg = ErrorMsg + "<%=Translate("Password must be at least 7 characters",Alt_Language,conn)%>, <%=Translate("maximum",Alt_Language,conn)%> 14\r\n";
  }
  
  str = dNT.Password.value;
  badchars = /[\[\]:;|=,+*?<>"\\\/]/
  D_BadChars = '\\ / [ ] : ; | = , + * ? < > "';              //"'
  if (badchars.test(str)) {
	  ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Password contains illegal characters",Alt_Language,conn)) & ": "%>" + D_BadChars + "\r\n";
  }

  if (dNT.Password_Change.value.length) {
    strlen = dNT.Password_Change.value.length;
    if (strlen < 7 || strlen > 14) {
        ErrorMsg = ErrorMsg + "<%=Translate("Password Change must be at least 7 characters",Alt_Language,conn)%>, <%=Translate("maximum",Alt_Language,conn)%> 14\r\n";
    }
    else {
      if (dNT.ID.value == "add" ) {
        str = dNT.Password_Change.value;
        badchars = /[\[\]:;|=,+*?<>"\\\/]/
        D_BadChars = '\\ / [ ] : ; | = , + * ? < > "';              //"'
        if (badchars.test(str)) {
      	  ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Password Change contains illegal characters",Alt_Language,conn)) & ": "%>" + D_BadChars + "\r\n";
        }
      }  
    }
  }
 
  if (! dNT.FirstName.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("First Name",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.LastName.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Surname",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Job_Title.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Job Title",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Phone.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Office Phone",Alt_Language,conn)%> (<%=Translate("Direct",Alt_Language,conn)%>)\r\n";
  }

  if (! dNT.Email.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Email",Alt_Language,conn)%> (<%=Translate("Direct",Alt_Language,conn)%>)\r\n";
  }
  
  if (! dNT.EMail_Method.value.length) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email Format",Alt_Language,conn))%>\r\n";  
  }

  if (! dNT.Connection_Speed.value.length) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Internet Connection Speed",Alt_Language,conn))%>\r\n";  
  }

  if (! dNT.Company.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Company Name",Alt_Language,conn)%>\r\n";
  }
  
  if (! dNT.Business_Address.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Office Address",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_City.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Office City",Alt_Language,conn)%>\r\n";
  }

  if ((! dNT.Business_State.value.length) && (! dNT.Business_State_Other.value.length)) {
    ErrorMsg = ErrorMsg + "<%=Translate("Office USA State or Canadian Province or N/A",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Postal_Code.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Office Postal Code",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Country.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Office Country",Alt_Language,conn)%>\r\n";
  }

  // check SubGroups checkbox group - note it is always an array
  for(ctr=0;ctr<dNT.CSubGroups.length;ctr++) {
    if (dNT.CSubGroups[ctr].checked) {
	  RadioChecked = 1;
      if ((dNT.CSubGroups[ctr].value == 'branch') && (! dNT.Fluke_ID.value.length)) {
        ErrorMsg = ErrorMsg + "<%=Translate("Group Affiliation - Branch Location Administrator without Fluke Customer Number",Alt_Language,conn)%>\r\n";
	    }
	  }
  }
  
  if (! RadioChecked) {
    // the user could be an administrator or domain, in which case they have a hidden field
	  if (! dNT.SubGroups.value.length) {
      ErrorMsg = ErrorMsg + "<%=Translate("Group Affiliation(s)",Alt_Language,conn)%>\r\n";
	  }
  }
  RadioChecked = 0;
  
  CheckSts = 0;
  for (ctr=0; ctr < dNT.CSubGroups.length; ctr++) {
    if (dNT.CSubGroups[ctr].value == 'wtbadm' || dNT.CSubGroups[ctr].value == 'wtbtsm, wtbaap' || dNT.CSubGroups[ctr].value == 'wtbtsm' || dNT.CSubGroups[ctr].value == 'wtbdis, wtbaap'  || dNT.CSubGroups[ctr].value == 'wtbdis') {
      if (dNT.CSubGroups[ctr].checked == true) {
        CheckSts = CheckSts + 1;
      }    
    }  
  }
  if (CheckSts > 1) {
    ErrorMsg = ErrorMsg + "<%=Translate("Please select only one WTB Administrator permission or none.",Alt_Language,conn)%>\r\n";    
  }
  
 
  CheckSts = 0;
  for (ctr = 0; ctr < dNT.CSubGroups.length; ctr++) {
    if (dNT.CSubGroups[ctr].value == 'wtbadm') {
      if (dNT.CSubGroups[ctr].checked == true) {
        CheckSts = CheckSts + 1;
        break;
      }
    }    
  }
  
  if (CheckSts == 1) {
    for (ctr0 = 0; ctr0 < dNT.CSubGroups.length; ctr0++) {
      if (dNT.CSubGroups[ctr0].value == 'account' || dNT.CSubGroups[ctr0].value == 'site' || dNT.CSubGroups[ctr0].value == 'administrator' || dNT.CSubGroups[ctr0].value == 'domain') {
        if (dNT.CSubGroups[ctr0].checked == true) {
          CheckSts = CheckSts + 1;
          break;  
        }
      }
    }              
  }    

  if (CheckSts == 1) {
    ErrorMsg = ErrorMsg + "<%=Translate("Account Administrator previlages must be granted to allow Where To Buy Super Administrator previlages.",Alt_Language,conn)%>\r\n";    
  }
  
  CheckSts = 0;

  <%  
    for i = 0 to 9
      if Aux_Required(i) = true then 
        select case Aux_Method(i)
		      case 0  ' Text
      			response.write "if (! dNT.Aux_" + Trim(CStr(i)) + ".value.length) {" & vbCrLf
      			response.write "  ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + dNT.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
      			response.write "}" & vbCrLF & vbCrLF
    		  case 1  ' Drop-Down
      			response.write "strVal = dNT.Aux_" + Trim(CStr(i)) + ".value; " & vbCrLF				
      			response.write "strChk = strVal.substring(0, 7); " & vbCrLF 
      			response.write "  if ((strVal == """") || (strChk.toLowerCase() == ""select"")) {" & vbCrLf
      			response.write "     ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + dNT.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
      			response.write "  }" & vbCrLF & vbCrLF				 
          case 2  ' Radio or Checkbox
            response.write "RadioChecked = 0;" & vbCrLf
            response.write "TestRadio = dNT.Aux_" + Trim(Cstr(i)) & ";" & vbCrLf
            response.write "for (ctr=0; ctr < TestRadio.length; ctr++) {" & vbCrLf
            response.write "  if (TestRadio[ctr].checked) {" & vbCrLf
            response.write "   RadioChecked = 1;" & vbCrLf
            response.write "   break;" & vbCrLf
            response.write "   }" & vbCrLf
            response.write "}" & vbCrLf            
      			response.write "if (RadioChecked == 0) {" & vbCrLf 
      			response.write "     ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + dNT.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
      			response.write "}" & vbCrLF & vbCrLF
        end select
      end if
    next  
  %>         
  
  if (dNT.PostalSame.checked) {
      dNT.Postal_City.value        = dNT.Business_City.value;            
      dNT.Postal_State.value       = dNT.Business_State.value;
      dNT.Postal_State_Other.value = dNT.Business_State_Other.value;
      dNT.Postal_Postal_Code.value = dNT.Business_Postal_Code.value;      
      dNT.Postal_Country.value     = dNT.Business_Country.value;            
      if (dNT.Postal_Address.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Box Number",Login_Language,conn))%>\r\n";
      }
  }
    
  if (dNT.ShippingSame.checked) {
      dNT.Shipping_MailStop.value    = dNT.Business_MailStop.value;  
      dNT.Shipping_Address.value     = dNT.Business_Address.value;
      dNT.Shipping_Address_2.value   = dNT.Business_Address_2.value;
      dNT.Shipping_City.value        = dNT.Business_City.value;            
      dNT.Shipping_State.value       = dNT.Business_State.value;
      dNT.Shipping_State_Other.value = dNT.Business_State_Other.value;
      dNT.Shipping_Postal_Code.value = dNT.Business_Postal_Code.value;      
      dNT.Shipping_Country.value     = dNT.Business_Country.value;            
  }

  if (dNT.Postal_State.value == "" && dNT.Postal_State_Other.value != "") {
    dNT.Postal_State.value = "ZZ";
  }

  if (dNT.Shipping_State.value == "" && dNT.Shipping_State_Other.value != "") {
    dNT.Shipping_State.value = "ZZ";
  }
  
  if (dNT.Fcm_ID.options[dNT.Fcm_ID.selectedIndex].value == 0) {
    dNT.send_email_fcm.checked = false;
  }
  
  if (ErrorMsg.length) {
    ErrorMsg = "<%=Translate("Please enter the missing information for following REQUIRED fields (or use N/A)",Alt_Language,conn)%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    return (false);
  }
  else {
  	// the receiving code is expecting the the hidden field "SubGroups", add CSubGroups to it
	
    for(ctr=0;ctr<dNT.CSubGroups.length;ctr++) {
      if (dNT.CSubGroups[ctr].checked) {
        if (dNT.SubGroups.value.length) {
          dNT.SubGroups.value += ', ' + dNT.CSubGroups[ctr].value;
  	    }
	      else {
          dNT.SubGroups.value = dNT.CSubGroups[ctr].value;
	      }
	    }
    }
  	return (true);
  } 
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

//-->
</SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="/SW-Common/SW-Calendar_PopUp.js"></SCRIPT>

<%
Call Disconnect_SiteWide
%>