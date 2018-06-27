<%
' --------------------------------------------------------------------------------------
' Author: K. D. Whitlock
' Date:   2/1/2000
' 06/19/2002 - Added Fields and Re-Ordered Form to work with Euro DCM
'
' --------------------------------------------------------------------------------------
' Update Profile
' --------------------------------------------------------------------------------------  

  Dim Disable_Flag
  
  if isblank(Cart_Mode) then
    Cart_Mode = False
    Disable_Flag = ""
  else
    Cart_Mode = True
    Disable_Flag = ""    ' Set to " Disabled" if you want to limit editing to selected field data
  end if

  SQL = "SELECT UserData.*, UserData.ID "
  SQL = SQL & "FROM UserData "
  SQL = SQL & "WHERE (((UserData.ID)=" & CLng(Login_ID) & "))"

  Set rsProfile = Server.CreateObject("ADODB.Recordset")
  rsProfile.Open sql, conn, 3, 3

  if not rsProfile.EOF then    

    response.write "<IMG SRC=""/images/lock.gif"" BORDER=0 WIDTH=37 ALIGN=""ABSMIDDLE"">" & Translate("This is a secure site connection to protect your personal information.",Login_Language,conn) & "<P>"

%>
    <FORM NAME="NTAccount" ACTION="/sw-common/SW-Profile_Admin.asp" METHOD="POST" onsubmit="return(CheckRequiredFields(this.form));" onKeyUp="highlight(event)" onClick="highlight(event)">
    <INPUT TYPE="Hidden" NAME="BackURL" Value="<%=BackURLSecure%>">                
    <INPUT TYPE="Hidden" NAME="Account_ID" VALUE="<%=Login_ID%>">
    <INPUT TYPE="Hidden" NAME="Region" VALUE="<%=rsProfile("Region")%>">    
    <INPUT TYPE="Hidden" NAME="ChangeID" VALUE="<%=Login_ID%>">
    <INPUT TYPE="Hidden" NAME="ChangeDate" Value="<%=Date%>">
    <INPUT TYPE="Hidden" NAME="Site_ID" Value="<%=Site_ID%>">            
    <INPUT TYPE="Hidden" NAME="Permit_Update" Value="<%=CInt(True)%>">
    <INPUT TYPE="Hidden" NAME="Login_Language" Value="<%=Login_Language%>">            
    <INPUT TYPE="Hidden" NAME="CM_ID" Value="<%=rsProfile("CM_ID")%>">        
    <%
    if CInt(Cart_Mode) = CInt(True) then
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Cart_Mode"" VALUE=""" & CInt(True) & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""NTLogin"" VALUE=""" & rsProfile("NTLogin") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Language"" VALUE=""" & rsProfile("Language") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Subscription"" VALUE="""
      if CInt(rsProfile("Subscription")) = True then response.write "on"
      response.write """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Password_Current"" VALUE="""">" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Password_Change"" VALUE="""">" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Password_Confirm"" VALUE="""">" & vbCrLf      
      response.write "<INPUT TYPE=""Hidden"" NAME=""Prefix"" VALUE=""" & rsProfile("Prefix") & """>" & vbCrLf
      response.write "<INPUT TYPE=""Hidden"" NAME=""Gender"" VALUE=""" & rsProfile("Gender") & """>" & vbCrLf
      response.write "<INPUT TYPE=""Hidden"" NAME=""Initials"" VALUE=""" & rsProfile("Initials") & """>"& vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Mobile_Phone"" VALUE=""" & rsProfile("Mobile_Phone") & """>" & vbCrLF
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Pager"" VALUE=""" & rsProfile("Pager") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""EMail_Method"" VALUE=""" & rsProfile("EMail_Method") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Connection_Speed"" VALUE=""" & rsProfile("Connection_Speed") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Company_Website"" VALUE=""" & rsProfile("Company_Website") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Email_2"" VALUE=""" & rsProfile("Email_2") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Business_Phone_2"" VALUE=""" & rsProfile("Business_Phone_2") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""PostalSame"" VALUE="""">" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Postal_Address"" VALUE=""" & rsProfile("Postal_Address") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Postal_City"" VALUE=""" & rsProfile("Postal_City") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Postal_State"" VALUE=""" & rsProfile("Postal_State") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Postal_State_Other"" VALUE=""" & rsProfile("Postal_State_Other") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Postal_Postal_Code"" VALUE=""" & rsProfile("Postal_Postal_Code") & """>" & vbCrLf
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Postal_Country"" VALUE=""" & rsProfile("Postal_Country") & """>" & vbCrLf
    else
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Cart_Mode"" VALUE=""" & CInt(False) & """>" & vbCrLf
    end if

    Call Table_Begin
    %>
       
    <TABLE WIDTH="100%" BORDER=0 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
    	<TR>
    		<TD WIDTH="100%" BGCOLOR="#EEEEEE" VALIGN=TOP>
  	  		<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>       
  		  		<TR>
            	<TD WIDTH="40%" COLSPAN=2 CLASS=NavLeftSelected1 NOWRAP>
                <%=Translate("Description",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallBold>(<%=Translate("Note",Login_Language,conn)%>:&nbsp;<IMG SRC="/images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> = <%=Translate("Required Information or use N/A",Login_Language,conn)%>)</SPAN>
              </TD>
    	        <TD WIDTH="60%" ALIGN=LEFT CLASS=NavLeftSelected1>
                <%=Translate("Account Profile Information",Login_Language,conn)%>
              </TD>
            </TR>

            <% if CInt(Cart_Mode) = CInt(False) then %>
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold>
              <%=Translate("Account Information",Login_Language,conn)%>
            </TD></TR>
            <% end if %>
                                
             <!-- NT Login -->
  
            <% if CInt(Cart_Mode) = CInt(False) then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Account User Name",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>&nbsp;</TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                  response.write "<INPUT CLASS=Medium TYPE=""Hidden"" NAME=""NTLogin"" VALUE=""" & rsProfile("NTLogin") & """>" & rsProfile("NTLogin")
                %>    
              </TD>
            </TR>
            <% end if %>
    
            <!-- Change Password -->
    
            <% if CInt(Cart_Mode) = CInt(False) then
                 if rsProfile("Password") <> "(hidden)" and not isblank(rsProfile("Password")) and rsProfile("NewFlag") <> CInt(True) and ((site_id = 3 and CInt(rsProfile("Region"))) <> 2 or site_id <> 3) then %>
          				<TR>
                  	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium><%=Translate("Current Password",Login_Language,conn)%>:</TD>
                  	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>&nbsp;</TD>                                                
          	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium><INPUT CLASS=Medium TYPE="PASSWORD" NAME="Password_Current" SIZE="50" MAXLENGTH="14" VALUE=""></TD>
                  </TR>

          				<TR>
                  	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium><%=Translate("New Password",Login_Language,conn)%>:</TD>
                  	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>&nbsp;</TD>                                                
          	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium><INPUT CLASS=Medium TYPE="PASSWORD" NAME="Password_Change" SIZE="50" MAXLENGTH="14" VALUE=""></TD>
                  </TR>
          				<TR>
                  	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium><%=Translate("New Password - Confirm",Login_Language,conn)%>:</TD>
                  	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>&nbsp;</TD>                                                
          	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium><INPUT CLASS=Medium TYPE="PASSWORD" NAME="Password_Confirm" SIZE="50" MAXLENGTH="14" VALUE=""></TD>
                  </TR>
            <%   else %>
                  <INPUT CLASS=Medium TYPE="Hidden" NAME="Password_Current" VALUE="">
                  <INPUT CLASS=Medium TYPE="Hidden" NAME="Password_Change" VALUE="">
                  <INPUT CLASS=Medium TYPE="Hidden" NAME="Password_Confirm" VALUE="">                                    
            <%   end if
               end if %>
  
            <!-- Preferred Language -->

            <% if CInt(Cart_Mode) = CInt(False) and Login_Language <> "elo" then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Preferred Language",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
  
                <SELECT Name="Language" CLASS=Medium>    
                <%
                SQL = "SELECT * FROM Language WHERE Language.Enable=-1" & " ORDER BY Language.Sort"
                Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                rsLanguage.Open SQL, conn, 3, 3
                                      
                Do while not rsLanguage.EOF
                  if rsProfile("Language") = rsLanguage("Code") then
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
            <%
            end if
            %>

             <!-- Subscription Service -->
             
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Subscription Service",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Checkbox" NAME="Subscription" <%if rsProfile("Subscription") = True then response.write " CHECKED"%>>
                &nbsp;&nbsp;<%=Translate("Site Newsletters by EMail",Login_Language,conn)%>
              </TD>
            </TR>
            <% end if %>
  
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Contact Information",Login_Language,conn)%></TD></TR>
  
             <!-- Name -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Name",Login_Language,conn)%>:&nbsp;&nbsp;<SPAN CLASS=Small><IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("First",Login_Language,conn)%>]&nbsp;&nbsp;[<%=Translate("Middle",Login_Language,conn)%>]&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>[<%=Translate("Surname",Login_Language,conn)%>],&nbsp;&nbsp;[<%=Translate("Suffix",Login_Language,conn)%>]</SPAN>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                <INPUT TYPE="Text" <%=Disable_Flag%> NAME="FirstName" SIZE="10" MAXLENGTH="50" VALUE="<%=rsProfile("FirstName")%>" CLASS=Medium>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" <%=Disable_Flag%> NAME="MiddleName" SIZE="6" MAXLENGTH="50" VALUE="<%=rsProfile("MiddleName")%>" CLASS=Medium>&nbsp;&nbsp;&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><INPUT TYPE="Text" <%=Disable_Flag%> NAME="LastName" SIZE="11" MAXLENGTH="50" VALUE="<%=rsProfile("LastName")%>" CLASS=Medium> <B>,</B>&nbsp;&nbsp;&nbsp;<INPUT TYPE="Text" <%=Disable_Flag%> NAME="Suffix" SIZE="2" MAXLENGTH="50" VALUE="<%=rsProfile("Suffix")%>" CLASS=Medium>
              </TD>
            </TR>
  
             <!-- Initials -->
    
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Initials",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                <INPUT TYPE="Text" <%=Disable_Flag%> NAME="Initials" SIZE="10" MAXLENGTH="10" VALUE="<%=rsProfile("Initials")%>" CLASS=Medium>
              </TD>
            </TR>
            <% end if %>
  
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <%=Translate("Gender",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
                <% 
                sValue = rsProfile("Gender")
                %>
                <SELECT CLASS=Medium NAME="Gender" CLASS=Medium>
                  <OPTION CLASS=Medium VALUE=""><%=Translate("Select",Login_Language,conn)%></OPTION>
                  <OPTION CLASS=Region2 VALUE="0"<% If sValue = "0" Then Response.Write " SELECTED" %>><%=Translate("Male",Login_Language,conn)%></OPTION>
                  <OPTION CLASS=Region3 VALUE="1"<% If sValue = "1" Then Response.Write " SELECTED" %>><%=Translate("Female",Login_Language,conn)%></OPTION>
                </SELECT>
              </TD>
            </TR>
            <% end if %>
  
             <!-- Job Title -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Job Title",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Job_Title" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Job_Title")%>">
              </TD>
            </TR>
  
             <!-- Business Phone -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Phone",Login_Language,conn)%>&nbsp;(<%=Translate("Direct",Login_Language,conn)%>):</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone")%>">&nbsp;&nbsp;<%response.write Translate("Extension",Login_Language,conn)%>: <INPUT CLASS=Medium TYPE="Text" NAME="Business_Phone_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone_Extension")%>">
              </TD>
            </TR>
  
             <!-- Mobile Phone -->
  
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Mobile",Login_Language,conn)%>&nbsp;<%=Translate("Phone",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Mobile_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Mobile_Phone")%>">
              </TD>
            </TR>
            <% end if %>
  
             <!-- Pager -->
  
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Pager",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Pager" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Pager")%>">
              </TD>
            </TR>
            <% end if %>
  
            <!-- Email -->
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%response.write Translate("EMail",Login_Language,conn)%> (<%=Translate("Direct",Login_Language,conn)%>):</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Email" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Email")%>">
              </TD>
            </TR>
          
            <!-- EMail Method -->
            
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Email Format",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
  
                <SELECT Name="EMail_Method" CLASS=Medium>
                  <OPTION Class=Medium VALUE="0"<%if rsProfile("Email_Method") = 0 then response.write " SELECTED"%>><%=Translate("Plain Text without Graphics",Login_Language,conn)%></OPTION>
                  <OPTION Class=Medium VALUE="1"<%if rsProfile("Email_Method") = 1 then response.write " SELECTED"%>><%=Translate("Rich Text with Graphics",Login_Language,conn)%></OPTION>
                </SELECT>                                  
  
              </TD>
            </TR>
            <% end if %>
  
            <!-- Connection Speed -->
  
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
                <% response.write Translate("Internet Connection Speed",Login_Language,conn) & ":"%>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                SQL = "SELECT Download_Time.* FROM Download_Time WHERE Download_Time.Enabled=" & CInt(True)
                SQL = SQL & " ORDER BY DownLoad_Time.bps, DownLoad_Time.Description"
                Set rsDownload = Server.CreateObject("ADODB.Recordset")
                rsDownload.Open SQL, conn, 3, 3
                response.write "<SELECT NAME=""Connection_Speed"">" & vbCrLf
                response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                do while not rsDownload.EOF
                  response.write "<OPTION Class=Medium VALUE=""" & rsDownload("ID") & """"
                  if not isblank(rsProfile("Connection_Speed")) then
                    if rsProfile("Connection_Speed") = rsDownLoad("ID") then response.write " SELECTED"
                  elseif rsDownload("ID") = 6 and (isblank(rsProfile("Connection_Speed")) or rsProfile("Connection_Speed") = 0) then
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
            <% end if %>

            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Company Information",Login_Language,conn)%></TD></TR>
    
            <!-- Company -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Company Name",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Company" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Company")%>">
              </TD>
            </TR>
    
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Company Website Address",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Company_Website" SIZE="50" MAXLENGTH="255" VALUE="<%=rsProfile("Company_Website")%>">
              </TD>
            </TR>
            <%
            end if
            %>
  
            <TR><TD COLSPAN=3 BGCOLOR="Silver" CLASS=MediumBold><%=Translate("Office Information",Login_Language,conn)%></TD></TR>
   
            <!-- Mail Stop -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Mail Stop",Login_Language,conn)%> / <%=Translate("Building Number",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_MailStop")%>">
              </TD>
            </TR>
    
            <!-- Business Address -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Top CLASS=Medium><%=Translate("Address",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT VALIGN=TOP CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                               
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Address")%>"><BR>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Address_2")%>">                
              </TD>
            </TR>
  
            <!-- Business City -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("City",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_City")%>">
              </TD>
            </TR>
  
            <!-- Business State -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("USA State or Canadian Province",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <% SValue = rsProfile("Business_State") %>
                <SELECT NAME="Business_State" <%=Disable_Flag%> CLASS=Medium>

                 <%
                 if isblank(rsProfile("Business_State")) then
                  response.write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"
                 end if
                 %>  
                  
                <!--#include virtual="/include/core_states.inc"-->
  
                </SELECT>
  
              </TD>
            </TR>
  
            <!-- Business State Other-->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><B><%=Translate("or",Login_Language,conn)%></B>&nbsp;<%=Translate("Other State, Province or Local",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_State_Other")%>">
              </TD>
            </TR>
  
            <!-- Business Postal Code -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Postal Code",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Postal_Code")%>">
              </TD>
            </TR>
  
            <!-- Business Country -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Country",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                Users_Country = rsProfile("Business_Country")
                Call Connect_FormDatabase
                Call displayCountryList("Business_Country",Users_Country,Translate("Select from List",Login_Language,conn),"Medium")
                Call Disconnect_FormDatabase
                %>
              </TD>
            </TR>
  
            <!-- Email 2 -->
    
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("EMail",Login_Language,conn)%>&nbsp;&nbsp;&nbsp;(<%=Translate("General Office",Login_Language,conn)%>):</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Email_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Email_2")%>">
              </TD>
            </TR>
            <%
            end if
            %>
  
            <!-- Business Phone 2 -->
    
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Phone",Login_Language,conn)%>&nbsp;(<%=Translate("General Office",Login_Language,conn)%>):</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_Phone_2" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone_2")%>">&nbsp;&nbsp;<%response.write Translate("Extension",Login_Language,conn)%>: <INPUT CLASS=Medium TYPE="Text" NAME="Business_Phone_2_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone_2_Extension")%>">
              </TD>
            </TR>
            <%
            end if
            %>
  
            <!-- Business Fax -->
    
            <% if CInt(Cart_Mode) = CInt(False) then %>  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Fax",Login_Language,conn)%>&nbsp;(<%=Translate("General Office",Login_Language,conn)%>):</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" <%=Disable_Flag%> NAME="Business_Fax" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Fax")%>">
              </TD>
            </TR>
            <%
            end if
            %>

            <% if CInt(Cart_Mode) = CInt(False) then %>  

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
                <INPUT TYPE="Text" <%=Disable_Flag%> NAME="Postal_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Postal_Address")%>" CLASS=Medium>
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
                <INPUT CLASS=NavLeftHighlight1 TYPE="BUTTON" VALUE="<%=Translate("Click Here",Login_Language,conn)%>" ONCLICK="Postal_Same();" onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'">
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
                <INPUT TYPE="Text" <%=Disable_Flag%> NAME="Postal_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Postal_City")%>" CLASS=Medium>
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
    
                <% SValue = rsProfile("Postal_State") %>              
                <SELECT <%=Disable_Flag%> CLASS=Medium NAME="Postal_State" CLASS=Medium>
    
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
                <INPUT TYPE="Text" <%=Disable_Flag%> NAME="Postal_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Postal_State_Other")%>" CLASS=Medium>
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
                <INPUT TYPE="Text" <%=Disable_Flag%> NAME="Postal_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Postal_Postal_Code")%>" CLASS=Medium>
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
                Users_Country = rsProfile("Postal_Country")
                Call Connect_FormDatabase
                Call displayCountryList("Postal_Country",Users_Country,Translate("Select from List",Login_Language,conn),"Medium")
                Call Disconnect_FormDatabase
                %>
              </TD>
            </TR>
            <% end if %>  
 
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
                <INPUT CLASS=NavLeftHighlight1 TYPE="BUTTON" VALUE="<%=Translate("Click Here",Login_Language,conn)%>" ONCLICK="Shipping_Same();"" onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'">
              </TD>
            </TR>
 
            <!-- Shipping Mail Stop -->
  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Mail Stop",Login_Language,conn)%> / <%=Translate("Building Number",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_MailStop")%>">
              </TD>
            </TR>
  
            <!-- Shipping Address -->
                  
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Top CLASS=Medium><%=Translate("Address",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT VALIGN=TOP CLASS=Medium>
                <%
                 if CInt(Cart_Mode) = CInt(False) then 
                  response.write "&nbsp;"
                else
                  response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                end if
                %>  
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_Address")%>"><BR>
                <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_Address_2")%>">
              </TD>
            </TR>
  
            <!-- Shipping City -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("City",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <%
                 if CInt(Cart_Mode) = CInt(False) then 
                  response.write "&nbsp;"
                else
                  response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                end if
                %>  
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_City")%>">
              </TD>
            </TR>
  
            <!-- Shipping State -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("USA State or Canadian Provdince",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <%
                 if CInt(Cart_Mode) = CInt(False) then 
                  response.write "&nbsp;"
                else
                  response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                end if
                %>  
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
  
                <% SValue = rsProfile("Shipping_State") %>              
                <SELECT NAME="Shipping_State" CLASS=Medium>
  
                 <%
                 if isblank(rsProfile("Shipping_State")) then
                  response.write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>"
                 end if
                 %>  
  
                <!--#include virtual="/include/core_states.inc"-->
  
                </SELECT>
              </TD>
            </TR>
  
            <!-- Shipping State Other-->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><B><%=Translate("or",Login_Language,conn)%></B> <%=Translate("Other State, Province or Local",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                &nbsp;
              </TD>                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_State_Other")%>">
              </TD>
            </TR>
  
            <!-- Shipping Postal Code -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Postal Code",Login_Language,conn)%>:</TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
                <%
                 if CInt(Cart_Mode) = CInt(False) then 
                  response.write "&nbsp;"
                else
                  response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                end if
                %>  
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_Postal_Code")%>">
              </TD>
            </TR>
  
            <!-- Shipping Country -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Country",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
                <%
                 if CInt(Cart_Mode) = CInt(False) then 
                  response.write "&nbsp;"
                else
                  response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"" ALIGN=ABSMIDDLE>"
                end if
                %>  
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                Users_Country = rsProfile("Shipping_Country")
                Call Connect_FormDatabase
                Call DisplayCountryList("Shipping_Country",Users_Country,Translate("Select from List",Login_Language,conn),"Medium")
                Call Disconnect_FormDatabase
                %>  
              </TD>
            </TR>                          

            <%
           
            ' Auxiliary Fields
            
            Dim Aux_Selection
            Dim Aux_Selection_Max
            Dim Aux_Required(9)
            Dim Aux_Method(9)
            for i = 0 to 9
              Aux_Required(i) = False
              Aux_Method(i)   = 0
            next  
            
            SQL = "SELECT Auxiliary.* FROM Auxiliary WHERE Auxiliary.Site_ID=" & CInt(Site_ID) & " AND Auxiliary.Enabled=" & CInt(True) & " AND Auxiliary.User_Edit=" & CInt(True) & " ORDER BY Auxiliary.Order_Num"
            Set rsAuxiliary = Server.CreateObject("ADODB.Recordset")
            rsAuxiliary.Open SQL, conn, 3, 3
            
            if not rsAuxiliary.EOF then
            
              if CInt(Cart_Mode) = CInt(False) then 
                response.write "<TR><TD COLSPAN=3 BGCOLOR=""Silver"" CLASS=MediumBold>" & Translate("Other Information",Login_Language,conn) & "</TD></TR>"
              end if  
  
              do while not rsAuxiliary.EOF
  
                response.write "<INPUT TYPE=""Hidden"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Required"" VALUE=""" & rsAuxiliary("Required") & """>"
    
                if rsAuxiliary("Enabled") = CInt(True) and rsAuxiliary("User_Edit") = CInt(True) then
  
                  if CInt(Cart_Mode) = CInt(False) then 
          				  response.write "<TR>" & vbCrLf
                  	response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=MIDDLE CLASS=Medium>"
                    response.write Translate(rsAuxiliary("Description"),Login_Language,conn) & ":"
                    response.write "</TD>" & vbCrLf
                    response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                  end if
                  if CInt(rsAuxiliary("Required")) = True then
                    Aux_Required(rsAuxiliary("Order_Num")) = True
                    Aux_Method(rsAuxiliary("Order_Num"))   = rsAuxiliary("Input_Method")
                    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Description"" VALUE=""" & rsAuxiliary("Description") & """>"
                     if CInt(Cart_Mode) = CInt(False) then 
                      response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10 ALIGN=ABSMIDDLE>"
                    end if
                  else
                     if CInt(Cart_Mode) = CInt(False) then 
                      response.write "&nbsp;"
                    end if
                  end if  
  
                  if CInt(Cart_Mode) = CInt(False) then 
                    response.write "</TD>" & vbCrLf
       	            response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
                  end if
  
                  Aux_Selection     = Split(rsAuxiliary("Radio_Text"),",")
                  Aux_Selection_Max = Ubound(Aux_Selection)
  
                  if CInt(Cart_Mode) = CInt(False) then 
                    Select Case rsAuxiliary("Input_Method")
                      Case 0      ' Text
                        response.write "<INPUT TYPE=""Text"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ SIZE=""50"" MAXLENGTH=""50"" VALUE=""" & rsProfile("Aux_" & Trim(rsAuxiliary("Order_Num"))) & """ CLASS=Medium>" & vbCrLf
                      Case 1      ' Drop-Down
                        response.write "<SELECT NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium>" & vbCrLf
                        response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                        for i = 0 to Aux_Selection_Max
                          response.write "<OPTION" 
                          if rsProfile("Aux_" & Trim(rsAuxiliary("Order_Num"))) = Trim(Aux_Selection(i)) then
                            response.write " SELECTED"
                          end if                            
                          response.write " CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "</OPTION>" & vbCrLf
                        next
                        response.write "</SELECT>" & vbCrLf
                      Case 2      ' Radio
                        for i = 0 to Aux_Selection_Max
                          response.write "<INPUT"
                          if rsProfile("Aux_" & Trim(rsAuxiliary("Order_Num"))) = Trim(Aux_Selection(i)) then
                            response.write " CHECKED"
                          end if                                                      
                          response.write " TYPE=RADIO NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium VALUE=""" & Trim(Aux_Selection(i)) & """>&nbsp;" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & "&nbsp;&nbsp;" & vbCrLf
                        next
                    end select
                        
                    response.write "</TD>" & vbCrLf
                    response.write "</TR>" & vbCrLf
                  else
                    response.write "<INPUT TYPE=""Hidden"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ VALUE=""" & rsProfile("Aux_" & Trim(rsAuxiliary("Order_Num"))) & """>" & vbCrLf
                  end if
                  
                end if
                
                rsAuxiliary.MoveNext
                
              loop
              
            else
              for x = 0 to 9
                Order_Number = "Aux_" & Trim(CStr(x))
                response.write "<INPUT TYPE=""Hidden"" NAME=""" & Order_Number & """ VALUE=""" & rsProfile(Order_Number) & """>" & vbCrLf                
              next
            end if
            
            rsAuxiliary.Close
            set rsAuxiliary = nothing
            %>             

            <!-- Navigation Buttons -->
            
            <TR>
              <TD COLSPAN=3>
                <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666" BORDER=0>
                  <TR>
                    <TD ALIGN=CENTER WIDTH="40%" CLASS=Medium>
                      &nbsp;
                    </TD>             
                    <TD ALIGN=LEFT WIDTH="60%" CLASS=Medium>&nbsp;
                      <% if CInt(Cart_Mode) = CInt(False) then %>
                      <INPUT TYPE="Submit" CLASS=NavLeftHighLight1 NAME="Submit" VALUE=" <%=Translate("Update",Login_Language,conn)%> " onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'">
                      <% else %>
                      <INPUT TYPE="Submit" CLASS=NavLeftHighLight1 NAME="Submit" VALUE=" <%=Translate("Return to Checkout",Login_Language,conn)%> ">                      
                      <% end if %>
                    </TD>
                  </TR>        
                </TABLE>
              </TD>
            </TR>        
          </TABLE>
        </TD>
      </TR>
    </TABLE>
    <% Call Table_End %>
  </FORM> 
  <BR>
  
<%
else 

  response.write Translate("Your Profile could not be found.  Please contact your Site Administrator by clicking on the Contact Us navigation button and describing this problem",Login_Language,conn) & ".<BR>"

end if
     
rsProfile.close
set rs=nothing

' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------

%>

<!-- #include virtual="/include/core_countries.inc"-->


<SCRIPT LANGUAGE=JAVASCRIPT>
<!--//

function CheckRequiredFields() {

  var ErrorMsg = "";
  var str;
  var strVal;
  var strChk;
  var badchars;
  var D_BadChars;
  var TestRadio;
  var ctr;
  var RadioChecked = 0;
  var dNT = document.NTAccount;
  var Cart_Mode = false;

  if (! dNT.Email.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("EMail",Alt_Language,conn)%> (<%=Translate("Primary",Alt_Language,conn)%>)\r\n";
  }
  if (dNT.Region.value != 2 && <%=CInt(Cart_Mode)%> == <%=CInt(False)%>) {
    if (dNT.Password_Current.value.length > 0 && dNT.Password_Change.value.length > 0 && dNT.Password_Confirm.value.length > 0) {
      if (dNT.Password_Change.value.length < 7 || dNT.Password_Change.value.length > 14 || dNT.Password_Confirm.value.length < 7 || dNT.Password_Confirm.value.length > 14) {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Change Password must be at least 7 characters",Alt_Language,conn)) & ", " & Translate("maximum",Alt_Language,conn) & " 14"%>\r\n";
        dNT.Password_Change.style.backgroundColor = "#FFB9B9";
      }  
      else if (dNT.Password_Change.value != dNT.Password_Confirm.value) {  
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Change Password does not match Change Password Confirm",Alt_Language,conn))%>\r\n";
        dNT.Password_Change.style.backgroundColor  = "#FFB9B9";
        dNT.Password_Confirm.style.backgroundColor = "#FFB9B9";
      }
    }
  }    

  if (! dNT.EMail_Method.value.length) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email Format",Alt_Language,conn))%>\r\n";  
  }

  if (! dNT.Connection_Speed.value.length) {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Internet Connection Speed",Alt_Language,conn))%>\r\n";  
  }

  if (! dNT.FirstName.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("First Name",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.LastName.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Last Name",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Company.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Company Name",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Job_Title.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Job Title",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Address.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Business Address",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_City.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Business City",Alt_Language,conn)%>\r\n";
  }

  if ((! dNT.Business_State.value.length) && (! dNT.Business_State_Other.value.length)) {
    ErrorMsg = ErrorMsg + "<%=Translate("Business State or Other",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Postal_Code.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Business Postal Code",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Country.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Business Country",Alt_Language,conn)%>\r\n";
  }

  if (! dNT.Business_Phone.value.length) {
    ErrorMsg = ErrorMsg + "<%=Translate("Business Phone",Alt_Language,conn)%> 1\r\n";
  }

  <%
    for i = 0 to 9
      if CInt(Aux_Required(i)) = CInt(true)  and CInt(Cart_Mode) = CInt(False) then 
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

    response.write "Cart_Mode = " & CInt(Cart_Mode) & ";" & vbCrLf & vbCrLf
  %>  

  if (Cart_Mode != 0) {
    if (! dNT.Shipping_Address.value.length) {
      ErrorMsg = ErrorMsg + "<%=Translate("Shipping Address",Alt_Language,conn)%>\r\n";
    }
  
    if (! dNT.Shipping_City.value.length) {
      ErrorMsg = ErrorMsg + "<%=Translate("Shipping City",Alt_Language,conn)%>\r\n";
    }
  
    if ((! dNT.Shipping_State.value.length) && (! dNT.Shipping_State_Other.value.length)) {
      ErrorMsg = ErrorMsg + "<%=Translate("Shipping State or Other",Alt_Language,conn)%>\r\n";
    }
  
    if (! dNT.Shipping_Postal_Code.value.length) {
      ErrorMsg = ErrorMsg + "<%=Translate("Shipping Postal Code",Alt_Language,conn)%>\r\n";
    }
  
    if (! dNT.Shipping_Country.value.length) {
      ErrorMsg = ErrorMsg + "<%=Translate("Shipping Country",Alt_Language,conn)%>\r\n";
    }
  }

  if (dNT.Postal_State.value == "" && dNT.Postal_State_Other.value != "") {
    dNT.Postal_State.value = "ZZ";
  }

  if (dNT.Shipping_State.value == "" && dNT.Shipping_State_Other.value != "") {
    dNT.Shipping_State.value = "ZZ";
  }

  if (ErrorMsg.length) {
    ErrorMsg = "<%=Translate("Please enter the missing information for following REQUIRED fields (or use N/A)",Alt_Language,conn)%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    return (false);
  }
  else {
  	return (true);
  } 
}

function Postal_Same() {

  var dNT = document.NTAccount;
  var ErrorMsg = "";

  dNT.Postal_City.value        = dNT.Business_City.value;            
  dNT.Postal_State.value       = dNT.Business_State.value;
  dNT.Postal_State_Other.value = dNT.Business_State_Other.value;
  dNT.Postal_Postal_Code.value = dNT.Business_Postal_Code.value;      
  dNT.Postal_Country.value     = dNT.Business_Country.value;            

  if (dNT.Postal_Address.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Box Number",Login_Language,conn))%>\r\n";
  }

  if (ErrorMsg != "") {
    ErrorMsg = "<%=KillQuote(Translate("Please enter the missing information for following REQUIRED fields (or use N/A (Not Applicable))",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    dNT.Postal_Address.focus();
    return false;
  }
  else {
    ErrorMsg = "<%=KillQuote(Translate("Your Office Address Information has been copied to Postal Address section of this form.",Alt_Language,conn) & "\n\n" & Translate("Please verify that your Postal Address information is correct.",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    dNT.Postal_Address.focus();
  	return true;
  } 
}

function Shipping_Same() {

  var dNT = document.NTAccount;
  var ErrorMsg = "";
    
  dNT.Shipping_MailStop.value    = dNT.Business_MailStop.value;
  dNT.Shipping_Address.value     = dNT.Business_Address.value;        
  dNT.Shipping_Address_2.value   = dNT.Business_Address_2.value;
  dNT.Shipping_City.value        = dNT.Business_City.value;            
  dNT.Shipping_State.value       = dNT.Business_State.value;
  dNT.Shipping_State_Other.value = dNT.Business_State_Other.value;
  dNT.Shipping_Postal_Code.value = dNT.Business_Postal_Code.value;      
  dNT.Shipping_Country.value     = dNT.Business_Country.value;            

  ErrorMsg = "<%=KillQuote(Translate("Your Office Address Information has been copied to Shipping Address section of this form.",Alt_Language,conn) & "\n\n" & Translate("Please verify that your Shipping Address information is correct.",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
  alert (ErrorMsg);

  dNT.Shipping_MailStop.focus();
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

