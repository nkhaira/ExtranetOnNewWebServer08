<%

' --------------------------------------------------------------------------------------
' Update Profile
' --------------------------------------------------------------------------------------  
  
  SQL = "SELECT UserData.*, UserData.ID "
  SQL = SQL & "FROM UserData "
  SQL = SQL & "WHERE (((UserData.ID)=" & CInt(Login_ID) & "))"

  Set rsProfile = Server.CreateObject("ADODB.Recordset")
  rsProfile.Open sql, conn, 3, 3

  if not rsProfile.EOF then    

    response.write "<FONT CLASS=Normal><IMG SRC=""/images/lock.gif"" BORDER=0 WIDTH=37 ALIGN=""ABSMIDDLE"">" & Translate("This is a secure site connection to protect your personal information.",Login_Language,conn) & "</FONT><BR><BR>"

%>

    <FORM NAME="NTAccount" ACTION="/sw-common/profile_admin.asp" METHOD="POST" onsubmit="return(CheckRequiredFields(this.form));">
    <INPUT TYPE="Hidden" NAME="BackURL" Value="<%=BackURLSecure%>">                
    <INPUT TYPE="Hidden" NAME="Account_ID" VALUE="<%=Login_ID%>">
    <INPUT TYPE="Hidden" NAME="ChangeID" VALUE="<%=Login_ID%>">
    <INPUT TYPE="Hidden" NAME="ChangeDate" Value="<%=Date%>">        

    <TABLE WIDTH="100%" BORDER=1 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
  	<TR>
  		<TD WIDTH="100%" BGCOLOR="#EEEEEE">
  			<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>
        
          <!-- Header -->
  				<TR>
          	<TD WIDTH="40%" BGCOLOR="Black" COLSPAN=2 CLASS=MediumBoldGold><%=Translate("Description",Login_Language,conn)%></TD>
  	        <TD WIDTH="60%" BGCOLOR="Black" ALIGN=LEFT CLASS=MediumBoldGold><%=Translate("Account Profile Information",Login_Language,conn)%></TD>
          </TR>
				  <TR>
        	  <TD BGCOLOR="Silver" COLSPAN=3 CLASS=MediumBold><%=Translate("Note",Login_Language,conn)%>:&nbsp;&nbsp;&nbsp;<IMG SRC="/images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> = <%=Translate("Required Information",Login_Language,conn)%>.</TD>
          </TR>        
                    
           <!-- NT Login -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Account Login User Name",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>&nbsp;</TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <%
                response.write "<INPUT CLASS=Medium TYPE=""Hidden"" NAME=""NTLogin"" VALUE=""" & rsProfile("NTLogin") & """>" & rsProfile("NTLogin") & " " & Login_Name
              %>    
            </TD>
          </TR>
  
          <!-- Change Password -->
  
          <% if rsProfile("Password") <> "(hidden)" and not isblank("Password") and rsProfile("NewFlag") <> True then %>
  				<!--TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium><%=Translate("Change Account Password",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>&nbsp;</TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium><INPUT CLASS=Medium TYPE="Text" NAME="Password_Change" SIZE="50" MAXLENGTH="14" VALUE=""></TD>
          </TR-->
          <% end if %>

           <!-- Email -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%response.write Translate("EMail",Login_Language,conn)%> <FONT SIZE="1">(<%=Translate("Primary",Login_Language,conn)%>)</FONT>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Email" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Email")%>">
            </TD>
          </TR>
        
         <!-- Connection Speed -->

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
          

          <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>
                      
           <!-- Name -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium>
              <%=Translate("Name",Login_Language,conn)%> <FONT SIZE=1>(<%=Translate("Prefix",Login_Language,conn)%>, <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><%=Translate("First",Login_Language,conn)%>, <%=Translate("Middle",Login_Language,conn)%>, <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><%=Translate("Last",Login_Language,conn)%>, <%=Translate("Suffix",Login_Language,conn)%>)</FONT>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium NOWRAP>
              <% 
              sValue = rsProfile("Prefix")
              %>
              <SELECT NAME="Prefix" CLASS=Medium>
                <OPTION VALUE=""><%=Translate("Select",Login_Language,conn)%></OPTION>
                <option value="Mr"<% If sValue = "Mr" Then Response.Write " SELECTED" %>><%=Translate("Mr",Login_Language,conn)%></OPTION>
                <option value="Ms"<% If sValue = "Ms" Then Response.Write " SELECTED" %>><%=Translate("Ms",Login_Language,conn)%></OPTION>
                <option value="Miss"<% If sValue = "Miss" Then Response.Write " SELECTED" %>><%=Translate("Miss",Login_Language,conn)%></OPTION>
                <option value="Mrs"<% If sValue = "Mrs" Then Response.Write " SELECTED" %>><%=Translate("Mrs",Login_Language,conn)%></OPTION>
                <option value="Dr"<% If sValue = "Dr" Then Response.Write " SELECTED" %>><%=Translate("Dr",Login_Language,conn)%></OPTION>
              </SELECT>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><INPUT CLASS=Medium TYPE="Text" NAME="FirstName" SIZE="11" MAXLENGTH="50" VALUE="<%=rsProfile("FirstName")%>">&nbsp;&nbsp;<INPUT CLASS=Medium TYPE="Text" NAME="MiddleName" SIZE="2" MAXLENGTH="50" VALUE="<%=rsProfile("MiddleName")%>">&nbsp;<IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE><INPUT CLASS=Medium TYPE="Text" NAME="LastName" SIZE="11" MAXLENGTH="50" VALUE="<%=rsProfile("LastName")%>">&nbsp;&nbsp;<INPUT CLASS=Medium TYPE="Text" NAME="Suffix" SIZE="5" MAXLENGTH="50" VALUE="<%=rsProfile("Suffix")%>">                                
            </TD>
          </TR>
  
           <!-- Company -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Company",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Company" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Company")%>">
            </TD>
          </TR>
  
           <!-- Job Title -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Job Title",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Job_Title" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Job_Title")%>">
            </TD>
          </TR>
  
           <!-- Mail Stop -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Mail Stop",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_MailStop" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_MailStop")%>">
            </TD>
          </TR>
  
          <!-- Business Address -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Address",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT VALIGN=TOP CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                               
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Address")%>"><BR>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Address_2")%>">                
            </TD>
          </TR>

           <!-- Business City -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=MIDDLE CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("City",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_City")%>">
            </TD>
          </TR>

           <!-- Business State -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business State or Province",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <% SValue = rsProfile("Business_State") %>
              <SELECT NAME="Business_State" CLASS=Medium>

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
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Other State, Province or Country",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_State_Other")%>">
            </TD>
          </TR>

           <!-- Business Postal Code -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Postal Code",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Postal_Code")%>">
            </TD>
          </TR>

           <!-- Business Country -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Country",Login_Language,conn)%>:</TD>
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

          <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>

           <!-- Business Phone -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Phone",Login_Language,conn)%> 1:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10" ALIGN=ABSMIDDLE>
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone")%>">&nbsp;&nbsp;<%response.write Translate("Extension",Login_Language,conn)%>: <INPUT CLASS=Medium TYPE="Text" NAME="Business_Phone_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone_Extension")%>">
            </TD>
          </TR>

           <!-- Business Phone 2 -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Business",Login_Language,conn)%>&nbsp;<%=Translate("Phone",Login_Language,conn)%> 2:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Business_Phone_2" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone_2")%>">&nbsp;&nbsp;<%response.write Translate("Extension",Login_Language,conn)%>: <INPUT CLASS=Medium TYPE="Text" NAME="Business_Phone_2_Extension" SIZE="10" MAXLENGTH="50" VALUE="<%=rsProfile("Business_Phone_2_Extension")%>">
            </TD>
          </TR>

           <!-- Mobile Phone -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Mobile",Login_Language,conn)%>&nbsp;<%=Translate("Phone",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Mobile_Phone" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Mobile_Phone")%>">
            </TD>
          </TR>

           <!-- Pager -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Pager",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Pager" SIZE="28" MAXLENGTH="50" VALUE="<%=rsProfile("Pager")%>">
            </TD>
          </TR>

           <!-- Email 2 -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium>E-Mail (<%=Translate("Secondary",Login_Language,conn)%>):</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT CLASS=Medium TYPE="Text" NAME="Email_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Email_2")%>"></FONT>
            </TD>
          </TR>              
          
          <!-- Ship Same -->
          <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP COLSPAN=2 CLASS=MediumRed><%=Translate("If Shipping Address is the same as Business Address",Login_Language,conn)%>:</TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT  CLASS=MediumRed>
              <INPUT CLASS=Medium TYPE="Checkbox" NAME="ShippingSame">&nbsp;&nbsp;<%=Translate("click checkbox",Login_Language,conn)%>
              </FONT>
            </TD>
          </TR>

          <!-- Shipping Mail Stop -->

  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("Mail Stop",Login_Language,conn)%>:
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
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("Address",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_Address" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_Address")%>"><BR>
              <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_Address_2" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_Address_2")%>">
            </TD>
          </TR>

           <!-- Shipping City -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("City",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_City" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_City")%>">
            </TD>
          </TR>

           <!-- Shipping State -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("State",Login_Language,conn)%> / <%=Translate("Province",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
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
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("Other State, Province or County",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
            </TD>                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_State_Other" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_State_Other")%>">
            </TD>
          </TR>


           <!-- Shipping Postal Code -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("Postal Code",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Text" NAME="Shipping_Postal_Code" SIZE="50" MAXLENGTH="50" VALUE="<%=rsProfile("Shipping_Postal_Code")%>">
            </TD>
          </TR>

           <!-- Shipping Country -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Shipping",Login_Language,conn)%>&nbsp;<%=Translate("Country",Login_Language,conn)%>:
            </TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=RIGHT CLASS=Medium>
              &nbsp;
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

          <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>

          <!-- Preferred Language -->
          
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
          
           <!-- Subscription Service -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=Middle CLASS=Medium><%=Translate("Subscription Service",Login_Language,conn)%>:</TD>
          	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
              &nbsp;
            </TD>                                                                
  	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
              <INPUT CLASS=Medium TYPE="Checkbox" NAME="Subscription" <%if rsProfile("Subscription") = True then response.write " CHECKED"%>>
            </TD>
          </TR>

          <!-- EMail Method -->
          
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
          
          <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>
          
          <!-- Auxiliary Fields -->

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

            do while not rsAuxiliary.EOF
              with response
              .write "<TR>"
            	.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=MIDDLE CLASS=Medium>"
              .write Translate(rsAuxiliary("Description"),Login_Language,conn) & ":"
              .write "</TD>"
              .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              end with
              if CInt(rsAuxiliary("Required")) = True then
                Aux_Required(rsAuxiliary("Order_Num")) = True
                Aux_Method(rsAuxiliary("Order_Num"))   = rsAuxiliary("Input_Method")
                response.write "<INPUT TYPE=""HIDDEN"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & "_Description"" VALUE=""" & rsAuxiliary("Description") & """>"                
                response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10 ALIGN=ABSMIDDLE>"
              else
                response.write "&nbsp;"
              end if  
              response.write "</TD>                                                "

 	            response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf

              Dim Aux_Selection
              Dim Aux_Selection_Max

              Aux_Selection     = Split(rsAuxiliary("Radio_Text"),",")
              Aux_Selection_Max = Ubound(Aux_Selection)
              
              Select Case rsAuxiliary("Input_Method")
                Case 0      ' Text
                  response.write "<INPUT TYPE=""Text"" NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ SIZE=""50"" MAXLENGTH=""50"" VALUE="""" CLASS=Medium>" & vbCrLf
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
                    response.write " TYPE=RADIO NAME=""Aux_" & Trim(rsAuxiliary("Order_Num")) & """ CLASS=Medium VALUE=""" & Translate(Trim(Aux_Selection(i)),Login_Language,conn) & """>&nbsp;" & Trim(Aux_Selection(i)) & "&nbsp;&nbsp;" & vbCrLf
                  next
              end select
                  
              response.write "</TD>" & vbCrLf
              response.write "</TR>" & vbCrLf
              
              rsAuxiliary.MoveNext
              
            loop
            
            response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>"
            
          end if
          
          rsAuxiliary.Close
          set rsAuxiliary = nothing
            
          if rsAuxFlag = True then
            response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray""></TD></TR>"
          end if                    

          %>             

          <!-- Navigation Buttons -->
  
          <TR>
          <TD COLSPAN=3>
            <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666">
              <TR>
                <TD ALIGN=CENTER WIDTH="25%" CLASS=Medium>
                  &nbsp;
                </TD>
                <TD ALIGN=CENTER WIDTH="25%" CLASS=Medium>
                  &nbsp;  
                </TD>             
                <TD ALIGN=LEFT WIDTH="25%" CLASS=Medium>
                  <INPUT CLASS=NAVLEFTHIGHLIGHT1 TYPE="Submit" VALUE=" <% response.write Translate("Update",Login_Language,conn)%> ">                
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
  </FORM> 
  <BR>
  
<%
else 

  response.write Translate("Your Profile could not be found.  Please contact your Site Administrator by clicking on the Contact Us navigation button and describing this problem",Login_Language,conn) & ".<BR>"

end if
     
rsProfile.close
set rs=nothing

' --------------------------------------------------------------------------------------
' Connection Includes
' --------------------------------------------------------------------------------------
%>
<!-- #include virtual="/Connections/Connection_FormData.asp" -->
<% 
' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------
%>

<!-- #include virtual="/include/core_countries.inc"-->


<SCRIPT LANGUAGE=JAVASCRIPT>
<!--

function CheckRequiredFields() {

  var ErrorMsg;
  ErrorMsg = "";

  var strVal;
  var strChk;
  var TestRadio;
  var ctr;
  var RadioChecked;
     
  if (document.NTAccount.Email.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email",Alt_Language,conn))%> (<%=ReplaceRSQuote(Translate("Primary",Alt_Language,conn))%>)\r\n";  
  }
  
  if (document.NTAccount.Connection_Speed.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Internet Connection Speed",Alt_Language,conn))%>\r\n";  
  }

  if (document.NTAccount.FirstName.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("First Name",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.LastName.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Last Name",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Company.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Company",Alt_Language,conn))%>\r\n";
  }
  
  if (document.NTAccount.Job_Title.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Job Title",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Business_Address.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Business",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Address",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Business_City.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Business",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("City",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Business_State.value == "" && document.NTAccount.Business_State_Other.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Business USA State or Canadian Province",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Business_Country.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Business",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Country",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Business_Country.value == "US" && document.NTAccount.Business_Postal_Code.value != "") {
    strVal = document.NTAccount.Business_Postal_Code.value;
    if (strVal.length != 5 && strVal.length != 10) {
      ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code - 5 digit or 5 digit + 4",Alt_Language,conn))%>\r\n";
    }
    else {
      for (var i=0; i < strVal.length; i++) {
        strChk = "" + strval.substring(i, i+1);
        if (valid.indexOf(strChk) == "-1") {
          ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code - Invalid Characters",Alt_Language,conn))%>\r\n";
        }
      }
    }  
  }
  
  if (document.NTAccount.Business_Postal_Code.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Business",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Postal Code",Alt_Language,conn))%>\r\n";
  }

  if (document.NTAccount.Business_Phone.value == "") {
    ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Business",Alt_Language,conn))%>\t<%=ReplaceRSQuote(Translate("Phone",Alt_Language,conn))%> 1\r\n";
  }

  <%  
    for i = 0 to 9
      if Aux_Required(i) = true then 
        select case Aux_Method(i)
    			case 0  ' Text
    				response.write "if (document.NTAccount.Aux_" + Trim(CStr(i)) + ".value == """") {" & vbCrLf
    				response.write "  ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + document.NTAccount.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
    				response.write "}" & vbCrLF & vbCrLF
    			case 1  ' Drop-Down
    				response.write "strVal = document.NTAccount.Aux_" + Trim(CStr(i)) + ".value; " & vbCrLF				
    				response.write "strChk = strVal.substring(0, 7); " & vbCrLF 
    				response.write "  if ((strVal == """") || (strChk.toLowerCase() == ""select"")) {" & vbCrLf
    				response.write "     ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + document.NTAccount.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
    				response.write "  }" & vbCrLF & vbCrLF				 
          case 2  ' Radio or Checkbox
            response.write "RadioChecked = 0;" & vbCrLf
            response.write "TestRadio = document.NTAccount.Aux_" + Trim(Cstr(i)) & ";" & vbCrLf
            response.write "for (ctr=0; ctr < TestRadio.length; ctr++) {" & vbCrLf
            response.write "  if (TestRadio[ctr].checked) {" & vbCrLf
            response.write "   RadioChecked = 1;" & vbCrLf
            response.write "   break;" & vbCrLf
            response.write "   }" & vbCrLf
            response.write "}" & vbCrLf            
    				response.write "if (RadioChecked == 0) {" & vbCrLf 
    				response.write "     ErrorMsg = ErrorMsg + ""\n" & Translate("Missing Answer to the following Question",Alt_Language,conn) & ":\r\n"" + document.NTAccount.Aux_" + Trim(CStr(i)) + "_Description.value + ""\r\n"";" & vbCrLF
    				response.write "}" & vbCrLF & vbCrLF
        end select
      end if
    next  
  %>         
    
  if (document.NTAccount.ShippingSame.checked) {
      document.NTAccount.Shipping_MailStop.value    = document.NTAccount.Business_MailStop.value;  
      document.NTAccount.Shipping_Address.value     = document.NTAccount.Business_Address.value;
      document.NTAccount.Shipping_Address_2.value   = document.NTAccount.Business_Address_2.value;
      document.NTAccount.Shipping_City.value        = document.NTAccount.Business_City.value;            
      document.NTAccount.Shipping_State.value       = document.NTAccount.Business_State.value;
      document.NTAccount.Shipping_State_Other.value = document.NTAccount.Business_State_Other.value;
      document.NTAccount.Shipping_Postal_Code.value = document.NTAccount.Business_Postal_Code.value;      
      document.NTAccount.Shipping_Country.value     = document.NTAccount.Business_Country.value;            
  }
    
  if (ErrorMsg != "") {
    ErrorMsg = "<%=KillQuote(Translate("Please enter the missing information for following REQUIRED fields (or use N/A (Not Applicable))",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    return false;
  }
  else {
  	return true;
  } 
}

//-->
</SCRIPT>

