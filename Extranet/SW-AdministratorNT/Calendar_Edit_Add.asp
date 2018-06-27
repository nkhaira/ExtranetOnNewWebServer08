<%
' --------------------------------------------------------------------------------------
' Add New Content Item - Include File for /sw-administratorNT/Calendar_Edit.asp
'
' Author: Kelly Whitlock
' --------------------------------------------------------------------------------------

server.ScriptTimeout=20

Write_Form_Show_Values = False
Call Get_Show_Values

if not isblank(request("Content_Group")) then
  Content_Group = request("Content_Group")
else
  Content_Group = 0
end if

if Content_Group > 0 and not Show_Calendar then
  Field_Editable = " DISABLED"        
else
  Field_Editable = ""
end if

if Content_Group > 0 and not isblank(request("Campaign")) then
  Campaign = request("Campaign")
else
  Campaign = 0
end if

' --------------------------------------------------------------------------------------
' Get Preset Field Values from Campaign
' --------------------------------------------------------------------------------------

Campaign_Status    = 0
Campaign_LDays     = 0
Campaign_BDate     = FormatDate(1,Date)
Campaign_EDate     = FormatDate(1,Date)
Campaign_VDate     = ""
Campaign_PEDate    = ""
Campaign_SubGroups = ""
Campaign_Country   = ""
Campaign_Subscription_Early = CInt(false)

if Campaign or Show_Calendar then

  SQL = "SELECT * FROM Calendar WHERE ID=" & Campaign
  Set rsCampaign = Server.CreateObject("ADODB.Recordset")
  rsCampaign.Open SQL, conn, 3, 3
  
  if not rsCampaign.EOF then
    Campaign_Status    = 0
    Campaign_LDays     = rsCampaign("LDays")
    Campaign_BDate     = rsCampaign("BDate")
    Campaign_EDate     = rsCampaign("EDate")
    Campaign_VDate     = rsCampaign("VDate")    
    Campaign_PEDate    = rsCampaign("PEDate")
    Campaign_SubGroups = rsCampaign("SubGroups")
    Campaign_Country   = rsCampaign("Country")
    Campaign_Subscription_Early = CInt(rsCampaign("Subscription_Early"))
  end if  

  rsCampaign.close
  set rsCampaign = nothing
  set SQL = nothing

end if

FormName = "AddContent"

%>
<FORM NAME="<%=FormName%>" ACTION="/SW-AdministratorNT/Calendar_admin.asp?FileUpEE_Flag=<%=CInt(FileUpEE_Flag)%>&FileUpEE_Remote_Flag=<%=CInt(FileUpEE_Remote_Flag)%>" METHOD="Post" ENCTYPE="MULTIPART/FORM-DATA" onKeyUp="highlight(event)" onClick="highlight(event)" onsubmit="return CheckRequiredFields(this.form);">
<INPUT Type="Hidden" NAME="Show_View" VALUE="<%=Show_View%>">
<INPUT TYPE="Hidden" NAME="ID" VALUE="<%=Calendar_ID%>">
<INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
<INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
<INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
<INPUT TYPE="Hidden" NAME="Clone" VALUE="0">
<INPUT TYPE="Hidden" NAME="Status" VALUE="0">
<INPUT TYPE="Hidden" NAME="PDate" VALUE="<%=FormatDate(1,Date())%>">
<INPUT TYPE="Hidden" NAME="UDate" VALUE="<%=Now()%>">                  
<INPUT TYPE="Hidden" NAME="Show_Calendar" VALUE="<%=CInt(Show_Calendar)%>">    
<INPUT TYPE="Hidden" NAME="Submitted_By" VALUE="<%=Admin_ID%>">
<INPUT TYPE="Hidden" NAME="Admin_Access" VALUE="<%=Admin_Access%>">
<INPUT TYPE="Hidden" NAME="ProgressID" VALUE="<%=ProgressID%>">
<INPUT Type="Hidden" Name="FileUpEE_Flag" VALUE="<%=CInt(FileUpEE_Flag)%>">
<INPUT Type="Hidden" Name="FileUpEE_Remote_Flag" VALUE="<%=CInt(FileUpEE_Remote_Flag)%>">
<INPUT TYPE="Hidden" NAME="Current_Time" VALUE="<%=Now()%>">
<%

' --------------------------------------------------------------------------------------
' Site Information and Paths
' --------------------------------------------------------------------------------------

' Get Names of Content Groupings

SQL = "SELECT * FROM Site WHERE ID=" & CInt(Site_ID)
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3
Site_Code = rsSite("Site_Code")
response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Site"" VALUE=""" & Site_Code & """>"
rsSite.close
set rsSite=nothing
set SQL = nothing

response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Site_POD"" VALUE=""" & Path_Site_POD & """>" & vbCrLf
    
if not Show_View and Content_Group > 0 and Show_Content_Group then
  if Content_Group > 0 or not isblank(request("Content_Group")) or Show_Calendar then
    response.write "<UL><LI><SPAN CLASS=SmallBold>" & Translate("You have designated this Content or Calendar Event item be included in the following Master Asset Container (MAC)",Login_Language,conn) & ": <SPAN CLASS=SmallBoldRed>"
    select case Content_Group
      case 1, 2
        response.write Translate(Code_8000_Name,Login_Language,conn)
      case 3, 4
        response.write Translate(Code_8001_Name,Login_Language,conn)
    end select
    response.write "</SPAN></LI>"
    response.write "<LI>"
    response.write Translate("Please read the following note(s):",Login_Language,conn)
    response.write "</SPAN></LI>"
  end if  
  if Content_Group > 0 and not Show_Calendar then
    response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 1</SPAN><SPAN CLASS=Small>: " & Translate("Certain fields may blocked from editing (indicated by a shaded color).  These values for these blocked fields are automatically updated when you update the Master Asset Container (MAC) that this asset or event is associated with.",Login_Language,conn) & "</SPAN></LI>"
  end if
  if not isblank(request("Content_Group")) then
    response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 2</SPAN><SPAN CLASS=Small>: " & Translate("Because you have changed the &quot;Content Grouping&quot; for an existing record, review your &quot;Groups allowed to view this information&quot; selections and &quot;Country Restrictions&quot;.",Login_Language,conn) & "</SPAN></LI>"
  end if
  if Content_Group > 0 and Show_Calendar then
    response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 3</SPAN><SPAN CLASS=Small>: " & Translate("If this &quot;Calendar Event&quot; will occur before the &quot;Pre-Announce&quot; or &quot;Beginning Date&quot; of the Master Asset Container (MAC) that it is associated with, ensure that the content is generic and does not provide specific details about the nature of the Master Asset Container (MAC), because this Calendar Event will be available for view prior to the actual &quot;Pre-Announce&quot; or &quot;Begining Date&quot; date of the Master Asset Container (MAC).",Login_Language,conn)
    response.write  "&nbsp;&nbsp;" & Translate("This restriction does not apply to a Calendar Event that will occur after the &quot;Pre-Announce&quot; or &quot;Beginning Date&quot;.",Login_Language,conn) & "</SPAN></LI>"
  end if
  if Content_Group > 0 then
    response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 4</SPAN><SPAN CLASS=Small>: " & Translate("Subscription Service and time of the subscription service email sent are controled by the Master Asset Container (MAC).",Login_Language,conn) & "</SPAN></LI>"
  end if
  if Content_Group > 0 or not isblank(request("Content_Group")) or Show_Calendar then
    response.write "</UL>"
  end if  
end if  

' --------------------------------------------------------------------------------------
' Build Add Item Form
' --------------------------------------------------------------------------------------

      Call Table_Begin
%>
      <TABLE WIDTH="100%" BORDER=0 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN="CENTER">
    	<TR>
    		<TD WIDTH="100%" BGCOLOR="#EEEEEE" CLASS=Medium>
    			<TABLE WIDTH="100%" CELLPADDING=4 Border=0>
          
            <!-- Header -->
    				<TR>
            	<TD WIDTH="40%" BGCOLOR="Silver" COLSPAN=2 CLASS=NavLeftSelected1>
                <%=Translate("Field Name / Description",Login_Language,conn)%>
              </TD>
    	        <TD WIDTH="60%" BGCOLOR="Silver" ALIGN=LEFT CLASS=NavLeftSelected1>
                <%=Translate("Content or Event Information",Login_Language,conn)%>
              </TD>
            </TR>
    				<TR>
            	<TD BGCOLOR="Silver" COLSPAN=2 CLASS=SmallBold>
                <%=Translate("Note",Login_Language,conn)%>:&nbsp;&nbsp;&nbsp;<IMG SRC="/images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> <%=Translate("Required Information",Login_Language,conn)%>.
              </TD>
              <TD BGCOLOR="Silver" VALIGN=TOP CLASS=Medium NOWRAP>
                <A HREF="#HELP" TITLE="Click for Help Information on Asset Field Data Entry"><IMG SRC="/images/help_button.gif" BORDER=0 ALIGN=RIGHT VALIGN=TOP></A>
                <INPUT TYPE="Submit" NAME="Nav_Main_Menu" VALUE=" <%=Translate("Main Menu",Login_Language,conn)%> " CLASS=Navlefthighlight1 ONCLICK="Menu_Button = true;">
              </TD>
            </TR>        
    
            <!-- Calendar Event ID -->
    
    		<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Content / Event ID Number",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER WIDTH="2%" CLASS=Medium>
                &nbsp;
              </TD>                                                                               
    	        <TD BGCOLOR="White" CLASS=MediumBoldRed>
              
                <%=UCase(Calendar_ID)%>
              </TD>
            </TR>

      		<!-- Content Grouping -->

            <%
            if Show_Content_Group then
    	    	  response.write "<TR>"
             	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
              response.write Translate("Content Grouping",Login_Language,conn) & ":"
              if (Category_Code < 8000 or Category_Code > 8999) and Admin_Access > 2  and Content_Group = 0 then
                response.write "&nbsp;&nbsp;<SPAN CLASS=SmallRed>(" & Translate("Select before completing the rest of this form",Login_Language,conn) & ")</SPAN>"
              end if
              response.write "</TD>"
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10>"
              response.write "</TD>"
    	        response.write "<TD BGCOLOR=""White"" CLASS=Medium>"
  
              if (Category_Code < 8000 or Category_Code > 8999) and Admin_Access > 2 then
              
                response.write "<SELECT CLASS=Medium LANGUAGE=""JavaScript"" ONCHANGE=""alert('" + Translate("Important Notice",Alt_Language,conn) & "\n\n" &_
                               Translate("If you are associating this asset with a MAC container and there is an expiration date (e.g. Ending Date <> Beginning Date or Move to Archive days <> 0), then this asset will archive on the same date as the controlling PI/C container.",Alt_Language,conn) & "\n\n" &_
                               Translate("To prevent this from occuring, first create an Individual Asset with the same Beginning Date as the MAC Container, clone it, then associate the cloned asset with the MAC container.",Alt_Language,conn) & "\n\n" &_
                               Translate("Use the -Only option and do not use the +Individual option otherwise you will have two identical assets shown in the Library Category for this asset.",Alt_Language,conn) &_
                               "'); window.location.href='/sw-administratorNT/Calendar_edit.asp?ID=add&Site_ID=" & Site_ID & "&Category_ID=" & Category_ID & "&Campaign=" & Campaign & "&Content_Group='+this.options[this.selectedIndex].value"" NAME=""Content_Group"">"
  
                SQL = "SELECT * FROM Content_Group ORDER BY ID"
                Set rsContent_Group = Server.CreateObject("ADODB.Recordset")
                rsContent_Group.Open SQL, conn, 3, 3
  
                do while not rsContent_Group.EOF
               	  response.write "<OPTION"
                  if CInt(Content_Group) = CInt(rsContent_Group("ID")) then
                    response.write " SELECTED"
                  end if
                  select case rsContent_Group("ID")
                    case 0
                      response.write " CLASS=Medium"
                    case 1,2
                      response.write " CLASS=Region1"
                    case 3,4
                      response.write " CLASS=Region2"
                    case else
                      response.write " CLASS=RegionX"
                  end select    
                  response.write " VALUE=""" & rsContent_Group("ID") & """>"
                  response.write Translate(Replace(Replace(rsContent_Group("Group_Name"),"Product Introduction",Code_8000_Name),"Campaign",Code_8001_Name),Login_Language,conn) & "</OPTION>" & vbCrLf
              	  rsContent_Group.MoveNext                   
                loop
  
                rsContent_Group.close
                set rsContent_Group=nothing
                set SQL = nothing
  
                response.write "</SELECT>"
                  
              else
                response.write "<INPUT TYPE=""HIDDEN"" NAME=""Content_Group"" VALUE=""0"">Individual"
              end if
              response.write "</TD>"
              response.write "</TR>"
            else
              response.write "<INPUT TYPE=""HIDDEN"" NAME=""Content_Group"" VALUE=""0"">"
            end if
            %>            
    
      			<!-- Content Grouping PI/Campaign Select-->

            <%
            if Content_Group > 0 and Show_Content_Group then
     	    	  response.write "<TR>" & vbCrLf
            	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"

              select case Content_Group
                case 1, 2   ' Product Introduction Kits
                  Code_X_Name = Translate(Code_8000_Name,Login_Language,conn)
                  response.write Code_X_Name & " " & Translate("Name",Login_Language,conn) & ":"
                  SQL = "SELECT * FROM Calendar WHERE Site_ID=" & Site_ID & " AND Code=8000 AND Link IS NULL ORDER BY Language, Title"
                case 3, 4   ' Campaigns
                  Code_X_Name = Translate(Code_8001_Name,Login_Language,conn)
                  response.write Code_X_Name & " " & Translate("Name",Login_Language,conn) & ":"
                  SQL = "SELECT * FROM Calendar WHERE Site_ID=" & Site_ID & " AND Code=8001 AND Link IS NULL ORDER BY Language, Title"
              end select
              
              if Campaign = 0 then
                response.write "&nbsp;&nbsp;<SPAN CLASS=SmallRed>(" & Translate("Select before completing the rest of this form",Login_Language,conn) & ")</SPAN>"
              end if

              response.write "</TD>" & vbCrLf
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10>"
              response.write "</TD>" & vbCrLf
    	        response.write "<TD BGCOLOR=""White"" CLASS=Medium>" & vbCrLf

              Set rsCampaign = Server.CreateObject("ADODB.Recordset")
              rsCampaign.Open SQL, conn, 3, 3

              if not rsCampaign.EOF then
                %>
                <SELECT CLASS=Medium LANGUAGE="JavaScript" ONCHANGE="window.location.href='/sw-administratorNT/Calendar_edit.asp?ID=add&Site_ID=<%=Site_ID%>&Category_ID=<%=Category_ID%>&Content_Group=<%=Content_Group%>&Campaign='+this.options[this.selectedIndex].value" NAME="Campaign">
                <%
                response.write "<OPTION VALUE=""0"">" & Translate("Select from this list",Login_language,conn) & "</OPTION>" & vbCrLf
                do while not rsCampaign.EOF
                  response.write "<OPTION"
                  
                  select case rsCampaign("Status")
                    case 0
                      response.write " CLASS=Review"
                    case 1
                      response.write " CLASS=Region1"                    
                    case 2
                      response.write " CLASS=Archive"                    
                  end select
                  
                  if CLng(Campaign) = CLng(rsCampaign("ID")) then response.write " SELECTED"
                  response.write " VALUE=""" & rsCampaign("ID") & """>"
                  select case Content_Group
                    case 1, 2   ' Product Introduction Kits
                      response.write "P"
                    case 3, 4   ' Campaigns
                      response.write "C"
                  end select
                  for i = Len(rsCampaign("ID")) + 1 to 5
                    response.write "0"
                  next  
                  response.write rsCampaign("ID") & " - " & UCase(rsCampaign("Language"))
                  if rsCampaign("Country") = "none" then
                    response.write " - "
                  else
                    response.write " x "
                  end if
                  response.write " " & Mid(rsCampaign("Title"),1,35) & "</OPTION>" & vbCrLf
              	  rsCampaign.MoveNext 
                loop
                  
                response.write "</SELECT>" & vbCrLf
              else
                response.write "<INPUT TYPE=""HIDDEN"" NAME=""Campaign"" VALUE=""0"">"
                response.write "<SPAN CLASS=MediumRed>" & Translate("None Available - Change to Individual or Contact Site Administrator", Login_Language,conn) & "</SPAN>" & vbCrLf
              end if
                   
              rsCampaign.close
              set rsCampaign = nothing
              set SQL = nothing

              response.write "</TD>" & vbCrLf
              response.write "</TR>" & vbCrLf
            end if
            %>            

            <!-- Category -->

            <% if Category_ID = false then %>

    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Category",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                  
    	        <TD BGCOLOR="White" CLASS=Medium>
                <SELECT CLASS=Medium LANGUAGE="JavaScript" ONCHANGE="window.location.href='Calendar_edit.asp?ID=<%=Calendar_ID%>&Site_ID=<%=Site_ID%>&Category_ID='+this.options[this.selectedIndex].value" NAME="Category_ID" ONFOCUS="Grouping_Name_Check();">    
              <%
                SQL = "SELECT * FROM Calendar_Category WHERE Site_ID=" & CInt(Site_ID) & " ORDER BY Category"
                Set rsCategory = Server.CreateObject("ADODB.Recordset")
                rsCategory.Open SQL, conn, 3, 3

                response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
                                  
                Do while not rsCategory.EOF            
               	  response.write "<OPTION CLASS=Medium VALUE=""" & rsCategory("ID") & """>" & Translate(rsCategory("Title"),Login_language,conn) & "</OPTION>" & vbCrLf
              	  rsCategory.MoveNext 
                loop
                   
                rsCategory.close
                set rsCategory=nothing
                set SQL = nothing
              %>          
                </SELECT>
              </TD>
            </TR>

            <% else %>
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%
                response.write  Translate("Category",Login_Language,conn) & ":"
                if Content_Group > 0 and Show_Calendar and Show_Content_Group then
                  response.write "&nbsp;&nbsp;<SPAN CLASS=SmallRed>(" & Translate("See Note 3 Above",Login_Language,conn) & ")</SPAN>"
                end if
                %>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                  
    	        <TD BGCOLOR="White" CLASS=Medium>
              <%
                Write_Form_Show_Values = True
                Call Get_Show_Values
              %>
              </TD>
            </TR>

             <!-- Sub-Category -->
            <% if Show_Sub_Category = true then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%
                response.write Translate("Sub-Category",Login_Language,conn) & ":"
            
                if Admin_Access >= 8 then
                  response.write "<BR><BR>"
                  response.write "<SPAN CLASS=Small>" & Translate("Note: New Sub-Categories must be in English",Login_Language,conn) & "</SPAN>"
                end if
                %>  
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <%
                response.write "<SELECT NAME=""Sub_Category_New"" LANGUAGE=""JavaScript"" CLASS=Medium ONFOCUS=""Grouping_Name_Check();"">"

                
                if Admin_Access >= 8 then
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list or enter new below",Login_Language,conn) & "</OPTION>"
                else
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list",Login_Language,conn) & "</OPTION>"                  
                end if

                ' Preset Sub-Categories
                
                SQL = "SELECT Content_Sub_Category.Site_ID, Content_Sub_Category.Sub_Category, Content_Sub_Category.Code, Content_Sub_Category.Language "
                SQL = SQL & "FROM Content_Sub_Category "
                SQL = SQL & "GROUP BY Content_Sub_Category.Site_ID, Content_Sub_Category.Sub_Category, Content_Sub_Category.Code, Content_Sub_Category.Language "
                SQL = SQL & "HAVING Content_Sub_Category.Site_ID=" & Site_ID & " "
                SQL = SQL & "AND Content_Sub_Category.Sub_Category<>'' "
                SQL = SQL & "AND Content_Sub_Category.Code=" & CInt(Category_Code) & " "
                SQL = SQL & "AND Content_Sub_Category.Language='eng'"
                
                'response.write "<P>" & SQL & "<P>"
                'response.flush
                
                Set rsSubCategoryPreset = Server.CreateObject("ADODB.Recordset")
                rsSubCategoryPreset.Open SQL, conn, 3, 3              

                if not rsSubCategoryPreset.EOF then
                  rsSCP = True
                  response.write "<OPTION CLASS=Medium VALUE="""">+++ " & Translate("Preset Sub-Categories",Login_Language,conn) & " +++"
                  Do while not rsSubCategoryPreset.EOF            
                 	  response.write "<OPTION CLASS=Medium VALUE=""" & rsSubCategoryPreset("Sub_Category") & """>" & Translate(rsSubCategoryPreset("Sub_Category"),Login_Language,conn) & "</OPTION>" & vbCrLf
                	  rsSubCategoryPreset.MoveNext 
                  loop
                else
                  rsSCP = False
                end if                     
                rsSubCategoryPreset.close
                set rsSubCategoryPreset = nothing
                set SQL = nothing

                ' Free Form Sub-Categories
                
                SQL = "SELECT Calendar.Site_ID, Calendar.Sub_Category, Calendar.Code, Calendar.Language "
                SQL = SQL & "FROM Calendar "
                SQL = SQL & "GROUP BY Calendar.Site_ID, Calendar.Sub_Category, Calendar.Code, Calendar.Language "
                SQL = SQL & "HAVING Calendar.Site_ID=" & Site_ID & " "
                SQL = SQL & "AND Calendar.Sub_Category<>'' "
                SQL = SQL & "AND Calendar.Code=" & CInt(Category_Code) & " "
                SQL = SQL & "AND Calendar.Language='eng'"
                
                Set rsSubCategory = Server.CreateObject("ADODB.Recordset")
                rsSubCategory.Open SQL, conn, 3, 3
                
                if rsSCP = True and not rsSubCategory.EOF then
                  response.write "<OPTION CLASS=Medium VALUE="""">+++ " & Translate("Alternate Sub-Categories",Login_Language,conn) & " +++"
                end if                     
                              
                Do while not rsSubCategory.EOF            
               	  response.write "<OPTION CLASS=Medium VALUE=""" & rsSubCategory("Sub_Category") & """>" & Translate(rsSubCategory("Sub_Category"),Login_Language,conn) & "</OPTION>" & vbCrLf
              	  rsSubCategory.MoveNext 
                loop
                   
                rsSubCategory.close
                set rsSubCategory=nothing
                set SQL = nothing

                response.write "</SELECT>"
                
                response.write "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF="""" onclick=""Category_Window=window.open('/sw-administratorNT/subcategory_list.asp?Site_ID=" & Site_ID & "&Language=" & Login_Language &  "','Category_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Category_Window.focus();return false;"" CLASS=Medium><SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Matrix",Login_language,conn) & "&nbsp;&nbsp;</SPAN></A>"                

                if Admin_Access >= 8 then                
                  response.write "<BR>"
                  response.write "<INPUT TYPE=""Text"" NAME=""Sub_Category"" SIZE=50 MAXLENGTH=255 VALUE="""" CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"">"
                end if
                %>
              </TD>
            </TR>
            <% end if %>
            
            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>
            
            <!-- Item Number -->
            <%
            if Show_Item_Number = True then
              if Site_ID < 10 then
                Site_ID_Pad = "0" & Trim(CStr(Site_ID))
              else
                Site_ID_Pad = Trim(CStr(Site_ID))
              end if
              SQLin = "SELECT MAX(Item_Number) + 1 AS Next_Number " &_
                      "FROM dbo.Calendar " &_
                      "WHERE (Item_Number >= 9" & Site_ID_Pad & "0000) AND (Item_Number <= 9" & Site_ID_Pad & "9999)"
              Set rsin = Server.CreateObject("ADODB.Recordset")
              rsin.Open SQLin, conn, 3, 3
              if not rsin.EOF then
                if isblank(rsin("Next_Number")) then
                Next_Number = "9" & Site_ID_Pad & "0000"
                else
                  Next_Number = rsin("Next_Number")
                end if
              else
                Next_Number = "9" & Site_ID_Pad & "0000"
              end if
              rsin.close
              set rsin  = nothing
              set SQLin = nothing
            %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%
                if not isblank(Next_Number) then
                  response.write Translate("Item / Reference Number",Login_Language,conn) & "&nbsp;1:&nbsp;<SPAN CLASS=SmallRed>(" & Translate("Oracle or Generic",Login_Language,conn) & ")</SPAN>"
                  response.write "&nbsp;<INPUT TYPE=""BUTTON"" CLASS=NavLeftHighlight1 LANGUAGE=""JavaScript"" NAME=""Generic"" VALUE=""&nbsp;" & Translate("Generic",Login_Language,conn) & "&nbsp;"" ONCLICK=""document." & FormName & ".Item_Number.value='" & Next_Number & "';document." & formname & ".Revision_Code.value='A';document." & formname & ".Revision_Code.focus();"">"
                else  
                  response.write Translate("Item / Reference Number",Login_Language,conn) & "&nbsp;1:&nbsp;<SPAN CLASS=SmallRed>(" & Translate("Oracle Literature Number",Login_Language,conn) & ")</SPAN>"
                end if
                %>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Item_Number" SIZE="50" MAXLENGTH="20" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONCHANGE="ck_item_number();" ONFOCUS="Grouping_Name_Check();">
                &nbsp;&nbsp;<%=Translate("Rev",Login_Language,conn)%>:&nbsp;
                <INPUT TYPE="Text" NAME="Revision_Code" SIZE="1" MAXLENGTH="4" VALUE="" CLASS=Medium>&nbsp;
                <INPUT TYPE="Hidden" NAME="Cost_Center" VALUE="">
                <%
                response.write "<INPUT TYPE=""Checkbox"" CHECKED NAME=""Item_Number_Show"" CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"">&nbsp;&nbsp;" & Translate("Show",Login_Language,conn)
                %>
              </TD>
            </TR>
              <% if (Category_Code < 8000 or Category_Code > 8999) and Show_Item_Number_2 then %>
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Item / Reference Number",Login_Language,conn)%>&nbsp;2:&nbsp;<SPAN CLASS=Small>(<%=Translate("Legacy",Login_Language,conn)%>)</SPAN>
                </TD>
               	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                              
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Item_Number_2" SIZE="50" MAXLENGTH="20" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">
                </TD>
              </TR>
              
              <% end if %>
              
              <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>

            <% else %> 
                <INPUT TYPE="Hidden" NAME="Item_Number" VALUE="">
                <INPUT TYPE="Hidden" NAME="Revision_Code" VALUE="">
                <% if Category_Code < 8000 or Category_Code > 8999 then %>
                  <INPUT TYPE="Hidden" NAME="Item_Number_Show" VALUE="off">
                  <INPUT TYPE="Hidden" NAME="Item_Number_2" VALUE="">
                <% end if %>                  
            <% end if %>

             <!-- Product -->
            <% if Show_Product_Series = true then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%=Translate("Product or Product Family",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <SELECT NAME="Product_New" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">
                <OPTION CLASS=Medium VALUE=""><%=Translate("Select from this list or enter new below",Login_Language,conn)%></OPTION>
                <%
                    SQL = "SELECT Calendar.Site_ID, Calendar.Product, Calendar.Language "
                    SQL = SQL & "FROM Calendar "
                    SQL = SQL & "GROUP BY Calendar.Site_ID, Calendar.Product, Calendar.Language "
                    SQL = SQL & "HAVING Calendar.Site_ID=" & Site_ID & " AND Calendar.Product<>'' AND Calendar.Language='eng'"
                    Set rsProduct = Server.CreateObject("ADODB.Recordset")
                    rsProduct.Open SQL, conn, 3, 3
                                  
                    Do while not rsProduct.EOF
                   	  response.write "<OPTION CLASS=Medium VALUE=""" & RestoreQuote(rsProduct("Product")) & """>" & RestoreQuote(rsProduct("Product")) & "</OPTION>" & vbCrLf
                  	  rsProduct.MoveNext 
                    loop
                       
                    rsProduct.close
                    set rsProduct=nothing
                    set SQL = nothing
                %>
                </SELECT>
                <BR>
                <INPUT TYPE="Text" NAME="Product" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">
              </TD>
            </TR>
            <% else %>
                <INPUT TYPE="Hidden" NAME="Product" VALUE="General">            
            <% end if %>
                                              
            <!-- Title -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Title",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium VALIGN=MIDDLE>
                <INPUT TYPE="Hidden" NAME="Title_B64" VALUE="">
                
                <%
                'Modified  by zensar for MaxLength validation.
                if Show_PID = true then
                     if  PID_System = 0 then  %>
                        <INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="128" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();"> 
                    <% elseif PID_System = 1 then  %>
                        <INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">  
                    <%  end if
                else%>
                   <INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">  
                <%
                End if 
                %>
                <!--<INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">-->
                <%'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> %>
              </TD>
            </TR>
            
    <%
            ' Description
            
            response.write "<TR>"
            response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
            response.write Translate("Description",Login_Language,conn) & ":"
            response.write "</TD>"
            response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
    	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium VALIGN=MIDDLE>"
            MaxLength = 2000 - 50
            response.write "<INPUT TYPE=""Hidden"" NAME=""Description_B64"" VALUE="""">"
            response.write "<TEXTAREA NAME=""Description"" COLS=53 ROWS=6 CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"" ONKEYUP=""if (this.value.length > " & MaxLength & ") { alert('You have exceeded the maximum characters allowed for this field.\r\n\nMaximum Characters allowed = " & MaxLength & "');this.value = this.value.substring(0," & MaxLength & ");};""></TEXTAREA>"
            Response.Write "<input type=hidden name=""opr"" value="""">"
            if Admin_Access >= 4 and Admin_RTE_Enabled = true then
              response.write "&nbsp;&nbsp;"
              RTE_Length = MaxLength
              RTE_Cols   = 53
              RTE_Rows   = 6
              Element = "Description"
              response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""HTML"" LANGUAGE=""JavaScript"" ONCLICK=""RTEditor_Open('" & FormName & "','" & Element & "','" & Site_ID & "','" & Site_Code & "','" & RTE_Length & "','" & RTE_Cols & "','" & RTE_Rows & "');"" TITLE=""Edit Field with the HTML Editor"">" & vbCrLf
            end if
              
            response.write "</TD>"
            response.write "</TR>"

            ' Special Instructions
            
            if Show_Special_Instructions then
              response.write "<TR>"
              response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              response.write Translate("Special Instructions",Login_Language,conn) & ":"
              response.write "</TD>"
              response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium VALIGN=MIDDLE>"
              MaxLength = 500 - 50
              response.write "<INPUT TYPE=""Hidden"" NAME=""Instructions_B64"" VALUE="""">"            
              response.write "<TEXTAREA NAME=""Instructions"" COLS=53 ROWS=6 CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"" ONKEYUP=""if (this.value.length > " & MaxLength & ") { alert('You have exceeded the maximum characters allowed for this field.\r\n\nMaximum Characters allowed = " & MaxLength & "');this.value = this.value.substring(0," & MaxLength & ");};""></TEXTAREA>"
              
              if Admin_Access >= 4 and Admin_RTE_Enabled = true then
                response.write "&nbsp;&nbsp;"
                RTE_Length = MaxLength
                RTE_Cols   = 53
                RTE_Rows   = 6
                Element = "Instructions"
                response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""HTML"" LANGUAGE=""JavaScript"" ONCLICK=""RTEditor_Open('" & FormName & "','" & Element & "','" & Site_ID & "','" & Site_Code & "','" & RTE_Length & "','" & RTE_Cols & "','" & RTE_Rows & "');"" TITLE=""Edit Field with the HTML Editor"">" & vbCrLf
              end if  
  
              response.write "</TD>"
              response.write "</TR>"
            end if

            ' Splash Header / Footer - PIK / Campaign
            
            if Category_Code >= 8000 and Category_Code <= 8999 then

              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>"

              ' Header
              response.write "<TR>"
              response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              select case Category_Code
                case 8000  ' PIK
                  response.write Translate(Code_8000_Name,Login_Language,conn)
                case 8001  ' Campaign
                  response.write Translate(Code_8001_Name,Login_Language,conn)
              end select
              response.write " - " & Translate("Splash Header",Login_Language,conn) & ":"
              response.write "</TD>"
              response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              MaxLength = 1000 - 50
              response.write "<INPUT TYPE=""Hidden"" NAME=""Splash_Header_B64"" VALUE="""">"
              response.write "<TEXTAREA NAME=""Splash_Header"" COLS=53 ROWS=8 CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"" ONKEYUP=""if (this.value.length > " & MaxLength & ") { alert('You have exceeded the maximum characters allowed for this field.\r\n\nMaximum Characters allowed = " & MaxLength & "');this.value = this.value.substring(0," & MaxLength & ");};""></TEXTAREA>"
              
              if Admin_Access >= 4 and Admin_RTE_Enabled = true then
                response.write "&nbsp;&nbsp;"
                RTE_Length = MaxLength
                RTE_Cols   = 53
                RTE_Rows   = 8
                Element = "Splash_Header"
                response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""HTML"" LANGUAGE=""JavaScript"" ONCLICK=""RTEditor_Open('" & FormName & "','" & Element & "','" & Site_ID & "','" & Site_Code & "','" & RTE_Length & "','" & RTE_Cols & "','" & RTE_Rows & "');"" TITLE=""Edit Field with the HTML Editor"">" & vbCrLf
              end if
                
              response.write "</TD>"
              response.write "</TR>"

              ' Introduction Letter ID Number
              response.write "<TR>"
            	response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              select case Category_Code
                case 8000  ' PI
                  response.write Translate(Code_8000_Name,Login_Language,conn) & " - " & Translate("Letter ID Number",Login_Language,conn)
                case 8001  ' Campaign
                  response.write Translate(Code_8001_Name,Login_Language,conn) & " - " & Translate("Letter ID Number",Login_Language,conn)
              end select
              response.write "</TD>"
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "&nbsp;"
              response.write "</TD>"
    	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<INPUT TYPE=""Text"" NAME=""Item_Number_2"" SIZE=""50"" MAXLENGTH=""20"" VALUE="""" CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"">"
              response.write "</TD>"
              response.write "</TR>"
              
              ' Footer
              response.write "<TR>"
              response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              select case Category_Code
                case 8000  ' PIK
                  response.write Translate(Code_8000_Name,Login_Language,conn)
                case 8001  ' Campaign
                  response.write Translate(Code_8001_Name,Login_Language,conn)
              end select
              response.write " - " & Translate("Splash Footer",Login_Language,conn) & ":"
              response.write "</TD>"
              response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              MaxLength = 500 - 50
              response.write "<INPUT TYPE=""Hidden"" NAME=""Splash_Footer_B64"" VALUE="""">"
              response.write "<TEXTAREA NAME=""Splash_Footer"" COLS=53 ROWS=8  CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"" ONKEYUP=""if (this.value.length > " & MaxLength & ") { alert('You have exceeded the maximum characters allowed for this field.\r\n\nMaximum Characters allowed = " & MaxLength & "');this.value = this.value.substring(0," & MaxLength & ");};""></TEXTAREA>"
              
              if Admin_Access >= 4 and Admin_RTE_Enabled = true then
                response.write "&nbsp;&nbsp;"
                RTE_Length = MaxLength
                RTE_Cols   = 53
                RTE_Rows   = 6
                Element = "Splash_Footer"
                response.write "<INPUT CLASS=NavLeftHighlight1 TYPE=""BUTTON"" VALUE=""HTML"" LANGUAGE=""JavaScript"" ONCLICK=""RTEditor_Open('" & FormName & "','" & Element & "','" & Site_ID & "','" & Site_Code & "','" & RTE_Length & "','" & RTE_Cols & "','" & RTE_Rows & "');"" TITLE=""Edit Field with the HTML Editor"">" & vbCrLf
              end if
                
              response.write "</TD>"
              response.write "</TR>"
              
            end if
            %>

            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>
            
            <!-- Location -->
    
            <% if Show_Location = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Location",Login_Language,conn)%> <SPAN CLASS=Small(<%=Translate("City",Login_Language,conn)%>, <%=Translate("State",Login_Language,conn)%> <%=Translate("Country",Login_Language,conn)%>)</SPAN>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Hidden" NAME="Location_B64" VALUE="">
                <INPUT TYPE="Text" NAME="Location" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">
              </TD>
            </TR>
            <% end if %>
    
            
            <%
            ' Image Store Locator ID
            if Show_ImageStore = True then
            
              SQLSite = "SELECT * FROM Site WHERE Site_Code='Image-Store'"
              Set rsSite = Server.CreateObject("ADODB.Recordset")
              rsSite.Open SQLSite, conn, 3, 3
  
              if not rsSite.EOF then
                Link_Name    = Replace(rsSite("URL"),"https://support.fluke.com","https://" & request.ServerVariables("SERVER_NAME"))
                Link_Name    = Replace(Link_Name,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))                           

              end if
              rsSite.close
              set rsSite = nothing
              set SQLSite = nothing

    				  response.write "<TR>" & vbCrLf
            	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
              response.write Translate("Image Store Reference Number",Login_Language,conn) & ":"
              response.write "</TD>" & vbCrLf
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "&nbsp;"
              response.write "</TD>" & vbCrLf
    	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium NOWRAP>"
              response.write "<INPUT TYPE=""Text"" NAME=""Image_Locator"" SIZE=""30"" MAXLENGTH=""255"" VALUE="""" CLASS=Medium LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"">"
              response.write "&nbsp;&nbsp;" & Translate("Search",Login_Language,conn) & ":&nbsp;"
              response.write "<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" CLASS=Medium><SPAN CLASS=NavLeftHighlight1 Title=""Search for Individual Image. Remember to Copy the Image's Object ID into the Image Store Reference Number field."">&nbsp;&nbsp;" & Translate("Individual",Login_language,conn) & "&nbsp;&nbsp;</SPAN></A>"
              response.write "&nbsp;&nbsp;" & Translate("New",Login_Language,conn) & ":&nbsp;"
              response.write "<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=NEW&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" CLASS=Medium><SPAN CLASS=NavLeftHighlight1 Title=""Create New Image Collection. Remember to Copy the Collection's Object ID into the Image Store Reference Number field."">&nbsp;&nbsp;" & Translate("Collection",Login_language,conn) & "&nbsp;&nbsp;</SPAN></A>"
              response.write "</TD>" & vbCrLf
              response.write "</TR>" & vbCrLf
              
              Show_Link = False
              Show_Link_PopUp_Disabled = False
              
            end if
            %>           

             <!-- Link -->
             
            <% if Show_Link = True or Show_Forum = True then %>    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("URL to Web Page",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                if Show_Forum = True then
                  %>
                  <INPUT TYPE="Text" NAME="URLLink" SIZE="50" MAXLENGTH="255" VALUE="http://<%=Request("SERVER_NAME")%>/SW-Forums/Default.asp" CLASS=Medium>
                  <%
                else
                  %>
                  <INPUT TYPE="Text" NAME="URLLink" SIZE="50" MAXLENGTH="255" VALUE="" CLASS=Medium LANGUAGE="JavaScript" ONFOCUS="Grouping_Name_Check();">
                  <%
                end if
                %>
              </TD>
            </TR>
            <% else %>
                <INPUT TYPE="HIDDEN" NAME="URLLink" VALUE="">            
            <% end if %>           
    
             <!-- Link PopUp Window Disable -->
             
            <% if Show_Link_PopUp_Disabled = True then %>    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("URL to Web Page Pop-Up Window",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<TABLE WIDTH=""100%"" BORDER=0>"
                response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"" CLASS=Medium></TD><TD CLASS=Medium>" & Translate("Disable",Login_Language,conn) & "</TD></TR>"                
                response.write "</TABLE>"
                %> 
              </TD>
            </TR>
            <% end if %>

             <!-- Include File -->
    
            <% if Show_Include = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Content File",Login_Language,conn)%> <SPAN CLASS=Small>(HTM or ASP)</SPAN>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Include"" VALUE=""" & Path_Include & """>"
                response.write "<INPUT TYPE=""File"" NAME=""Include"" SIZE=""30"" MAXLENGTH=""255"" CLASS=Medium onblur=""Check_Filename(this);"" LANGUAGE=""JavaScript"" ONFOCUS=""Grouping_Name_Check();"">"
                response.write "<INPUT TYPE=""Hidden"" NAME=""Include_Existing"" VALUE="""">"
                %>
              </TD>
            </TR>
            <% end if %>

             <!-- Upload File -->
    
            <% if Show_File = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium VALIGN=TOP>
                <%
                response.write Translate("Asset File",Login_Language,conn) & " - " & Translate("(LOW Resolution)",Login_Language,conn) & ": "
                response.write "<SPAN CLASS=SmallRed>(" & Translate("Virus Scan <B>prior</B> to uploading file",Login_Language,conn) & ")</SPAN>"
                %>                 
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File"" VALUE=""" & Path_File & """>"
                response.write "<INPUT TYPE=""File"" NAME=""File_Name"" SIZE=""30"" MAXLENGTH=""255"" CLASS=Medium onblur=""Check_Filename(this);"">"
                 '****Ri - 514 Commited By Zensar
                'if CInt(Path_Site_Secure) = CInt(true) then               
                 ' response.write "&nbsp;&nbsp;<INPUT TYPE=HIDDEN NAME=""Secure_Stream"" VALUE=""on"">"
                'else
                 ' response.write "&nbsp;&nbsp;<INPUT TYPE=CHECKBOX NAME=""Secure_Stream"" TITLE=""" & Translate("Please see note in help section below before using this option.",Alt_Language,conn) & """>&nbsp;" & Translate("Secure Stream",Login_Language,conn)
                'end if  

                response.write "<INPUT TYPE=""Hidden"" NAME=""File_Existing"" VALUE="""">"
                %>
              </TD>
            </TR>
            <% else
                response.write "<INPUT TYPE=""Hidden"" NAME=""File_Name"" VALUE="""">"                
                response.write "<INPUT TYPE=""Hidden"" NAME=""File_Existing"" VALUE="""">"                                
               end if %>

            <!-- Upload POD File -->
    
            <% if Show_File_POD = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%
                response.write Translate("Everett - Marketing Communications Use Only.",Login_Language,conn) & "<BR>"
                response.write Translate("Asset File",Login_Language,conn) & " - " & Translate("(POD Resolution)",Login_Language,conn) & ": "
                response.write "<SPAN CLASS=SmallRed>(" & Translate("Virus Scan <B>prior</B> to uploading file",Login_Language,conn) & ")</SPAN>"
                %>                 
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File_POD"" VALUE=""" & Path_File_POD & """>"
                response.write "<INPUT TYPE=""File"" NAME=""File_Name_POD"" SIZE=""30"" MAXLENGTH=""255"" CLASS=Medium onblur=""Check_Filename(this);"" ONCHANGE=""Check_POD_File(this);"">"
                response.write "<INPUT TYPE=""Hidden"" NAME=""File_Existing_POD"" VALUE="""">"                
                %>
              </TD>
            </TR>
            <% else
                response.write "<INPUT TYPE=""Hidden"" NAME=""File_Name_POD"" VALUE="""">"                
                response.write "<INPUT TYPE=""Hidden"" NAME=""File_Existing_POD"" VALUE="""">"                
               end if %>

             <!-- Thumbnail File -->
    
            <% if Show_Thumbnail = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium VALIGN=TOP>
                <%=Translate("Thumbnail File",Login_Language,conn)%> <SPAN CLASS=Small> - (GIF or JPG): </SPAN>
                <%response.write "&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallRed>(" & Translate("Virus Scan <B>prior</B> to uploading file",Login_Language,conn) & ")</SPAN>"%>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Thumbnail"" VALUE=""" & Path_Thumbnail & """>"
                response.write "<INPUT TYPE=""Hidden"" NAME=""Thumbnail_Existing"" VALUE="""">"
                response.write "<INPUT TYPE=""File"" NAME=""Thumbnail"" SIZE=""30"" MAXLENGTH=""255"" CLASS=Medium onblur=""Check_Filename(this);"">"
                response.write "&nbsp;&nbsp;<INPUT TYPE=""Checkbox"" NAME=""Thumbnail_Request"">&nbsp;&nbsp;" & Translate("Request Thumbnail",Login_Language,conn)
                %>
              </TD>
            </TR>
            <% else
                 response.write "<INPUT TYPE=""Hidden"" NAME=""Thumbnail"" VALUE="""">"
               end if %>
                                
            <% if Show_Location = True or Show_Link = True or Show_File = True or Show_File_POD = True or Show_Include = True or Show_Thumbnail = True then %>
              <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                        
            <% end if %>
            
            <!-- Forum ID -->
    
            <% if Show_Forum = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Forum ID Number",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Forum_ID" SIZE="50" MAXLENGTH="10" VALUE="" CLASS=Medium>
              </TD>
            </TR>
            
            <!-- Forum Moderated -->

    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Forum Moderated",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""Forum_Moderated"">"

                  response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("Moderator Name",Login_Language,conn) & ": "
                  
                  response.write "<SELECT NAME=""Forum_Moderator_ID"">" & vbCrLf
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLF
                  response.write "<OPTION CLASS=Medium VALUE=""""></OPTION>" & vbCrLF

                  ' Forum Moderators (Primary)
                  SQL = "SELECT UserData.ID, UserData.SubGroups, UserData.FirstName, UserData.LastName, UserData.Region FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.Subgroups LIKE '%forum%' ORDER BY UserData.LastName, UserData.FirstName"
                  Set rsModerator = Server.CreateObject("ADODB.Recordset")
                  rsModerator.Open SQL, conn, 3, 3

                  if not rsModerator.EOF then

                    response.write "<OPTION CLASS=NavLeftHighlight1 VALUE="""">" & Translate("Primary",Login_Language,conn) & "</OPTION>" & vbCrLf

                    do while not rsModerator.EOF
                      response.write "<OPTION CLASS=Region" & rsModerator("Region") & "NavMedium VALUE=""" & rsModerator("ID") & """"
                      response.write ">" & rsModerator("LastName") & ", " & rsModerator("FirstName") & "</OPTION>" & vbCrLf
                      
                      rsModerator.MoveNext
                      
                    loop
                    
                  end if
                  
                  rsModerator.close
                  set rsModerator = nothing
                  set SQL = nothing
                   
                  ' Forum Moderators (Alternate)
                  SQL = "SELECT UserData.ID, UserData.SubGroups, UserData.FirstName, UserData.LastName, UserData.Region FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND (UserData.SubGroups LIKE '%administrator%' OR UserData.SubGroups LIKE '%content%') ORDER BY UserData.LastName, UserData.FirstName"
                  Set rsModerator = Server.CreateObject("ADODB.Recordset")
                  rsModerator.Open SQL, conn, 3, 3

                  if not rsModerator.EOF then

                    response.write "<OPTION CLASS=NavLeftHighlight1 VALUE="""">" & Translate("Alternates",Login_Language,conn) & "</OPTION>" & vbCrLf

                    do while not rsModerator.EOF
                      response.write "<OPTION CLASS=Region" & rsModerator("Region") & "NavMedium VALUE=""" & rsModerator("ID") & """"
                      response.write ">" & rsModerator("LastName") & ", " & rsModerator("FirstName") & "</OPTION>" & vbCrLf
                      
                      rsModerator.MoveNext
                      
                    loop
                    
                  end if
                  
                  rsModerator.close
                  set rsModerator = nothing
                  set SQL = nothing
                  
                  response.write "</SELECT>" & vbCrLf & vbCrLf

                %>
              </TD>
            </TR>

            <% end if %>
            
            <% if Show_Forum = True then %>
              <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                        
            <% end if %>
            
            <%
              ' Override MAC Dates
              if Content_Group > 0 and not Show_Calendar and Show_Content_Group then
      				  response.write "<TR>"
         				response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
         				response.write Translate("Override MAC Date",Login_Language,conn) & ":"
                response.write "</TD>"
               	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
                response.write "&nbsp;<INPUT TYPE=""Checkbox"" NAME=""SubGroups"""
'                if instr(1,lcase(rs("SubGroups")),lcase("nomac")) > 0 then
'                  response.write " CHECKED"
'                end if  
                response.write " CLASS=Medium VALUE=""nomac"""
                response.write " LANGUAGE=""JavaScript"" ONCLICK=""MAC_Date_Override();"""
                response.write ">&nbsp;&nbsp;" & Translate("Enable",Login_language,conn)
                response.write "</TD>"
                response.write "</TR>"
              end if

              ' Pre Announcement Days before BDate
              
            if Show_Date_Basic = false then 
    				  response.write "<TR>"
       				response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
       				response.write Translate("Pre-Announce",Login_Language,conn) & ":"
              response.write "</TD>"
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
    	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<INPUT " & Field_Editable & " TYPE=""Text"" NAME=""LDays"" SIZE=""30"" MAXLENGTH=""3"" VALUE=""" & Campaign_LDays & """ CLASS=Medium>&nbsp;&nbsp;" & Translate("days before",Login_Language,conn)
              response.write "</TD>"
              response.write "</TR>"
            end if
            %>

             <!-- Beginning Date -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>                
                <%=Translate("Beginning Date",Login_Language,conn)%> <SPAN CLASS=Small>(mm/dd/yyyy)</SPAN>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT VALIGN=TOP CLASS=Medium>
                <INPUT <%=Field_Editable%> TYPE="Text" NAME="BDate" SIZE="30" MAXLENGTH="10" VALUE="<%=Campaign_BDate%>" CLASS=Medium>&nbsp;&nbsp;
                <%
                if isblank(trim(Field_Editable)) then
                  %>
                  <A HREF="javascript:void()" LANGUAGE="JavaScript" onClick="window.dateField = document.<%=FormName%>.BDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false"><IMG SRC="/images/calendar/calendar_icon.gif" BORDER=0 HEIGHT="21"ALIGN=TOP></A>&nbsp;&nbsp;
                  <%
                end if
                if Show_Date_Basic = false then
                  response.write Translate("through",Login_Language,conn)
                end if
                %>
              </TD>
            </TR>
            
            <% if Show_Date_Basic = false then %>
            
             <!-- Ending Date -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>                
                <%=Translate("Ending Date",Login_Language,conn)%> <SPAN CLASS=Small>(mm/dd/yyyy)</SPAN>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT <%=Field_Editable%> TYPE="Text" NAME="EDate" SIZE="30" MAXLENGTH="10" VALUE="<%=Campaign_EDate%>" CLASS=Medium>&nbsp;&nbsp;
                <%
                if isblank(trim(Field_Editable)) then
                  %>
                  <A HREF="javascript:void()" LANGUAGE="JavaScript" onClick="window.dateField = document.<%=FormName%>.EDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false"><IMG SRC="/images/calendar/calendar_icon.gif" BORDER=0 HEIGHT="21"ALIGN=TOP></A>&nbsp;&nbsp;
                  <%
                end if
                response.write Translate("then",Login_Language,conn)
                %>
              </TD>
            </TR>        
                
             <!-- Post Announcement Days After EDate -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Move to Archive",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="XDays" SIZE="30" MAXLENGTH="3" VALUE="0" CLASS=Medium>&nbsp;&nbsp<%=Translate("days after ending date",Login_Language,conn)%>
              </TD>
            </TR>
            <% else %>
                <INPUT TYPE="Hidden" NAME="LDays" VALUE="0" CLASS=Medium>
                <INPUT TYPE="Hidden" NAME="EDate" VALUE="" CLASS=Medium>            
                <INPUT TYPE="Hidden" NAME="XDays" VALUE="0" CLASS=Medium>
            <% end if %>
                
<%if 1=2 then%>
            <!-- EEF Embargo Date -->
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Digital Library Release Date",Login_Language,conn) & " <SPAN Class=Small>(mm/dd/yyyy):<BR>(" & Translate("Leave blank if same as Beginning Date",Login_Language,conn) & ")</SPAN>"%>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<INPUT " & Field_Editable & " TYPE=""Text"" NAME=""VDate"" SIZE=""30"" MAXLENGTH=""10"" VALUE=""" & Campaign_VDate & """ CLASS=Medium>&nbsp;&nbsp;"
                if isblank(trim(Field_Editable)) then
                  %>
                  <A HREF="javascript:void()" LANGUAGE="JavaScript" onClick="window.dateField = document.<%=FormName%>.VDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false"><IMG SRC="/images/calendar/calendar_icon.gif" BORDER=0 HEIGHT="21"ALIGN=TOP></A>&nbsp;&nbsp;
                  <%
                end if
              response.write "</TD>"
            response.write "</TR>"
            
end if            
            %>

            <!-- Public Embargo Date -->
            <% if Show_Date_PRD then %>
      				<TR>
              	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                  <%=Translate("Public Release Date",Login_Language,conn) & " <SPAN Class=Small>(mm/dd/yyyy):<BR>(" & Translate("Leave blank if same as Beginning Date",Login_Language,conn) & ")</SPAN>"%>
                </TD>
               	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                              
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  response.write "<INPUT " & Field_Editable & " TYPE=""Text"" NAME=""PEDate"" SIZE=""30"" MAXLENGTH=""10"" VALUE=""" & Campaign_PEDate & """ CLASS=Medium>&nbsp;&nbsp;"
                  if isblank(trim(Field_Editable)) then
                    %>
                    <A HREF="javascript:void()" LANGUAGE="JavaScript" onClick="window.dateField = document.<%=FormName%>.PEDate;calendar = window.open('/sw-common/sw-calendar_picker.asp','cal','WIDTH=200,HEIGHT=250');return false"><IMG SRC="/images/calendar/calendar_icon.gif" BORDER=0 HEIGHT="21"ALIGN=TOP></A>&nbsp;&nbsp;
                    <%
                  end if
                response.write "</TD>"
              response.write "</TR>"
            else %>
              <INPUT TYPE="HIDDEN" NAME="PEDate" VALUE="">
            <% end if %>

            <!-- Mark as Confidential -->            

            <% if Show_Mark_Confidential then %>
      				   <TR>
              	<TD BGCOLOR="#EEEEEE">
                  <%=Translate("Mark as Confidential",Login_Language,conn)%>:
                </TD>
               	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                              
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  response.write "<TABLE WIDTH=""100%"" BORDER=0>"
                  response.write "<TR>"
                  response.write "<TD WIDTH=20 BGCOLOR=""Red"" CLASS=Medium>"
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""Confidential"" CLASS=Medium>"
                  response.write "</TD>"
                  response.write "<TD CLASS=Small>&nbsp;</TD></TR>"
                  response.write "</TABLE>"
                  %> 
                </TD>
              </TR>
            <% end if %>

            <TR>
              	<TD BGCOLOR="#EEEEEE">
                  <%=Translate("Price list access codes(separated by comma):<br>",Login_Language,conn)%>
                  <%="<SPAN Class=Small>(" & Translate("This field is applicable for Price Lists category only",Login_Language,conn) & ")</SPAN>"%>
                </TD>
               	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                              
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <%
                  response.write "<TABLE WIDTH=""100%"" BORDER=0>"
                  response.write "<TR>"
                  response.write "<TD WIDTH=20 CLASS=Medium>"
                  response.write "<INPUT TYPE=""textbox"" size = ""100"" MAXLENGTH = ""1000"" NAME=""txtAccessCode"" CLASS=Medium>"
                  response.write "</TD>"
                  response.write "<TD CLASS=Small>&nbsp;</TD></TR>"
                  response.write "</TABLE>"
                  %> 
                </TD>
            </TR>

            <!-- Language -->
            
           <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>            
            <!-- PCAT Interface -->
            <%
            ''response.write "Show_PID:-" & CInt(Show_PID)
            ''response.write "PID_System:-" & CInt(PID_System)
              if Show_PID = true then
                if  PID_System = 0 then 
            %>
                  <!-- #include virtual="/sw-administratorNT/SW-PCAT_FNET.asp" -->
            <%  elseif PID_System = 1 then %>
                  <!-- #include virtual="/sw-administratorNT/SW-PCAT_FIND.asp"-->
            <%  end if
              end if
            %> 
            
            <%
            ''Nitin Code Changes Start
            with Response
    .write "<TR>" & vbCrLf
    .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    .write Translate("Partner Portal Categories",Login_Language,conn) & ":"
    .write "</TD>" & vbCrLf

    'Required Icon or Space
    .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
    .write      "&nbsp;" 
    .write "</TD>"

    .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"

    .write "<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0>" & vbCrLf
    .write "<TR>" & vbCrLf
    
    .write "<TD CLASS=Medium WIDTH=""48%"">" & vbCrlf
    .write "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"" height=""80"" WIDTH=""100%"" CLASS=NavLeftHighlight1>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD CLASS=Medium>"
    .write        Translate("Available Categories",Login_Language,conn)
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD>" & vbCrLf
    if showpcat=true then  
        .write "<SELECT  size=""5"" multiple name=""PCat_APortalCats"" CLASS=Medium>" & vbCrLf
    else
        .write "<SELECT  size=""5"" multiple name=""PCat_APortalCats"" CLASS=Medium>" & vbCrLf
    end if

    strSQL123 = "SELECT ID, Code, Title FROM dbo.Calendar_Category WHERE (Site_ID = " & CInt(Site_ID) & ") AND (Code <> 9999)"
    Set rsCategory1 = Server.CreateObject("ADODB.Recordset")
    rsCategory1.Open strSQL123, conn, 3, 3
    
    if not rsCategory1.EOF then
      Do while not rsCategory1.EOF
        .Write "<OPTION value=""" & rsCategory1("ID") & """>" & rsCategory1("Title") & "</OPTION>" & vbCrLf
        rsCategory1.MoveNext
      loop
    end if

    rsCategory1.close
    set rsCategory1 = nothing

    .write "			</SELECT>" & vbCrLf
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>	
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "<TD width=""2%"">" & vbCrLf
    .write "<TABLE height=""61"" cellSpacing=""1"" cellPadding=""1"" border=""0"">" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "  	<TD CLASS=Medium ALIGN=CENTER>" & vbCrLf
    .write "      &nbsp;<INPUT type=""button"" value="">"" CLASS=NavLeftHighlight1 name=""btnAPortalCats"" onclick=""AddRemoveOptions('PCat_APortalCats','PCat_SPortalCats')"">&nbsp;" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "	<TR>" & vbCrLf
    .write "		<TD CLASS=Medium ALIGN=CENTER>" & vbCrLf

    .write "      &nbsp;<INPUT type=""button"" value=""<"" CLASS=NavLeftHighlight1 name=""btnRPortalCats"" onclick=""RemoveOption('PCat_SPortalCats')"">&nbsp;" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "<TD WIDTH=""48%"">" & vbCrLf
    .write "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"" height=""80"" "" WIDTH=100%"" CLASS=NavLeftHighlight1>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "		<TD CLASS=Medium>" & vbCrLf
    .write        Translate("Selected Categories",Login_Language,conn) & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "	<TR>" & vbCrLf
    .write "    <TD CLASS=Medium>" & vbCrLf
    
    %>
    <% 
    if showpcat=true then
       .write " <SELECT LANGUAGE=""JavaScript"" multiple size=""5"" NAME=""PCat_SPortalCats"" CLASS=""Medium"">"
	   else
       .write " <SELECT LANGUAGE=""JavaScript"" multiple size=""5"" NAME=""PCat_SPortalCats"" CLASS=""Medium"">"
    end if
    
    strSQL123 = "SELECT AC.AssetId, AC.CategoryId, AC.CreateDate, AC.CreatedBy, CC.Title FROM dbo.Asset_Category AS AC INNER JOIN dbo.Calendar_Category AS CC ON AC.CategoryId = CC.ID AND AC.AssetId = 0"
    if (Request.QueryString("ID") <> "" and Request.QueryString("ID") <> "add") then
      strSQL123 = "SELECT AC.AssetId, AC.CategoryId, AC.CreateDate, AC.CreatedBy, CC.Title FROM dbo.Asset_Category AS AC INNER JOIN dbo.Calendar_Category AS CC ON AC.CategoryId = CC.ID AND AC.AssetId = " & Request.QueryString("ID")
    end if
    
    Set rsCategory1 = Server.CreateObject("ADODB.Recordset")
    rsCategory1.Open strSQL123, conn, 3, 3

    if not rsCategory1.EOF then
      Do while not rsCategory1.EOF
        .Write "<OPTION value=""" & rsCategory1("CategoryId") & """>" & rsCategory1("Title") & "</OPTION>" & vbCrLf
        rsCategory1.MoveNext
      loop
    end if

    rsCategory1.close
    set rsCategory1 = nothing

    ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  %>
    </SELECT>
    <% .write "  </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf

    .write "</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "</TR>" & vbCrLf            

    end with        
            ''Nitin Code Changes End
    %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%=Translate("Language",Login_Language,conn)%>:
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                              
                <SELECT Name="Content_Language" CLASS=Medium>    
                <%
             	  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf                                
                
                if Site_Id = 82 then
                      SQL = "SELECT * FROM Language WHERE Oracle_Enable=-1 OR Enable=-1 and Code != 'enb'  ORDER BY Language.Sort"
                    else
                      SQL = "SELECT * FROM Language WHERE Oracle_Enable=-1 OR Enable=-1 ORDER BY Language.Sort"
                 end if 
                Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                rsLanguage.Open SQL, conn, 3, 3
                                      
                Do while not rsLanguage.EOF
                  if "elo" = rsLanguage("Code") then                
                  elseif "eng" = rsLanguage("Code") then
                 	  response.write "<OPTION CLASS=Medium SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn)
                  else
                 	  response.write "<OPTION CLASS=Medium VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn)
                  end if
                  if CInt(rsLanguage("Enable")) = CInt(True) then
                    response.write " +"
                  end if                    
                  response.write "</OPTION>" & vbCrLf                  
              	  rsLanguage.MoveNext 
                loop
                
                rsLanguage.close
                set rsLanguage=nothing
                set SQL = nothing
                %>
                </SELECT>
                <%response.write "<SPAN CLASS=SMALL>+ " & Translate("Indicates a supported infrastructure language.",Login_Language,conn) & "</SPAN>"%>              
              </TD>
            </TR>  
            
            <!-- Added on 28th Oct 2009, for Business Unit and Ad Pixel -->  
            <% if Show_Marketing_Automation = True then %> 
            <!-- Business Unit --> 
            <TR>
                <TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                    <%=Translate("Business Unit",Login_Language,conn)%>:
                </TD>
                <TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                    &nbsp;
                </TD> 
                <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                    <SELECT Name="Business_Unit" CLASS=Medium> 
                       <% 
                        response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf                                
                        SQL = "SELECT BusinessUnitCode, (BusinessUnitDesc + ' (' + BusinessUnitCode + ')') As BusinessUnitDesc FROM Business_Units ORDER BY BusinessUnitDesc"
                        Set rsBusinessUnits = Server.CreateObject("ADODB.Recordset")
                        rsBusinessUnits.Open SQL, conn, 3, 3
                        
                        Do while not rsBusinessUnits.EOF
                          response.write "<OPTION CLASS=Medium VALUE=""" & rsBusinessUnits("BusinessUnitCode") & """>" & Translate(rsBusinessUnits("BusinessUnitDesc"),Login_Language,conn)
                          response.write "</OPTION>" & vbCrLf                  
          	              rsBusinessUnits.MoveNext 
                        loop
                        
                        rsBusinessUnits.close
                        set rsBusinessUnits=nothing
                        set SQL = nothing
                       %> 
                    </SELECT>
                    <%response.write "<SPAN CLASS=SMALL>" & Translate("Used for query string grouping in Eloqua.",Login_Language,conn) & "</SPAN>"%>              
                </TD>
            </TR>  
            
            <!-- Ad Tracking Pixel --> 
            <TR>
                <TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                    <%=Translate("Ad Tracking Pixel",Login_Language,conn)%>:
                </TD>
                <TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                    &nbsp;
                </TD> 
                <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                    <%
                      response.write "<INPUT TYPE=""textbox"" size = ""75"" MAXLENGTH = ""500"" NAME=""txtAdPixel"" CLASS=Medium>"
                    %> 
                </TD>
            </TR>    
            <% end if %>  
            <!-- End-->      
            
            <!-- Post via Subscription Service -->                
                         
            <% if Show_Subscription = True then %>

            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                           
 
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Send Notice via Subscription Service",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                 &nbsp;
              </TD>                                                                  
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<TABLE WIDTH=""100%"" BORDER=0>"
                response.write "<TR>"
                response.write "<TD WIDTH=20 CLASS=Medium>"
                response.write "<INPUT TYPE=""Checkbox"" NAME=""Subscription"" CLASS=Medium" & Field_Editable & ">"
                response.write "</TD>"
                response.write "<TD CLASS=Medium>" & Translate("Subscription Service",Login_Language,conn)
                ' Evening Bulk EMail
                response.write "&nbsp;&nbsp;&nbsp;9:00pm PST&nbsp;"
                response.write "<INPUT TYPE=""Radio"" NAME=""Subscription_Early"" CLASS=Medium "
                if CInt(Campaign_Subscription_Early) = CInt(False) then
                  response.write "CHECKED "
                end if
                response.write "VALUE=""0"" " & Field_Editable & ">"
                ' Afternoon Bulk EMail
                response.write "&nbsp;&nbsp;&nbsp;12:00pm PST (noon)&nbsp;"
                response.write "<INPUT TYPE=""Radio"" NAME=""Subscription_Early"" CLASS=Medium "
                if CInt(Campaign_Subscription_Early) = CInt(True) then
                  response.write "CHECKED "
                end if
                response.write "VALUE=""-1"" " & Field_Editable & ">"
                response.write "</TD></TR>"
                response.write "</TABLE>"
                %> 
              </TD>
            </TR>
            <% end if %>       
           
            <!-- NT Sub-Groups -->         
            
            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                              

    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%=Translate("Select Groups allowed to view this information",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                 <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                                  
    	        <TD BGCOLOR="White" CLASS=Medium>               
          
              <%
                response.write "<TABLE WIDTH=""100%"">"

                ' Electronic Email Fulfillment EEF End-User Oracle

                if (Category_Code < 8000 or Category_Code > 8999) and Show_Item_Number = True then

                  ' Electronic Email Fulfillment - End-User Oracle
                  response.write "<TR><TD WIDTH=20 CLASS=Medium>"
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"""
                  if Preset_EEF = true then response.write " CHECKED "
                  response.write " CLASS=Medium VALUE=""view""></TD>"
                  response.write "<TD CLASS=Medium BGCOLOR=""#FF9966"">&nbsp;" & Translate("Available to Electronic Email Fulfillment",Login_Language,conn) & " - <SPAN CLASS=Small>(" & Translate("End-User Oracle",Login_Language,conn) & ")</SPAN></TD>"
                  response.write "</TR>"

                  ' Electronic Digital Library
                  response.write "<TR><TD WIDTH=20 CLASS=Medium>"
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"""
                  if Preset_FDL = true then response.write " CHECKED "
                  response.write " CLASS=Medium VALUE=""fedl""></TD>"
                  response.write "<TD CLASS=Medium BGCOLOR=""#FF9966"">&nbsp;" & Translate("Available to Electronic Fulfillment",Login_Language,conn) & " - <SPAN CLASS=Small>(" & Translate("End-User Digital Library",Login_Language,conn) & ")</SPAN></TD>"
                  response.write "</TR>"

                  ' Shopping Cart - Order Literature
                  
                  if Show_Shopping_Cart then
                    response.write "<TR><TD WIDTH=20 CLASS=Medium>"
                    response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"""
                    response.write " CLASS=Medium VALUE=""shpcrt""></TD>"
                    response.write "<TD CLASS=Medium BGCOLOR=""#FF9966"">&nbsp;" & Translate("Exclude From Literature Order Shopping Cart",Login_Language,conn) & "</TD>"
                    response.write "</TR>"
                  end if
                  
                  response.write "<TR><TD HEIGHT=8 WIDTH=20></TD><TD HEIGHT=8></TD></TR>" & vbCrLf
  
                end if

                ' Regional Groups
'               response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""all"" CLASS=Medium></TD><TD CLASS=Medium><B>" & Translate("All Groups in all Regions",Login_Language,conn) & "</B></TD></TR>"

                if isblank(Admin_Region) then Admin_Region = 1
                
                for i = 0 to 1
                
                  Select case i
                    case 0
                      SQL = "SELECT SubGroups.*, SubGroups.Order_Num "
                      SQL = SQL & "FROM SubGroups "
                      SQL = SQL & "WHERE SubGroups.Site_ID=" & Site_ID & " AND SubGroups.Region=" & Admin_Region & " AND SubGroups.Enabled=" & CInt(true)
                      SQL = SQL & "ORDER BY SubGroups.Order_Num"                
                    case else
                      SQL = "SELECT SubGroups.*, SubGroups.Order_Num "
                      SQL = SQL & "FROM SubGroups "
                      SQL = SQL & "WHERE SubGroups.Site_ID=" & Site_ID & " AND SubGroups.Region<>" & Admin_Region & " AND SubGroups.Enabled=" & CInt(true)
                      SQL = SQL & "ORDER BY SubGroups.Order_Num"                  
                  end select  
                            
                  Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
                  rsSubGroups.Open SQL, conn, 3, 3
                
                  if not rsSubGroups.EOF then
                                                      
                    Do while not rsSubGroups.EOF
  
                      if RegionValue <> Mid(rsSubGroups("Code"),1,1) then
                        RegionValue = Mid(rsSubGroups("Code"),1,1)
                        select case UCase(RegionValue)
                          case "U"                    ' United States
                            RegionColorPointer = 1
                          case "E"                    ' Europe
                            RegionColorPointer = 2
                          case "I"                    ' Intercon
                            RegionColorPointer = 3
                          case "N"                    ' Entitlements
                            RegionColorPointer = 4
                          case else
                            RegionColorPointer = 0
                        end select    
                        Region = Region + 1
                        if Region <= 4 then
                          response.write "<TR><TD HEIGHT=8 WIDTH=20></TD><TD HEIGHT=8></TD></TR>"
                          if rsSubGroups.RecordCount >= 2 then
                            response.write "<TR><TD CLASS=Medium BGCOLOR="""
                            if RegionColorPointer = 4 then
                              response.write "#FFCC99"                            
                            elseif Region <> Admin_Region then
                              response.write "Yellow"
                            else
                              response.write "Green"
                            end if
                            response.write """>"
                            if RegionColorPointer <> 4 then
                              response.write "<INPUT TYPE=""Checkbox"""
                              response.write " ONCLICK=""SubGroups_" & Trim(CStr(RegionColorPointer)) & "_Check();"""
                              response.write " NAME=""SubGroups_" & Trim(CStr(RegionColorPointer)) & """ CLASS=Medium>"
                            end if
                            if RegionColorPointer <> 4 then
                              response.write "</TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;<B>" & Translate("All Groups for this Section",Login_Language,conn) & "</B></TD></TR>"
                            else
                              response.write "</TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;<B>" & Translate("Entitlements",Login_Language,conn) & "</B></TD></TR>"                            
                            end if  
                          end if
                        elseif Region > 4 then
                          Region = 4
                        end if
                      end if
                     
                      Default_Select = False
                      if (Category_Code < 8000 or Category_Code > 8999) and not isblank(Campaign_SubGroups) then
                        if instr(1,Campaign_Subgroups,rsSubGroups("Code")) > 0 then
                          Default_Select = True
                        end if
                      elseif rsSubGroups("Default_Select") = CInt(True) then
                        Default_Select = True
                      end if
                      'Below code updated, gold silo aggregation(check box to radio button) for Fnet AMS(siteId 82)
                      'other portal site remains same  -- by zensar(11/17/08)  
                   if Site_Id = 82 then
                      if Default_Select = True then
                        response.write "<TR><TD CLASS=Medium"
                        if Default_Select = True then
                          response.write " BGCOLOR=""#FF0000"""
                        end if  
                        response.write ">"
                            response.write "<INPUT TYPE=""radio"""
                            if Region <> Admin_Region and RegionColorPointer <> 4 then
                              response.write " ONCLICK=""SubGroups_Check();"""
                            end if
                            response.write " NAME=""EntSubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED CLASS=Medium></TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;" & rsSubGroups("X_Description") & "</TD></TR>"
                      else
                            response.write "<TR><TD CLASS=Medium>"
                            response.write "<INPUT TYPE=""radio"""
                            if Region <> Admin_Region and RegionColorPointer <> 4 then
                              response.write " ONCLICK=""SubGroups_Check();"""
                            end if
                            response.write " NAME=""EntSubGroups"" VALUE=""" & rsSubGroups("Code") & """ CLASS=Medium></TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;" & rsSubGroups("X_Description") & "</TD></TR>"
                      end if
                      
                      
                      
                   else
                    if Default_Select = True then
                        response.write "<TR><TD CLASS=Medium"
                        if Default_Select = True then
                          response.write " BGCOLOR=""#FF0000"""
                        end if  
                        response.write ">"
                        response.write "<INPUT TYPE=""Checkbox"""
                        if Region <> Admin_Region and RegionColorPointer <> 4 then
                          response.write " ONCLICK=""SubGroups_Check();"""
                        end if
                        response.write " NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED CLASS=Medium></TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;" & rsSubGroups("X_Description") & "</TD></TR>"
                      else
                        response.write "<TR><TD CLASS=Medium>"
                        response.write "<INPUT TYPE=""Checkbox"""
                        if Region <> Admin_Region and RegionColorPointer <> 4 then
                          response.write " ONCLICK=""SubGroups_Check();"""
                        end if
                        response.write " NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """ CLASS=Medium></TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;" & rsSubGroups("X_Description") & "</TD></TR>"
                      end if
                   end if   
                    
                   ''Added on 29th Oct 2009
                   if Show_Marketing_Automation = True and rsSubGroups("Code") = "nfre" then
                        response.write "<TR><TD CLASS=Medium>"
                        response.write "<INPUT TYPE=""radio"""
                        if Region <> Admin_Region and RegionColorPointer <> 4 then
                          response.write " ONCLICK=""SubGroups_Check();"""
                        end if
                        response.write " NAME=""EntSubGroups"" VALUE=""" & rsSubGroups("Code") & """ CLASS=Medium></TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;Form Processing"
                        response.write "&nbsp;&nbsp;&nbsp;&nbsp; URL:&nbsp;&nbsp;"
                        response.write "<INPUT TYPE=""textbox"" size = ""50"" MAXLENGTH = ""500"" NAME=""txtFormProcessingURL"" CLASS=Medium></TD></TR>"
                   end if
                   ''end

                  rsSubGroups.MoveNext 
                loop

                    rsSubGroups.close
                    set rsSubGroups=nothing
                    set SQL = nothing

                  else
                    Region = Region + 1
                  end if
                  
                next
                
                if Admin_Access >= 8 then
                  response.write "<TR><TD HEIGHT=8 WIDTH=20>&nbsp;</TD><TD HEIGHT=8>&nbsp;</TD></TR>" & vbCrLF
                  response.write "<TR><TD HEIGHT=8 WIDTH=20>&nbsp;</TD><TD HEIGHT=8 CLASS=Small>" & Translate("Restricted Groups",Login_Language,conn) & " - " & Translate("Do not use the groups listed below for regular content items.",Login_Language,conn) & "</TD></TR>" & vbCrLF

                  SQL = "SELECT SubGroups.*, SubGroups.Order_Num "
                  SQL = SQL & "FROM SubGroups "
                  SQL = SQL & "WHERE SubGroups.Site_ID=0" & " AND SubGroups.Enabled=" & CInt(True)
                  SQL = SQL & "ORDER BY SubGroups.Order_Num"
                  
                  Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
                  rsSubGroups.Open SQL, conn, 3, 3

                  do while not rsSubGroups.EOF
                    if rsSubGroups("Enabled") = True then
                      response.write "<TR>" & vbCrLF
                      response.write "<TD WIDTH=20 CLASS=Medium><SPAN FACE=""Arial"" SIZE=2><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """></SPAN></TD>" & vbCrLF
                      response.write "<TD CLASS=Medium BGCOLOR=""#669999"">&nbsp;" & Translate(rsSubGroups("X_Description"),login_language,conn) & "</TD>" & vbCrLF
                      response.write "</TR>" & vbCrLF
                    end if
                    rsSubGroups.MoveNext
                  loop  

                  rsSubGroups.close
                  set rsSubGroups=nothing
                  set SQL = nothing

                end if
              end if
              
              response.write "</TABLE>"

            response.write "</TD>"
            response.write "</TR>"
          %>
          
          <!-- Restricted to Countries -->

           <% if Show_Country_Restrictions then %>
           
             <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>            
             
      				<TR>
              	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                  <%
                  response.write "<INPUT TYPE=""Radio"" NAME=""Country_Reset"" VALUE="""" CLASS=Medium CHECKED ONCLICK=""document." & FormName & ".Country.value='';"">"
                  response.write Translate("No Country Restrictions",Login_Language,conn) & ":<BR>"
                  response.write "<INPUT TYPE=""Radio"" NAME=""Country_Reset"" VALUE="""" CLASS=Medium>"
                  response.write Translate("Include only these Countries",Login_Language,conn) & ":<BR>"
                  response.write "<INPUT TYPE=""Radio"" NAME=""Country_Reset"" VALUE=""0"" CLASS=Medium>"
                  response.write Translate("Exclude only these Countries",Login_Language,conn) & ":"
                  %>
                </TD>
               	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                              
       	        <TD BGCOLOR="White" ALIGN=LEFT VALIGN=TOP CLASS=Medium>
                  <%
                  Users_Country = Campaign_Country
  
                  Call Connect_FormDatabase
                  Call DisplayCountryList("Country",Users_Country,"","Medium")
                  Call Disconnect_FormDatabase
                  response.write "<P>" & vbCrLf
                  response.write "<SPAN CLASS=Small>(" & Translate("Multi-Select Drop-Down",Login_Language,conn) & ")<BR>"
                  response.write Translate("Use [CTRL] + [LEFT MOUSE] to select multiple countries.",Login_Language,conn)
                  response.write "</SPAN>" & vbCrLf
                  %>  
                </TD>
              </TR>
            
            <% end if %>                        
            
           <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                                       
           <!-- Approver Selection -->
           <% if Admin_Access <= 2 or ((Admin_Access = 4 or Admin_Access >= 8) and Show_Submission_Approve = true) then
          
                SQL =       "SELECT Approvers.* "
                SQL = SQL & "FROM Approvers "
                SQL = SQL & "WHERE Approvers.Site_ID=" & Site_ID & " AND (Approvers.Approver_ID Is Not Null OR Approvers.Approver_ID <> 0)"
                SQL = SQL & "ORDER BY Approvers.Order_Num"
          
                Set rsApprovers = Server.CreateObject("ADODB.Recordset")
                rsApprovers.Open SQL, conn, 3, 3
                
                if not rsApprovers.EOF then %>
                     
        				<TR>
                	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                    <%=Translate("Select Group to Approve this Submission",Login_Language,conn)%>:
                  </TD>
                 	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                    <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
                  </TD>                                                              
        	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                    <% 
                    response.write "<SELECT NAME=""Review_By_Group"" CLASS=Medium>" & vbCrLf
                    if Admin_Access = 4 or Admin_Access>=8 then
                      response.write "<OPTION CLASS=NavLeftHighlight1 VALUE=""0"">" & Translate("Approval by Current Administrator",Login_Language,conn) & "</OPTION>" & vbCrLf
                    else
                      response.write "<OPTION CLASS=Medium VALUE="""" SELECTED>" & Translate("Select from this list",Login_Language,conn) & "</OPTION>" & vbCrLf
                    end if
                    
                    do while not rsApprovers.EOF
                      response.write "<OPTION CLASS=Medium VALUE=""" & rsApprovers("ID") & """>" & rsApprovers("Description") & "</OPTION>" & vbCrLf
                      rsApprovers.MoveNext
                    loop
                    response.write "</SELECT>" & vbCrLf
                    %>                    
                  </TD>
                </TR>                        
                <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                                       
              <% end if
                 rsApprovers.Close
                 set rsApprovers = nothing
                 set SQL = nothing
                 
            else
            
              response.write "<INPUT TYPE=""HIDDEN"" NAME=""Review_By_Group"" VALUE=""0"">"
                 
            end if
            
            ' --------------------------------------------------------------------------------------                                
            ' Navigation Buttons
            ' --------------------------------------------------------------------------------------
    
            response.write "<TR>"
            response.write "<TD COLSPAN=3 CLASS=Medium>"
            response.write "<TABLE WIDTH=""100%"" CELLPADDING=2 BGCOLOR=""Silver"">"
            response.write "<TR>"
            response.write "<TD ALIGN=CENTER WIDTH=""25%"" CLASS=Medium>"
            response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Main_Menu"" VALUE="" " & Translate("Main Menu",Login_Language,conn) & " "" CLASS=Navlefthighlight1 ONCLICK=""Menu_Button=true;"">"
            response.write "</TD>"
            response.write "<TD ALIGN=CENTER WIDTH=""25%"" CLASS=Medium>"
            response.write "&nbsp;"
            response.write "</TD>"
            response.write "<TD ALIGN=CENTER WIDTH=""25%"" CLASS=Medium>"
            if Category_ID <> false then
                'Modified  by zensar for onclick event on 14-06-2006.
                if Show_PID = true then
                  if  PID_System = 0 then  
                    response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Update"" onclick =""return setOperation('U')"" VALUE="" " & Translate("Save",Login_Language,conn) & " / " & Translate("Update",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
                  elseif PID_System = 1 then  
                    response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Update"" VALUE="" " & Translate("Save",Login_Language,conn) & " / " & Translate("Update",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
                  end if
                else
                  response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Update"" VALUE="" " & Translate("Save",Login_Language,conn) & " / " & Translate("Update",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
                end if 
            end if
            response.write "</TD>"
            response.write "<TD ALIGN=CENTER WIDTH=""25%"" CLASS=Medium>"
            response.write "&nbsp;"
            response.write "</TD>"
            response.write "</TR>"
            response.write "</TABLE>"
            response.write "</TD>"
          response.write "</TR>"
         response.write "</TABLE>"
        response.write "</TD>"
      response.write "</TR>"
    response.write "</TABLE>"
    Call Table_End
    response.write "</FORM>"
    response.write "<BR><BR>"      

%>
