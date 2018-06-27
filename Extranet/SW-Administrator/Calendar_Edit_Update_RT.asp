<!--#include virtual="/include/RTEditor/class.devedit.asp" -->
<%
' --------------------------------------------------------------------------------------
' Edit Content Item - Include File for /SW-Administrator/Calendar_Edit.asp
'
' Author: K. D. Whitlock
' --------------------------------------------------------------------------------------

  SQL =       "SELECT Calendar.* "
  SQL = SQL & "FROM Calendar LEFT JOIN Calendar_Category ON (Calendar.Category_ID = Calendar_Category.ID) AND (Calendar.Site_ID = Calendar_Category.Site_ID) "
  SQL = SQL & "WHERE (((Calendar.ID)=" & CInt(Calendar_ID) & "))"
 
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open SQL, conn, 3, 3

  if not rs.EOF then
    Clone = rs("Clone")
  else
    Clone = 0
  end if       
    
  if not rs.EOF then
  
    Category_ID = rs("Category_ID")
    Write_Form_Show_Values = False
    Call Get_Show_Values

    if not isblank(request("Content_Group")) then
      Content_Group = request("Content_Group")
    else
      Content_Group = rs("Content_Group")
    end if

    if Content_Group > 0 and not Show_Calendar then
      Field_Editable = " DISABLED"        
    else
      Field_Editable = ""  
    end if

    if Content_Group > 0 and not isblank(request("Campaign")) then
      Campaign = request("Campaign")
    elseif Content_Group > 0 then
      Campaign = rs("Campaign")
    else
      Campaign = 0
    end if 

    if CInt(Show_View) = CInt(True) then

      ' Replicate what user sees at site for one record - i.e., visual feedback to submitter

      Set rsCalendar = rs.clone
               
'      SQL = "SELECT Calendar.* FROM Calendar WHERE Calendar.ID=" & CInt(Calendar_ID)
'      Set rsCalendar = Server.CreateObject("ADODB.Recordset")
'      rsCalendar.Open SQL, conn, 3, 3
      
      if not rsCalendar.EOF then           
        
        Dim Record_Number
        Record_Number = 0

        response.write "<FONT CLASS=MediumBold>" & Translate("Content or Event",Login_Language,conn) & "</FONT><BR>"
         
        response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
          
        Write_Form_Show_Values = False
        Show_Detail            = True
        Category_ID = rs("Category_ID")
        Call Get_Show_Values
        Call Update_Fields
        Call Display_Category_Item
          
        response.write "</TABLE>"
        response.write "<FONT CLASS=SmallBold><FONT Color=""DarkGray"">" & Translate("Note",Login_language,conn) & ": </FONT></FONT><FONT CLASS=Small><FONT COLOR=DarkGray>" & Translate("Clicking on one of the above icons to view, download, or email the asset, you may need to logon to the site in the new pop-up window.  Once site logon is complete, repeat the icon selection from this screen to see the asset.",Login_Language,conn) & "</FONT></FONT>"
        response.write "<P>"
          
      end if
        
      rsCalendar.close
      set rsCalendar = nothing
        
      if not isblank(rs("Splash_Header")) then
        response.write "<FONT CLASS=MediumBold>" & Translate("Splash Header",Login_Language,conn) & "</FONT><BR>"
        response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=4 BGCOLOR=""#F3F3F3"">" & vbCrLf
        response.write "<TR>"
        if not isblank(rs("Thumbnail")) then
          response.write "<TD CLASS=Medium WIDTH=""80"" VALIGN=TOP><IMG SRC=""/" & Site_Code & "/" & rs("Thumbnail") & """ BORDER=1 WIDTH=80></TD>"
        end if
        response.write "<TD CLASS=Medium VALIGN=TOP>" & RestoreQuote(rs("Splash_Header")) & "</TD>"
        response.write "</TR>"
        response.write "</TABLE>"
        response.write "<P>"
      end if
      
      ' Display Introduction Letter
      if not isblank(rs("Item_Number_2")) then

        SQL = "SELECT Calendar.* FROM Calendar WHERE Calendar.ID=" & CInt(rs("Item_Number_2"))
        Set rsCalendar = Server.CreateObject("ADODB.Recordset")
        rsCalendar.Open SQL, conn, 3, 3
      
        if not rsCalendar.EOF then           
        
          Record_Number = 0
  
          response.write "<FONT CLASS=MediumBold>" & Translate("Introduction Letter",Login_Language,conn) & "</FONT><BR>"
          response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
          
          Write_Form_Show_Values = False
          Show_Detail            = True
          Category_ID = rs("Category_ID")
          Call Get_Show_Values
          Call Update_Fields
          Call Display_Category_Item
          
          response.write "</TABLE>"

        end if
        
        rsCalendar.close
        set rsCalendar = nothing
        
      end if  

      if not isblank(rs("Splash_Footer")) then response.write "<BR><BR>"
              
      if not isblank(rs("Splash_Footer")) then
        response.write "<FONT CLASS=MediumBold>" & Translate("Splash Footer",Login_Language,conn) & "</FONT><BR>"
        response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=4 BGCOLOR=""#F3F3F3"">" & vbCrLf
        response.write "<TR><TD CLASS=Medium VALIGN=TOP>" & RestoreQuote(rs("Splash_Footer")) & "</TD></TR>"
        response.write "</TABLE>"
        response.write "<P>"
      end if

    end if

    FormName = "EditContent"
    %>
      
    <FORM NAME="<%=FormName%>" ACTION="calendar_admin.asp" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA" onKeyUp="highlight(event)" onClick="highlight(event)">
    <INPUT Type="Hidden" NAME="Show_View" VALUE="<%=Show_View%>">
    <INPUT TYPE="Hidden" NAME="ID" VALUE="<%=Calendar_ID%>">
    <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
    <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
    <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
    <INPUT TYPE="Hidden" NAME="PDate" VALUE="<%=rs("PDate")%>">
    <INPUT TYPE="Hidden" NAME="UDate" VALUE="<%=Now()%>">
    <INPUT TYPE="Hidden" NAME="Show_Calendar" VALUE="<%=CInt(Show_Calendar)%>">    
    <INPUT TYPE="Hidden" NAME="Admin_Access" VALUE="<%=Admin_Access%>">                                          
    <INPUT TYPE="Hidden" NAME="Admin_ID" VALUE="<%=Admin_ID%>">
    <INPUT TYPE="Hidden" NAME="ProgressID" VALUE="<%=ProgressID%>">                                          
    
    <%
    if Clone > 0 then
      response.write "<INPUT TYPE=""Hidden"" NAME=""Clone"" VALUE=""" & Clone & """>" & vbCrLf
    else
      response.write "<INPUT TYPE=""Hidden"" NAME=""Clone"" VALUE=""" & Calendar_ID & """>" & vbCrLf
    end if  

    if isblank(rs("Submitted_By")) then
      Submitted_By_Current = Admin_ID
      response.write "<INPUT TYPE=""Hidden"" NAME=""Submitted_By"" VALUE=""" & Admin_ID & """>" & vbCrLf
    else
      Submitted_By_Current = rs("Submitted_By")
      response.write "<INPUT TYPE=""Hidden"" NAME=""Submitted_By"" VALUE=""" & rs("Submitted_By") & """>" & vbCrLf
    end if
        
    ' Determine Site Path based on Site_ID Number 

    SQL = "SELECT * FROM Site WHERE ID=" & CInt(Site_ID)
    Set rsSite = Server.CreateObject("ADODB.Recordset")
    rsSite.Open SQL, conn, 3, 3
    Site_Code = rsSite("Site_Code")
    response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Site"" VALUE=""" & Site_Code & """>" & vbCrLf
    rsSite.close
    set rsSite=nothing            

    response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Site_POD"" VALUE=""" & Path_Site_POD & """>" & vbCrLf    

    if not Show_View and Content_Group > 0 then
      if Content_Group > 0 or not isblank(request("Content_Group")) or Show_Calendar then
        response.write "<UL><LI><SPAN CLASS=SmallBold>" & Translate("You have designated this Content or Calendar Event item be included in the following multiple asset container (MAC)",Login_Language,conn) & ": <SPAN CLASS=SmallBoldRed>"
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
        response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 1</SPAN><SPAN CLASS=Small>: " & Translate("Certain fields may blocked from editing (indicated by a shaded color).  These values for these blocked fields are automatically updated when you update the multiple asset container that this asset or event is associated with.",Login_Language,conn) & "</SPAN></LI>"
      end if
      if not isblank(request("Content_Group")) then
        response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 2</SPAN><SPAN CLASS=Small>: " & Translate("Because you have changed the &quot;Content Grouping&quot; for an existing record, review your &quot;Groups allowed to view this information&quot; selections.",Login_Language,conn) & "</SPAN></LI>"
      end if
      if Content_Group > 0 and Show_Calendar then
        response.write "<LI><SPAN CLASS=SmallBoldRed>" & Translate("Note",Login_Language,conn) & " 3</SPAN><SPAN CLASS=Small>: " & Translate("If this &quot;Calendar Event&quot; will occur before the &quot;Pre-Announce&quot; or &quot;Beginning Date&quot; of the master multiple asset container that it is associated with, ensure that the content is generic and does not provide specific details about the nature of the multiple asset container, because this Calendar Event will be available for view prior to the actual &quot;Pre-Announce&quot; or &quot;Begining Date&quot; date of the multiple asset container.",Login_Language,conn)
        response.write  "&nbsp;&nbsp;" & Translate("This restriction does not apply to a Calendar Event that will occur after the &quot;Pre-Announce&quot; or &quot;Beginning Date&quot;.",Login_Language,conn) & "</SPAN></LI>"
      end if
      if Content_Group > 0 or not isblank(request("Content_Group")) or Show_Calendar then
        response.write "</UL>"
      end if  
    end if  
    %>
      
    <TABLE WIDTH="100%" BORDER=1 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
      <TR>
    	  <TD WIDTH="100%" BGCOLOR="#EEEEEE">
    			<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>          
          
            <!-- Header -->
     				<TR>
             	<TD WIDTH="50%" BGCOLOR="Black" COLSPAN=2 CLASS=MediumBold>
                 <FONT COLOR="#FFCC00"><%=Translate("Description",Login_Language,conn)%></FONT>
              </TD>
     	        <TD WIDTH="50%" BGCOLOR="Black" ALIGN=LEFT CLASS=MediumBold>
                 <FONT COLOR="#FFCC00"><%=Translate("Content or Event Information",Login_Language,conn)%></FONT>
              </TD>
            </TR>
    				<TR>
            	<TD BGCOLOR="Silver" COLSPAN=2 CLASS=MediumBold>
                <B><%=Translate("Note",Login_Language,conn)%>:&nbsp;&nbsp;&nbsp;<IMG SRC="/images/required.gif" BORDER=0 HEIGHT="10" WIDTH="10"> = <%=Translate("Required Information",Login_Language,conn)%>.
              </TD>
              <TD BGCOLOR="Silver" VALIGN=TOP CLASS=Medium NOWRAP>                              
                <A HREF="#HELP">
                <IMG SRC="/images/help_button.gif" BORDER=0 ALIGN=RIGHT VALIGN=TOP>
                </A>
                <INPUT TYPE="Submit" NAME="Nav_Main_Menu" VALUE=" <%=Translate("Main Menu",Login_Language,conn)%> " CLASS=Navlefthighlight1 Title="Return to Administrators Main Menu.">&nbsp;&nbsp;&nbsp;&nbsp;
                <INPUT TYPE="Button" Title="Show/Hide Items in Site View Mode." onClick="location.href='/SW-Administrator/Calendar_Edit.asp?ID=<%=Calendar_ID%>&Site_ID=<%=Site_ID%>&Show_View=<%if CInt(Show_View) = CInt(True) then response.write CInt(False) else response.write CInt(True)%>'" VALUE="<%if CInt(Show_View) = CInt(True) then response.write Translate("Hide Site View",Login_Language,conn) else response.write Translate("Site View",Login_Language,conn)%>" CLASS=Navlefthighlight1></A>
                <%
                if CInt(rs("Code")) >= 8000 AND CInt(rs("Code")) <= 8999 then
                  response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
                  response.write "<INPUT TYPE=""Button"" onClick=""location.href='/SW-Administrator/Site_Utility.asp?ID=Site_Utility&Campaign=" & rs("ID") & "&Site_ID=" & Site_ID & "&Utility_ID=50&View=4'"" CLASS=Navlefthighlight1 Title=""List Individual Assets belonging to this Container."" VALUE=""" & Translate("List Assets",Login_Language,conn) & """>"
                end if
                response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
                response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Update"" VALUE="" " & Translate("Update",Login_Language,conn) & " "" CLASS=Navlefthighlight1 TITLE=""Save Changes to Record"" onclick=""startupload()"">"
                %>
              </TD>
            </TR>        
              
            <!-- Calendar Event ID -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" WIDTH="48%" CLASS=Medium>
                <%=Translate("Content or Event ID Number",Login_Language,conn)%>:
              </TD>
              <TD BGCOLOR="#EEEEEE" ALIGN=CENTER WIDTH="2%" CLASS=Medium>
                &nbsp;
              </TD>                                              
    	        <TD BGCOLOR="White" WIDTH="50%" CLASS=Medium>              
                <%
                response.write "<TABLE WIDTH=""100%"">"
                  response.write "<TR>"

                    response.write "<TD WIDTH=""10%"" CLASS=Medium>"
                    if rs("Locked") = True then
                      response.write "<FONT COLOR=""Red"">"
                    else
                      response.write "<FONT COLOR=""Gray"">"
                    end if
                    response.write rs("ID") & "</FONT>"
  
                    response.write "</TD>"
                
                    response.write "<TD WIDTH=""10%"" CLASS=Medium>"
                    if Clone > 0 then
                      response.write "<FONT COLOR=""Gray"">[" & Clone & "]</FONT>"
                    else
                      response.write "&nbsp;"
                    end if
                    response.write "</TD>"

                    response.write "<TD WIDTH=""20%"" ALIGN=Right CLASS=Medium>"
                    response.write Translate("Status",Login_Language,conn) & ":&nbsp;"

                    if Admin_Access >=8 then
                      if rs("Locked") = True then
                         response.write "<INPUT TYPE=""Checkbox"" NAME=""Locked"" CHECKED CLASS=Medium>"
                        else
                         response.write "<INPUT TYPE=""Checkbox"" NAME=""Locked"" CLASS=Medium></TD><TD CLASS=Medium>"
                      end if
                    end if                                                                                           

                    response.write "&nbsp;"
                    response.write "</TD>"

                    ' Status Review
                    
                    response.write "<TD WIDTH=""20%"" BGCOLOR="
                    if isblank(rs("Status")) or rs("Status") = 0 then
                      response.write """Yellow"""
                    else
                      response.write """#EEEEEE"""
                    end if
                    response.write " CLASS=Medium>"

                      if Show_Calendar then
                        response.write "<INPUT TYPE=""RADIO"" NAME=""STATUS"" VALUE=""0"""
                        if isblank(rs("Status")) or rs("Status") = 0 then response.write " CHECKED"
                        response.write " CLASS=Medium>&nbsp;&nbsp;"
                        if not rs("Status") then
                          response.write "<B>" & Translate("Review",Login_Language,conn) & "</B>"
                        else
                          response.write Translate("Review",Login_Language,conn)
                        end if  
                      else
                        response.write "<INPUT " & Field_Editable & " TYPE=""RADIO"" NAME=""STATUS"" VALUE=""0"""
                        if isblank(rs("Status")) or rs("Status") = 0 then response.write " CHECKED"
                        response.write " CLASS=Medium>&nbsp;&nbsp;"
                        if not rs("Status") then
                          response.write "<B>" & Translate("Review",Login_Language,conn) & "</B>"
                        else
                          response.write Translate("Review",Login_Language,conn)
                        end if  
                      end if
                                           
'                      if Content_Group > 0 and not Show_Calendar then
'                        response.write "<INPUT TYPE=""Hidden"" NAME=""Status"" VALUE=" & rs("Status") & ">"
'                      end if

                    response.write "</TD>"
                      
                    ' Status Live

                    response.write "<TD WIDTH=""20%"" BGCOLOR="
                    if rs("Status") = 1 then
                      response.write """#00CC00"""
                    else
                      response.write """#EEEEEE"""
                    end if  
                    response.write " CLASS=Medium>"

                      if Admin_Access = 4 or Admin_Access >=8 then
                        if Show_Calendar then
                          response.write "<INPUT TYPE=""RADIO"" NAME=""STATUS"" Value=""1"""
                          if rs("Status") = 1 then response.write " CHECKED"
                          response.write " CLASS=Medium>&nbsp;&nbsp;"
                          if rs("Status") then
                            response.write "<B>" & Translate("Live",Login_Language,conn) & "</B>"
                          else
                            response.write Translate("Live",Login_language,conn)
                          end if  
                        else
                          response.write "<INPUT " & Field_Editable & " TYPE=""RADIO"" NAME=""STATUS"" Value=""1"""
                          if rs("Status") = 1 then response.write " CHECKED"
                          response.write " CLASS=Medium>&nbsp;&nbsp;"
                          if rs("Status") then
                            response.write "<B>" & Translate("Live",Login_Language,conn) & "</B>"
                          else
                            response.write Translate("Live",Login_language,conn)
                          end if  
                        end if  
                      else
                        response.write "&nbsp;"
                      end if

                    response.write "</TD>"
                  
                    ' Status Archive
                    
                    response.write "<TD WIDTH=""20%"" BGCOLOR="
                    if rs("Status") = 2 then
                      response.write """#AAAACC"""
                    else
                      response.write """#EEEEEE"""
                    end if
                    response.write " CLASS=Medium>"

                      if Admin_Access = 4 or Admin_Access >=8 then
                        if Show_Calendar then
                          response.write "<INPUT TYPE=""RADIO"" NAME=""STATUS"" Value=""2"""
                          if rs("Status") = 2 then response.write " CHECKED"
                          response.write " CLASS=Medium>&nbsp;&nbsp;"
                          if rs("Status") then
                            response.write "<B>" & Translate("Archive",Login_Language,conn) & "</B>"
                          else
                            response.write Translate("Archive",Login_Language,conn)
                          end if  
                        else
                          response.write "<INPUT " & Field_Editable & " TYPE=""RADIO"" NAME=""STATUS"" Value=""2"""
                          if rs("Status") = 2 then response.write " CHECKED"
                          response.write " CLASS=Medium>&nbsp;&nbsp;"
                          if rs("Status") then
                            response.write "<B>" & Translate("Archive",Login_Language,conn) & "</B>"
                          else
                            response.write Translate("Archive",Login_Language,conn)
                          end if  
                        end if  
                      else
                        response.write "&nbsp;"
                      end if

'                      if Content_Group > 0 and not Show_Calendar then
'                        response.write "<INPUT TYPE=""HIDDEN"" NAME=""STATUS"" Value=""" & rs("Status") & """>"
'                      end if
                      %>  
                    </TD>
                  </TR>
                </TABLE>
              </TD>              
            </TR>

            <%
            ' Content Grouping

    	    	  response.write "<TR>" & vbCfLf
            	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
              response.write Translate("Content Grouping",Login_Language,conn) & ":"
              response.write "</TD>" & vbCfLf
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=10 HEIGHT=10>"
              response.write "</TD>" & vbCfLf
    	        response.write "<TD BGCOLOR=""White"" CLASS=Medium>"

              if (Category_Code < 8000 or Category_Code > 8999) and Admin_Access > 2 then
              %>
                <SELECT CLASS=Medium LANGUAGE="JavaScript" ONCHANGE="window.location.href='/SW-Administrator/Calendar_edit.asp?ID=<%=rs("ID")%>&Site_ID=<%=Site_ID%>&Content_Group='+this.options[this.selectedIndex].value" NAME="Content_Group">
              <%
                SQL = "SELECT * FROM Content_Group ORDER BY ID"
                Set rsContent_Group = Server.CreateObject("ADODB.Recordset")
                rsContent_Group.Open SQL, conn, 3, 3
                                              
                Do while not rsContent_Group.EOF
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

                response.write "</SELECT>"
            else
              response.write Translate("Individual",Login_Language,conn)
              response.write "<INPUT TYPE=""HIDDEN"" NAME=""Content_Group"" VALUE=""0"">"
              Content_Group = 0
            end if  

            response.write "</TD>" & vbCfLf
            response.write "</TR>" & vbCfLf

      			' Content Grouping PI/Campaign Select

            if Content_Group > 0 and Admin_Access > 2 then
            
     	    	  response.write "<TR>" & vbCrLf
            	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"

              select case Content_Group
                case 1, 2   ' Product Introduction Kits
                  
                
                  response.write Translate("Product Introduction Name",Login_Language,conn) & ":"
                  SQL = "SELECT * FROM Calendar WHERE Site_ID=" & Site_ID & " AND Code=8000 AND Link IS NULL ORDER BY Product"
                case 3, 4   ' Campaigns
                  SQL = "SELECT * FROM Calendar WHERE Site_ID=" & Site_ID & " AND Code=8001 AND Link IS NULL ORDER BY Product"
                  response.write Translate("Campaign Name",Login_Language,conn) & ":"
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
                response.write "<SELECT CLASS=Medium NAME=""Campaign"">" & vbCrLf
                response.write "<OPTION VALUE=""0"">" & Translate("Select from this list",Login_language,conn) & "</OPTION>" & vbCrLf
                Do while not rsCampaign.EOF
                  response.write "<OPTION"
                  if rs("Campaign") = rsCampaign("ID") then response.write " SELECTED"
                  response.write " CLASS=Medium VALUE=""" & rsCampaign("ID") & """>"
                  select case Content_Group
                    case 1, 2   ' Product Introduction Kits
                      response.write "P"
                    case 3, 4   ' Campaigns
                      response.write "C"
                  end select
                  for i = Len(rsCampaign("ID")) + 1 to 5
                    response.write "0"
                  next  
                  response.write rsCampaign("ID") & " - " & Mid(rsCampaign("Title"),1,35) & "</OPTION>" & vbCrLf
              	  rsCampaign.MoveNext 
                loop
                  
                response.write "</SELECT>" & vbCrLf
              else
                response.write "<SPAN CLASS=MediumRed>" & Translate("None Available - Change to Individual or Contact Site Administrator", Login_Language,conn) & "</SPAN>" & vbCrLf
              end if
                   
              rsCampaign.close
              set rsCampaign = nothing
              
              response.write "</TD>" & vbCrLf
              response.write "</TR>" & vbCrLf
            end if
            %>            

            <!-- Category -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%
                response.write  Translate("Category",Login_Language,conn) & ":"
                if Content_Group > 0 and Show_Calendar then
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
                Category_ID = rs("Category_ID")
                Call Get_Show_Values
                %>
              </TD>                                              
            </TR>

             <!-- Sub-Category -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%
                response.write Translate("Sub-Category",Login_Language,conn) & ":"

                if Admin_Access >= 8 then
                  response.write "<BR><BR>"
                  response.write "<FONT CLASS=Small>" & Translate("Note: New Sub-Categories must be in English",Login_Language,conn) & "</FONT>"
                end if
                %>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <%
                if not isblank(rs("Sub_Category")) then
                  if RestoreQuote(rs("Sub_Category")) <> Translate(RestoreQuote(rs("Sub_Category")),Login_Language,conn) then
                    response.write "<FONT CLASS=Small>" & Translate(RestoreQuote(rs("Sub_Category")),Login_Language,conn) & "</FONT><BR>"
                  end if
                end if

                response.write "<SELECT NAME=""Sub_Category_New"" CLASS=Medium>"
                if Admin_Access >=8 then
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list or enter new below",Login_Language,conn) & "</OPTION>"
                else  
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list",Login_Language,conn) & "</OPTION>"
                end if
                  
                SQL       = "SELECT Content_Sub_Category.Site_ID, Content_Sub_Category.Sub_Category, Content_Sub_Category.Code, Content_Sub_Category.Language "
                SQL = SQL & "FROM Content_Sub_Category "
                SQL = SQL & "GROUP BY Content_Sub_Category.Site_ID, Content_Sub_Category.Sub_Category, Content_Sub_Category.Code, Content_Sub_Category.Language "
                SQL = SQL & "HAVING Content_Sub_Category.Site_ID=" & Site_ID & " "
                SQL = SQL & "AND Content_Sub_Category.Sub_Category IS NOT NULL "
                SQL = SQL & "AND Content_Sub_Category.Code=" & CInt(Category_Code) & " "
                SQL = SQL & "AND Content_Sub_Category.Language='eng'"

                Set rsSubCategoryPreset = Server.CreateObject("ADODB.Recordset")
                rsSubCategoryPreset.Open SQL, conn, 3, 3              
                                                
                if not rsSubCategoryPreset.EOF then
                  rsSCP = True
                  response.write "<OPTION CLASS=Medium VALUE="""">+++ " & Translate("Preset Sub-Categories",Login_Language,conn) & " +++"                      
                  Do while not rsSubCategoryPreset.EOF            
                 	  response.write "<OPTION CLASS=Medium VALUE=""" & rsSubCategoryPreset("Sub_Category") & """"
                    if LCase(rs("Sub_Category")) = LCase(rsSubCategoryPreset("Sub_Category")) then
                      response.write " SELECTED"
                    end if  
                    response.write ">" & rsSubCategoryPreset("Sub_Category") & "</OPTION>" & vbCrLF                 
                	  rsSubCategoryPreset.MoveNext 
                  loop
                else
                  rsSCP = False
                end if                     
                rsSubCategoryPreset.close
                set rsSubCategoryPreset = nothing

                ' Free Form Sub-Categories
                
                SQL       = "SELECT Calendar.Site_ID, Calendar.Sub_Category, Calendar.Code, Calendar.Language "
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
               	  response.write "<OPTION CLASS=Medium VALUE=""" & RestoreQuote(rsSubCategory("Sub_Category")) & """"
                  if LCase(rs("Sub_Category")) = LCase(rsSubCategory("Sub_Category")) then
                    response.write " SELECTED"
                  end if  
                  response.write ">" & RestoreQuote(rsSubCategory("Sub_Category")) & "</OPTION>" & vbCrLf
              	  rsSubCategory.MoveNext 
                loop
                   
                rsSubCategory.close
                set rsSubCategory=nothing
                    
                response.write "</SELECT>"

                response.write "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF="""" onclick=""Category_Window=window.open('/sw-administrator/subcategory_list.asp?Site_ID=" & Site_ID & "&Language=" & Login_Language &  "','Category_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Category_Window.focus();return false;"" CLASS=Medium Title=""Category / Sub Category - Matrix Listing""><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Matrix",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"

                if Admin_Access >= 8 then
                  response.write "<BR>"
                  response.write "<INPUT TYPE=""Text"" NAME=""Sub_Category"" SIZE=50 MAXLENGTH=255 VALUE=""" & RestoreQuote(rs("Sub_Category")) & """ CLASS=Medium>"                
                end if
                %>  
              </TD>
            </TR>
                                                
            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>

             <!-- Product -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%=Translate("Product or Product Family",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <INPUT TYPE="Text" NAME="Product" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("PRODUCT"))%>" CLASS=Medium>                
                <BR>
                <SELECT NAME="Product_New" CLASS=Medium>
                <OPTION CLASS=Medium VALUE=""><%=Translate("Enter Above or Select from this List",Login_Language,conn)%></OPTION>
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
                %>
                </SELECT>
              </TD>
            </TR>
            
            <!-- Title -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Title",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("TITLE"))%>" CLASS=Medium>
              </TD>
            </TR>
    
    				<%
            ' Description
            
            response.write "<TR>"
            response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
            response.write Translate("Description",Login_Language,conn) & ":"
            response.write "</TD>"
            response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
    	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
            response.write "<TEXTAREA NAME=""Description"" COLS=53 ROWS=6 MAXLENGTH=""4000"" CLASS=Medium>" & RestoreQuote(rs("DESCRIPTION")) & "</TEXTAREA>"
            response.write "</TD>"
            response.write "</TR>"

            ' Special Instructions
            
            response.write "<TR>"
            response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
            response.write Translate("Special Instructions",Login_Language,conn) & ":"
            response.write "</TD>"
            response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
    	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
            response.write "<TEXTAREA NAME=""Instructions"" COLS=53 ROWS=6 MAXLENGTH=""4000"" CLASS=Medium>" & RestoreQuote(rs("Instructions")) & "</TEXTAREA>"
            response.write "</TD>"
            response.write "</TR>"

            ' Splash Header / Footer - PIK / Campaign
            
            if Category_Code >= 8000 and Category_Code <= 8999 then

              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>"

              ' Header
              response.write "<TR>"
              response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              select case Category_Code
                case 8000  ' PIK
                  response.write Translate("Product Introduction",Login_Language,conn)
                case 8001  ' Campaign
                  response.write Translate("Campaign",Login_Language,conn)
              end select
              response.write " - " & Translate("Splash Header",Login_Language,conn) & ":"
              response.write "</TD>"
              response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<TEXTAREA NAME=""Splash_Header"" COLS=53 ROWS=8 MAXLENGTH=""4000"" CLASS=Medium>" & RestoreQuote(rs("Splash_Header")) & "</TEXTAREA>"
              response.write "</TD>"
              response.write "</TR>"

              ' Introduction Letter ID Number
              response.write "<TR>"
            	response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              select case Category_Code
                case 8000  ' PIK
                  response.write Translate("Product Introduction - Letter ID Number",Login_Language,conn)
                case 8001  ' Campaign
                  response.write Translate("Campaign Introduction - Letter ID Number",Login_Language,conn)
              end select
              
              response.write "</TD>"
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "&nbsp;"
              response.write "</TD>"
    	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<INPUT TYPE=""Text"" NAME=""Item_Number_2"" SIZE=""50"" MAXLENGTH=""50"" VALUE=""" & rs("Item_Number_2") & """ CLASS=Medium>"
              response.write "</TD>"
              response.write "</TR>"

              ' Footer
              response.write "<TR>"
              response.write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>"
              select case Category_Code
                case 8000  ' PIK
                  response.write Translate("Product Introduction",Login_Language,conn)
                case 8001  ' Campaign
                  response.write Translate("Campaign",Login_Language,conn)
              end select
              response.write " - " & Translate("Splash Footer",Login_Language,conn) & ":"
              response.write "</TD>"
              response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<TEXTAREA NAME=""Splash_Footer"" COLS=53 ROWS=8 MAXLENGTH=""4000"" CLASS=Medium>" & RestoreQuote(rs("Splash_Footer")) & "</TEXTAREA>"
              response.write "</TD>"
              response.write "</TR>"
              
            end if
            %>

            <!-- Item Number -->
            
            <% if Show_Item_Number = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%=Translate("Item / Reference Number",Login_Language,conn)%>&nbsp;1:&nbsp;<SPAN CLASS=SmallRed>(<%=Translate("Oracle / Mfgpro",Login_Language,conn)%>)</SPAN>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Item_Number" SIZE="50" MAXLENGTH="7" VALUE="<%=RestoreQuote(rs("Item_Number"))%>" CLASS=Medium>&nbsp;&nbsp;<%=Translate("Rev",Login_Language,conn)%>:&nbsp;<INPUT TYPE="Text" NAME="Revision_Code" SIZE="4" MAXLENGTH="10" VALUE="<%=RestoreQuote(rs("Revision_Code"))%>" CLASS=Medium>&nbsp;
                <%
                response.write "<INPUT TYPE=""Checkbox"" NAME=""Item_Number_Show"""
                if CInt(rs("Item_Number_Show")) = CInt(True) then
                  response.write " CHECKED"
                end if
                response.write " CLASS=Medium>&nbsp;&nbsp;" & Translate("Show",Login_Language,conn)
                %>
              </TD>
            </TR>
            
              <% if Category_Code < 8000 then %>
      				<TR>
              	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                  <%=Translate("Item / Reference Number",Login_Language,conn)%>&nbsp;2:&nbsp;<SPAN CLASS=Small>(<%=Translate("Legacy",Login_Language,conn)%>)</SPAN>
                </TD>
               	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                  &nbsp;
                </TD>                                                              
      	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                  <INPUT TYPE="Text" NAME="Item_Number_2" SIZE="50" MAXLENGTH="50" VALUE="<%=RestoreQuote(rs("Item_Number_2"))%>" CLASS=Medium>
                </TD>
              </TR>
              <% end if %>                     
            
            <% end if %>

            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>    

            <!-- Location -->
    
            <% if Show_Location = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Location",Login_Language,conn)%> <FONT CLASS=Small>(<%=Translate("City",Login_Language,conn)%>, <%=Translate("State",Login_Language,conn)%> <%=Translate("Country",Login_Language,conn)%>)</FONT>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT TYPE="Text" NAME="Location" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("LOCATION"))%>" CLASS=Medium>
              </TD>
            </TR>
            <% end if %>
            
            <%
            ' Image Store Locator ID
            if Show_ImageStore = True then
    				  response.write "<TR>" & vbCrLf
            	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
              response.write Translate("Image Store Reference Number",Login_Language,conn) & ":"
              response.write "</TD>" & vbCrLf
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
              response.write "&nbsp;"
              response.write "</TD>" & vbCrLf
    	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<INPUT TYPE=""Text"" NAME=""Image_Locator"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & rs("Image_Locator") & """ CLASS=Medium>"
              
              SQLSite = "SELECT * FROM Site WHERE Site_Code='Image-Store'"
              Set rsSite = Server.CreateObject("ADODB.Recordset")
              rsSite.Open SQLSite, conn, 3, 3
  
              if not rsSite.EOF then
                Link_Name    = Replace(rsSite("URL"),"https://support.fluke.com","https://" & request.ServerVariables("SERVER_NAME"))
                Link_Name    = Replace(Link_Name,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))                           
              end if
              rsSite.close
              set rsSite = nothing
              
              if not isblank(rs("Image_Locator")) then
                if isnumeric(rs("Image_Locator")) then
                  response.write "&nbsp;&nbsp;<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=" & rs("Image_Locator") & "&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" Title=""View Individual Image (Site View)"" CLASS=Medium><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("View",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"                
                  response.write "&nbsp;&nbsp;" & Translate("Search",Login_Language,conn) & ":&nbsp;"
                  response.write "<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" CLASS=Medium><FONT CLASS=NavLeftHighlight1 Title=""Search for Individual Image. Remember to Copy the Image's Object ID into the Image Store Reference Number field."">&nbsp;&nbsp;" & Translate("Individual",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"
                elseif LCase(rs("Image_Locator")) = "search" then
                  response.write "&nbsp;&nbsp;<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=" & rs("Image_Locator") & "&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" Title=""View Image Store Search (Site View)"" CLASS=Medium><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("View",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"                
                else
                  response.write "&nbsp;&nbsp;<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=" & rs("Image_Locator") & "&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" Title=""View Image Collection (Site View)"" CLASS=Medium><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("View",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"                
                  response.write "&nbsp;&nbsp;" & Translate("Edit",Login_Language,conn) & ":&nbsp;"
                  response.write "<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=edit:" & rs("Image_Locator") & "&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" CLASS=Medium><FONT CLASS=NavLeftHighlight1 Title=""Edit/Add to Existing Image Collection."">&nbsp;&nbsp;" & Translate("Collection",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"
                end if  
              end if
              
              if isblank(rs("Image_Locator")) then
                response.write "&nbsp;&nbsp;" & Translate("Search",Login_Language,conn) & ":&nbsp;"
                response.write "<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" CLASS=Medium><FONT CLASS=NavLeftHighlight1 Title=""Search for Individual Image. Remember to Copy the Image's Object ID into the Image Store Reference Number field."">&nbsp;&nbsp;" & Translate("Individual",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"
                response.write "&nbsp;&nbsp;" & Translate("New",Login_Language,conn) & ":&nbsp;"
                response.write "<A HREF="""" onclick=""Image_Store=window.open('" & Link_Name & "/Default.asp?Site_ID=" & Site_ID & "&Locator=NEW&KillBackURL=-1','Image_Store','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Image_Store.focus();return false;"" CLASS=Medium><FONT CLASS=NavLeftHighlight1 Title=""Create New Image Collection. Remember to Copy the Collection's Object ID into the Image Store Reference Number field."">&nbsp;&nbsp;" & Translate("Collection",Login_language,conn) & "&nbsp;&nbsp;</FONT></A>"
              end if
        
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
                <INPUT TYPE="Text" NAME="Link" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("LINK"))%>" CLASS=Medium>
              </TD>
            </TR>
            <% else %>
              <INPUT TYPE="HIDDEN" NAME="Link" VALUE="<%=RestoreQuote(rs("LINK"))%>">
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
                <% response.write "<TABLE WIDTH=""100%"" BORDER=0>"
                
                 if rs("Link_PopUp_Disabled") = True then
                    response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"" CHECKED CLASS=Medium></TD><TD CLASS=Medium>" & Translate("Disable",Login_Language,conn) & "</TD></TR>"
                   else
                    response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"" CLASS=Medium></TD><TD CLASS=Medium>" & Translate("Disable",Login_Language,conn) & "</TD></TR>"
                 end if
                
                 response.write "</TABLE>"
                 %>
              </TD>
            </TR>
            <% end if %>

             <!-- Include File -->
    
            <% if Show_Include = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Content File",Login_Language,conn)%> <FONT CLASS=Small>(HTM or ASP)</FONT>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Include"" VALUE=""" & Path_Include & """>"
                if isblank(rs("Include")) then
                  response.write "<INPUT TYPE=""File"" NAME=""Include"" SIZE=""30"" MAXLENGTH=""50"" CLASS=Medium onblur=""Check_Filename(this);""><BR>"

                  response.write "<SELECT NAME=""Include_Existing"" CLASS=Medium>" & vbCrLf
  
                  SQL = "SELECT Calendar.Site_ID, Calendar.Include FROM Calendar GROUP BY Calendar.Site_ID, Calendar.Include HAVING (((Calendar.Site_ID)=" & Site_ID & ") AND ((Calendar.Include) Is Not Null Or Not (Calendar.Include)=''))"
  
                  Set rsInclude = Server.CreateObject("ADODB.Recordset")
                  rsInclude.Open SQL, conn, 3, 3
  
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list or upload new above",Login_Language,conn) & "</OPTION>" & vbCrLF
                  Do while not rsInclude.EOF            
                 	  response.write "<OPTION CLASS=Medium VALUE=""" & LCase(rsInclude("Include")) & """>" & LCase(rsInclude("Include")) & "</OPTION>" & vbCrLf
                	  rsInclude.MoveNext 
                  loop
                  response.write "<OPTION CLASS=Medium VALUE="""">-----------------------------------------------------------------------</OPTION>" & vbCrLf
                  response.write "</SELECT>" & vbCrLf
                  
                  rsInclude.close
                  set rsInclude=nothing
                  
                else
                  response.write "<INPUT TYPE=""Text"" NAME=""Include"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & LCase(rs("Include")) & """ CLASS=Medium onblur=""Check_Filename(this);"">&nbsp;&nbsp"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Include_Existing"" VALUE=""" & LCase(rs("Include")) & """>"                 
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Include_Size"" VALUE=""" & rs("Include_Size") & """>"
                  %>
                  <A HREF="http://<%=Request("SERVER_NAME")%>/<%=Site_Code%>/<%=rs("Include")%>" TARGET="" onclick="openit_mini('http://<%=Request("SERVER_NAME")%>/<%=Site_Code%>/<%=rs("Include")%>','Vertical');return false;"><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;<%=Translate("View",Login_Language,conn)%>&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;&nbsp;
                  <%
                  response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_Include"" CLASS=Medium>&nbsp;&nbsp;" & Translate("Unattach File",Login_Language,conn)
                  if not isblank(rs("Include_Size")) then
                    if isnumeric(rs("Include_Size")) and rs("Include_Size") <> 0 then
                      response.write "<BR><FONT CLASS=Small>" & Translate("File Size",Login_Language,conn) & ": " & CInt(CDbl(rs("Include_Size") / 1024)) & " KB</FONT>"
                    end if  
                  end if                    
                end if
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
                if isblank(rs("File_Name")) then
                  response.write "<FONT CLASS=SmallRed>(" & Translate("View Locally and Virus Scan <B>Prior</B> to uploading file",Login_Language,conn) & ")</FONT>"
                end if
                %>  
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File"" VALUE=""" & Path_File & """>"

                if isblank(rs("File_Name")) then

                  response.write "<INPUT TYPE=""File"" NAME=""File_Name"" SIZE=""30"" MAXLENGTH=""50"" CLASS=Medium onblur=""Check_Filename(this);"">"

if 1=2 then
                  response.write "<BR><SELECT NAME=""File_Existing"" CLASS=Medium>" & vbCrLf
  
                  SQL = "SELECT Calendar.Site_ID, Calendar.File_Name, Calendar.ID, Calendar.Category_ID, "  & vbcrlf &_
                          "REVERSE(LEFT(REVERSE(Calendar.File_Name),CHARINDEX('/',REVERSE(Calendar.File_Name))-1)) AS MyFile " & vbcrlf &_
                        "FROM Calendar " & vbcrlf &_
                        "GROUP BY Calendar.Site_ID, Calendar.File_Name, Calendar.Category_ID, Calendar.ID " & vbcrlf &_
                        "HAVING ((Calendar.Site_ID=" & Site_ID & ") "

                        select case Admin_Access
                          case 4,8,9  ' Allow Administrators to see all files for all Categories
                          case else   ' Limit Selection to Category only for Submitters
                            SQL = SQL & "AND (Calendar.Category_ID=" & CInt(rs("Category_ID")) & ") "
                        end select    

                  SQL = SQL & "AND (Calendar.File_Name Is Not Null Or Not Calendar.File_Name='')) " & vbcrlf &_
                        "ORDER BY MyFile"
  
                  Set rsFile = Server.CreateObject("ADODB.Recordset")
                  rsFile.Open SQL, conn, 3, 3
  
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list or upload new above",Login_Language,conn) & "</OPTION>" & vbCrLf
                  response.write "<OPTION CLASS=Medium VALUE="""">" & "</OPTION>" & vbCrLf                  

                  Do while not rsFile.EOF
                    if instr(1,LCase(rsFile("File_Name")),".zip") = 0 then            

                      response.write "<OPTION CLASS=Medium VALUE=""" & rsFile("ID") & """>"
                      
                      if Instr(1, rsFile("File_Name"), "/") > 0 then
                        response.write LCase(Mid(rsFile("File_Name"), InStrRev(rsFile("File_Name"), "/") + 1))
                      else
                        response.write LCase(rsFile("File_Name"))
                      end if
                      response.write "</OPTION>" & vbCrLf  
                    end if  
                	  rsFile.MoveNext 
                  loop                   
                  response.write "<OPTION CLASS=Medium VALUE="""">-----------------------------------------------------------------------</OPTION>" & vbCrLf
                  response.write "</SELECT>" & vbCrLf
                  
                  rsFile.close
                  set rsFile=nothing
else
                  response.write "<INPUT TYPE=""HIDDEN"" NAME=""File_Existing"">"
end if
                                  
                else

                  response.write "<INPUT TYPE=""Text"" NAME=""File_Name"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & rs("File_Name") & """ CLASS=Medium onblur=""Check_Filename(this);"">&nbsp;&nbsp"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""File_Existing"" VALUE=""" & LCase(rs("File_Name")) & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""File_Size"" VALUE=""" & rs("File_Size") & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Archive_Existing"" VALUE=""" & LCase(rs("Archive_Name")) & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Archive_Size"" VALUE=""" & rs("Archive_Size") & """>"
                  %>
                  <A HREF="http://<%=Request("SERVER_NAME")%>/<%=Site_Code%>/<%=rs("File_Name")%>" TARGET="" onclick="openit_mini('http://<%=Request("SERVER_NAME")%>/<%=Site_Code%>/<%=rs("File_Name")%>','Vertical');return false;"><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;<%=Translate("View",Login_Language,conn)%>&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;&nbsp;
                  <%
                  response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_File"" CLASS=Medium>&nbsp;&nbsp;" & Translate("Unattach File",Login_Language,conn)
                  if not isblank(rs("File_Size")) then
                    if isnumeric(rs("File_Size")) and rs("File_Size") <> 0 then
                      response.write "<BR><FONT CLASS=Small>" & Translate("File Size",Login_Language,conn) & ": " & FormatNumber((rs("File_Size") / 1024),0) & " KB</FONT>"
                    end if
                    if isnumeric(rs("Archive_Size")) and rs("Archive_Size") <> 0 then
                      response.write "&nbsp;&nbsp;|&nbsp;&nbsp;<FONT CLASS=Small>" & Translate("Compressed Size",Login_Language,conn) & ": " & FormatNumber((rs("Archive_Size") / 1024),0) & " KB</FONT>"
                    end if  
                    if (isnumeric(rs("File_Size")) and rs("File_Size") <> 0) and (isnumeric(rs("Archive_Size")) and rs("Archive_Size") <> 0) then
                      response.write "&nbsp;&nbsp;|&nbsp;&nbsp;<FONT CLASS=Small>" & Translate("Compression",Login_Language,conn) & ": " & FormatNumber(((1 - (CDbl(rs("Archive_Size")) / CDbl(rs("File_Size")))) * 100),0) & " %</FONT>"
                    end if
                  end if  
                end if
                %>
              </TD>
            </TR>
            <% end if %>

            <!-- Upload POD File -->
    
            <% if Show_File_POD = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium VALIGN=TOP>
                <%
                response.write  Translate("Asset File",Login_Language,conn) & " - " & Translate("(POD Resolution)",Login_Language,conn) & ": "
                if isblank(rs("File_Name_POD")) then
                  response.write "<FONT CLASS=SmallRed>(" & Translate("View Locally and Virus Scan <B>Prior</B> to uploading file",Login_Language,conn) & ")</FONT>"
                end if
                %>  
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File_POD"" VALUE=""" & Path_File_POD & """>"

                if isblank(rs("File_Name_POD")) then

                  response.write "<INPUT TYPE=""File"" NAME=""File_Name_POD"" SIZE=""30"" MAXLENGTH=""50"" CLASS=Medium onblur="" Check_Filename(this);"">"
if 1=2 then
                  response.write "<BR>"
                  response.write "<SELECT NAME=""File_Existing_POD"" CLASS=Medium>" & vbCrLf
  
                  SQL = "SELECT Calendar.Site_ID, Calendar.File_Name_POD, Calendar.ID, Calendar.Category_ID, "  & vbcrlf &_
                          "REVERSE(LEFT(REVERSE(Calendar.File_Name_POD),CHARINDEX('/',REVERSE(Calendar.File_Name_POD))-1)) AS MyFile " & vbcrlf &_
                        "FROM Calendar " & vbcrlf &_
                        "GROUP BY Calendar.Site_ID, Calendar.File_Name_POD, Calendar.Category_ID, Calendar.ID " & vbcrlf &_
                        "HAVING ((Calendar.Site_ID=" & Site_ID & ") "

                        select case Admin_Access
                          case 4,8,9  ' Allow Administrators to see all files for all Categories
                          case else   ' Limit Selection to Category only for Submitters
                            SQL = SQL & "AND (Calendar.Category_ID=" & CInt(rs("Category_ID")) & ") "
                        end select    

                  SQL = SQL & "AND (Calendar.File_Name_POD Is Not Null Or Not Calendar.File_Name_POD='')) " & vbcrlf &_
                        "ORDER BY MyFile"
  
                  Set rsFile = Server.CreateObject("ADODB.Recordset")
                  rsFile.Open SQL, conn, 3, 3
  
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list or upload new above",Login_Language,conn) & "</OPTION>" & vbCrLf
                  response.write "<OPTION CLASS=Medium VALUE="""">" & "</OPTION>" & vbCrLf                  

                  Do while not rsFile.EOF
                    if instr(1,LCase(rsFile("File_Name_POD")),".zip") = 0 then            

                      response.write "<OPTION CLASS=Medium VALUE=""" & rsFile("ID") & """>"
                      
                      if Instr(1, rsFile("File_Name_POD"), "/") > 0 then
                        response.write LCase(Mid(rsFile("File_Name_POD"), InStrRev(rsFile("File_Name_POD"), "/") + 1))
                      else
                        response.write LCase(rsFile("File_Name_POD"))
                      end if
                      response.write "</OPTION>" & vbCrLf  
                    end if  
                	  rsFile.MoveNext 
                  loop                   
                  response.write "<OPTION CLASS=Medium VALUE="""">-----------------------------------------------------------------------</OPTION>" & vbCrLf
                  response.write "</SELECT>" & vbCrLf
                  
                  rsFile.close
                  set rsFile=nothing
else
                  response.write "<INPUT TYPE=""HIDDEN"" NAME=""File_Existing_POD"">"
end if
                else

                  response.write "<INPUT TYPE=""Text"" NAME=""File_Name_POD"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & rs("File_Name_POD") & """ CLASS=Medium onblur=""Check_Filename(this);"">&nbsp;&nbsp"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""File_Existing_POD"" VALUE=""" & LCase(rs("File_Name_POD")) & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""File_Size_POD"" VALUE=""" & rs("File_Size_POD") & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Archive_Existing_POD"" VALUE=""" & LCase(rs("Archive_Name_POD")) & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Archive_Size_POD"" VALUE=""" & rs("Archive_Size_POD") & """>"
                  %>
                  <A HREF="http://<%=Request("SERVER_NAME")%>/<%=rs("File_Name_POD")%>" TARGET="" onclick="openit_mini('http://<%=Request("SERVER_NAME")%>/<%=rs("File_Name_POD")%>','Vertical');return false;"><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;<%=Translate("View",Login_Language,conn)%>&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;&nbsp;
                  <%
                  response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_File_POD"" CLASS=Medium>&nbsp;&nbsp;" & Translate("Unattach File",Login_Language,conn)
                  if not isblank(rs("File_Size_POD")) then
                    if isnumeric(rs("File_Size_POD")) and rs("File_Size_POD") <> 0 then
                      response.write "<BR><FONT CLASS=Small>" & Translate("File Size",Login_Language,conn) & ": " & FormatNumber(Cdbl(rs("File_Size_POD") / 1024),0) & " KB</FONT>"
                    end if
                    if isnumeric(rs("Archive_Size_POD")) and rs("Archive_Size_POD") <> 0 then
                      response.write "&nbsp;&nbsp;|&nbsp;&nbsp;<FONT CLASS=Small>" & Translate("Compressed Size",Login_Language,conn) & ": " & FormatNumber(CDbl(rs("Archive_Size_POD") / 1024),0) & " KB</FONT>"
                    end if  
                    if (isnumeric(rs("File_Size_POD")) and rs("File_Size_POD") <> 0) and (isnumeric(rs("Archive_Size_POD")) and rs("Archive_Size_POD") <> 0) then
                      response.write "&nbsp;&nbsp;|&nbsp;&nbsp;<FONT CLASS=Small>" & Translate("Compression",Login_Language,conn) & ": " & FormatNumber(((1 - (CDbl(rs("Archive_Size_POD")) / CDbl(rs("File_Size_POD")))) * 100),0) & " %</FONT>"
                    end if
                  end if  
                end if
                %>
              </TD>
            </TR>
            <% end if %>

             <!-- Thumbnail File -->
    
            <% if Show_Thumbnail = True then %>
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium VALIGN=Top>
                <%if not isblank(rs("Thumbnail")) then response.write "<IMG SRC=""/" & Site_Code & "/" & rs("Thumbnail") & """ BORDER=1 ALIGN=RIGHT WIDTH=80>"%>
                <%=Translate("Thumbnail File",Login_Language,conn)%> <FONT CLASS=Small> - (GIF or JPG): </FONT>
                <%if isblank(rs("Thumbnail")) then response.write "&nbsp;&nbsp;&nbsp;<FONT CLASS=SmallRed>(" & Translate("View Locally and Virus Scan <B>Prior</B> to uploading file",Login_Language,conn) & ")</FONT>"%>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                                            
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>                
                <%
                response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Thumbnail"" VALUE=""" & Path_Thumbnail & """>"

                if isblank(rs("Thumbnail")) then
                  response.write "<INPUT TYPE=""File"" NAME=""Thumbnail"" SIZE=""30"" MAXLENGTH=""50"" CLASS=Medium onblur=""Check_Filename(this);"">"
if 1=2 then
                  response.write "<BR>"
                  response.write "<SELECT NAME=""Thumbnail_Existing"" CLASS=Medium>" & vbCrLf
  
                  SQL = "SELECT Calendar.Site_ID, Calendar.Thumbnail FROM Calendar GROUP BY Calendar.Site_ID, Calendar.Thumbnail HAVING (((Calendar.Site_ID)=" & Site_ID & ") AND ((Calendar.Thumbnail) Is Not Null Or Not (Calendar.Thumbnail)=''))"
  
                  Set rsThumbnail = Server.CreateObject("ADODB.Recordset")
                  rsThumbnail.Open SQL, conn, 3, 3
  
                  response.write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from this list or upload new above",Login_Language,conn) & "</OPTION>" & vbCrLf
                  Do while not rsThumbnail.EOF            
                 	  response.write "<OPTION CLASS=Medium VALUE=""" & LCase(rsThumbnail("Thumbnail")) & """>" & LCase(rsThumbnail("Thumbnail")) & "</OPTION>" & vbCrLf
                	  rsThumbnail.MoveNext 
                  loop                   
                  response.write "<OPTION CLASS=Medium VALUE="""">-----------------------------------------------------------------------</OPTION>" & vbCrLf
                  response.write "</SELECT><BR>"
                  
                  rsThumbnail.close
                  set rsThumbnail=nothing
else
                  response.write "<INPUT TYPE=""HIDDEN"" NAME=""Thumbnail_Existing"">"
end if                                   
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""Thumbnail_Request"""
                  if CInt(rs("Thumbnail_Request")) = CInt(True) then
                    response.write " CHECKED"
                  end if  
                  response.write ">&nbsp;&nbsp;" & Translate("Request Thumbnail",Login_Language,conn)

                else
                  response.write "<INPUT TYPE=""Text"" NAME=""Thumbnail"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & rs("Thumbnail") & """ CLASS=Medium onblur=""Check_Filename(this);"">&nbsp;&nbsp"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Thumbnail_Existing"" VALUE=""" & LCase(rs("Thumbnail")) & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Thumbnail_Size"" VALUE=""" & rs("Thumbnail_Size") & """>"
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Thumbnail_Request"" VALUE=""off"">"
                  %>
                  <A HREF="http://<%=Request("SERVER_NAME")%>/<%=Site_Code%>/<%=rs("Thumbnail")%>" TARGET="" onclick="openit_mini('http://<%=Request("SERVER_NAME")%>/<%=Site_Code%>/<%=rs("Thumbnail")%>','Vertical');return false;"><FONT CLASS=NavLeftHighlight1>&nbsp;&nbsp;<%=Translate("View",Login_Language,conn)%>&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;&nbsp;
                  <%
                  response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_Thumbnail"" CLASS=Medium>&nbsp;&nbsp;" & Translate("Unattach File",Login_Language,conn)
                  if not isblank(rs("Thumbnail_Size")) then
                    if isnumeric(rs("Thumbnail_Size")) and rs("Thumbnail_Size") <> 0 then
                      response.write "<BR><FONT CLASS=Small>" & Translate("File Size",Login_Language,conn) & ": " & FormatNumber(CDbl(rs("Thumbnail_Size") / 1024),0) & " KB</FONT>"
                    end if  
                  end if  

                end if
                %>
              </TD>
            </TR>
            <% end if %>
            
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
                <INPUT TYPE="Text" NAME="Forum_ID" SIZE="50" MAXLENGTH="10" VALUE="<%=RestoreQuote(rs("Forum_ID"))%>" CLASS=Medium>
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
                  response.write "<INPUT CLASS=Medium TYPE=""Checkbox"" NAME=""Forum_Moderated"""
                  if CInt(rs("Forum_Moderated")) = CInt(True) then
                    response.write " CHECKED"
                  end if  
                  response.write ">"
                  
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
                      if rs("Forum_Moderator_ID") = rsModerator("ID") then
                        response.write " SELECTED"
                      end if
                      response.write ">" & rsModerator("LastName") & ", " & rsModerator("FirstName") & "</OPTION>" & vbCrLf
                      
                      rsModerator.MoveNext
                      
                    loop
                    
                  end if
                  
                  rsModerator.close
                  set rsModerator = nothing
                   
                  ' Forum Moderators (Alternate)
                  SQL = "SELECT UserData.ID, UserData.SubGroups, UserData.FirstName, UserData.LastName, UserData.Region FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND (UserData.SubGroups LIKE '%administrator%' OR UserData.SubGroups LIKE '%content%') ORDER BY UserData.LastName, UserData.FirstName"
                  Set rsModerator = Server.CreateObject("ADODB.Recordset")
                  rsModerator.Open SQL, conn, 3, 3

                  if not rsModerator.EOF then

                    response.write "<OPTION CLASS=NavLeftHighlight1 VALUE="""">" & Translate("Alternates",Login_Language,conn) & "</OPTION>" & vbCrLf

                    do while not rsModerator.EOF
                      response.write "<OPTION CLASS=Region" & rsModerator("Region") & "NavMedium VALUE=""" & rsModerator("ID") & """"
                      if rs("Forum_Moderator_ID") = rsModerator("ID") then
                        response.write " SELECTED"
                      end if
                      response.write ">" & rsModerator("LastName") & ", " & rsModerator("FirstName") & "</OPTION>" & vbCrLf
                      
                      rsModerator.MoveNext
                      
                    loop
                    
                  end if
                  
                  rsModerator.close
                  set rsModerator = nothing
                  
                  response.write "</SELECT>" & vbCrLf & vbCrLf
                %>
              </TD>
            </TR>
            <% end if %>
            
            <% if Show_Forum = True then %>
              <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>                        
            <% end if %>

           
            <%
              ' Pre Announcement Days before BDate               
    				  response.write "<TR>"
       				response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
       				response.write Translate("Pre-Announce",Login_Language,conn) & ":"
              response.write "</TD>"
             	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
    	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              response.write "<INPUT " & Field_Editable & " TYPE=""Text"" NAME=""LDAYS"" SIZE=""30"" MAXLENGTH=""3"" VALUE=" & rs("LDAYS") & " CLASS=Medium>&nbsp;&nbsp;" & Translate("days before",Login_Language,conn)
              if Content_Group > 0 and not Show_Calendar then
                response.write "<INPUT TYPE=""Hidden"" NAME=""LDAYS"" VALUE=" & rs("LDAYS") & ">"
              end if
              response.write "</TD>"
              response.write "</TR>"
            %>

             <!-- Beginning Date -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>                
                <%=Translate("Beginning Date",Login_Language,conn)%> <FONT CLASS=Small>(mm/dd/yyyy)</FONT>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT VALIGN=TOP CLASS=Medium>
                <INPUT <%=Field_Editable%> TYPE="Text" NAME="BDate" SIZE="30" MAXLENGTH="10" VALUE="<%=FormatDate(1, rs("BDATE"))%>" CLASS=Medium>&nbsp;&nbsp;
                <%
                if Content_Group > 0 and not Show_Calendar then
                  response.write "<INPUT TYPE=""Hidden"" NAME=""BDate"" VALUE=" & FormatDate(1,rs("BDATE")) & ">"               
                else %>
                	<A HREF="javascript:ShowCalendar(document.EditContent.myCalendarImage1, document.EditContent.BDate, null, '<%=Date()%>', '<%=DateAdd("m",12,Date())%>')">
                  <IMG ALIGN=TOP BORDER=0 HEIGHT=21 ID=myCalendarImage1 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></A>&nbsp;&nbsp;
                <%
                end if %>  
                <%=Translate("through",Login_Language,conn)%>
              </TD>
            </TR>
        
             <!-- Ending Date -->
    
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>                
                <%=Translate("Ending Date",Login_Language,conn)%> <FONT CLASS=Small>(mm/dd/yyyy)</FONT>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <INPUT <%=Field_Editable%> TYPE="Text" NAME="EDate" SIZE="30" MAXLENGTH="10" VALUE="<%=FormatDate(1, rs("EDATE"))%>" CLASS=Medium>&nbsp;&nbsp;
                <%
                if Content_Group > 0 and not Show_Calendar then
                  response.write "<INPUT TYPE=""Hidden"" NAME=""EDate"" VALUE=" & FormatDate(1,rs("EDate")) & ">"
                else %>
                	<A HREF="javascript:ShowCalendar(document.EditContent.myCalendarImage2, document.EditContent.EDate, null, '<%=Date()%>', '<%=DateAdd("m",12,Date())%>')">
                  <IMG ALIGN=TOP BORDER=0 HEIGHT=21 ID=myCalendarImage2 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></A>&nbsp;&nbsp;
                <%
                end if %>
                <%=Translate("then",Login_Language,conn)%>
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
                <INPUT TYPE="Text" NAME="XDays" SIZE="30" MAXLENGTH="3" VALUE="<%=rs("XDays")%>" CLASS=Medium>&nbsp;&nbsp<%=Translate("days after ending date",Login_Language,conn)%>
              </TD>
            </TR>       

            <!-- Public Embargo Date -->
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%=Translate("Public Release Date",Login_Language,conn) & " <FONT Class=Small>(mm/dd/yyyy):<BR>(" & Translate("Leave blank if same as Beginning Date",Login_Language,conn) & ")</FONT>"%>
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <%
                if IsDate(rs("PEDate")) then
                  if CDate(rs("PEDate")) <> CDate("01/01/1900") then
                    if Content_Group > 0 and not Show_Calendar then
                      response.write "<INPUT DISABLED TYPE=""TEXT"" NAME=""PEDate"" SIZE=""30"" MAXLENGTH=""10"" VALUE=""" & FormatDate(1, rs("PEDATE")) & """ CLASS=Medium>&nbsp;&nbsp;"
                    else
                      response.write "<INPUT TYPE=""TEXT"" NAME=""PEDate"" SIZE=""30"" MAXLENGTH=""10"" VALUE=""" & FormatDate(1, rs("PEDATE")) & """ CLASS=Medium>&nbsp;&nbsp;"
                    end if
                  end if  
                else                    
                  if Content_Group > 0 and not Show_Calendar then
                    response.write "<INPUT DISABLED TYPE=""TEXT"" NAME=""PEDate"" SIZE=""30"" MAXLENGTH=""10"" VALUE="""" CLASS=Medium>&nbsp;&nbsp;"
                  else
                    response.write "<INPUT TYPE=""TEXT"" NAME=""PEDate"" SIZE=""30"" MAXLENGTH=""10"" VALUE="""" CLASS=Medium>&nbsp;&nbsp;"
                  end if  
                end if
                if Content_Group > 0 and not Show_Calendar then
                  response.write "<INPUT TYPE=""HIDDEN"" NAME=""PEDate"" VALUE=" & rs("PEDate") & ">"
                else %>
                	<A HREF="javascript:ShowCalendar(document.EditContent.myCalendarImage3, document.EditContent.PEDate, null, '<%=Date()%>', '<%=DateAdd("m",12,Date())%>')">
                  <IMG ALIGN=TOP BORDER=0 HEIGHT=21 ID=myCalendarImage3 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></A>
                <%
                end if %>
                  
              </TD>
            </TR>

            <!-- Mark as Confidential -->            
            
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
                response.write "<TR><TD WIDTH=20"
                if rs("Confidential") = True then
                  response.write " BGCOLOR=""red"""
                end if
                response.write " CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""Confidential"""
                if rs("Confidential") = True then
                  response.write " CHECKED"
                end if
                response.write " CLASS=Medium></TD><TD CLASS=Small>&nbsp;</TD></TR>"
                response.write "</TABLE>"
                %> 
              </TD>
            </TR>
                   
            <!-- Language -->
            
           <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>            
                       
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <FONT FACE="Arial" SIZE=2 ><%=Translate("Language",Login_Language,conn)%>:</FONT>
              </TD>
            	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER>
                &nbsp;
              </TD>                                                                
    	        <TD BGCOLOR="White" ALIGN=LEFT>
                              
                <SELECT Name="Content_Language" CLASS=Medium>    
                <%
                SQL = "SELECT * FROM Language WHERE Language.Enable=-1" & " ORDER BY Language.Sort"
                Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                rsLanguage.Open SQL, conn, 3, 3
                                      
                Do while not rsLanguage.EOF
                  if lcase(rs("Language")) = rsLanguage("Code") then
                 	  response.write "<OPTION CLASS=Medium SELECTED VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
                  else
                 	  response.write "<OPTION CLASS=Medium VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
                  end if
              	  rsLanguage.MoveNext 
                loop
                
                rsLanguage.close
                set rsLanguage=nothing
                %>
                </SELECT>              
                </FONT>
              </TD>
            </TR>
            
            <!-- Post via Subscription Service -->
    
            <% if Show_Subscription = True then %>

            <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>            
           
      		  <TR>
            	<TD BGCOLOR="#EEEEEE">
                <%=Translate("Send Notice via Subscription Service",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>
                <% response.write "<TABLE WIDTH=""100%"" BORDER=0>"
                
                 if rs("Subscription") = True then
                    response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""Subscription"" CHECKED CLASS=Medium></TD><TD CLASS=Medium>" & Translate("Subscription Service",Login_Language,conn) & "</TD></TR>"
                   else
                    response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""Subscription"" CLASS=Medium></TD><TD CLASS=Medium>" & Translate("Subscription Service",Login_Language,conn) & "</TD></TR>"
                 end if
                
                 response.write "</TABLE>"
                %> 
              </TD>
            </TR>
                      
            <% end if %>
                     
            <!-- NT Sub-Groups -->
            
            <TR><TD COLSPAN=3 BGCOLOR="Gray"></TD></TR>
            
    				<TR>
            	<TD BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=Medium>
                <%=Translate("Select Groups allowed to view this information",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER VALIGN=TOP CLASS=Medium>
                <IMG SRC="/images/required.gif" Border=0 WIDTH="10" HEIGHT="10">
              </TD>                                                              
    	        <TD BGCOLOR="White" CLASS=Medium>
          
              <%

                response.write "<TABLE WIDTH=""100%"">" & vbCrLF
         
                ' End-User
                
                if (Category_Code < 8000 or Category_Code > 8999) and Show_Item_Number = True then
                
                  response.write "<TR><TD WIDTH=20 CLASS=Medium>"
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"""
                  if instr(1,lcase(rs("SubGroups")),lcase("view")) > 0 then
                    response.write " CHECKED"
                  end if  
                  response.write " CLASS=Medium VALUE=""view""></TD>"
                  response.write "<TD CLASS=Medium BGCOLOR=""#FF9966"">&nbsp;" & Translate("Available to Electronic Fulfillment",Login_Language,conn) & " - <SPAN CLASS=Small>(" & Translate("End-User Viewable",Login_Language,conn) & ")</SPAN></TD>"
                  response.write "</TR>"

                  ' Shopping Cart - Order Literature

                  response.write "<TR><TD WIDTH=20 CLASS=Medium>"
                  response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"""
                  if instr(1,lcase(rs("SubGroups")),lcase("shpcrt")) > 0 then
                    response.write " CHECKED"
                  end if  
                  response.write " CLASS=Medium VALUE=""shpcrt""></TD>"
                  response.write "<TD CLASS=Medium BGCOLOR=""#FF9966"">&nbsp;" & Translate("Exclude from Literature Order Shopping Cart",Login_Language,conn) & "</TD>"
                  response.write "</TR>"

                  response.write "<TR><TD HEIGHT=8 WIDTH=20></TD><TD HEIGHT=8></TD></TR>" & vbCrLf
                end if  

                ' Regional Groups                  

                if isblank(Admin_Region) then Admin_Region = 1
                
                for i = 0 to 1
                
                  Select case i
                    case 0
                      SQL = "SELECT SubGroups.*, SubGroups.Order_Num "
                      SQL = SQL & "FROM SubGroups "
                      SQL = SQL & "WHERE SubGroups.Site_ID=" & Site_ID & " AND SubGroups.Region=" & Admin_Region & " AND SubGroups.Enabled=" & CInt(True)
                      SQL = SQL & "ORDER BY SubGroups.Order_Num"
                      
                      Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
                      rsSubGroups.Open SQL, conn, 3, 3
                                            
'                      if instr(1,lcase(rs("SubGroups")),lcase("all")) > 0 then
'                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""all"" CHECKED CLASS=Medium></TD><TD CLASS=MediumRed>&nbsp;<B>" & Translate("All Groups in all Regions",Login_Language,conn) & "</B></TD></TR>" & vbCrLF
'                      else
'                        response.write "<TR><TD WIDTH=20 CLASS=Medium><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""all"" CLASS=Medium></TD><TD CLASS=Medium>&nbsp;<B>" & Translate("All Groups in all Regions",Login_Language,conn) & "</B></TD></TR>" & vbCrLF
'                      end if               
                      
                    case else
                      SQL = "SELECT SubGroups.*, SubGroups.Order_Num "
                      SQL = SQL & "FROM SubGroups "
                      SQL = SQL & "WHERE SubGroups.Site_ID=" & Site_ID & " AND SubGroups.Region<>" & Admin_Region & " AND SubGroups.Enabled=" & CInt(True)
                      SQL = SQL & "ORDER BY SubGroups.Order_Num"
                      
                      Set rsSubGroups = Server.CreateObject("ADODB.Recordset")                      
                      rsSubGroups.Open SQL, conn, 3, 3
                      
                  end select                   
                  
                  if not rsSubGroups.EOF then
                                                       
                    Do while not rsSubGroups.EOF
                       
                      if RegionValue <> Mid(rsSubGroups("Code"),1,1) then
                        RegionValue = Mid(rsSubGroups("Code"),1,1)
                        select case UCase(RegionValue)
                          case "U"
                            RegionColorPointer = 1
                          case "E"
                            RegionColorPointer = 2
                          case "I"
                            RegionColorPointer = 3
                          case else
                            RegionColorPointer = 0
                        end select    
                        Region = Region + 1
                        if Region <= 3 then
                          response.write "<TR><TD HEIGHT=8 WIDTH=20></TD><TD HEIGHT=8></TD></TR>" & vbCrLF
                          if rsSubGroups.RecordCount >= 2 then
                            response.write "<TR><TD CLASS=Medium BGCOLOR="""
                            if Region <> Admin_Region then
                              response.write "Yellow"
                            else
                              response.write "Green"
                            end if
                            response.write """>"
                            response.write "<INPUT TYPE=""Checkbox"""
                            response.write " ONCLICK=""SubGroups_" & Trim(CStr(RegionColorPointer)) & "_Check();"""
                            response.write " NAME=""SubGroups_" & Trim(CStr(RegionColorPointer)) & """ CLASS=Medium>"
                            response.write "</TD><TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;<B>" & Translate("All Groups for this Region",Login_Language,conn) & "</B></TD></TR>" & vbCrLF
                          end if
                        elseif Region > 3 then
                          Region = 3
                        end if
                      end if
                        
                      if instr(1,lcase(rs("SubGroups")),lcase(rsSubGroups("Code"))) > 0 then
                        response.write "<TR>" & vbCrLF
                        response.write "<TD WIDTH=20 CLASS=Medium>" & vbCrLF
                        response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED CLASS=Medium>"
                        response.write "</TD>" & vbCrLF
                        response.write "<TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;" & rsSubGroups("X_Description") & "</TD></TR>" & vbCrLF
                      elseif rsSubGroups("Enabled") = True then
                        response.write "<TR>" & vbCrLF
                        response.write "<TD WIDTH=20 CLASS=Medium>"
                        response.write "<INPUT TYPE=""Checkbox"""
                        if Region <> Admin_Region then
                          response.write " ONCLICK=""SubGroups_Check();"""
                        end if
                        response.write " NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """></TD>" & vbCrLF
                        response.write "<TD CLASS=Medium BGCOLOR=""" & RegionColor(RegionColorPointer) & """>&nbsp;" & rsSubGroups("X_Description") & "</TD>" & vbCrLF
                        response.write "</TR>" & vbCrLF
                      end if
                  
                  	  rsSubGroups.MoveNext 
  
                    loop
                    
                    rsSubGroups.close
                    set rsSubGroups=nothing
                    
                  else
                    region = region + 1
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
                    if instr(1,lcase(rs("SubGroups")),lcase(rsSubGroups("Code"))) > 0 then
                      response.write "<TR>" & vbCrLF
                      response.write "<TD WIDTH=20 CLASS=Medium>" & vbCrLF
                      response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED CLASS=Medium>"
                      response.write "</TD>" & vbCrLF
                      response.write "<TD CLASS=Medium BGCOLOR=""#669999"">&nbsp;" & Translate(rsSubGroups("X_Description"),Login_Language,conn) & "</TD></TR>" & vbCrLF
                    elseif rsSubGroups("Enabled") = True then
                      response.write "<TR>" & vbCrLF
                      response.write "<TD WIDTH=20 CLASS=Medium><FONT FACE=""Arial"" SIZE=2>"
                      response.write "<INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """></TD>" & vbCrLF
                      response.write "<TD CLASS=Medium BGCOLOR=""#669999"">&nbsp;" & Translate(rsSubGroups("X_Description"),Login_Language,conn) & "</TD>" & vbCrLF
                      response.write "</TR>" & vbCrLF
                    end if
                    rsSubGroups.MoveNext
                  loop  
                  
                  rsSubGroups.close
                  set rsSubGroups=nothing
                    
                end if  
                
                response.write "</TABLE>" & vbCrLF & vbCrLF
                                   
                %>
              </TD>
            </TR>
            
          <!-- Restricted to Countries -->

           <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>            
           
    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=Medium>
                <%
                response.write "<INPUT TYPE=""Radio"" NAME=""Country_Reset"" VALUE=""none"" CLASS=Medium ONCLICK=""document." & FormName & ".Country.value='';"""
                if isblank(rs("Country")) or instr(1,LCase(rs("Country")),"none") > 0 then
                  response.write " CHECKED"
                end if
                response.write ">" & vbCrLf
                response.write Translate("No Country Restrictions",Login_Language,conn) & ":<BR>" & vbCrLf & vbCrLf

                response.write "<INPUT TYPE=""Radio"" NAME=""Country_Reset"" VALUE="""" CLASS=Medium"
                if not isblank(rs("Country")) and instr(1,LCase(rs("Country")),"0") = 0 and instr(1,LCase(rs("Country")),"none") = 0 then
                  response.write " CHECKED"
                end if
                response.write ">" & vbCrLf
                response.write Translate("Include only these Countries",Login_Language,conn) & ":<BR>" & vbCrLf

                response.write "<INPUT TYPE=""Radio"" NAME=""Country_Reset"" VALUE=""0"" CLASS=Medium"
                if not isblank(rs("Country")) and (instr(1,LCase(rs("Country")),"0") > 0 and instr(1,LCase(rs("Country")),"none") = 0) then
                  response.write " CHECKED"
                end if
                response.write ">" & vbCrLf
                response.write Translate("Exclude only these Countries",Login_Language,conn) & ":" & vbCrLf
                %>
                
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT VALIGN=TOP CLASS=Medium>
                <%
                Users_Country = rs("Country")

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

    				<TR>
            	<TD BGCOLOR="#EEEEEE" CLASS=MediumRed>
                <%=Translate("Currently Selected Restricted Countries (abbr.)",Login_Language,conn)%>:
              </TD>
             	<TD BGCOLOR="#EEEEEE" ALIGN=CENTER CLASS=Medium>
                &nbsp;
              </TD>                                                              
    	        <TD BGCOLOR="White" ALIGN=LEFT CLASS=Medium>

                <%
                if not isblank(rs("Country")) and rs("Country") <> "none" then
                  if instr(1,rs("Country"),"0, ") = 1 then
                    response.write mid(rs("Country"),3)
                  else  
                    response.write rs("Country")
                  end if  
                else
                  response.write Translate("All Countries can view this Item or Event",Login_Language,conn)
                end if
                %> 
              </TD>
            </TR>
    
            <TR><TD COLSPAN=3 BGCOLOR="Gray" CLASS=Medium></TD></TR>
 
            <!-- Approver Selection -->
           
            <%
            if Admin_Access = 2 or Admin_Access = 4 or Admin_Access >=8 then
          
              SQL =       "SELECT Approvers.* "
              SQL = SQL & "FROM Approvers "
              SQL = SQL & "WHERE Approvers.Site_ID=" & Site_ID & " AND (Approvers.Approver_ID Is Not Null OR Approvers.Approver_ID <> 0) "
              SQL = SQL & "ORDER BY Approvers.Order_Num"
  
              Set rsApprovers = Server.CreateObject("ADODB.Recordset")
              rsApprovers.Open SQL, conn, 3, 3
              
              if not rsApprovers.EOF then
                   
                response.write "<TR>"
                	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium VALIGN=TOP>"
                  response.write Translate("Group Assigned to Approve this Submission",Login_Language,conn) & ":"
                  response.write "</TD>"
                 	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                  response.write "&nbsp;"
                  response.write "</TD>"
        	        response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"

                    response.write "<SELECT NAME=""Review_By_Group"" CLASS=Medium>" & vbCrLf
    
                    if Admin_Access = 4 or Admin_Access>=8 then
                      response.write "<OPTION CLASS=NavLeftHighlight1 VALUE=""0"">" & Translate("Approval by Current Administrator",Login_Language,conn) & "</OPTION>" & vbCrLf
                    end if
                    
                    do while not rsApprovers.EOF
                      response.write "<OPTION CLASS=Region" & rsApprovers("Region") & "NavMedium VALUE=""" & rsApprovers("ID") & """"
                      if not isblank(rs("Review_By_Group")) then
                        if CInt(rs("Review_By_Group")) = CInt(rsApprovers("ID")) then response.write " SELECTED"
                      end if  
                      response.write ">" & rsApprovers("Description") & "</OPTION>" & vbCrLf
                      rsApprovers.MoveNext
                    loop
    
                    response.write "</SELECT>" & vbCrLf
                    response.write "<BR>"
                    response.write "<INPUT TYPE=""Checkbox"" NAME=""Send_EMail_Admin"" CLASS=Medium>&nbsp;&nbsp;" & Translate("Request Review of Submission by Email",Login_Language,conn)

                  response.write "</TD>"
                response.write "</TR>"

              end if
              
              rsApprovers.Close
              set rsApprovers = nothing
              
              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>" & vbCrLf
            
            end if

            ' Re-Assign Submitter / Owner
            
            if Admin_Access = 4 or Admin_Access >=8 then
                               
              ' First check to see if Current Submitter Account is still valid else default to this admin
              SQL =       "SELECT UserData.* "
              SQL = SQL & "FROM UserData "
              SQL = SQL & "WHERE UserData.Site_ID=" & Site_ID & " "
              SQL = SQL & "AND UserData.ID=" & Submitted_By_Current & " "
              SQL = SQL & "AND (UserData.SubGroups LIKE '%administrator%' OR UserData.SubGroups LIKE '%content%' OR UserData.SubGroups LIKE '%submitter%') "

              Set rsSubmitters = Server.CreateObject("ADODB.Recordset")
              rsSubmitters.Open SQL, conn, 3, 3
          
              if rsSubmitters.EOF then
                Submitted_By_Current = Admin_ID
              end if  
              
              rsSubmitters.close
              set rsSubmitters = nothing
              
              ' List all Content Admins
              SQL =       "SELECT UserData.* "
              SQL = SQL & "FROM UserData "
              SQL = SQL & "WHERE (UserData.Site_ID=" & Site_ID & " OR UserData.Site_ID=0)"
              SQL = SQL & "AND (UserData.SubGroups LIKE '%domain%' OR UserData.SubGroups LIKE '%administrator%' OR UserData.SubGroups LIKE '%content%' OR UserData.SubGroups LIKE '%submitter%') "
              SQL = SQL & "ORDER BY UserData.LastName, UserData.FirstName"

              Set rsSubmitters = Server.CreateObject("ADODB.Recordset")
              rsSubmitters.Open SQL, conn, 3, 3
          
              if not rsSubmitters.EOF then
              
          		  response.write "<TR>"
                response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
                response.write Translate("Reassign Owner of this Content to",Login_Language,conn) & ":"
                response.write "</TD>"
                response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
                response.write "&nbsp;"
                response.write "</TD>"
        	      response.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
              
                response.write "<SELECT NAME=""Submitted_By_New"" CLASS=Medium>" & vbCrLf
                response.write "<OPTION CLASS=NavLeftHighlight1 VALUE=""" & Admin_ID & """>(-) " & Translate("Unassigned",Login_Language,conn) & "</OPTION>"                  

                do while not rsSubmitters.EOF
                  response.write "<OPTION CLASS=Region" & rsSubmitters("Region") & "NavMedium VALUE=""" & rsSubmitters("ID") & """"
                  if CInt(Submitted_By_Current) = CInt(rsSubmitters("ID")) then response.write " SELECTED"
                  response.write ">"
                  if instr(1,rsSubmitters("SubGroups"),"domain") > 0 then
                    response.write "(D) "
                  elseif instr(1,rsSubmitters("SubGroups"),"administrator") > 0 then
                    response.write "(A) "
                  elseif instr(1,rsSubmitters("SubGroups"),"content") > 0 then
                    response.write "(C) "
                  elseif instr(1,rsSubmitters("SubGroups"),"submitter") > 0 then
                    response.write "(S) "
                  else                    
                    response.write "(-) "
                  end if  
                  response.write rsSubmitters("LastName") & " " & rsSubmitters("FirstName") & "</OPTION>" & vbCrLf
                  rsSubmitters.MoveNext
                loop
                
                response.write "</SELECT>" & vbCrLf
                
              end if
              
              rsSubmitters.close
              set rsSubmitters = nothing  
              
              response.write "</TD>"
              response.write "</TR>"

            end if
            
            if Admin_Access = 2 or Admin_Access = 4 or Admin_Access >=8 then
              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>"            
            end if


            if Admin_Access = 4 or Admin_Access >=8 then
            
              for i = 0 to 1
      				  response.write "<TR>"
              	response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
                if i = 0 then
                  response.write Translate("Creation Date / Time",Login_Language,conn) & ":"
                else
                  response.write Translate("Last Update Date / Time",Login_Language,conn) & ":"
                end if  
                response.write "</TD>"
               	response.write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>"
      	        response.write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium ALIGN=RIGHT>"
                if i = 0 then
                  Response.write rs("PDate") & " PST"
                else
                  Response.write rs("UDate") & " PST"
                end if
                response.write "</TD>"
                response.write "</TR>"
              next
              
              response.write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>"            
                            
            end if  

            ' --------------------------------------------------------------------------------------                                
            ' Navigation Buttons
            ' --------------------------------------------------------------------------------------

            response.write "<TR>"
            response.write "<TD COLSPAN=3 CLASS=Medium>"
            response.write "<TABLE WIDTH=100% CELLPADDING=2 BGCOLOR=""#666666"">"
            response.write "<TR>"

            response.write "<TD ALIGN=CENTER WIDTH=""26%"" CLASS=Medium>"
            response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Main_Menu"" VALUE="" " & Translate("Main Menu",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
            response.write "</TD>"

            response.write "<TD ALIGN=CENTER WIDTH=""12%"" CLASS=Medium>"
            if (Clone = 0 and not rs("Locked")) or (Clone = 0 and Admin_Access >= 8) then
              response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Clone"" VALUE="" " & Translate("Clone",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
            else
              response.write "&nbsp;"
            end if
            response.write "</TD>"

            response.write "<TD ALIGN=CENTER WIDTH=""12%"" CLASS=Medium>"
            if (Clone = 0 and not rs("Locked")) or (Clone = 0 and Admin_Access >= 8) then
              response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Duplicate"" VALUE="" " & Translate("Duplicate",Login_Language,conn) &  " "" CLASS=Navlefthighlight1>"
            else
              response.write "&nbsp;"
            end if
            response.write "</TD>"

            response.write "<TD ALIGN=CENTER WIDTH=""25%"" CLASS=Medium>"
            response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Update"" VALUE="" " & Translate("Update",Login_Language,conn) & " "" CLASS=Navlefthighlight1 onclick=""startupload()"">"
            response.write "</TD>"
            response.write "<TD ALIGN=CENTER WIDTH=""25%"" CLASS=Medium>"
            response.write "<INPUT TYPE=""Submit"" NAME=""Nav_Delete"" VALUE="" " & Translate("Delete",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
            response.write "</TD>"

            response.write "</TR>"
            response.write "</TABLE>"
            response.write "</TD>"
            response.write "</TR>"
          response.write "</TABLE>"
        response.write "</TD>"
      response.write "</TR>"
    response.write "</TABLE>"
    response.write "</FORM>"
    response.write "<BR><BR>"
    
  end if

  rs.close
  set rs=nothing

%>