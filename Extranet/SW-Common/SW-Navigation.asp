
<!-- Top Navigation -->

<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER=0 BGCOLOR="#CCCCCC" CLASS=TABLEBACKGROUND ID="Table1">

  <% if Top_Navigation = True then %>

  <TR>
    <TD HEIGHT=17 WIDTH=128>
    
      <TABLE BORDER=0 HEIGHT=17 WIDTH=128 CELLPADDING=0 CELLSPACING=0 CLASS=TABLEBACKGROUND ID="Table2">
        <TR>
          <% ThisCID = Button(0) %>
          <TD CLASS=<% if CID = ThisCID then response.write "SmallBoldLine BGCOLOR=""#CCCCCC""" else response.write "NavTop"%>>
            &nbsp;&nbsp;<A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=CIN%>&CINN=<%=CINN%>" CLASS=<% if CID = ThisCID then response.write "NavTopHighlight" else response.write "NavTop"%> TITLE="<%=Button_Help(0)%>"><IMG SRC="/images/home.gif" WIDTH=21 HEIGHT=10 BORDER=0 VSPACE=0 ALT="<%=Button_Help(0)%>"><%=Button_Title(0)%></A>&nbsp;&nbsp;
          </TD>

        </TR>
      </TABLE>
    
    </TD>
    <TD HEIGHT=17>
    
      <TABLE BORDER=0 HEIGHT=17 CELLPADDING=0 CELLSPACING=0 CLASS=TABLEBACKGROUND ID="Table3">
        <TR>

          <% for Button_Number = 1 to 23 ' Exclude Contact Us, Messages, Profile, Site Stats and Admin
                if not isblank(Button_Title(Button_Number)) then
                  select case Button_Number
                    case 6  ' Forums
                    case 7  ' Order Inquiry
                      response.write "<TD WIDTH=1><IMG SRC=""/images/1x1Line.gif"" WIDTH=1 HEIGHT=16></TD>"
                      ThisCID = Button(Button_Number)
                      %>    
                      <TD NO WRAP VALIGN=MIDDLE CLASS=<% if CID = ThisCID then response.write "NavTopHighlight BGCOLOR=""#CCCCCC""" else response.write "NavTop"%>>
                        &nbsp;&nbsp;<A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=CIN%>&CINN=<%=CINN%>" CLASS=<% if CID = ThisCID then response.write "NavTopHighlight" else response.write "NavTop"%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>&nbsp;&nbsp;
                      </TD>
                      <%
                    case else    
                      response.write "<TD WIDTH=1><IMG SRC=""/images/1x1Line.gif"" WIDTH=1 HEIGHT=16></TD>"
                      ThisCID = Button(Button_Number)
                      %>    
                      <TD NO WRAP VALIGN=MIDDLE CLASS=<% if CID = ThisCID then response.write "NavTopHighlight BGCOLOR=""#CCCCCC""" else response.write "NavTop"%>>
                        &nbsp;&nbsp;<A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=CIN%>&CINN=<%=CINN%>" CLASS=<% if CID = ThisCID then response.write "NavTopHighlight" else response.write "NavTop"%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>&nbsp;&nbsp;
                      </TD>
                      <%
                  end select
                end if
             next

            if not isblank(Button_Title(27)) then
              response.write "<TD WIDTH=1><IMG SRC=""/images/1x1Line.gif"" WIDTH=1 HEIGHT=16></TD>"            
              %>    
              <TD NOWRAP VALIGN=MIDDLE CLASS="NavTop">
                &nbsp;&nbsp;<A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%if Top_Navigation = True then response.write "false" else response.write "true"%>&CID=<%=CID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=CIN%>&CINN=<%=CINN%>" CLASS="NavTop" TITLE="<%=Button_Help(27)%>"><%=Button_Title(27)%></A>&nbsp;&nbsp;
              </TD>
              <%
            end if
          %>          
          
        </TR>
      </TABLE>
    
    </TD>
    
    <!-- Date -->

    <% if Show_Date = True then %>   
    <TD HEIGHT=17 WIDTH=150 ALIGN="RIGHT" VALIGN=Top>
      <FONT FACE="Arial" SIZE=1 COLOR="White"><B><%response.write Translate(WeekDayName(WeekDay(Date),False,wbSunday),Login_Language,conn) & " - " & Day(Date) & " " & Translate(MonthName(Month(Date)),Login_Language,conn) & " " & Year(Date)%></B></FONT>
    </TD>
    <% end if %>
              
  </TR>
  
  <% end if ' end of "if Top_Navigation = True" %>
  
  <TR>
    <TD <%if Top_Navigation = True then response.write " COLSPAN=3"%> HEIGHT=6 CLASS=TopColorBar><IMG SRC="/images/1x1trans.gif" HEIGHT=6 BORDER=0 VSPACE=0></TD>
  </TR>
  <%if cstr(utility_id) <> "" then %>
  <tr><td <%if Top_Navigation = True then response.write " COLSPAN=3"%> HEIGHT=6 bgcolor="White"><br></td></tr>
  <TR>
    <TD <%if Top_Navigation = True then response.write " COLSPAN=3"%> HEIGHT=6 bgcolor="White" align=center CLASS=SmallBoldRed><DIV ID="Standby" NAME="Standby" style="visibility: visible"><%=trim(Translate("Standby... Retrieving Records",Login_Language,conn))%></DIV></TD>
  </TR>
  <%end if %>
</TABLE>


<!-- Side Navigation Rows and Container -->

<TABLE WIDTH="100%" HEIGHT="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0" ID="Table4">
  <TR VALIGN=TOP>  

    <!-- Side NAVIGATION ROWS -->

    <% if Side_Navigation = True then %>
    
    <TD WIDTH=4><IMG SRC="/Images/Spacer.gif" WIDTH=4></TD>
    <TD WIDTH=128 VALIGN=TOP BGCOLOR=White CLASS=Small>
    <IMG SRC="/Images/Spacer.gif" WIDTH=128 HEIGHT=4><BR>
    
    <%
    if Top_Navigation = false or (Top_Navigation = True and CID >=9002 and CID <=9003) then
      Call Nav_Border_Begin
    end if 
    %>
    
    <TABLE WIDTH=128 BORDER=0 CELLPADDING=2 CELLSPACING=0 ID="Table5">

      <%
        ' Home - 9000 and What's New - 9001
        
        for Button_Number = 0 to 1

          if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then%>                

            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
            
            <TR>
              <% ThisCID=Button(Button_Number) %>

              <!--TD WIDTH=8></TD-->
              <TD CLASS=<% if CID=ThisCID and CIN = 0 then response.write "NavLeftSelected1 BGCOLOR=""" & Contrast & """" else response.write """NavLeft1"" BGCOLOR=""White"""%>>
                <% if ThisCID = Button(0) then response.write "<IMG SRC=""/images/home.gif"" WIDTH=21 HEIGHT=10 BORDER=0 VSPACE=0 ALT=""Home"" ALIGN=RIGHT>"%><A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=0&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>
          <%
          end if
        next       
                
        ' Calendar - 9002

        for Button_Number = 2 to 2

          if Top_Navigation = False and Button_Title(Button_Number) <> "" then%>                

            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
            
            <TR>
              <% ThisCID=Button(Button_Number) %>

              <!--TD WIDTH=8></TD-->
              <TD CLASS=<% if CID=ThisCID and CIN = 0 then response.write "NavLeftSelected1 BGCOLOR=""" & Contrast & """" else response.write """NavLeft1"" BGCOLOR=""White"""%>>
                <!--IMG SRC="/images/calendar_button.gif" WIDTH=16 BORDER=0 VSPACE=0 ALT="Calendar" ALIGN=RIGHT-->
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>
          <%
          end if
        next
        
        ' Calendar Sub-Categories
        
        if CID < Button(0) or CID = Button(2) then
                
          if Access_Level = 0 then        ' Filter Navigation Buttons for User Only to Categories that have Active or Archive content
	
            SQL =       "SELECT Calendar_Category.ID, Calendar_Category.Code, Calendar_Category.Title, Calendar_Category.Separator " & vbCrLf
            SQL = SQL & "FROM Calendar_Category, Calendar " & vbCrLf
            SQL = SQL & "WHERE Calendar_Category.Code = Calendar.Code " & vbCrLf
            SQL = SQL & "AND ((Calendar.Site_ID=" & Site_ID & ") "
            SQL = SQL & "AND (Calendar.Status=1 Or Calendar.Status=2) " & vbCrLf            
            SQL = SQL & "AND (Calendar.SubGroups LIKE '%all%'"        
            for i = 0 to UserSubGroups_Max
              SQL = SQL & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
            next
            SQL = SQL & ") "        

            SQL = SQL & "AND ((Calendar.LDate<='" & Date & "' AND Calendar.XDAYS=0) OR (Calendar.LDate<='" & Date & "' AND Calendar.XDate>'" & Date & "')) "          

            SQL = SQL & "AND (Calendar.Language='eng' OR Calendar.Language='" & Login_Language & "') "                  


            SQL = SQL & "AND ((Calendar.Country) Like '%none%' " & vbCrLf
            SQL = SQL & "OR (Calendar.Country) Like '%" & Login_Country & "%')) " & vbCrLf
            SQL = SQL & "AND (Calendar.Content_Group=0 OR Calendar.Content_Group=1 OR Calendar.Content_Group=3) "
            SQL = SQL & "GROUP BY Calendar_Category.ID, Calendar_Category.Code, Calendar_Category.Title, Calendar_Category.Separator, Calendar_Category.Sort, Calendar_Category.Enabled, Calendar_Category.Calendar_View, Calendar_Category.Site_ID " & vbCrLf
            SQL = SQL & "HAVING ((Calendar_Category.Enabled=" & CInt(True) & ") AND (Calendar_Category.Site_ID=" & Site_ID & ") AND (Calendar_Category.Calendar_View=" & CInt(True) & ")) "  & vbCrLf
            SQL = SQL & "ORDER BY Calendar_Category.Sort, Calendar_Category.Title" & vbCrLf
          else
            SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=-1" & " AND Calendar_Category.Calendar_View=-1" & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
          end if
            
          Set rsCategory = Server.CreateObject("ADODB.Recordset")
          rsCategory.Open SQL, conn, 3, 3
                                              
          Do while not rsCategory.EOF
  
            ThisCIN = rsCategory("Code")
            
            response.write "  <TR>" & vbCrLf

            response.write "    <TD CLASS="
            if CIN = ThisCIN then
              response.write "NavLeftSelected1 BGCOLOR=""" & Contrast & """>" & vbCrLf
            else
              response.write """NavLeft1"" BGCOLOR=""White"">" & vbCrLf
            end if
  
            response.write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>" & vbCrLf
            response.write "  <TR>" & vbCrLf
            response.write "    <TD WIDTH=10>&nbsp;</TD>" & vbCrLf
            response.write "    <TD>"
              
            response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID
            response.write "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & rsCategory("Code") & "&CINN=" & rsCategory("ID") & """ CLASS="
      
            if CIN = ThisCIN then
              response.write "NavLeftSelected1"
            else
              response.write """NavLeft1"" BGCOLOR=""White"""
            end if
            response.write " TITLE=""" & rsCategory("Title") & """>" & Translate(rsCategory("Title"),Login_Language,conn) & "</A>"
            
            response.write "</TD>" & vbCrLf
            response.write "</TR>" & vbCrLf
  
            ' Separator
  
            if rsCategory("Separator") = True then
              response.write "<TR>"
              response.write "<TD HEIGHT=1><IMG SRC=""/images/1X1LINE.GIF"" WIDTH=""100%"" HEIGHT=1></TD>"
              response.write "</TR>"                
            end if
  
            response.write "</TABLE>" & vbCrLf
            response.write "</TD>" & vbCrLf                                         
            response.write "</TR>" & vbCrLf                               
        	  rsCategory.MoveNext
                       
          loop
        
          rsCategory.close
          set rsCategory=nothing
        
        end if

        ' Library - 9003
        
        for Button_Number = 3 to 3

          if Top_Navigation = False and Button_Title(Button_Number) <> "" then%>                

            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
            
            <TR>
              <% ThisCID=Button(Button_Number) %>

              <!--TD WIDTH=8></TD-->
              <TD CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1 BGCOLOR=""" & Contrast & """" else response.write """NavLeft1"" BGCOLOR=""White"""%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=0&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>
          <%
          end if
        next       
        
        ' Library Sub-Categories
        if Site_Id =3 then
              showPriceListLink=false
              session("PriceListLink")= false
              sqlRegion = "select  region from userdata where ntlogin like '%" & Login_Name & "%' and site_id=" & Site_ID & " and region=1"
			  Set rsRegion = Server.CreateObject("ADODB.Recordset")
			  rsRegion.Open sqlRegion, conn, 3, 3
					
			  if (rsRegion.EOF=false) then
				showPriceListLink = true
				session("PriceListLink")= true
			  end if  
			  set  rsRegion = nothing
        else
			  showPriceListLink=true
        end if
		                                
        if CID < Button(0) or CID = Button(3) then

        if Access_Level = 0 then        
          ' Filter Navigation Buttons for User Only to Categories that have Active or Archive content
          SQL =       "SELECT Calendar_Category.ID, Calendar_Category.Code, Calendar_Category.Title, Calendar_Category.Separator " & vbCrLf
          SQL = SQL & "FROM Calendar_Category, Calendar " & vbCrLf
          SQL = SQL & "WHERE Calendar_Category.Code = Calendar.Code " & vbCrLf
          SQL = SQL & "AND ((Calendar.Site_ID=" & Site_ID & ") "
          SQL = SQL & "AND (Calendar.Status=1 Or Calendar.Status=2) " & vbCrLf
          SQL = SQL & "AND (Calendar.SubGroups LIKE '%all%'"        
          for i = 0 to UserSubGroups_Max
            SQL = SQL & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
          next
          SQL = SQL & ") "

          SQL = SQL & "AND ((Calendar.LDate<='" & Date & "' AND Calendar.XDAYS=0) OR (Calendar.LDate<='" & Date & "' AND Calendar.XDate>'" & Date & "')) "          

          SQL = SQL & "AND (Calendar.Language='eng' OR Calendar.Language='" & Login_Language & "') "                  
 
          SQL = SQL & "AND ((Calendar.Country) Like '%none%' " & vbCrLf
          SQL = SQL & "OR (Calendar.Country) Like '%" & Login_Country & "%')) " & vbCrLf
                      SQL = SQL & "AND (Calendar.Content_Group=0 OR Calendar.Content_Group=1 OR Calendar.Content_Group=3) " & vbCrLf
          SQL = SQL & "GROUP BY Calendar_Category.ID, Calendar_Category.Code, Calendar_Category.Title, Calendar_Category.Separator, Calendar_Category.Sort, Calendar_Category.Enabled, Calendar_Category.Calendar_View, Calendar_Category.Site_ID " & vbCrLf
          SQL = SQL & "HAVING ((Calendar_Category.Enabled=" & CInt(True) & ") AND (Calendar_Category.Site_ID=" & Site_ID & ") AND (Calendar_Category.Calendar_View=" & CInt(False) & ")) "  & vbCrLf
          SQL = SQL & "ORDER BY Calendar_Category.Sort, Calendar_Category.Title" & vbCrLf         
        else  
          SQL =       "SELECT Calendar_Category.* " & vbCrLf
          SQL = SQL & "FROM Calendar_Category " & vbCrLf
          SQL = SQL & "WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " " & vbCrLf
          SQL = SQL & "AND Calendar_Category.Enabled=" & CInt(True) & " " & vbCrLf
          SQL = SQL & "AND Calendar_Category.Calendar_View=" & CInt(False) & " " & vbCrLf
          SQL = SQL & "ORDER BY Calendar_Category.Sort, Calendar_Category.Title" & vbCrLf
        end if  
        
        Set rsCategory = Server.CreateObject("ADODB.Recordset")
        rsCategory.Open SQL, conn, 3, 3

        Do while not rsCategory.EOF
            if not(cstr(trim(rsCategory("ID")))="117" and showPriceListLink = false) then
				ThisCIN = rsCategory("Code")
				response.write "  <TR>" & vbCrLf
				response.write "    <TD CLASS="
				if CIN = ThisCIN then
					response.write "NavLeftSelected1 BGCOLOR=""" & Contrast & """>" & vbCrLf
				else
					response.write """NavLeft1"" BGCOLOR=""White"">" & vbCrLf
				end if

				response.write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>" & vbCrLf
				response.write "  <TR>" & vbCrLf
				response.write "    <TD WIDTH=10>&nbsp;</TD>" & vbCrLf
				response.write "    <TD>"
		          
				response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID
				response.write "&SCID=" & SCID & "&PCID=0" & "&CIN=" & rsCategory("Code") & "&CINN=" & rsCategory("ID") & """ CLASS="
				if CIN = ThisCIN then
					response.write "NavLeftSelected1"
				else
					response.write """NavLeft1"" BGCOLOR=""White"""
				end if
		          
				response.write " TITLE=""" & rsCategory("Title") & """>" & Translate(rsCategory("Title"),Login_Language,conn) & "</A>"
				response.write "</TD>" & vbCrLf
				response.write "</TR>" & vbCrLf

				' Separator
				if rsCategory("Separator") = True then
					response.write "<TR>"
					response.write "<TD HEIGHT=1><IMG SRC=""/images/1X1LINE.GIF"" WIDTH=""100%"" HEIGHT=1></TD>"
					response.write "</TR>"                
				end if

				response.write "</TABLE>" & vbCrLf
				response.write "</TD>" & vbCrLf                                         
				response.write "</TR>" & vbCrLf 
			end if                              
      	    rsCategory.MoveNext
        loop
        
        rsCategory.close
        set rsCategory=nothing
        
        end if

        ' Search - 9004

        Button_Number = 4

        if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then%>                

          <!-- Level 1 Menu Item -->
          <TR>
            <!--TD></TD-->
            <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
          </TR>
          
          <TR>
            <% ThisCID=Button(Button_Number) %>
            <!--TD WIDTH=8></TD-->
            <TD CLASS=NavLeftSelected1>
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=0&CIN=<%=0%>&KeySearch=<%=KeySearch%>&BolSearch=<%=BolSearch%>&CINN=<%=0%>&SearchDB=<%=CINN%>" CLASS=NavLeftSelected1 TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
            </TD>
          </TR>
            
        <%
        end if 
        
        ' Brand Sites - 9005
        
        Button_Number = 5

        if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then
          
          Recriprocity_Status = False
                  
                SQL = "SELECT  dbo.UserData.Site_ID, dbo.UserData.NTLogin, dbo.UserData.NewFlag, dbo.UserData.ExpirationDate, dbo.Site.Enabled, dbo.Site.Site_Description " &_
                      "FROM    dbo.UserData LEFT OUTER JOIN " &_
                      "        dbo.Site ON dbo.UserData.Site_ID = dbo.Site.ID " &_
                      "WHERE   (dbo.UserData.Site_ID<>" & Site_ID & ") AND (dbo.Site.Enabled = -1) AND (dbo.UserData.NewFlag = 0) AND (dbo.UserData.NTLogin = '" & Login_Name & "') AND " &_
                      "        (dbo.UserData.ExpirationDate >= CONVERT(DATETIME, '" & Date() & "', 102)) " &_
                      "ORDER BY dbo.Site.Site_Description"

                Set rsSite_NewFlag = Server.CreateObject("ADODB.Recordset")
                rsSite_NewFlag.Open SQL, conn, 3, 3
      
                do while not rsSite_NewFlag.EOF
                  if rsSite_NewFlag("NewFlag") = CInt(False) and CDate(rsSite_NewFlag("ExpirationDate")) >= Date then
                    Recriprocity_Status = True
                    exit do
                  end if
                  rsSite_NewFlag.MoveNext
                loop
      
                rsSite_NewFlag.close
                set rsSite_NewFlag = nothing
            
          if CInt(Recriprocity_Status) = CInt(True) then
            %>

            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
          
            <TR>
              <% ThisCID=Button(Button_Number) %>
              <!--TD WIDTH=8></TD-->
              <TD CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1 BGCOLOR=""" & Contrast & """" else response.write """NavLeft1"" BGCOLOR=""White"""%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>             
            <%
          
          end if
          
        end if
        
       ' Navigation Buttons - 9006 Forums

        Button_Number = 6

        if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then
          ThisCID=Button(Button_Number)
			    if CID = ThisCID and CIN = 0 then
			  	  but_cl_col  = "CLASS=""NavLeftSelected1"" BGCOLOR=""" & Contrast & """"
  				  href_cl_col = "CLASS=""NavLeftSelected1"""
  			  else
  			  	but_cl_col  = "CLASS=""NavLeft1"" BGCOLOR=""White"""
  				  href_cl_col = "CLASS=""NavLeft1"" BGCOLOR=""White"""
  			  end if
  	  	  %>                

          <!-- Level 1 Menu Item -->
          <TR>
            <!--TD></TD-->
            <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
          </TR>
          
          <TR>
            <!--TD WIDTH=8></TD-->
            <TD <%=but_cl_col%>>
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" <%=href_cl_col%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
            </TD>
          </TR>
            
          <%
        end if
        
       ' Navigation Buttons - 9007 Order Inquiry

        Button_Number = 7

        if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then
          'Modified on 13-May-2009 for Nvision.
          'if CInt(Order_Inquiry)  = CInt(True) or _
          '   CInt(Order_Entry)    = CInt(True) or _
          '   CInt(Price_Delivery) = CInt(True) or _
          '   CInt(Shopping_Cart)  = CInt(True) then

          if CInt(Order_Inquiry)  = CInt(True) or _
             CInt(Order_Entry)    = CInt(True) or _
             CInt(Price_Delivery) = CInt(True) then
            
            ThisCID=Button(Button_Number)
  			    if CID = ThisCID and CIN = 0 then
  			  	  but_cl_col  = "CLASS=""NavLeftSelected1"" BGCOLOR=""" & Contrast & """"
    				  href_cl_col = "CLASS=""NavLeftSelected1"""
    			  else
    			  	but_cl_col  = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    				  href_cl_col = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    			  end if
    	  	  %>                

            <!-- Level 1 Menu Item -->
            <TR>
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
          
            <TR>
              <TD <%=but_cl_col%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" <%=href_cl_col%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>
          <%
          end if

        end if
        
       ' Navigation Buttons - 9008 Price & Delivery

        Button_Number = 8
        
        if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then
          ThisCID=Button(Button_Number)
          %>
          <!-- Level 1 Menu Item -->
          <TR>
            <!--TD></TD-->
            <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
          </TR>
          
          <TR>
            <!--TD WIDTH=8></TD-->
            <TD <%=but_cl_col%>>
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" <%=href_cl_col%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
            </TD>
          </TR>
          <%

        end if

       ' Navigation Buttons - 9009 WTB Distributor StoreFront Editor

        Button_Number = 9

        if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then
        
          if Instr(1,Login_Subgroups,"wtbdis") > 0 or Instr(1,Login_Subgroups,"wtbtsm") > 0 or Instr(1,Login_Subgroups,"wtbadm") or Access_Level = 9 then
          
            ThisCID=Button(Button_Number)
            %>
            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
            
            <TR>
              <!--TD WIDTH=8></TD-->
              <TD <%=but_cl_col%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" <%=href_cl_col%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>
            <%
          end if  
  
        end if

       ' Navigation Buttons through 9015 are reserved for Gatway Applications
        
       ' Navigation Buttons - 9015 through 9025

        for Button_Number = 10 to 25

          if Top_Navigation = False and not isblank(Button_Title(Button_Number)) then
            ThisCID=Button(Button_Number)
    			  if CID = ThisCID and CIN = 0 then
    			  	but_cl_col  = "CLASS=""NavLeftSelected1"" BGCOLOR=""" & Contrast & """"
    				  href_cl_col = "CLASS=""NavLeftSelected1"""
    			  else
    			  	but_cl_col  = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    				  href_cl_col = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    			  end if
		        %>                

            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
            
            <TR>
              <!--TD WIDTH=8></TD-->
              <TD <%=but_cl_col%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" <%=href_cl_col%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
              </TD>
            </TR>
              
            <%
          end if
        next
        
        ' Messages
        
        Buton_Number = 26
        
        if Top_Navigation = False and not isblank(Button_Title(26)) then%>                

          <!-- Level 1 Menu Item -->
          <TR>
            <!--TD></TD-->
            <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
          </TR>
          
          <TR>
            <%
            ThisCID=Button(26)
            if Message_Count > 0 and CID <> ThisCID then
              response.write "<TD CLASS=""NavLeft1"" BGCOLOR=""Red"">"
            else
              if CID = ThisCID then
                response.write "<TD CLASS=""NavLeftSelected1"" BGCOLOR=""" & Contrast & """>"
              else
                response.write "<TD CLASS=""NavLeft1"" BGCOLOR=""White"">"
              end if
            end if
            %>  
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftSelected1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
            </TD>
          </TR>
          <%
        end if

        ' Site Statistics

        if Top_Navigation = False and Button_Title(28) <> "" then

          if Access_Level >= 4 then

            ThisCID = Button(28)
    			  if CID = ThisCID and CIN = 0 then
    			  	but_cl_col  = "CLASS=""NavLeftSelected1"" BGCOLOR=""" & Contrast & """"
    				  href_cl_col = "CLASS=""NavLeftSelected1"""
    			  else
    			  	but_cl_col  = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    				  href_cl_col = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    			  end if
            %>
            
            <!-- Level 1 Menu Item -->
            
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
            
            <TR>
              <TD <%=but_cl_col%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=0&CINN=<%=0%>" <%=href_cl_col%> TITLE="<%=Button_Help(28)%>"><%=Button_Title(28)%></A>
              </TD>
            </TR>
            <%
                                                              
          end if  

        end if

        ' Administrators - 9029
        
        if Top_Navigation = False and Button_Title(29) <> "" then

          if Access_Level > 0 then
            ThisCID = Button(29)
    			  if CID = ThisCID and CIN = 0 then
    			  	but_cl_col  = "CLASS=""NavLeftSelected1"" BGCOLOR=""" & Contrast & """"
    				  href_cl_col = "CLASS=""NavLeftSelected1"""
    			  else
    			  	but_cl_col  = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    				  href_cl_col = "CLASS=""NavLeft1"" BGCOLOR=""White"""
    			  end if
            %>        

            <!-- Level 1 Menu Item -->
            <TR>
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
  
            <TR>
              <TD <%=but_cl_col%>>
                <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" TITLE="<%=Button_Help(29)%>" <%=href_cl_col%>><%=Button_Title(29)%></A>
              </TD>
            </TR>
            
            <%
          end if            

        end if
        
        ' Logoff
        
        if Top_Navigation = False then%>                

          <!-- Level 1 Menu Item -->
          <TR>
            <!--TD></TD-->
            <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
          </TR>
                      
          <TR>
            <!--TD WIDTH=8></TD-->
            <TD CLASS="NavLeft1" BGCOLOR="White">
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<% if Top_Navigation = True then response.write "false" else response.write "true"%>&CID=9999&SCID=0&PCID=0&CIN=0&CINN=0" CLASS=NavLeft1 TITLE="Logoff from Site"><%=Translate("Logoff",Login_Language,conn)%></A>
            </TD>
          </TR>

          <%
        end if
        
        ' Navigation

        if Top_Navigation = False and not isblank(Button_Title(27)) then%>                

          <!-- Level 1 Menu Item -->
          <TR>
            <!--TD></TD-->
            <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
          </TR>
          
          <TR>
            <% ThisCID=Button(27) %>
            <!--TD WIDTH=8></TD-->
            <TD CLASS="NavLeft1" BGCOLOR="White">
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<% if Top_Navigation = True then response.write "false" else response.write "true"%>&CID=<%=CID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" CLASS="NavLeft1" TITLE="<%=Button_Help(27)%>"><%=Button_Title(27)%></A>
            </TD>
          </TR>
            
          <%
        end if
                
        
        ' Language

        Language_ID = 0

        if Top_Navigation = False then 
          if Login_Language <> "elo" then
          %>
  
            <!-- Level 1 Menu Item -->
            <TR>
              <!--TD></TD-->
              <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
            </TR>
  
            <TR>
              <!--TD WIDTH=8></TD-->
              <TD NOWRAP CLASS=NavLeft1>            
                <%
  
                response.write "<FORM NAME=""LanguageForm"">"
                response.write Translate("Language",Login_Language,conn) & "<BR>"
                SQL = "SELECT * FROM Language WHERE Language.Enable=" & CInt(True) & " ORDER BY Language.Sort"
                Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                rsLanguage.Open SQL, conn, 3, 3
                response.write "<SELECT NAME=""Language"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&SortBy=" & SortBy & "&Show_Detail=" & Show_Detail & "&Language='+this.options[this.selectedIndex].value"">" & vbCrLf                      
                
                Do while not rsLanguage.EOF
                 	  response.write "<OPTION VALUE=""" & rsLanguage("Code") & """"
                    if LCase(rsLanguage("Code")) = LCase(Login_Language) then
                     response.write " SELECTED"
                     Language_ID = rsLanguage("ID")
                    end if 
                    if LCase(rsLanguage("Code")) <> "eng" then
                      response.write " CLASS=Region2"
                    else
                      response.write " CLASS=Small"
                    end if  
                    response.write ">" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>"              
              	  rsLanguage.MoveNext 
                loop
                
                rsLanguage.close
                set rsLanguage=nothing
                
                Select case Access_Level
                  case 2,4,8,9
                    if Session("ShowTranslation") = True then
                      response.write "<OPTION VALUE=""XOF"" CLASS=Translate>Translation View Off</OPTION>"
                    elseif Session("ShowTranslation") = False then
                      response.write "<OPTION VALUE=""XON"" CLASS=Translate>Translation View On</OPTION>"
                    end if
                end select
                                        
                response.write "</SELECT><BR>"
                response.write "<IMG SRC=""/images/1X1Trans.GIF""  WIDTH=""100%"" HEIGHT=6><BR>"
                response.write "<IMG SRC=""/images/1X1LINE.GIF""  WIDTH=""100%"" HEIGHT=1>"
                response.write "</FORM>"
                if Access_Level >= 8 then
                  response.write "<SPAN CLASS=NavLeft1>Users Online:&nbsp;&nbsp;[" & Application("UserOnLine") & "]</SPAN>"
                end if  
                %>
              </TD>
            </TR>
          
            <%
          end if
        end if
        %>        
        
      </TABLE>

      <%
      if Top_Navigation = false or (Top_Navigation = True and CID >=9002 and CID <=9003) then
        Call Nav_Border_End
      end if 
      %>
      
    </TD>
    
    <% end if %>
    
    <!-- END LEFT NAVIGATION ROWS -->  

    <!-- BEGIN CONTENT CONTAINER-->

    <TD VALIGN="top" CLASS=Normal WIDTH="100%">

    <% if isnumeric(Content_Width) then            
         response.write "<DIV ALIGN=CENTER>" & vbCrLf
         response.write "<TABLE WIDTH=""" & Content_Width & "%"">" & vbCrLf
         response.write "  <TR>" & vbCrLf
         response.write "    <TD CLASS=NORMAL VALIGN=""TOP"" WIDTH=""100%"">" & vbCrLf
       end if
    %> 
    
    <BR CLEAR=ALL>
    