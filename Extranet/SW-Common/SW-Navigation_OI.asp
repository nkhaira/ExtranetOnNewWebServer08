<!-- Side Navigation Rows and Container -->

<TABLE WIDTH="100%" HEIGHT="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0">
  <TR VALIGN=TOP>  
    <TD WIDTH=4><IMG SRC="/Images/Spacer.gif" WIDTH=4></TD>
    <TD WIDTH=128 VALIGN=TOP BGCOLOR=White CLASS=Small>
      <IMG SRC="/Images/Spacer.gif" WIDTH=128 HEIGHT=4><BR>
    
      <% Call Nav_Border_Begin %>
    
      <TABLE WIDTH=128 BORDER=0 CELLPADDING=2 CELLSPACING=0>

        <!-- Home - Level 1 Menu Item -->

        <%=ThisCID = 0%>
        <TR>
          <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
        </TR>
            
        <TR>
           <TD CLASS=<% if CID=ThisCID and CIN=0 then response.write "NavLeftHighlight1 BGCOLOR=""" & Contrast & """" else response.write """NavLeft1"" BGCOLOR=""White"""%>>
             <%response.write "<IMG SRC=""/images/home.gif"" WIDTH=21 HEIGHT=10 BORDER=0 VSPACE=0 ALT=""Home"" ALIGN=RIGHT>"%>
             <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=0&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftHighlight1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
           </TD>
         </TR>

         <!-- Order Inquiry - Level 1 Menu Item -->
         <TR>
           <TD HEIGHT=1><IMG SRC="/images/1X1LINE.GIF"  WIDTH="100%" HEIGHT=1></TD>
         </TR>
            
         <TR>
           <TD CLASS=<% if CID=ThisCID and CIN = 0 then response.write "NavLeftHighlight1 BGCOLOR=""" & Contrast & """" else response.write """NavLeft1"" BGCOLOR=""White"""%>>
              <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>&Language=<%=Login_Language%>&NS=<%=Top_Navigation%>&CID=<%=ThisCID%>&SCID=<%=SCID%>&PCID=<%=PCID%>&CIN=<%=0%>&CINN=<%=0%>" CLASS=<% if CID = ThisCID and CIN = 0 then response.write "NavLeftHighlight1" else response.write """NavLeft1"" BGCOLOR=""White"""%> TITLE="<%=Button_Help(Button_Number)%>"><%=Button_Title(Button_Number)%></A>
           </TD>
          </TR>

          <%        
          response.write "  <TR>" & vbCrLf
          response.write "    <TD CLASS="
          if CIN = ThisCIN then
            response.write "NavLeftHighlight1 BGCOLOR=""" & Contrast & """>" & vbCrLf
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
            response.write "NavLeftHighlight1"
          else
            response.write """NavLeft1"" BGCOLOR=""White"""
          end if
          response.write " TITLE=""" & rsCategory("Title") & """>" & Translate(rsCategory("Title"),Login_Language,conn) & "</A>"
          response.write "</TD>" & vbCrLf
          response.write "</TR>" & vbCrLf
  
          response.write "</TABLE>" & vbCrLf
          response.write "</TD>" & vbCrLf                                         
          response.write "</TR>" & vbCrLf                               
%>
        
      </TABLE>
      
      <% Call Nav_Border_End %>
      
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
    