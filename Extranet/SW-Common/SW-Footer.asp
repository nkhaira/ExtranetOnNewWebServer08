
      <!-- END CONTENT -->      
      
      <% if isnumeric(Content_Width) then
          response.write "    </TD>" & vbCrLf
          response.write "  </TR>" & vbCrLf
          response.write "</TABLE>" & vbCrLf
          response.write "</DIV>" & vbCrLf
         end if
      %>                 

      <!-- END CONTENT CONTAINER-->

    </TD>
  </TR>

  <!-- Begin Footer -->

  <TR>
           
    <% if Side_Navigation = True then %>
    <TD><IMG SRC="/Images/Spacer.gif" WIDTH=4 BGCOLOR=WHITE></TD>
    <TD BGCOLOR=WHITE>&nbsp;</TD>
    <% end if %>
    <TD ALIGN="CENTER" VALIGN="TOP" CLASS=Small BGCOLOR=WHITE>
      <%
      if not Footer_Disabled then
        response.write "&copy; 1995-" & DatePart("yyyy",Date) & " " & Translate("Fluke Corporation",Login_Language,conn) & " - " & Translate("All rights reserved",Login_Language,conn) & "."
      end if
      if Access_Level >= 8 then
        Page_Timer = Now() - Page_Timer_Begin
        response.write "<BR><SPAN CLASS=Small>Server Compilation Time: [" & FormatTime(Page_Timer) & "]</SPAN>" & vbCrLf & vbCrLf
      end if  
      %>
    </TD>
  </TR>

  <!-- End Footer -->

</TABLE>

</BODY>
</HTML>

<% response.write VbCrLf & vbCrLf & "<!-- Whitlock's SiteWide Content Server SWCS 4.0 " & Now & " PST -->"%>