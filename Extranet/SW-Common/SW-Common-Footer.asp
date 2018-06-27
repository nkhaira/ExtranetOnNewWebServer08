
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
    <TD>&nbsp;</TD>
    <% end if %>
    
    <TD><IMG SRC="/Images/Spacer.gif" WIDTH=4></TD>
    <TD ALIGN="CENTER" VALIGN="TOP" CLASS=Small>
      <%
      response.write "&copy; 1995-" & DatePart("yyyy",Date) & " Fluke Corporation - " & "All rights reserved" & "."
      %>
    </TD>
  </TR>

  <!-- End Footer -->

</TABLE>

</BODY>
</HTML>
