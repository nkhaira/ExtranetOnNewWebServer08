      
      <% if Content_Width <> "" then
          response.write "</TD></TR></TABLE></CENTER>"
         end if
      %>                 

      <!-- End Contents -->

      </FONT>
    </TD>
  </TR>

  <!-- Begin Footer -->

  <TR>
    
    <% if Navigation = True then %>
    <TD>
      &nbsp;
    </TD>
    <% end if %>
    
    <TD ALIGN="CENTER" VALIGN="TOP">
      
      <FONT SIZE="1" FACE="Verdana,Arial,Helvetica">
      <%
      response.write "&copy; 1995-" & DatePart("yyyy",Date) & " Fluke Corporation -  All rights reserved."
      %>
      </FONT>
    </TD>
  </TR>

  <!-- End Footer -->

</TABLE>

</BODY>
</HTML>
