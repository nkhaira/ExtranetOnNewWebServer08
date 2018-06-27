<%
if isblank(Border_Toggle) then Border_Toggle = 0

sub Table_Begin()
    response.write "<TABLE BORDER=""" & Border_Toggle & """ CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder VSPACE=""0"" HSPACE=""0"">" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif""            BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLr
    response.write "  <TR>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD VALIGN=""top"" WIDTH=""100%"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "    </TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>"
    response.write "</TABLE>" & vbCrLf
end sub

'--------------------------------------------------------------------------------------

sub Nav_Border_Begin()
    response.write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" CLASS=NavBorder>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLr
    response.write "      <TR>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "        <TD VALIGN=""top"" CLASS=NavBorder>" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Nav_Border_End()
    response.write "        </TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" WIDTH=""8""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "      <TR>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "        <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT=""""></TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf
    response.write "    </TABLE>" & vbCrLf
end sub  
%>
