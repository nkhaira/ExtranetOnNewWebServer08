<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/Pop-Up.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

Dim seq
seq = 0
if isblank(request.form("Password")) then
  seq = 0
elseif request.form("Password") = "Xtreme" then
  seq = 1
end if

Call Connect_SiteWide

Dim Site_ID
Site_ID          = 3
Login_Language = "eng"

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")
Screen_Title     = "Selling Fluke solutions has never been easier"
Bar_Title        = "Selling Fluke solutions has never been easier"
Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

Logo             = rsSite("Logo")  
Logo_Left        = rsSite("Logo_Left")

Dim DocURL, Item_Number
DocURL = "/find_it.asp?Document="

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->


<%
if seq = 0 then
%>

<DIV ALIGN=CENTER>
<FORM METHOD=POST NAME=Password ACTION="/promotions/grainger/default.asp">
  <% Call Nav_Border_Begin %>

  <TABLE BGCOLOR="FFCC00"" CELLSPACING=10>
    <TR>
      <TD CLASS=MediumBold>Password</TD>
      <TD CLASS=Medium><INPUT TYPE=Password Name=Password></TD>
    </TR>
    <TR>
      <TD Class=MediumBold>&nbsp;</TD>
      <TD Class=MediumBold><INPUT TYPE=SUBMIT NAME=SUBMIT VALUE="Click to Logon"></TD>
    </TR>
  </TABLE>
  <% Call Nav_Border_End %>
</DIV>

<%
elseif seq = 1 then
%>

  <P> 	 	 
  <SPAN CLASS=Heading3>For a wider library of application notes and cases studies visit: <A HREF="http://www.fluke.com">www.fluke.com</A></SPAN>
  <BR>
  <BR>
  <BR>
  </P>
  
  <DIV ALIGN=CENTER>
  <SPAN CLASS=Medium>Click on the <IMG SRC="/Images/Button_Down.gif" BORDER=0 width=16 VSPACE=0 ALIGN=ABSMIDDLE> icon to view the document.
  <P>
  <% Call Nav_Border_Begin %>
  	 
  <TABLE BGCOLOR=White CELLSPACING=10>
    <TR>
      <TD CLASS=MediumBold COLSPAN=4>Selling Solutions</TD>
    </TR>
 
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>
        Preventive Maintenance (PPM) and Predictive Maintenance (PdM) - Solution selling for productivity
      </TD>
      <TD CLASS=Medium>
        <% Item_Number = 2647262
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>
    </TR>
   	 	 
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>
        Power Quality - Solution selling for energy management
      </TD>
      <TD CLASS=Medium>
        <% Item_Number = 2647296
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>
    </TR>
   	 	 
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>
        Indoor Air Quality - Solution selling for indoor environmental quality
      </TD>
      <TD CLASS=Medium>
        <% Item_Number = 2647281
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>
    </TR>
   	 	 
    <TR>
      <TD CLASS=MediumBold COLSPAN=4>&nbsp;</TD>
    </TR>













    <TR>
      <TD CLASS=MediumBold COLSPAN=4>&nbsp;</TD>
    </TR>
    <TR>    
      <TD CLASS=MediumBold COLSPAN=4>Find your Fluke Territory Sales Manager</TD>
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>Fluke Territory Sales Manager Map</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2647122
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
   	 	 
    <TR>
      <TD CLASS=MediumBold COLSPAN=4>&nbsp;</TD>
    </TR>
    <TR>    
      <TD CLASS=MediumBold COLSPAN=4>Application Notes and Case Studies for Your Customers</TD>
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>ROI Calculator</TD>
      <TD CLASS=Medium>
        <% Item_Number = 9030104
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>The Basics of PdM</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2534401
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>Developing a Preventative Maintenance Inspection Program - 10 Steps</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2534191
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>PdM in Utilites: Case study of Thermal Imaging at a Coal Plant</TD>
      <TD CLASS=Medium>
        <% Item_Number = 251965
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>The Cost of Poor Power Quality</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2391563
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>Six Ways to Save More with Power Quality	</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2435490
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>Grow Your Business with Indoor Air Quality</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2457379
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>Indoor Air Quality: Diagnosing and Fixing an Ancient Problem</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2429205
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
   	 	 
    <TR>
      <TD CLASS=MediumBold COLSPAN=4>&nbsp;</TD>
    </TR>
    <TR>    
      <TD CLASS=MediumBOLD COLSPAN=4>New tools for you - Preventative Maintenance Products</TD>
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 5YE60 - Ti30 Thermal Imager</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2529814
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 1BE59 - 1587 Insulation Multimeters</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2532391
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 5YB32 - 570 Series Infrared Thermometers</TD>
      <TD CLASS=Medium>
        <% Item_Number = 9030102
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 5YE68 -  60 Series Mini Infrared Thermometers</TD>
      <TD CLASS=Medium>
        <% Item_Number = 9030103
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
   	 	 
    <TR>
      <TD CLASS=MediumBold COLSPAN=4>&nbsp;</TD>
    </TR>
    <TR>    
      <TD CLASS=MediumBold COLSPAN=4>New tools for you - Indoor Air Quality Products</TD>
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;</TD>
      <TD CLASS=Medium>No. 5BB60 - 983 Particle Counter</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2447045
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 5YE63 - 971 Temperature Humidity Meter</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2545349
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
   	 	 
    <TR>
      <TD CLASS=MediumBold COLSPAN=4>&nbsp;</TD>
    </TR> 
    <TR>    
      <TD CLASS=MediumBold COLSPAN=4>New Tools for you - Power Quality Products</TD>
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 4YE90 - 430 Series Three-Phase  Analyzer</TD>
      <TD CLASS=Medium>
        <% Item_Number = 2276752
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
    <TR>
      <TD CLASS=Medium>&nbsp;&nbsp;&nbsp;</TD>
      <TD CLASS=Medium>No. 1RH78 - 43B Single Phase Analyzer</TD>
      <TD CLASS=Medium>
        <% Item_Number = 1266142
           Call Write_Link
        %>
      </TD>
      <% Call Write_Thumb %>      
    </TR>
  
  </TABLE>
  <% Call Nav_Border_End %>
  </DIV>

<%
end if
%>
  
<P>
<BR>
<BR>

<!--#include virtual="/SW-Common/SW-Footer.asp"--> 
<%
Call Disconnect_SiteWide

sub Write_Link
  if Item_Number <> 0 then
    response.write "<A HREF=""javascript:void(0);"" TITLE=""View Document"
    response.write """ onclick=""openit('" & DocURL & Item_Number & "','Vertical');return false;"">"
    response.write "<IMG SRC=""/Images/Button_Down.gif"" BORDER=0 width=16 VSPACE=0 ALIGN=ABSMIDDLE>"
    response.write"</A>"
  end if
end sub

sub Write_Thumb
  if Item_Number <> 0 then
    SQLT = "SELECT Thumbnail FROM Calendar WHERE Item_Number='" & Item_Number & "' AND Thumbnail IS NOT NULL ORDER By Revision_Code Desc, BDate Desc"
    Set rsThumb = Server.CreateObject("ADODB.Recordset")
    rsThumb.Open SQLT, conn, 3, 3
    response.write "<TD>"
    if not rsThumb.EOF then
      response.write "<IMG SRC=""/" & Site_Code & "/" & rsThumb("Thumbnail") & """ width=40 border=1>"
    else
      response.write "&nbsp;"  
    end if
    response.write "</TD>"
    
    rsThumb.close
    set rsThumb = nothing  
  end if
end sub

%>
