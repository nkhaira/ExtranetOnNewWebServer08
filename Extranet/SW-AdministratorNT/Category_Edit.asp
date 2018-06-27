<%
'
' Author: D. Whitlock
'
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim BackURL
Dim HomeURL
Dim Site_ID
Dim Calendar_ID
Dim Category_ID

Dim Path_Site
Dim Path_Include
Dim Path_File
Dim Path_Thumbnail

Dim Show_Location
Dim Show_Link
Dim Show_Link_PopUp_Disabled
Dim Show_File
Dim Show_File_POD
Dim Show_Include
Dim Show_Thumbnail
Dim Show_Subscription
Dim Show_Calendar

BackURL       = request("HomeURL")
HomeURL       = request("BackURL")

Site_ID       = CInt(request("Site_ID"))
Calendar_ID   = request("ID")
Category_ID   = CInt(request("Category_ID"))

Path_Include  = "Content"
Path_File     = "Download"
Path_Thumbnail= "Photos"

Show_Location = false
Show_Link = false
Show_Link_PopUp_Disabled = false
Show_File = false
Show_File_POD = false
Show_Include = false
Show_Thumbnail = false
Show_Subscription = false
Show_Calendar = false

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

Screen_Title = "Sales Extranet - Category Administrator Screen"
Navigation = false
Content_Width = 95  ' Percent

%>

<!--#include virtual="/include/sw-header.asp"-->
<%

if IsNumeric(Calendar_ID) then

  SQL = "SELECT Calendar.*, Calendar_Category.Category AS Category_Name, Calendar.ID "
  SQL = SQL & "FROM Calendar LEFT JOIN Calendar_Category ON (Calendar.Category_ID = Calendar_Category.ID) AND (Calendar.Site_ID = Calendar_Category.Site_ID) "
  SQL = SQL & "WHERE (((Calendar.ID)=" & Clng(Calendar_ID) & "))"

  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sql, conn, 3, 3

  if not rs.EOF then    

    %>
    <FORM ACTION="Calendar_admin.asp" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA">
    <INPUT TYPE="Hidden" NAME="ID" VALUE="<%=Calendar_ID%>">
    <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
    <INPUT TYPE="Hidden" NAME="BackURL" VALUE="<%=BackURL%>">
    <INPUT TYPE="Hidden" NAME="HomeURL" VALUE="<%=HomeURL%>">
    <%  

    ' Determine Site Path based on Site_ID Number 

    SQL = "SELECT * FROM Site WHERE ID=" & CInt(Site_ID)
    Set rsSite = Server.CreateObject("ADODB.Recordset")
    rsSite.Open SQL, conn, 3, 3
    response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Site"" VALUE=""" & rsSite("Site Code") & """>"
    rsSite.close
    set rsSite=nothing            

    %> 
    <TABLE WIDTH="95%" BORDER=1 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
  	<TR>
  		<TD WIDTH="100%" BGCOLOR="#EEEEEE">
  			<TABLE WIDTH="100%" CELLPADDING=4 BORDER=0>
        
          <!-- Header -->
  				<TR>
          	<TD WIDTH="50%" BGCOLOR="Black">
              <FONT FACE="Arial" SIZE=2 COLOR="#FFCC00"><B>Description</B></FONT>
            </TD>
  	        <TD WIDTH="50%" BGCOLOR="Black" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2 COLOR="#FFCC00"><B>Data</B></FONT>
            </TD>
          </TR>
  				<TR>
          	<TD BGCOLOR="Silver">
              <FONT FACE="Arial" SIZE=2>Note: <B>Bold</B> Fields are Required.</FONT>
            </TD>
            <TD BGCOLOR="Silver">
            <A HREF="#HELP"><IMG SRC="images/help_button.gif" BORDER=0 ALIGN=RIGHT VALIGN=TOP></A>
            </TD>
          </TR>        
              
          <!-- Calendar Event ID -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Calendar Event ID Number:</FONT>
            </TD>
  	        <TD BGCOLOR="White">
              <FONT FACE="Arial" SIZE=2 COLOR="Gray"><%=rs("ID")%></FONT>          
            </TD>
          </TR>
  
          <!-- Category -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2><B>Category:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White"><FONT FACE="Arial" SIZE=2>          

              <%
              SQL = "SELECT * FROM Calendar_Category WHERE Site_ID=" & CInt(Site_ID) & " and ID=" & CInt(rs("Category_ID"))
              Set rsCategory = Server.CreateObject("ADODB.Recordset")
              rsCategory.Open SQL, conn, 3, 3
                                
              if not rsCategory.EOF then           
                
                  Show_Location = CInt(rsCategory("Location"))
                  Show_Link = CInt(rsCategory("Link"))
                  Show_Link_PopUp_Disabled = CInt(rsCategory("Link_PopUp_Disabled"))                    
                  Show_File = CInt(rsCategory("File"))
                  Show_File_POD = CInt(rsCategory("File_POD"))
                  Show_Include = CInt(rsCategory("Include"))                    
                  Show_Thumbnail = CInt(rsCategory("Thumbnail"))
                  Show_Subscription = CInt(rsCategory("Subscription"))
                  Show_Calendar = CInt(RsCategory("Calendar_View"))
                  
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" "                        
                  response.write "VALUE=""" & rsCategory("ID") & """ >" & rsCategory("Title")
                  Path_File = rsCategory("Category")                    
              else
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" VALUE="""">"
              end if                            
                           
              rsCategory.close
              set rsCategory=nothing

              %>
              </FONT>          
            </TD>
          </TR>

           <!-- Product -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP>
              <FONT FACE="Arial" SIZE=2><B>Product or Product Family:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Product" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("PRODUCT"))%>"></FONT>
              <BR>
              <SELECT NAME="Product_New">
              <OPTION VALUE="">Enter above or select from this list</OPTION>
                <%
                SQL = "SELECT Calendar.Product "
                SQL = SQL & "FROM Calendar "
                SQL = SQL & "GROUP BY Calendar.Site_ID, Calendar.Product "
                SQL = SQL & "HAVING Calendar.Site_ID=" & Site_ID & " AND Calendar.Product<>''"
                Set rsProduct = Server.CreateObject("ADODB.Recordset")
                rsProduct.Open SQL, conn, 3, 3
                                
                Do while not rsProduct.EOF            
               	  response.write "<OPTION VALUE=""" & RestoreQuote(rsProduct("Product")) & """>" & RestoreQuote(rsProduct("Product")) & "</OPTION>"
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
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2><B>Title:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("TITLE"))%>"></FONT>
            </TD>
          </TR>
  
          <!-- Short Description -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP>
              <FONT FACE="Arial" SIZE=2>Description:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <TEXTAREA NAME="Description" COLS=42 ROWS=6  MAXLENGTH="2048"><%=RestoreQuote(rs("DESCRIPTION"))%></TEXTAREA></FONT>
            </TD>
          </TR>
  
          <!-- Location -->
  
          <% if Show_Location = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2><B>Location <FONT SIZE=1>(City, State Country)</FONT>:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Location" SIZE="50" MAXLENGTH="255" VALUE="<%=rs("LOCATION")%>"></FONT>
            </TD>
          </TR>
          <% end if %>
          
           <!-- Link -->
  
          <% if Show_Link = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>URL to Web Site:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Link" SIZE="50" MAXLENGTH="255" VALUE="<%=RestoreQuote(rs("LINK"))%>"></FONT>
            </TD>
          </TR>
          <% end if %>

           <!-- Link PopUp Window Disable-->

          <% if Show_Link_PopUp_Disabled = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>URL to Web Site Pop-Up Window:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <% response.write "<TABLE WIDTH=""100%"" BORDER=0>"
              
               if rs("Subscription") = True then
                  response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"" VALUE=""" & True & """ CHECKED></TD><TD><FONT FACE=""Arial"" SIZE=2>Disable</FONT></TD></TR>"
                 else
                  response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"" VALUE=""" & False & """></TD><TD><FONT FACE=""Arial"" SIZE=2>Disable</FONT></TD></TR>"
               end if
              
               response.write "</TABLE>"
              %> 
              </FONT>                        
            </TD>
          </TR>
          <% end if %>       
          
          <!-- Include File -->
  
          <% if Show_Include = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Content File <FONT SIZE=1>(HTM or ASP)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Include"" VALUE=""" & Path_Include & """>"
              if isblank(rs("Include")) then
                response.write "<INPUT TYPE=""File"" NAME=""Include"" SIZE=""30"" MAXLENGTH=""50"">"
              else
                response.write "<INPUT TYPE=""Text"" NAME=""Include"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & rs("Include") & """>&nbsp;&nbsp"
                response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_Include"" VALUE=""yes"">&nbsp;&nbsp;Unattach File"
              end if
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>
          
          <!-- Upload File -->
  
          <% if Show_File = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Document File <FONT SIZE=1>(PDF, TXT, DOC, PPT, XLS, MDB)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File"" VALUE=""" & Path_File & """>"
              if isblank(rs("File")) then
                response.write "<INPUT TYPE=""File"" NAME=""File"" SIZE=""30"" MAXLENGTH=""50"">"
              else
                response.write "<INPUT TYPE=""Text"" NAME=""File"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & rs("File") & """>&nbsp;&nbsp"
                response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_File"" VALUE=""yes"">&nbsp;&nbsp;Unattach File"
              end if
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>

          <!-- Thumbnail File -->
  
          <% if Show_Thumbnail = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Image File <FONT SIZE=1>(GIF or JPG)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Thumbnail"" VALUE=""" & Path_Thumbnail & """>"
              if isblank(rs("Thumbnail")) then
                response.write "<INPUT TYPE=""File"" NAME=""Thumbnail"" SIZE=""30"" MAXLENGTH=""50"">"
              else
                response.write "<INPUT TYPE=""Text"" NAME=""Thumbnail"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & rs("Thumbnail") & """>&nbsp;&nbsp"
                response.write "<INPUT TYPE=""CHECKBOX"" NAME=""Delete_Thumbnail"" VALUE=""yes"">&nbsp;&nbsp;Unattach File"
              end if
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>
          
          <!-- Pre Announcement Days before BDate -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Pre-Announce:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="LDAYS" SIZE="30" MAXLENGTH="3" VALUE="<%=rs("LDAYS")%>">&nbsp;&nbsp;days before</FONT>
            </TD>
          </TR>

          <!-- Beginning Date -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>
              <A HREF="calendar_grid.asp" TARGET="calendar" onclick="openit('calendar_grid.asp?view=1&date=<%=Replace(rs("BDATE"), "/", "-")%>','Vertical');return false;"><IMG SRC="images/calendar_button.gif" BORDER=0 ALIGN=TOP></A>
              &nbsp;&nbsp;<B>Beginning Date</B> <FONT SIZE=1>(mm/dd/yyyy)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT VALIGN=TOP><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="BDate" SIZE="30" MAXLENGTH="35" VALUE="<%=rs("BDATE")%>">&nbsp;&nbsp;through</FONT>
            </TD>
          </TR>
      
          <!-- Ending Date -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>
              <A HREF="calendar_grid.asp" TARGET="calendar" onclick="openit('calendar_grid.asp?view=1&date=<%=Replace(rs("EDATE"), "/", "-")%>','Vertical');return false;"><IMG SRC="images/calendar_button.gif" BORDER=0 ALIGN=TOP></A>
              &nbsp;&nbsp;<B>Ending Date</B> <FONT SIZE=1>(mm/dd/yyyy)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="EDate" SIZE="30" MAXLENGTH="35" VALUE="<%=rs("EDATE")%>">&nbsp;&nbsp;then</FONT>
            </TD>
          </TR>        
  
          <!-- Post Announcement Days After EDate -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Post-Announce:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="XDays" SIZE="30" MAXLENGTH="3" VALUE="<%=rs("XDays")%>">&nbsp;&nbspdays after ending date</FONT>
            </TD>
          </TR>       

          <!-- Post via Subscription Service -->
  
          <% if Show_Subscription = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Send Notice via Subscripiton Service:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <% response.write "<TABLE WIDTH=""100%"" BORDER=0>"
              
               if rs("Subscription") = True then
                  response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Subscription"" VALUE=""" & True & """ CHECKED></TD><TD><FONT FACE=""Arial"" SIZE=2>Subscription Service</FONT></TD></TR>"
                 else
                  response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Subscription"" VALUE=""" & False & """></TD><TD><FONT FACE=""Arial"" SIZE=2>Subscription Service</FONT></TD></TR>"
               end if
              
               response.write "</TABLE>"
              %> 
              </FONT>                        
            </TD>
          </TR>
          <% end if %>       

          <!-- NT Sub-Groups -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP>
              <FONT FACE="Arial" SIZE=2><B>Select Groups allowed<BR>to view this information:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White"><FONT FACE="Arial" SIZE=2>               
        
            <%

              SQL = "SELECT SubGroups.*, SubGroups.Site "
              SQL = SQL & "FROM SubGroups "
              SQL = SQL & "WHERE (((SubGroups.Site)=" & CInt(Site_ID) & ")) "
              SQL = SQL & "ORDER BY SubGroups.Site, SubGroups.Order_Num"

              Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
              rsSubGroups.Open SQL, conn, 3, 3
              
              if not rsSubGroups.EOF then
              
                response.write "<TABLE WIDTH=""100%"">"

                Do while not rsSubGroups.EOF
                              
                  if instr(1,lcase(rs("SubGroups")),lcase(rsSubGroups("Code"))) > 0 then
                    response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED></TD><TD><FONT FACE=""Arial"" SIZE=2>" & rsSubGroups("Description") & "</FONT></TD></TR>"
                  elseif rsSubGroups("Enabled") <> True then
                    'response.write "<TR><TD WIDTH=20>&nbsp;</TD><TD><FONT FACE=""Arial"" SIZE=2 COLOR=""Gray"">" & rsSubGroups("Description") & "</FONT></TD></TR>"
                  elseif rsSubGroups("Enabled") = True then
                    response.write "<TR><TD WIDTH=20><FONT FACE=""Arial"" SIZE=2><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """></TD><TD><FONT FACE=""Arial"" SIZE=2>" & rsSubGroups("Description") & "</FONT></TD></TR>"
                  end if
              
              	  rsSubGroups.MoveNext 

                loop
              
                response.write "</TABLE>"
              
              end if  
                 
              rsSubGroups.close
              set rsSubGroups=nothing

              %>
        
              </SELECT>
              </FONT>          
            </TD>
          </TR>
  
          <!-- Navigation Buttons -->
  
          <TR>
          <TD COLSPAN=2>
            <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666">
              <TR>
                <TD ALIGN=CENTER WIDTH="25%">
                  <INPUT TYPE="Submit" NAME="Action" VALUE=" Back ">
                </TD>
                <TD ALIGN=CENTER WIDTH="25%">
                  &nbsp;
                </TD>             
                <TD ALIGN=CENTER WIDTH="25%">
                  <INPUT TYPE="Submit" NAME="Action" VALUE=" Update ">
                </TD>
                <TD ALIGN=CENTER WIDTH="25%">
                  <INPUT TYPE="Submit" NAME="Action" VALUE=" Delete ">
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
  <BR><BR>      
    
  <%
  
  end if

       
  rs.close
  set rs=nothing

  Call Disconnect_SiteWide
    
end if
%>

<!-- Add Record -->

<%
if Calendar_ID= "add" then
  %>

  <FORM ACTION="Calendar_admin.asp" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA">
  <INPUT TYPE="Hidden" NAME="ID" VALUE="<%=Calendar_ID%>">
  <INPUT TYPE="Hidden" NAME="Site_ID" VALUE="<%=Site_ID%>">
  <INPUT TYPE="Hidden" NAME="BackURL" VALUE="Calendar_Edit.asp">


  <TABLE WIDTH="95%" BORDER=1 BORDERCOLOR="GRAY" CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
  	<TR>
  		<TD WIDTH="100%" BGCOLOR="#EEEEEE">
  			<TABLE WIDTH="100%" CELLPADDING=4>
        
          <!-- Header -->
  				<TR>
          	<TD WIDTH="50%" BGCOLOR="Black">
              <FONT FACE="Arial" SIZE=2 COLOR="#FFCC00"><B>Description</B></FONT>
            </TD>
  	        <TD WIDTH="50%" BGCOLOR="Black" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2 COLOR="#FFCC00"><B>Data</B></FONT>
            </TD>
          </TR>
  				<TR>
          	<TD BGCOLOR="Silver">
              <FONT FACE="Arial" SIZE=2>Note: <B>Bold</B> Fields are Required.</FONT>
            </TD>
            <TD BGCOLOR="Silver">
            <A HREF="#HELP"><IMG SRC="images/help_button.gif" BORDER=0 ALIGN=RIGHT VALIGN=TOP></A>
            </TD>
          </TR>        
  
          <!-- Calendar Event ID -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Calendar Event ID Number:</FONT>
            </TD>
  	        <TD BGCOLOR="White">
              <FONT FACE="Arial" SIZE=2 COLOR="Gray"><%=UCase(Calendar_ID)%></FONT>          
            </TD>
          </TR>
  
          <!-- Category -->

          <% if Category_ID = false then %>

  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2><B>Category:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White"><FONT FACE="Arial" SIZE=2>
              <SELECT LANGUAGE="JavaScript" ONCHANGE="window.location.href='Calendar_edit.asp?ID=<%=Calendar_ID%>&Site_ID=<%=Site_ID%>&Category_ID='+this.options[this.selectedIndex].value" NAME="Category_ID">    
            <%
              SQL = "SELECT * FROM Calendar_Category WHERE Site_ID=" & CInt(Site_ID) & " ORDER BY Category"
              Set rsCategory = Server.CreateObject("ADODB.Recordset")
              rsCategory.Open SQL, conn, 3, 3

              response.write "<OPTION VALUE="""">Select from list</OPTION>"
                                
              Do while not rsCategory.EOF            
             	  response.write "<OPTION VALUE=""" & rsCategory("ID") & """>" & rsCategory("Title") & "</OPTION>"              
            	  rsCategory.MoveNext 
              loop
                 
              rsCategory.close
              set rsCategory=nothing
            %>
        
              </SELECT>
              </FONT>          
            </TD>
          </TR>


          <% else %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2><B>Category:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White"><FONT FACE="Arial" SIZE=2>          
            <%                        
              SQL = "SELECT * FROM Calendar_Category WHERE Site_ID=" & CInt(Site_ID) & " and ID=" & CInt(Category_ID)
              Set rsCategory = Server.CreateObject("ADODB.Recordset")
              rsCategory.Open SQL, conn, 3, 3
                                
              if not rsCategory.EOF then           
                
                  Show_Location = CInt(rsCategory("Location"))
                  Show_Link = CInt(rsCategory("Link"))
                  Show_Link_PopUp_Disabled = CInt(rsCategory("Link_PopUp_Disabled"))                    
                  Show_File = CInt(rsCategory("File"))
                  Show_Include = CInt(rsCategory("Include"))
                  Show_Thumbnail = CInt(rsCategory("Thumbnail"))
                  Show_Subscription = CInt(rsCategory("Subscription"))
                  Show_Calendar = CInt(RsCategory("Calendar_View"))
                                      
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" "                        
                  response.write "VALUE=""" & rsCategory("ID") & """ >" & rsCategory("Title")                    
              else
                  response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" VALUE="""">"
              end if                            
                           
              rsCategory.close
              set rsCategory=nothing
            %>
              </FONT>          
            </TD>
          </TR>

          <!-- Title -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2><B>Title:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Title" SIZE="50" MAXLENGTH="255" VALUE=""></FONT>
            </TD>
          </TR>

          <!-- Short Description -->
          
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP>
              <FONT FACE="Arial" SIZE=2>Description:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <TEXTAREA NAME="Description" COLS=42 ROWS=6 MAXLENGTH="2048"></TEXTAREA></FONT>
            </TD>
          </TR>    

           <!-- Product -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP>
              <FONT FACE="Arial" SIZE=2><B>Product or Product Family:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Product" SIZE="50" MAXLENGTH="255" VALUE=""></FONT>
              <BR>
              <SELECT NAME="Product_New">
              <OPTION VALUE="">Enter above or select from this list</OPTION>
              <%
                  SQL = "SELECT Calendar.Product "
                  SQL = SQL & "FROM Calendar "
                  SQL = SQL & "GROUP BY Calendar.Site_ID, Calendar.Product "
                  SQL = SQL & "HAVING Calendar.Site_ID=" & Site_ID & " AND Calendar.Product<>''"
                  Set rsProduct = Server.CreateObject("ADODB.Recordset")
                  rsProduct.Open SQL, conn, 3, 3
                                
                  Do while not rsProduct.EOF            
                 	  response.write "<OPTION VALUE=""" & rsProduct("Product") & """>" & rsProduct("Product") & "</OPTION>"                  
                	  rsProduct.MoveNext 
                  loop
                     
                  rsProduct.close
                  set rsProduct=nothing
              %>
              </SELECT>
            </TD>
          </TR>
          
          <!-- Location -->
  
          <% if Show_Location = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Location <FONT SIZE=1>(City, State Country)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Location" SIZE="50" MAXLENGTH="255" VALUE=""></FONT>
            </TD>
          </TR>
          <% end if %>
  
           <!-- Link -->
           
          <% if Show_Link = True then %>    
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>URL to Web Site:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="Link" SIZE="50" MAXLENGTH="255" VALUE=""></FONT>
            </TD>
          </TR>
          <% end if %>
          
  
           <!-- Post via Subscription Service -->

          <% if Show_Link_PopUp_Disabled = True then %>    
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>URL to Web Site Pop-Up Window:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <%
              response.write "<TABLE WIDTH=""100%"" BORDER=0>"
              response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Link_PopUp_Disabled"" VALUE=""" & false & """></TD><TD><FONT FACE=""Arial"" SIZE=2>Disable</FONT></TD></TR>"                
              response.write "</TABLE>"
              %> 
              </FONT>                        
            </TD>
          </TR>
          <% end if %>       

           <!-- Include File -->
  
          <% if Show_Include = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Content File <FONT SIZE=1>(HTM or ASP)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Include"" VALUE=""" & Path_Include & """>"
              response.write "<INPUT TYPE=""File"" NAME=""Include"" SIZE=""30"" MAXLENGTH=""50"">"
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>

           <!-- Upload File -->
  
          <% if Show_File = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Document File <FONT SIZE=1>(PDF, TXT, DOC, PPT, XLS, MDB)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File"" VALUE=""" & Path_File & """>"
              response.write "<INPUT TYPE=""File"" NAME=""File"" SIZE=""30"" MAXLENGTH=""50"">"
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>

           <!-- Upload File -->
  
          <% if Show_File_POD = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>POD Document File <FONT SIZE=1>(PDF, TXT, DOC, PPT, XLS, MDB)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_File_POD"" VALUE=""" & Path_File & """>"
              response.write "<INPUT TYPE=""File"" NAME=""File_POD"" SIZE=""30"" MAXLENGTH=""50"">"
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>

           <!-- Thumbnail File -->
  
          <% if Show_Thumbnail = True then %>
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Image File <FONT SIZE=1>(GIF or JPG)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT>
              <FONT FACE="Arial" SIZE=2>
              <%
              response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Thumbnail"" VALUE=""" & Path_Thumbnail & """>"
              response.write "<INPUT TYPE=""File"" NAME=""Thumbnail"" SIZE=""30"" MAXLENGTH=""50"">"
              %>
              </FONT>
            </TD>
          </TR>
          <% end if %>
                              
           <!-- Pre Announcement Days before BDate -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Pre-Announce:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="LDAYS" SIZE="30" MAXLENGTH="3" VALUE="0">&nbsp;&nbsp;days before</FONT>
            </TD>
          </TR>

           <!-- Beginning Date -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>
              <A HREF="calendar_grid.asp" TARGET="calendar" onclick="openit('calendar_grid.asp?view=1&date=<%=Replace(Date(), "/", "-")%>','Vertical');return false;"><IMG SRC="images/calendar_button.gif" BORDER=0 ALIGN=TOP></A>
              &nbsp;&nbsp;<B>Beginning Date</B> <FONT SIZE=1>(mm/dd/yyyy)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT VALIGN=TOP><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="BDate" SIZE="30" MAXLENGTH="35" VALUE="<%=Replace(Date(),"-","/")%>">&nbsp;&nbsp;through</FONT>
            </TD>
          </TR>
          
           <!-- Ending Date -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>
              <A HREF="calendar_grid.asp" TARGET="calendar" onclick="openit('calendar_grid.asp?view=1&date=<%=Replace(Date(), "/", "-")%>','Vertical');return false;"><IMG SRC="images/calendar_button.gif" BORDER=0 ALIGN=TOP></A>
              &nbsp;&nbsp;<B>Ending Date</B> <FONT SIZE=1>(mm/dd/yyyy)</FONT>:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="EDate" SIZE="30" MAXLENGTH="35" VALUE="<%=Replace(Date(),"-","/")%>">&nbsp;&nbsp;then</FONT>
            </TD>
          </TR>        
              
           <!-- Post Announcement Days After EDate -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Post-Announce:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <INPUT TYPE="Text" NAME="XDays" SIZE="30" MAXLENGTH="3" VALUE="0">&nbsp;&nbspdays after ending date</FONT>
            </TD>
          </TR>       
  
           <!-- Post via Subscription Service -->

          <% if Show_Subscription = True then %>    
  				<TR>
          	<TD BGCOLOR="#EEEEEE">
              <FONT FACE="Arial" SIZE=2>Send Notice via Subscripiton Service:</FONT>
            </TD>
  	        <TD BGCOLOR="White" ALIGN=LEFT><FONT FACE="Arial" SIZE=2>
              <%
              response.write "<TABLE WIDTH=""100%"" BORDER=0>"
              response.write "<TR><TD WIDTH=20><INPUT TYPE=""Checkbox"" NAME=""Subscription"" VALUE=""" & True & """ CHECKED></TD><TD><FONT FACE=""Arial"" SIZE=2>Subscription Service</FONT></TD></TR>"                
              response.write "</TABLE>"
              %> 
              </FONT>                        
            </TD>
          </TR>
          <% end if %>       

          <!-- NT Sub-Groups -->
  
  				<TR>
          	<TD BGCOLOR="#EEEEEE" VALIGN=TOP>
              <FONT FACE="Arial" SIZE=2><B>Select Groups allowed<BR>to view this information:</B></FONT>
            </TD>
  	        <TD BGCOLOR="White"><FONT FACE="Arial" SIZE=2>               
        
            <%
              SQL = "SELECT SubGroups.*, SubGroups.Site "
              SQL = SQL & "FROM SubGroups "
              SQL = SQL & "WHERE (((SubGroups.Site)=" & CInt(Site_ID) & ")) "
              SQL = SQL & "ORDER BY SubGroups.Site, SubGroups.Order_Num"

              Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
              rsSubGroups.Open SQL, conn, 3, 3
              
              if not rsSubGroups.EOF then
              
                response.write "<TABLE WIDTH=""100%"">"
                                
                Do while not rsSubGroups.EOF
                              
                  if rsSubGroups("Enabled") = True and rsSubGroups("Default") = True then
                    response.write "<TR><TD><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """ CHECKED></TD><TD><FONT FACE=""Arial"" SIZE=2>" & rsSubGroups("Description") & "</FONT></TD></TR>"
                  elseif rsSubGroups("Enabled") <> True then
                    'response.write "<TR><TD>&nbsp;</TD><TD><FONT FACE=""Arial"" SIZE=2 COLOR=""Gray"">" & rsSubGroups("Description") & "</FONT></TD></TR>"
                  elseif rsSubGroups("Enabled") = True and rsSubGroups("Default") <> True then
                    response.write "<TR><TD><FONT FACE=""Arial"" SIZE=2><INPUT TYPE=""Checkbox"" NAME=""SubGroups"" VALUE=""" & rsSubGroups("Code") & """></TD><TD><FONT FACE=""Arial"" SIZE=2>" & rsSubGroups("Description") & "</FONT></TD></TR>"
                  end if
              
              	  rsSubGroups.MoveNext 

                loop
              
                response.write "</TABLE>"
              
              end if  
                 
              rsSubGroups.close
              set rsSubGroups=nothing

              %>
        
              </SELECT>
              </FONT>          
            </TD>
          </TR>

        <% end if %>    
          
          <!-- Navigation Buttons -->
  
          <TR>
          <TD COLSPAN=2>
            <TABLE WIDTH=100% CELLPADDING=2 BGCOLOR="#666666">
              <TR>
                <TD ALIGN=CENTER WIDTH="25%">
                  <INPUT TYPE="Submit" NAME="Action" VALUE=" Back ">
                </TD>
                <TD ALIGN=CENTER WIDTH="25%">
                  &nbsp;
                </TD>             
                <TD ALIGN=CENTER WIDTH="25%">
                  <% if Category_ID <> false then %>
                  <INPUT TYPE="Submit" NAME="Action" VALUE=" Update ">
                  <% end if %>
                </TD>
                <TD ALIGN=CENTER WIDTH="25%">
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
     
  <%
  
  Call Disconnect_SiteWide
  
end if
%>

<!--#include virtual="/include/sw-footer.asp"-->
<!--#include virtual="/include/sw-header.asp"-->

