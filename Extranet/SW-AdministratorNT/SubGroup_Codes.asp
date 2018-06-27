<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%
response.expires=0

Dim Site_ID
Dim Asset_ID
Dim Asset_Country
Dim Asset_SubGroups

if not isblank(request("Site_ID")) then
  Site_ID = CInt(request("Site_ID"))
else
  response.redirect "/register/default.asp"
end if  

if not isblank(request("Asset_ID")) then
  Asset_ID = CLng(request("Asset_ID"))
else
  Asset_ID = 0
end if  

Dim RegionColor(4)
RegionColor(0) = "#0000CC"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"
RegionColor(4) = "#FFCC99"

Call Connect_SiteWide

' the whole page is different based on whether or not there is an asset id...
if Asset_ID then
    SQL = "SELECT Title, SubGroups, Country, ID from Calendar WHERE ID=" & Asset_ID
    Set rsAsset = Server.CreateObject("ADODB.Recordset")
    rsAsset.Open SQL, conn, 3, 3

    if not rsAsset.EOF then
      Asset_SubGroups = LCase(rsAsset("SubGroups"))
      Asset_Country   = LCase(rsAsset("Country"))
    else
	  ' this is a problem - they gave us an asset id but we can't find it
      Asset_ID = 0
    end if
    
    %>
    <!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
    <%
    Dim Top_Navigation        ' True / False
    Dim Side_Navigation       ' True / False
    Dim Screen_Title          ' Window Title
    Dim Bar_Title             ' Black Bar Title
    Dim Content_width	        ' Percent
    
    Screen_Title    = Site_Description & " - " & Translate("Asset Distribution Information",Login_Language,conn)
    Bar_Title       = Site_Description & "<BR><SPAN CLASS=MediumBoldGold>" & Translate("Asset Distribution Information",Login_Language,conn) & "</SPAN>"
    Top_Navigation  = False
    Side_Navigation = True
    Content_Width   = 95
    %>
    <!--#include virtual="/SW-Common/SW-Header.asp"-->
    <!--#include virtual="/SW-Common/SW-Common-No-Navigation.asp"-->
    <%
	
	if Asset_ID = 0 then
	  with response
	    .write "<P><BR><SPAN CLASS=Medium><LI>"
		.write Translate("This item cannot be found in the database.",Login_Language,conn)
		.write "</SPAN></LI>" & vbCrLf
	  end with
      %>
      <!--#include virtual="/SW-Common/SW-Footer.asp"-->
      <%
	  response.end
	  
	end if
	
    response.write "<P>"
    response.write "<SPAN CLASS=MediumBold>" & Translate("Title",Login_Language,conn) & ": "
    response.write rsAsset("Title") & "</SPAN><P>"
    
    rsAsset.close
    set rsAsset = nothing
	
    ' if Asset_SubGroups contains "all" then let's short circuit the process
	if Instr(1,Asset_SubGroups,"all") > 0 then
	  response.write "<P><BR><SPAN CLASS=Medium><LI>" & Translate("This item can viewed by all users.",Login_Language,conn) & "</SPAN></LI>" & vbCrLf
	  
	else
	  
      response.write "<SPAN CLASS=MediumBold>" & Translate("Distribution Groups",Login_Language,conn) & "</SPAN><BR>" & vbCrLf
      response.write "<SPAN CLASS=Small>" & Translate("This document is available for view or download to following Distribution Groups.",Login_Language,conn) & "</SPAN><P>" & vbCrLf
	  ' display the groups - let's just get the for this asset
	  Asset_SubGroups = Replace(Asset_SubGroups," ","")
	  aGroups = split(Asset_SubGroups,",")
	  Asset_SubGroups = join(aGroups,"','")
	  
	  SQL = "SELECT SubGroups.* FROM SubGroups" & vbcrlf &_
	    "WHERE SubGroups.Site_ID=" & Site_ID & vbcrlf &_
	    "AND SubGroups.Enabled=" & CInt(True) & vbcrlf &_
	    "AND SubGroups.Code in ('" & Asset_SubGroups & "')" & vbcrlf &_
	    "ORDER BY SubGroups.Order_Num"
	  Write_Subgroups(SQL)
	end if
	
    if Asset_Country = "none" then
      with response
	    .write "<P><BR><SPAN CLASS=Medium><LI>"
		.write Translate("There are no country restrictions for this item.",Login_Language,conn)
		.write "</SPAN></LI>" & vbCrLf
	  end with
	else  
	  Write_Countrys
    end if
	
    %>
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
    <%
else
    response.write "<HTML>" & vbCrLf
    response.write "<HEAD>" & vbCrLf
    response.write "<TITLE>Group Codes</TITLE>" & vbCrLf
    response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
    response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso8859-1"">" & vbCrLf
    response.write "</HEAD>" & vbCrLf
    response.write "<BODY BGCOLOR=""White"">" & vbCrLf & vbCrLf
	
    ' Display SubGroup Code Table
    
    SQL = "SELECT SubGroups.* FROM SubGroups" & vbcrlf &_
	  "WHERE SubGroups.Site_ID=" & Site_ID & vbcrlf &_
	  "AND SubGroups.Enabled=" & CInt(True) &  vbcrlf &_
	  "ORDER BY SubGroups.Order_Num"
	
	Write_SubGroups(SQL)
    response.write "</BODY>" & vbcrlf & "</HTML>"
end if

Call Disconnect_SiteWide

' --------------------------- end of main -----------------------------------

sub Write_Subgroups(ssql)

    Set rsSubGroups = Server.CreateObject("ADODB.Recordset")
    rsSubGroups.Open sSQL, conn, 3, 3
    
    if not rsSubGroups.EOF then
      
	  Table_Start Translate("Code",Login_Language,conn),_
	    Translate("Group Description",Login_Language,conn)
      
      do while not rsSubGroups.EOF
	  	Table_row RegionColor(rsSubGroups("Region")),rsSubGroups("Code"),_
			Translate(rsSubGroups("X_Description"),Login_Language,conn)
        
        rsSubGroups.MoveNext
      loop
      
	  Table_Done
    end if
    
    rsSubGroups.close
    set rsSubGroups = nothing
	
end sub

sub Write_Countrys
	with response
      .write "<P><BR><SPAN CLASS=MediumBold>"
	  .write Translate("Country Restrictions",Login_Language,conn) & "</SPAN><BR>" & vbCrLf
      .write "<SPAN CLASS=Small>"
	  .write Translate("This document is restricted for view or download in the following countries only.",Login_Language,conn) & "</SPAN><P>" & vbCrLf      
	end with
    
    SQL = "SELECT Country.Abbrev, Country.Name, Country.Region FROM Country ORDER BY Country.Name"
    Set rsCountry = Server.CreateObject("ADODB.Recordset")
    rsCountry.Open SQL, conn, 3, 3
    
    if not rsCountry.EOF then
	  Table_Start Translate("Code",Login_Language,conn),Translate("Country",Login_Language,conn)
      
      do while not rsCountry.EOF
        if instr(1, UCase(Asset_Country), UCase(rsCountry("Abbrev"))) > 0 then
		  Table_Row RegionColor(rsCountry("Region")),rsCountry("Abbrev"),_
		      Translate(rsCountry("Name"),Login_Language,conn)
        end if  
        
        rsCountry.MoveNext
      loop
      
      Table_Done
    end if
    
    rsCountry.close
    set rsCountry = nothing
end sub

sub Table_start(strOne,strTwo)
  	with response
    Call Table_Begin
	  .write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=0 CELLSPACING=0 "
	  .write "BORDERCOLOR=""#666666"" BGCOLOR=""#666666"">" & vbCrLf
      .write "  <TR>" & vbCrLf
      .write "    <TD>" & vbCrLf
      .write "      <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
      .write "        <TR>" & vbCrLf
      .write "          <TD WIDTH=""20%"" BGCOLOR=""Red"" ALIGN=CENTER CLASS=SmallBoldWhite>"
	  .write strOne & "</TD>" & vbCrLf
      .write "          <TD WIDTH=""80%"" BGCOLOR=""#000000"" ALIGN=CENTER CLASS=SmallBoldGold>"
	  .write strTwo & "</TD>" & vbCrLf
      .write "    		</TR>" & vbCrLf
	end with
end sub

sub Table_row(rcolor,rcode,rdesc)
  with response
    .write "<TR>" & vbCrLf
    .write "<TD ALIGN=CENTER CLASS=Medium BGCOLOR=""" & rcolor & """>" & rcode & "</TD>" & vbCrLf
    .write "<TD ALIGN=Left CLASS=Medium BGCOLOR=""" & rcolor & """>" & rdesc & "</TD>" & vbCrLf
	.write "</TR>" & vbCrLf
  end with
end sub

sub Table_Done
  response.write "</TABLE>" & vbCrLf & "</TD>" & vbCrLf & "</TR>" & vbCrLf & "</TABLE>" & vbCrLf
  Call Table_End
end sub
%>
