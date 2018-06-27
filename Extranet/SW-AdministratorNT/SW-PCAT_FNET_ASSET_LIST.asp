<%@ language="VBScript" codepage="65001" %>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/sw-administratorNT/SW-PCAT_FNET_IISSERVER.asp"-->
<%
on error resume next
Dim strHttpReferer
Dim Site_ID
Dim Record_Limit
Dim rsCalendar
Dim InvisionLogin

strHttpReferer = Request.ServerVariables("HTTP_REFERRER")
server.ScriptTimeout =3600
PCID     = CInt(request("PCID"))
'Page Sequence
if PCID  = Null or PCID = 0 then PCID = 0
Session("BackURL_Calendar") = ""
Record_Limit= 100
Site_ID = request.QueryString("Site_ID")
if trim(Site_ID) = "" then
    ShowRetrieveError("Invalid Site Id")
    Response.End
end if

InvisionLogin = false
' --------------------------------------------------------------------------------------
Call Connect_SiteWide
' --------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------
' Associate to Product Success / Failure Dialog
' --------------------------------------------------------------------------------------
if trim(Request.Form("username")) = "" then
%>
  <!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%
else
  InvisionLogin = true
end if

set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
'Pass the actual url here 
Call Products.open("POST",striisserverpath,0,0,0)
Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
Call Products.send("operation=PA")
strProducts=Products.responseXML.XML
set objxml=Server.CreateObject("msxml2.DomDocument")
Call objxml.loadxml(strProducts)
FormName = "AssetList"
%>
<script language="javascript">
<!--
    var FormName = document.<%=FormName%>
//-->
</script>
<%
Screen_TitleX = Translate("List Assets",Login_Language,conn)
SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3
Site_Code        = rsSite("Site_Code")     
Screen_Title     = rsSite("Site_Description") & " - " & Screen_TitleX
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=MediumBoldGold>" & Screen_TitleX & "</FONT>"
Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent
Logo             = rsSite("Logo")  
Logo_Left        = rsSite("Logo_Left")
%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<form name="<%=FormName%>" method="post">
    <p>
        <table cellpadding="0" align="CENTER" width="100%" border="0" id="ContentTable">
            <tr>
                <td width="20%">
                    <%if InvisionLogin = false then %>
                        <table border="0" cellpadding="0" align="LEFT" cellspacing="0" border="0" class="NAVBORDER">
                            <tr>
                                <td background="/images/SideNav_TL_corner.gif">
                                    <img src="/images/Spacer.gif" border="0" width="8" height="6" alt=""></td>
                                <td>
                                    <img src="/images/Spacer.gif" border="0" height="6" alt=""></td>
                                <td background="/images/SideNav_TR_corner.gif">
                                    <img src="/images/Spacer.gif" border="0" width="8" height="6" alt=""></td>
                            </tr>
                            <tr>
                                <td><img src="/images/Spacer.gif" width="8"></td>
                                <td valign="top" class="NAVBORDER">
                                    <a href="<%="/SW-ADMINISTRATOR/DEFAULT.ASP?SITE_ID=" & Site_ID & "&ASSOCIATE=" & Associate%>"
                                        class="NavLeftHighlight1">&nbsp;Main Menu&nbsp;</a>
                                </td>
                                <td><img src="/images/Spacer.gif" width="8" alt="" /></td>
                            </tr>
                            <tr>
                                <td background="/images/SideNav_BL_corner.gif">
                                    <img src="/images/Spacer.gif" border="0" width="8" height="6" alt=""></td>
                                <td>
                                    <img src="/images/Spacer.gif" border="0" height="6" alt=""></td>
                                <td background="/images/SideNav_BR_corner.gif">
                                    <img src="/images/Spacer.gif" border="0" width="8" height="6" alt=""></td>
                            </tr>
                        </table>
                    <%end if %>
                </td>
                <td width="80%"></td>
            </tr>
            <tr><td></td><td></td></tr>
            <tr><td></td><td></td></tr>
            <tr><td></td><td></td></tr>
            <tr>
                <td class="SmallBold"><%=Translate("Products",Login_Language,conn) %></td>
                <td align="left">
                    <% 
                         set objcol = objxml.selectsingleNode("Info")
  
                         if not(objcol is nothing) then
      		                 set objcolProducts = objcol.firstChild
                             set objcol = objcolProducts.firstChild
      		                 if (objcol is nothing) then
        	                       response.clear
      			                   ShowRetrieveError(objcolProducts.nodevalue)
        	                 end if
                             response.write "<SELECT Class=small name=""PCat_LProducts"" CLASS=Medium onchange=""RetrieveAssets(this.options[this.selectedIndex].value);this.style.cursor='Wait'"">" & vbCrLf
                             Response.Write "<OPTION value=""0"">All Products</OPTION>"
                             for icol = 0 to objcol.childnodes.length-1
  		                        set objinfo  = objcol.childnodes(icol)
                                set objinfo1 = objinfo.firstchild
          	                    set objinfo2 = objinfo.lastchild
                                if trim(Request("PCat_LProducts")) = trim(objinfo1.text) then 
                                    response.write "<OPTION selected Class=Medium value=""" & objinfo1.text & """ title=""" & objinfo2.text & """>" & objinfo2.text & "</OPTION>" & vbCrLf  
                                else
                                    response.write "<OPTION Class=Medium value=""" & objinfo1.text & """ title=""" & objinfo2.text & """>" & objinfo2.text & "</OPTION>" & vbCrLf  
                                end if                                    
      	                     next 
		                     response.write "</SELECT>" & vbCrLf
                             if err.number<>0 then
       		                    ShowRetrieveError(err.Description)
                             end if
      	                 end if
                    %>
                </td>
            </tr>
            <tr>
                <td class="SmallBold">
                    <%=Translate("Language",Login_Language,conn) %>
                </td>
                <td align="left">
                    <select name="PCat_Language" style="width=207" class="small" onchange="RetrieveAssets(this.options[this.selectedIndex].value);this.style.cursor='Wait'">
                        <%
                    Dim rs
                    set rs = conn.execute("exec PCAT_FNET_GETLANGUAGES")
                    response.Write "<option value=""0"">All Languages</option>"
                    do while not(rs.EOF)
                        if trim(Request("PCat_Language")) = trim(rs.fields("Code").value) then
                        %>
                            <option Class=Medium selected value="<%=rs.fields("Code").value%>">
                                <%=Translate(rs.fields("Description").value,Login_Language,conn) %>
                            </option>
                        <% 
                        else
                        %>
                            <option Class=Medium value="<%=rs.fields("Code").value%>">
                                <%=Translate(rs.fields("Description").value,Login_Language,conn) %>
                            </option>
                        <% 
                        end if  
                        rs.MoveNext
                    loop
                        %>
                    </select>
                </td>
            </tr>
            <tr>
                <td class="SmallBold">
                    <%=Translate("Content Category",Login_Language,conn) %>
                </td>
                <td align="left">
                    <select name="PCat_Category" style="width=207" class="small" onchange="RetrieveAssets(this.options[this.selectedIndex].value);this.style.cursor='Wait'">
                    <%
                    set rs = conn.execute("exec PCAT_FNET_GETCATEGORIES " & Site_ID)
                    response.Write "<option value=""0"">All Categories</option>"
                    do while not(rs.EOF)
                        if trim(Request("PCat_Category")) = trim(rs.fields("id").value) then
                        %>
                        <option Class=Medium selected value="<%=rs.fields("id").value%>">
                            <%=Translate(rs.fields("Title").value,Login_Language,conn) %>
                        </option>
                        <% 
                        else
                        %>
                        <option Class=Medium value="<%=rs.fields("id").value%>">
                            <%=Translate(rs.fields("Title").value,Login_Language,conn) %>
                        </option>
                        <%
                        end if   
                        rs.MoveNext
                    loop
                    %>
                    </select>
                </td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td align="left"><input class="small" type="button" value="Clear all filters" onclick="ClearFilters()" /></td>
            </tr>
            <tr>
                <td class="SmallBoldRed">
                    <div id="Standby" name="Standby" style="visibility: hidden">
                        <%=trim(Translate("Standby... Retrieving Assets.",Login_Language,conn))%>
                    </div>
                </td>
                <td><input type="hidden" name="username" value="<%=request.form("username")%>" />
                    &nbsp;</td>
            </tr>
        </table>
        <%
        strProducts = ""
        if trim(request("PCat_LProducts")) <> "0" and trim(request("PCat_LProducts")) <> "" then
          set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")

          'Pass the actual url here
          Call Products.open("POST",striisserverpath,0,0,0)
          Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
          Call Products.send("operation=PAssets&assetPid=" & request("PCat_LProducts"))

          strProducts=Products.responseXML.XML
          set objxml=Server.CreateObject("msxml2.domdocument")
          Call objxml.loadxml(strProducts)
          set products = nothing
          strProducts = objxml.text
        end if
       
        Cat_Id =0 
        if trim(request("PCat_Category")) <> "0" and trim(request("PCat_Category")) <> "" then
            Cat_Id = trim(request("PCat_Category"))
        end if
      
	    SQL ="EXEC PCAT_FNET_GETASSETS " & Site_ID & ",'" &  strProducts & "','" & trim(Request("PCat_Language")) & "'," & Cat_Id
        set rsCalendar = Server.CreateObject("ADODB.Recordset")
        set rsCalendar = conn.execute(SQL)
        
        Asset_ID_Old = 0
        Asset_Record_Count = 0
     
        if not( rsCalendar.eof) then Asset_Record_Count = rsCalendar.Fields(0).Value
        
        set rsCalendar =rsCalendar.NextRecordset
        Record_Count = Asset_Record_Count
        Record_Pages  = Record_Count \ Record_Limit
        if Record_Count mod Record_Limit > 0 then Record_Pages = Record_Pages + 1
        Page_QS = "Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&PCat_Language=" & Request("PCat_Language") & _
        "&PCat_Category=" & Request("PCat_Category") & "&PCat_LProducts=" & Request("PCat_LProducts") 
        xPCID = 1
        Record_Number = 0
    
        if not(rscalendar.EOF ) then
          if Record_Limit * (PCID - 1) > 0 then
            rsCalendar.Move (Record_Limit * (PCID - 1))
          end if 
        end if 
        Record_Number = 1
        Response.Write "<Div name=""AssetDiv"" id=""AssetDiv"">"
        response.write "<SPAN CLASS=SMALL>" & Translate("Total Assets ",Login_Language,conn) & ": " & Asset_Record_Count
        response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("Current Date/Time",Login_Language,conn) & ": " & Now() & " PST</SPAN><P>"
        Call RS_Page_Navigation()
        response.write "<br>"
        Call Table_Begin
        %>
        <table align="CENTER" cellpadding="2" width="100%" id="AssetData" class="tableborder">
        <tr>
         <td bgcolor="#666666" align="left" width="5%" class="smallboldwhite">
            <%=Translate("ID",Login_Language,conn)%>
         </td>
         <td bgcolor="#666666" align="left" width="30%" class="smallboldwhite">
            <%=Translate("Title",Login_Language,conn)%>
         </td>
         <td bgcolor="#666666" align="left" width="43%" class="smallboldwhite">
            <%=Translate("Description",Login_Language,conn)%>
         </td>
         <td bgcolor="#666666" align="left" width="5%" class="smallboldwhite">
            <%=Translate("Item Number",Login_Language,conn)%>
         </td>
         <td bgcolor="#666666" align="left" width="5%" class="smallboldwhite">
            <%=Translate("LNG",Login_Language,conn)%>
         </td>
         <td bgcolor="#666666" align="left" width="2%" class="smallboldwhite">
            <%=Translate("Revision",Login_Language,conn)%>
         </td>
         <td bgcolor="#666666" align="left" width="10%" class="smallboldwhite">
            <%=Translate("Live Date",Login_Language,conn)%>
         </td>
        </tr>
      <% 
       Categoryid =""
       do while not rsCalendar.EOF and Record_Number <= Record_Limit
            if Categoryid <> rscalendar.Fields("Category_Id").value then
                Categoryid = rscalendar.Fields("Category_Id").value
                %>
                    <tr>
                        <td bgcolor="#FFFFFF" align="left" colspan="6" class="small"><b>
                            <%=rsCalendar.Fields("Category_Name")%></b>
                        </td>
                    </tr>
            <%end if %>    
            <tr>
                <td bgcolor="#FFFFFF" align="left" class="small">
                  <%=rsCalendar.Fields("Id") %> 
                </td>
                <td bgcolor="#FFFFFF" align="left"  class="small">
                    <%=rsCalendar.Fields("Title")%>
                </td>
                <td bgcolor="#FFFFFF" align="left"  class="small">
                    <%=rsCalendar.Fields("Description")%>
                </td>
                <td bgcolor="#FFFFFF" align="left"  class="small">
                    <%Response.Write "<a href=""http://www.flukenetworks.com/FNet/en-us/findit?Document=" & rsCalendar.Fields("Item_number") & """>" & rsCalendar.Fields("Item_number") & "</a>"%>
                </td>
                <td bgcolor="#FFFFFF" align="left"  class="small">
                    <%=UCase(Left(rsCalendar.Fields("Language") &"", 1)) & LCase(Mid(rsCalendar.Fields("Language")&"", 2))%>
                </td>
                <td bgcolor="#FFFFFF" align="left"  class="small">
                    <%=rsCalendar.Fields("Revision_code")%>
                </td>
                <td bgcolor="#FFFFFF" align="left"  class="small">
                    <%=rsCalendar.Fields("bdate")%>
                </td>
            </tr>
            <% Record_Number = Record_Number + 1
               rscalendar.MoveNext
       loop
    response.write "</TABLE>"
    Call Table_End
    Response.Write "<br>"
    Call RS_Page_Navigation
	Response.Write "</Div>"
%>
</form>
<%
  if err.number <> 0 then
    Response.Clear
    ShowRetrieveError(err.Description)
  end if
%>
<!--#include virtual="/sw-common/sw-footer.asp"-->
<%

sub ShowRetrieveError(strmessage)
    BackURL= "/sw-administrator/default.asp?Site_ID=" & Site_ID 
    response.write "<HTML>" & vbCrLf
	response.write "<HEAD>" & vbCrLf
	response.write "<TITLE>Information</TITLE>" & vbCrLf
    response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf    
	response.write "</HEAD>"
	response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
	response.write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """ >" & vbCrLf
	response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """ NAME=""1"">" & vbCrLf
	response.write "<DIV ALIGN=CENTER>"
	Call Nav_Border_Begin
	response.write "<TABLE CELLPADDING=10 ><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
	response.write strmessage & "<br><br>"
	response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
	response.write "</TD></TR></TABLE>" & vbCrLf
	Call Nav_Border_End
	response.write "</FORM>" & vbCrLf
	response.write "</DIV>"
	response.write "</BODY>"
	response.write "</HTML>"
    on error goto 0
    response.end
end sub	

' --------------------------------------------------------------------------------------
' Record Set Page Navigation
' --------------------------------------------------------------------------------------
Sub RS_Page_Navigation
  Page_QS = "Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&PCat_Language=" & Request("PCat_Language") & _
    "&PCat_Category=" & Request("PCat_Category") & "&PCat_LProducts=" & Request("PCat_LProducts")
  if PCID = 0 then PCID = 1

  ltEnabled = 0
  
  if Record_Pages > 1 then

    Call Nav_Border_Begin
    
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Page", Login_Language, conn) & ": &nbsp;</SPAN>"

  	if PCID = 1 then
  		Call RS_Page_Numbers
    	Response.Write "<A HREF=""javascript:ChangeLocation('" & "sw-pcat_fnet_asset_list.asp?" & Page_QS & "&PCID=" & PCID + 1 &  "')"" CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "&nbsp;&gt;&gt;&nbsp;</A>"
        response.write "&nbsp;&nbsp;"
  	else
  		if PCID = Record_Pages then
            ltEnabled = 1
  		    Response.Write "<A HREF=""javascript:ChangeLocation('" & "sw-pcat_fnet_asset_list.asp?" & Page_QS & "&PCID=" & PCID - 1 &  "')"" CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
            response.write "&nbsp;&lt;&lt;&nbsp</A>&nbsp;&nbsp;"
    	    Call RS_Page_Numbers
  		else
            ltEnabled = 1
  			Response.Write "<A HREF=""javascript:ChangeLocation('" & "sw-pcat_fnet_asset_list.asp?" & Page_QS & "&PCID=" & PCID - 1 &  "')"" CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
            response.write "&nbsp;&lt;&lt;&nbsp;</A>&nbsp;&nbsp;"
    		Call RS_Page_Numbers
  			Response.Write "<A HREF=""javascript:ChangeLocation('" & "sw-pcat_fnet_asset_list.asp?" & Page_QS & "&PCID=" & PCID + 1 &  "')"" CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
            response.write "&nbsp;&gt;&gt;&nbsp;</A>"
  		end if
  	end if
    Call Nav_Border_End
  end if
End Sub

' --------------------------------------------------------------------------------------
' Record Set Page Numbers
' --------------------------------------------------------------------------------------
Sub RS_Page_Numbers
  iBreak = 0
  for i = 1 to Record_Pages
  	if i = PCID then
	  	Response.Write "<A HREF=""javascript:ChangeLocation('" & "sw-pcat_fnet_asset_list.asp?" & Page_QS & "&PCID=" & i &  "')"" CLASS=NAVLEFTHIGHLIGHT1>"
      response.write "&nbsp;"
      if i < 10 then response.write "&nbsp;&nbsp;"
      response.write CStr(i) & "&nbsp;</A>"
      if iBreak = 19 - (ltEnabled) then
        response.write "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        iBreak = -1
        ltEnabled = 0
      else
        response.write "&nbsp;&nbsp;"
      end if  
  	else
	  	Response.Write "<A HREF=""javascript:ChangeLocation('" & "sw-pcat_fnet_asset_list.asp?" & Page_QS & "&PCID=" & i &  "')"" CLASS=NavTopHighLight>"
        response.write  "&nbsp;"
        if i < 10 then response.write "&nbsp;&nbsp;"
          response.write CStr(i) & "&nbsp;</A>"
        if iBreak = 19 - (ltEnabled) then
          response.write "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
          iBreak = -1
          ltEnabled = 0        
        else
          response.write "&nbsp;&nbsp;"
        end if  
  	end if
    iBreak = iBreak + 1
  next
end sub
%>

<Script language="JAVASCRIPT">
<!--
  
  var sid =  <%=Site_ID%>
  var relate =  '<%=Associate%>'                               
  function ChangeLocation(strLocation)
  {
    document.getElementById("Standby").style.visibility = "visible";
    document.<%=FormName%>.action =strLocation;
    document.<%=FormName%>.submit();
    //window.location.href= strLocation;
  }
  
  function ClearFilters()
  {
      if (document.<%=FormName%>.PCat_LProducts.selectedIndex != 0 || document.<%=FormName%>.PCat_Category.selectedIndex !=0
      || document.<%=FormName%>.PCat_Language.selectedIndex!=0)
      {
          document.<%=FormName%>.PCat_LProducts.selectedIndex = 0;
          document.<%=FormName%>.PCat_Category.selectedIndex = 0;
          document.<%=FormName%>.PCat_Language.selectedIndex = 0;
          document.getElementById("Standby").style.visibility = "visible";
  	      document.<%=FormName%>.action="SW-PCAT_FNET_ASSET_LIST.asp?Site_id=" + sid;
  	      document.<%=FormName%>.submit();
  	  }
  }  
  function RetrieveAssets(productId) 
  {
    document.getElementById("Standby").style.visibility = "visible";
  	document.<%=FormName%>.action="SW-PCAT_FNET_ASSET_LIST.asp?Site_id=" + sid;
  	document.<%=FormName%>.submit();
  } 
 
//-->
</script>
