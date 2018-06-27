<%

' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/1/2003
'             SiteWide Standard Page Header
' --------------------------------------------------------------------------------------
Response.Buffer=True
Page_Timer_Begin = Now()
Response.CharSet = "utf-8"
Response.CodePage=65001 
response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN "">" & vbCrLf & vbCrLf
response.write "<!-- Whitlock's SiteWide Content Management System SW-CMS 3.0 " & Now & " PST -->" & vbCrLf & vbCrLf

response.write "<HTML>" & vbCrLf

response.write "<HEAD>" & vbCrLf

response.write "<TITLE>" & vbCrLf

if not isblank(Screen_Title) then
  response.write Screen_Title & vbCrLf
else
  response.write Translate("Extranet Support Site",Login_Language,conn) & vbCrLf
end if
      
response.write "</TITLE>" & vbCrLf

Meta_Charset  = "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">"
Meta_Language = "English"
Meta_ISO_Code = "en"

if IsObject(conn) then
  SQL = "SELECT Language.* FROM Language WHERE Language.Code='" & Login_Language & "'"
  Set rsMeta = Server.CreateObject("ADODB.Recordset")
  rsMeta.Open SQL, conn, 3, 3
  Meta_Charset  = rsMeta("Meta_Charset")
  Meta_Language = rsMeta("Description")
  Meta_ISO_Code = rsMeta("Code")  
  rsMeta.close
  set rsMeta = nothing

end if    

response.write "<META NAME=""AUTHOR"" CONTENT=""Kelly Whitlock - Kelly.Whitlock@fluke.com"">" & vbCrLf

if Cart_Mode = True or instr(LCase(request.ServerVariables("SCRIPT_NAME")),"csv.") > 0 then
  response.write "<META HTTP-EQUIV=""Expires"" CONTENT=""0"">" & vbCrLf
  response.write "<META HTTP-EQUIV=""PRAGMA"" CONTENT=""NO-CACHE"">" & vbCrLf
  response.write "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache"">" & vbCrLf
end if

'response.write "<META HTTP-EQUIV=""Edge-Control"" CONTENT=""no-store"">" & vbCrLf

response.write Meta_Charset & vbCrLf 

if Site_ID = 100 and instr(1,LCase(request.ServerVariables("SERVER_NAME")),"flukenetworks") > 0 then
  response.write "<LINK REL=STYLESHEET HREF=""/portweb/SW-Style.css"">" & vbCrLf
  Logo_Left = true
  Logo = "/images/FlukeNetworks-Logo.gif"
elseif FileExists(Site_Code & "\" & "SW-Style.css") then %><%   
  response.write "<LINK REL=STYLESHEET HREF=""/" & Site_Code & "/SW-Style.css"">" & vbCrLf
else
  response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
end if

' End of HEAD
%>




<!-- below css added for RI :1624 (Datapaq). added by : nkulkarn on 14 Jun 2011

<!--[if IE]><!-->  
<style type="text/css">
.TopHeader{padding-bottom:0px;}


</style> <!--<![endif]--> 

<!--[if !IE]><!--> 
<style type="text/css">
.TopHeader{padding-bottom:20px;}


</style> <!--<![endif]--> 
<!-- Below code added for RI#1595 (Google code update) -->
<script type="text/javascript">
var _gaq = _gaq || [];
_gaq.push(['_setAccount', 'UA-3420170-1']);
_gaq.push(['_setDomainName', '.fluke.com']);
_gaq.push(['_setAllowLinker', true]);
_gaq.push(['_setAllowHash', false]);
_gaq.push(['_trackPageview']);
(function() {
var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
})();




</script>

<%

response.write "</HEAD>" & vbCrLf

response.write "<BODY BGCOLOR=""White"" TOPMARGIN=""0"" LEFTMARGIN=""0"" MARGINWIDTH=""0"" MARGINHEIGHT=""0"" "
response.write "LINK =""#000000"" "
response.write "VLINK=""#000000"" "
response.write "ALINK=""#000000"" "
response.write "LANGUAGE=""Javascript"" "
response.write "ONLOAD=""" & OnLoadCode & """ "
response.write "ONUNLOAD=""" & OnUNLoadCode & """ "
response.write "ONFOCUS=""" & OnFocusCode & """ "
response.write "ONBLUR=""" & OnBlurCode & """"
response.write ">"
%>

<A NAME="VERY_TOP"></A>

<!-- Top Header -->
<% if not Logo_Left then %>
  <TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0" CLASS="TopBlackBar" >
    <TR>
      <TD HEIGHT=56 VALIGN=TOP ALIGN=LEFT>
  
      <!-- Top Navigation -->
  
    	  <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
      		<TR>
            <!-- Spacer -->
      			<TD WIDTH="1" ALIGN="LEFT" VALIGN="TOP" HEIGHT="16"></TD>
            
            <!-- Title -->
      			<TD WIDTH="100%">
      				<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
      					<TR>
      						<TD>
                    <%
				
                    if isblank(Bar_Title) then
                      Bar_Title = ""
                    elseif instr(1,LCase(Bar_Title),"smallboldgold") > 0 then
                      Bar_Title = Mid(Bar_Title,1,instr(1,LCase(Bar_Title),"smallboldgold")-1) & "SmallBoldBar" & Mid(Bar_Title,instr(1,LCase(Bar_Title),"smallboldgold")+13)
				

                    end if

                    Bar_Title = Replace(Bar_Title,"<BR>","<BR>&nbsp;")				
                    if instr(LCase(Site_Company),"fluke") > 0 AND (instr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") > 0 OR instr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla/5") > 0 OR instr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla/6") > 0) then
                      'response.write "<img src=""/images/fluke_newworld.gif"" border=0 width=342 height=52>" & vbCrLf
                      response.write "<div style=""position:absolute; left:6; top:14; filter: alpha(opacity=100);"">" & vbCrLf

                    elseif instr(1,UCase(Bar_Title),"<FIND_IT_HEADER>") then
                      response.write "<div style=""position:absolute; left:24; top:0;"">" & vbCrLf
                    else  
                      response.write "<div style=""position:absolute; left:6; top:14;"">" & vbCrLf                  
                    end if
  
                    response.write "&nbsp;<SPAN CLASS=Heading3Fluke>"
                    if not isblank(Bar_Title) then
                      response.write Bar_Title
                    else
                      response.write Translate("Extranet Support Site",Login_Language,conn)
                    end if
                    response.write "</SPAN>"
                    response.write "</DIV>"
                    %>
                  </TD>
      					</TR>
      				</TABLE>
      			</TD>
            
            <!-- Logo -->
      			<TD ALIGN="RIGHT" VALIGN="TOP">
      				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
                <TR>
                  <TD VALIGN="MIDDLE" >
                    <BR>
                <%
                if isblank(Logo) then
			    %>
				    <%
				    
				    if Site_ID = 1 then
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    elseif Site_ID = 5 then
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.png" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    elseif Site_ID = 4 then
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    else
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FlukeLogo3.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    end if
                    %>
                <%
                else
                %>
                    <%
			        if Site_ID = 1 then
			        %>
				        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
			        <%
			        elseif Site_ID = 5 then
			        %>
				        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.png" WIDTH=134 HEIGHT=44 BORDER=0></A>
			        <%
			        elseif Site_ID = 4 then
			        %>
				        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
			        <%
			        else
			        %>
                        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="<%=Logo%>" BORDER=0></A>
                    <%
                    end if
                    %>
                <%
                end if
                %>
                    <BR>                  
          				</TD>
                </TR>
              </TABLE>
      			</TD>         
      		</TR>
      	</TABLE>
      </TD>
    </TR>
  </TABLE>

  <% else %>

  <TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0" CLASS="TopBlackBar" >
  <!--<TABLE WIDTH="100%" CELLPADDING=0 CELLSPACING=0 BORDER="0" id="tblHeader" >-->
    <TR>
      <TD HEIGHT=56 VALIGN=TOP ALIGN=LEFT>
  
      <!-- Top Navigation -->
   <%
            
				    if Site_ID = 29 then
				    
				    %>
                 <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" CLASS="TopHeader" >
                  
                   <%
				   else
				   
				    %>
				    <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0"  >
				       <%
				   end if
				   
				    %>
    	 
      		<TR>
            <!-- Spacer -->
      			<TD WIDTH="1" ALIGN="LEFT" VALIGN="TOP" HEIGHT="16"></TD>
            
            <!-- Logo -->
      			<TD ALIGN="RIGHT" VALIGN="TOP">
      				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
                <TR>
                
                <%
            
				    if Site_ID = 29 then
				    
				    %>
                  <TD VALIGN="LEFT" style="padding-left:8px;" >
                  
                   <%
				   else
				   
				    %>
				     <TD VALIGN="LEFT">
				       <%
				   end if
				   
				    %>
                    <BR>
                    <%
                    
                if isblank(Logo) then
			    %>
				    <%
				    if Site_ID = 1 then
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    elseif Site_ID = 5 then
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.png" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    elseif Site_ID = 4 then
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    else
				    %>
					    <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FlukeLogo3.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
				    <%
				    end if
                    %>
                <%
                else
                %>
                    <%
			        if Site_ID = 1 then
			        %>
				        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
			        <%
			        elseif Site_ID = 5 then
			        %>
				        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.png" WIDTH=134 HEIGHT=44 BORDER=0></A>
			        <%
			        elseif Site_ID = 4 then
			        %>
				        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="/images/FCal_RGB_NB_REV_134px.gif" WIDTH=134 HEIGHT=44 BORDER=0></A>
			        <%
			        else
			        %>
                        <A HREF="<%=HomeURL%>?Site_ID=<%=Site_ID%>" TARGET="VERY_TOP"><IMG SRC="<%=Logo%>" BORDER=0></A>
                    <%
                    end if
                    %>
                <%
                end if
                %>
                    <BR>                  
          				</TD>
                </TR>
              </TABLE>
      			</TD>         
            
            <!-- Title -->
      			<TD WIDTH="100%">
      				<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
      					<TR>
      						<TD ALIGN=RIGHT>
                    <%
                    if isblank(Bar_Title) then
                      Bar_Title = ""
                    elseif instr(1,LCase(Bar_Title),"smallboldgold") > 0 then
                      Bar_Title = Mid(Bar_Title,1,instr(1,LCase(Bar_Title),"smallboldgold")-1) & "SmallBoldBar" & Mid(Bar_Title,instr(1,LCase(Bar_Title),"smallboldgold")+13)
                    end if
                    
                    Bar_Title = Replace(Bar_Title,"<BR>","&nbsp;&nbsp;&nbsp;&nbsp;<BR>")
                    
                    response.write "<SPAN CLASS=Heading3Fluke>"
                    if not isblank(Bar_Title) then
                      response.write Bar_Title & "&nbsp;&nbsp;&nbsp;&nbsp;"
                    else
                      response.write Translate("Extranet Support Site",Login_Language,conn) & "&nbsp;&nbsp;&nbsp;&nbsp;"
                    end if
                    response.write "</SPAN>"
                    response.write "</DIV>"
                    %>
                  </TD>
      					</TR>
      				</TABLE>
      			</TD>
      		</TR>
      	</TABLE>
      </TD>
    </TR>
  </TABLE>
  
  <% end if %>
<!-- END HEADER -->