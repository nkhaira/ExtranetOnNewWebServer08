<%@ Language="VBScript" CODEPAGE="65001" %>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/sw-administratorNT/SW-PCAT_FNET_IISSERVER.asp"-->
<%

' --------------------------------------------------------------------------------------
on error resume next
Session("BackURL_Calendar") = ""
Dim Site_ID
Site_ID        = request.QueryString("Site_ID")

Dim Associate
Associate        = request.QueryString("Associate")
if isblank(Associate) then
	Associate = true
end if

' --------------------------------------------------------------------------------------
Call Connect_SiteWide
' --------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------
' Associate to Product Success / Failure Dialog
' --------------------------------------------------------------------------------------

if request.form("operation") = "S" then

	set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")

	'Pass the actual url here

	Call Products.open("POST",striisserverpath,0,0,0)
	Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	Call Products.send("operation=Associate&assetPid=" & request.form("PCat_NLProducts") & "&assetId=" & request.form("AssetId") & "," & strAssets)

	strProducts=Products.responseXML.XML
	set objxml=Server.CreateObject("msxml2.domdocument")

	Call objxml.loadxml(strProducts)

	BackURL= "/sw-administrator/SW-PCAT_FNET_PROD_ASSET_REL.asp?Site_ID=" & Site_ID & "&Associate=false&operation=A" & "&PLCat_Category=" & mid(Request("Letter"),1,1) & "&PCat_LProducts=" & Request("PCat_NLProducts") & "&Result=True"

	response.write "<HTML>" & vbCrLf
	response.write "<HEAD>" & vbCrLf
	response.write "<TITLE>Result</TITLE>" & vbCrLf
  response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
	response.write "</HEAD>"
	response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
	response.write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """ >" & vbCrLf
	response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """ NAME=""1"">" & vbCrLf
	response.write "<DIV ALIGN=CENTER>"
	
 Call Nav_Border_Begin
	response.write "<TABLE CELLPADDING=10 ID=""Table2""><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
	'response.write objxml.text
	if objxml.text="True" then
				Response.Redirect(BackURL)
	else
				response.write Translate("Unable to associate Assets to Product.",Login_Language,conn) & "<br><br>"
				BackURL= "/sw-administrator/SW-PCAT_FNET_PROD_ASSET_REL.asp?Site_ID=" & Site_ID & "&Associate=True&operation=A" & "&PLCat_Category=" & mid(Request("PLCat_Category"),1,1) & "&PCat_LProducts=" & Request("PCat_LProducts") & _
				"&PNLCat_Category=" & Request("PNLCat_Category") & "&PCat_NLProducts=" & Request("PCat_NLProducts") 
	end if
	response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
	response.write "</TD></TR></TABLE>" & vbCrLf
	Call Nav_Border_End
	
 response.write "</FORM>" & vbCrLf
	response.write "</DIV>"
	response.write "</BODY>"
	response.write "</HTML>"
	on error goto 0
	response.end
  	
end if

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")

'Pass the actual url here

Call Products.open("POST",striisserverpath,0,0,0)
Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
Call Products.send("operation=PA")

strProducts=Products.responseXML.XML
set objxml=Server.CreateObject("msxml2.domdocument")

Call objxml.loadxml(strProducts)


FormName = "PrdRelationsForm"
%>
<script language=javascript>
<!--
var FormName = document.<%=FormName%>
//-->
</script>
<%

if Request("Associate") ="True" then
				Screen_TitleX = Translate("Bulk Association of Assets to New Product",Login_Language,conn)
else
				Screen_TitleX = Translate("List Assets Associated with Product",Login_Language,conn)
end if	

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
 
<FORM NAME="<%=FormName%>" method="post">
<P>
  <TABLE CELLPADDING=0 ALIGN=CENTER WIDTH="100%" BORDER=0 ID="ContentTable"  >  
    <TR>
      <TD WIDTH=45%>
        <TABLE BORDER="0" CELLPADDING="0" ALIGN=LEFT CELLSPACING="0" BORDER=0 CLASS=NAVBORDER ID="Table1">
          <TR>
          	<TD BACKGROUND="/images/SideNav_TL_corner.gif"><IMG SRC="/images/Spacer.gif" BORDER="0" WIDTH="8" HEIGHT="6" ALT=""></TD>
          	<TD><IMG SRC="/images/Spacer.gif" BORDER="0" HEIGHT="6" ALT=""></TD>
          	<TD BACKGROUND="/images/SideNav_TR_corner.gif"><IMG SRC="/images/Spacer.gif" BORDER="0" WIDTH="8" HEIGHT="6" ALT=""></TD>
          </TR>
          <TR>
          	<TD><IMG SRC="/images/Spacer.gif" WIDTH="8"></TD>
          	<TD VALIGN="top" CLASS=NAVBORDER>
          		<A HREF="<%="/SW-ADMINISTRATOR/DEFAULT.ASP?SITE_ID=" & Site_ID & "&ASSOCIATE=" & Associate%>" CLASS=NavLeftHighlight1>&nbsp;Main Menu&nbsp;</A>
            </TD>
          	<TD><IMG SRC="/images/Spacer.gif" WIDTH="8"></TD>
          </TR>
          <TR>
          	<TD BACKGROUND="/images/SideNav_BL_corner.gif"><IMG SRC="/images/Spacer.gif" BORDER="0" WIDTH="8" HEIGHT="6" ALT=""></TD>
          	<TD><IMG SRC="/images/Spacer.gif" BORDER="0" HEIGHT="6" ALT=""></TD>
          	<TD BACKGROUND="/images/SideNav_BR_corner.gif"><IMG SRC="/images/Spacer.gif" BORDER="0" WIDTH="8" HEIGHT="6" ALT=""></TD>
          </TR>
        </TABLE>
      </TD>
      <TD WIDTH=10%></TD>
      <TD WIDTH=45%></TD>
    </TR>
    <TR><TD></TD><TD></TD><TD></TD></TR>
    <TR><TD></TD><TD></TD><TD></TD></TR>
    <TR><TD></TD><TD></TD><TD></TD></TR>
    <TR>
      <TD ALIGN=LEFT WIDTH=45%>
        <%
							 '1
        set objcol = objxml.selectsingleNode("Info")
  
        if not(objcol is nothing) then
      		set objcolProducts = objcol.firstChild
            set objcol = objcolProducts.firstChild
      		if (objcol is nothing) then
        	    response.clear
      			ShowRetrieveError(objcolProducts.nodevalue)
        	end if

         strletters = ""
  
         response.write "<SELECT  style=""display:none;""  name=""PCat_LinkedProducts"" CLASS=Medium TITLE=""" & Translate("Select first alpha character of Product name.",Login_Language,conn) & """>" & vbCrLf

         for icol = 0 to objcol.childnodes.length-1
  		    	  set objinfo  = objcol.childnodes(icol)
         		set objinfo1 = objinfo.firstchild
          	set objinfo2 = objinfo.lastchild

        	  if isnumeric(left(objinfo2.text,1))=false then
	  	         if instr(1,ucase(strletters),ucase(left(objinfo2.text,1))) <=0 then
        									strletters=strletters & ucase(left(objinfo2.text,1)) 
  			    		  end if 
         		else
            	blnnumbers=true       
      			  end if
    
	          response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      		next 

		      response.write "</SELECT>" & vbCrLf

         	if err.number<>0 then
          		ShowRetrieveError(err.Description)
         	end if

      	end if
       
      	stralphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
      	strfound = ""

      	for icnt = 1 to len(stralphabets)
          if instr(1,ucase(strletters),mid(ucase(stralphabets),icnt,1))>0 then
      			strfound = strfound & mid(ucase(stralphabets),icnt,1)
          end if
      	next
  
        response.write "<SELECT name=""PLCat_Category"" CLASS=Medium onchange=""ChangeProducts(this.options[this.selectedIndex].value,'PCat_LinkedProducts','PCat_LProducts')"" TITLE="""& Translate("Select first alpha character of Product name.",Login_Language,conn) & """>" & vbCrLf    
        response.write "<OPTION Value=""0"" SELECTED>" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
        
      		if blnnumbers=true then
      								if request("PLCat_Category") ="1" then
  															response.write "<OPTION Value=""1"" selected>0 - 9</OPTION>" & vbCrLf
      								else
  															response.write "<OPTION Value=""1"" >0 - 9</OPTION>" & vbCrLf
      								end if
      		end if

        for icnt=1 to len(strfound)
          					if ucase(trim(request("PLCat_Category"))) =ucase(mid(trim(strfound),icnt,1)) then
																				response.write "<OPTION " & " selected " & " Value=""" & mid(strfound,icnt,1) & """ >" & mid(strfound,icnt,1) & "</OPTION>" & vbCrLf
      									else
        												response.write "<OPTION Value=""" & mid(strfound,icnt,1) & """ >" & mid(strfound,icnt,1) & "</OPTION>" & vbCrLf	
      									end if
        next
  
      	 response.write "</SELECT>" & vbCrLf
        %>
  
      </TD>
      <TD WIDTH=10%></TD>
      <TD WIDTH=45% ALIGN=LEFT>
        <%
        '2
        if Associate = "True" then
												blnnumbers = false
        			set objcol=objxml.selectsingleNode("Info")
			      		if not(objcol is nothing) then
																set objcolProducts=objcol.firstChild
           					set objcol=objcolProducts.lastChild
				      		
																if (objcol is nothing) then
      												response.clear
      												ShowRetrieveError(objcolProducts.nodevalue)
																end if
				      
      										set objcolnotrelated = objcol
      										strletters=""
      
																response.write "<SELECT  style=""display:none;""   name=""PCat_NotLinkedProducts"" CLASS=Medium TITLE=""Select first alpha character of Product name."">" & vbCrLf
      
																for icol=0 to objcol.childnodes.length-1
      														set objinfo  = objcol.childnodes(icol)
      														set objinfo1 = objinfo.firstchild
      														set objinfo2 = objinfo.lastchild
											      	
      														if isnumeric(left(objinfo2.text,1))=false then
      																	if instr(1,ucase(strletters),ucase(left(objinfo2.text,1))) <=0 then
      																				strletters = strletters & ucase(left(objinfo2.text,1))
      																	end if 
       													else
      																	blnnumbers=true       
      														end if
      														response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
       									next 
						      
																response.write "</SELECT>" & vbCrLf
      
															 if err.number<>0 then
      													ShowRetrieveError(err.Description)
      										end if
      				end if
        
      				stralphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
      				strfound = ""
			      
										for icnt=1 to len(stralphabets)
      							if instr(1,ucase(strletters),mid(ucase(stralphabets),icnt,1))>0 then
       										strfound=strfound & mid(stralphabets,icnt,1)
      							end if
      				next

      				if objcol.childnodes.length > 10 then
      							response.write "<SELECT name=""PNLCat_Category"" CLASS=Medium onchange=""ChangeProducts(this.options[this.selectedIndex].value,'PCat_NotLinkedProducts','PCat_NLProducts')"">" & vbCrLf    
       						response.write "<OPTION Value=""0"" SELECTED>" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
					        
													if blnnumbers = true then
      													if request("PNLCat_Category") ="1" then
      																	response.write "<OPTION Value=""1"" selected>0 - 9</OPTION>" & vbCrLf
      													else
      																	response.write "<OPTION Value=""1"" >0 - 9</OPTION>" & vbCrLf
      													end if
      							end if
					      		
													for icnt = 1 to len(strfound)
      															if ucase(trim(request("PNLCat_Category"))) =ucase(mid(trim(strfound),icnt,1)) then
      																				response.write "<OPTION " & " selected " & " Value=""" & mid(strfound,icnt,1) & """ >" & mid(strfound,icnt,1) & "</OPTION>" & vbCrLf
																					else
																										response.write "<OPTION Value=""" & mid(strfound,icnt,1) & """ >" & mid(strfound,icnt,1) & "</OPTION>" & vbCrLf
																					end if 
      							next
					      		
													response.write "</SELECT>" & vbCrLf
													response.write "<input type=""hidden"" value=""true"" name=""displayed"">"
      				else
			    						response.write "<input type=""hidden"" value=""false"" name=""displayed"">"
      				end if
     	else
      				response.write "<input type=""hidden"" value=""false"" name=""displayed"">"
     	end if
      %>
    </TD>
  </TR>

  <TR>
    <TD WIDTH=45% BGCOLOR="#FFFFFF" VALIGN=TOP CLASS=MEDIUM>&nbsp;</TD>
    <TD WIDTH=10%></TD>
      <% if Associate = "True" then %>
    	  <TD WIDTH=45% BGCOLOR="#FFFFFF" VALIGN=TOP CLASS=MEDIUM>&nbsp;</TD>
      <%end if%>
  </TR>
  
  <TR>
    <TD WIDTH=45% BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=MEDIUM><%=Translate("Active Products with Assets:",Login_Language,conn)%></TD>
    <TD WIDTH=10%></TD>
      <% if Associate = "True" then %>
    	  <TD WIDTH=45% BGCOLOR="#EEEEEE" VALIGN=TOP CLASS=MEDIUM><%=Translate("Associate Assets to Product:",Login_Language,conn)%></TD>
      <%end if%>
  </TR>

  <TR> 
    <TD ALIGN=LEFT WIDTH=45%>
 	  <SELECT  SIZE="5"  NAME="PCat_LProducts" STYLE="width=400" CLASS=MEDIUM onchange="RetrieveAssets(this.options[this.selectedIndex].value);this.style.cursor='Wait'">
      <%
      '3
      if request("operation")="A" then
											set objcol=objxml.selectsingleNode("Info")
      					if not(objcol is nothing) then
      						set objcolProducts=objcol.firstChild
      						set objcol=objcolProducts.firstChild
      											if (objcol is nothing) then
      														response.clear
    																ShowRetrieveError(objcolProducts.nodevalue)
																	end if
											  
      											for icol=0 to objcol.childnodes.length-1
      														set objinfo=objcol.childnodes(icol)
      														set objinfo1=objinfo.firstchild
      														set objinfo2=objinfo.lastchild
      														if  ucase(mid(trim(objinfo2.text),1,1)) = ucase(trim(request("PLCat_Category")))   then
			      																	if objinfo1.text = request("PCat_LProducts") then
      																								response.write "<OPTION selected value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																				else
      																								response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																				end if
																				elseif isnumeric(mid(trim(objinfo2.text),1,1)) and request("PLCat_Category") = "1" then
																						if cint(mid(objinfo2.text,1,1)) >=0 and cint(mid(objinfo2.text,1,1)) <=9 then
																												if objinfo1.text = request("PCat_LProducts") then
      																										response.write "<OPTION selected value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																						else
      																										response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																						end if
      																	end if
																					end if									
      											next 
											
      											if err.number<>0 then
      												ShowRetrieveError(err.Description)
      											end if
      						end if
  				end if
      %>

      </SELECT>
   
      <INPUT TYPE=HIDDEN NAME="operation" VALUE="">
    </TD>
    
    <TD WIDTH=10%></TD>
  
    <TD WIDTH=45% ALIGN=LEFT>
					 
      <% if Associate = "True" then %>
    										<SELECT  SIZE="5" NAME="PCat_NLProducts" CLASS=MEDIUM ID="Select1" STYLE="width=400">
	      								<%
															if objcolnotrelated.childnodes.length <= 10 then
    															for icol=0 to objcolnotrelated.childnodes.length-1
  	  																				set objinfo=objcolnotrelated.childnodes(icol)
    																					set objinfo1=objinfo.firstchild
    																					set objinfo2=objinfo.lastchild
    																					if Request("PCat_NLProducts") = objinfo1.text then
    																									response.write "<OPTION selected value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
    																					else
    																									response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
																									end if    																		
  	  														next 
  	  										else
  	  													set objcol=objxml.selectsingleNode("Info")
      												if not(objcol is nothing) then
      																		set objcolProducts=objcol.firstChild
      																		set objcol=objcolProducts.lastchild
      																		if (objcol is nothing) then
      																					response.clear
    																							ShowRetrieveError(objcolProducts.nodevalue)
																								end if
																		  
      																		for icol=0 to objcol.childnodes.length-1
      																					set objinfo=objcol.childnodes(icol)
      																					set objinfo1=objinfo.firstchild
      																					set objinfo2=objinfo.lastchild
      																					if  ucase(mid(trim(objinfo2.text),1,1)) = ucase(trim(request("PNLCat_Category")))   then
			      																							if objinfo1.text = request("PCat_NLProducts") then
      																														response.write "<OPTION selected value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																										else
      																														response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																										end if
																											elseif isnumeric(mid(trim(objinfo2.text),1,1)) and request("PNLCat_Category") = "1" then
																															if cint(mid(objinfo2.text,1,1)) >=0 and cint(mid(objinfo2.text,1,1)) <=9 then
																																				if objinfo1.text = request("PCat_NLProducts") then
      																																		response.write "<OPTION selected value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																														else
      																																		response.write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,70) & "</OPTION>" & vbCrLf  
      																														end if
      																										end if
																												end if									
      																		next 
																		
      																		if err.number<>0 then
      																								ShowRetrieveError(err.Description)
      																		end if
      													end if				
  													end if
															%>
  									</SELECT> 
  		 <%else		 
											if Request("Result") = "True" then
														ShowRetrieveError("Assets have been successfully Associated. See listing below.")
											end if
	     end if %>
    </TD>
  </TR>

  <TR> 
    <TD ALIGN=LEFT WIDTH=45% CLASS=SmallBoldRed>
    <DIV ID="Standby" NAME="Standby" style="visibility: hidden"><%=trim(Translate("Standby... Retrieving Assets for this Product",Login_Language,conn))%></DIV>
    </TD>
    <TD WIDTH=10%>&nbsp;</TD>
    <TD WIDTH=45% ALIGN=LEFT>
      <% if Associate = "True" then %>
				<INPUT CLASS=NAVLEFTHIGHLIGHT1 onmouseover="this.className='NavLeftButtonHover'" onmouseout="this.className='Navlefthighlight1'" TYPE=BUTTON VALUE="    GO    " NAME="btnGo" onclick="setRelationship()" ID="Button1">&nbsp;&nbsp;<span style="visibility: hidden" CLASS=SmallBoldRed ID="Processing" NAME="Processing" ><%=Translate("Standby... Associating Assets to this Product",Login_Language,conn)%></span>
      <% end if %>
    </TD>
  </TR>
</TABLE>

<BR>

  <%
  rowfound=false
  if request("operation") = "A" then     

    set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")

    'Pass the actual url here
    Call Products.open("POST",striisserverpath,0,0,0)
    Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    Call Products.send("operation=PAssets&assetPid=" & request("PCat_LProducts"))

    strProducts=Products.responseXML.XML
    set objxml=Server.CreateObject("msxml2.domdocument")

    Call objxml.loadxml(strProducts)

    if (objxml.text <> "NF") then
      SQL = "SELECT ID, Clone, Product, Title, Item_Number, Language, PID, Revision_Code, CASE dbo.Calendar.Clone WHEN 0 THEN dbo.Calendar.ID ELSE dbo.Calendar.Clone END AS PC_Order " &_
            "FROM Calendar " &_
            "  LEFT OUTER JOIN " &_
            "    dbo.[Language] ON dbo.Calendar.[Language] = dbo.[Language].Code " &_
            "WHERE PID IN (" & objxml.text & ") " &_
            "ORDER BY PC_Order, dbo.[Language].Sort"
          


      SQL = "SELECT dbo.Calendar.ID, dbo.Calendar.Clone, dbo.Calendar.Product, dbo.Calendar.Title, dbo.Calendar.Item_Number, dbo.Calendar.Language, dbo.Calendar.Status," &_
            "       dbo.Calendar.PID, dbo.Calendar.Revision_Code, " &_
            "        dbo.Calendar_Category.Title Category, " &_
            "       CASE dbo.Calendar.Clone WHEN 0 THEN dbo.Calendar.ID ELSE dbo.Calendar.Clone END AS PC_Order " &_
            "FROM   Calendar " &_
            "       LEFT OUTER JOIN dbo.[Language] ON dbo.Calendar.[Language] = dbo.[Language].Code " &_
            "       LEFT OUTER JOIN dbo.Calendar_Category ON dbo.Calendar.Category_ID = dbo.Calendar_Category.ID " &_
            "WHERE dbo.Calendar.PID IN (" & objxml.text & ") " &_
            "ORDER BY Category, PC_Order, dbo.[Language].Sort"
		set objxml = nothing
    	set rsCalendar = Server.CreateObject("ADODB.Recordset")
    	rsCalendar.Open SQL, conn, 3, 3
    
      Asset_ID_Old = 0
      Asset_Record_Count = 0
    
      do while not rsCalendar.EOF
        if rsCalendar("ID") <> Asset_ID_Old then
          Asset_Record_Count = Asset_Record_Count + 1
          Asset_ID_Old = rsCalendar("ID")
        end if
        rsCalendar.MoveNext
      loop

    rsCalendar.MoveFirst
    Response.Write "<Div name=""AssetDiv"" id=""AssetDiv"">"
    response.write "<SPAN CLASS=SMALL>" & Translate("Total Assets Associated with this Product",Login_Language,conn) & ": " & Asset_Record_Count
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Translate("Current Date/Time",Login_Language,conn) & ": " & Now() & " PST</SPAN><P>"
    
    Call Table_Begin
    %>

    <TABLE ALIGN=CENTER CELLPADDING=2  WIDTH="100%" ID="ContentTable" CLASS=TABLEBORDER>        
      <TR ID="ContentHeader1">
        
      <% if Associate = "false" then %>
        <TD BGCOLOR="#FF0000" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("Action",Login_Language,conn)%></TD>       
      <% else %>
        <TD BGCOLOR="#666666" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("Associate",Login_Language,conn)%></TD>  
      <% end if %>		
  
        <TD BGCOLOR="#666666" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("ID",Login_Language,conn)%></TD>      
        <TD BGCOLOR="#666666" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("PID",Login_Language,conn)%></TD>    
        <TD BGCOLOR="#666666" ALIGN="LEFT"   WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("Product",Login_Language,conn)%></TD>   
        <TD BGCOLOR="#666666" ALIGN="LEFT"              CLASS=SMALLBOLDWHITE><%=Translate("Title",Login_Language,conn)%></TD>
        <TD BGCOLOR="#666666" ALIGN="LEFT"   WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("Category",Login_Language,conn)%></TD>            
        <TD BGCOLOR="#666666" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("Item",Login_Language,conn)%></TD>    
        <TD BGCOLOR="#666666" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("Rev",Login_Language,conn)%></TD>    
        <TD BGCOLOR="#666666" ALIGN="CENTER" WIDTH="1%" CLASS=SMALLBOLDWHITE><%=Translate("LNG",Login_Language,conn)%></TD> 
      </TR>
    <%

      Category_Old = ""
      Asset_ID_Old = 0

      PC_Color     = "#FFFFFF"
      nPC_Color    = "#FFFFFF"      
      Old_PC_Order = 0
      

	    do while not (rsCalendar.EOF )
								rowfound= True
        if rsCalendar("ID") <> Asset_ID_Old then
     
          strLink="<A HREF=""javascript:Redirect("& rsCalendar.fields("ID").value &")"" Title=""Edit Asset"" CLASS=Navlefthighlight1>&nbsp;&nbsp;Edit&nbsp;&nbsp;</A>"
          %>
      		<TR>
            <% if Associate = "false" then 
    											if cstr(trim(Request("LastID")))=cstr(trim(rsCalendar.fields("ID"))) then %>
    														<TD  BGCOLOR="YELLOW" ALIGN=CENTER VALIGN=CENTER CLASS=SMALL TITLE="Click to Edit this Asset"><%=strLink%></TD>       
    											<%else%>
    														<TD  BGCOLOR="#FFFFFF" ALIGN=CENTER VALIGN=CENTER CLASS=SMALL TITLE="Click to Edit this Asset"><%=strLink%></TD>       
    											<%end if%>
       	    <% else
            
                if isnumeric(rsCalendar.fields("clone").value)= false or rsCalendar.fields("clone").value=0 then

                  response.write "<TD TITLE=""Check to associate or Uncheck to disassociate this Asset to a Product.  This association includes all alternate language versions."" BGCOLOR="""
                  if Old_PC_Order <> rsCalendar("PC_Order") then
                    if nPC_Color="#FFFFFF" then
                      nPC_Color = "#EEEEEE"
                    else
                      nPC_Color = "#FFFFFF"
                    end if
                  end if
                  response.write nPC_Color & """ ALIGN=""CENTER"" CLASS=Small>"
        	  		  %>
                  <INPUT TYPE=CHECKBOX NAME="AssetId" VALUE="<%=rsCalendar.fields("PID").value%>" ID="Checkbox1" CHECKED onclick="ChangeStatus(<%=rsCalendar.fields("PID").value%>)"></TD>
        	        <%
                else
                  response.write "<TD TITLE=""Alternate Language Versions"" BGCOLOR="""
                  response.write nPC_Color & """ ALIGN=""CENTER"" CLASS=Small>" & "<INPUT TYPE=CHECKBOX NAME=""AssetId"" VALUE=""" & rsCalendar.fields("PID").value & """ disabled CHECKED>"  & "</TD>"
      	        end if
         	    end if

              Status = rsCalendar("Status")
              
              select case Status
                case 1        
                  response.write "<TD TITLE=""Asset ID Status=LIVE"" BGCOLOR=""#00CC00"" ALIGN=""CENTER"" CLASS=Small>"
                case 2
                  response.write "<TD TITLE=""Asset ID Status=ARCHIVE"" BGCOLOR=""#AAAAFF"" ALIGN=""CENTER"" CLASS=Small>"
                case else
                  response.write "<TD TITLE=""Asset ID Status=REVIEW"" BGCOLOR=""Yellow"" ALIGN=""CENTER"" CLASS=Small>"
              end select
              response.write rsCalendar("ID")
              %>
            </TD>  
            <%
            response.write "<TD TITLE=""Parent ID"" BGCOLOR="""
            if Old_PC_Order <> rsCalendar("PC_Order") then
              if PC_Color="#FFFFFF" then
                PC_Color = "#EEEEEE"
              else
                PC_Color = "#FFFFFF"
              end if
              Old_PC_Order = rsCalendar("PC_Order")
            end if
            response.write PC_Color & """ ALIGN=""CENTER"" CLASS=Small>"
            
            if rsCalendar.fields("clone").value <> 0 then
              response.write rsCalendar.fields("clone").value
            else
              response.write "&nbsp;"
            end if
            %>
            </TD>  

      			<TD ALIGN="LEFT"   TITLE="Product or Product Series" BGCOLOR="#FFFFFF" CLASS=SMALL NOWRAP><%=rsCalendar.fields("Product").value%></TD>    
            <TD ALIGN="LEFT"   TITLE="Title of Asset" BGCOLOR="#FFFFFF" CLASS=SMALL><%=rsCalendar.fields("Title").value%></TD>   
            <TD ALIGN="LEFT"   TITLE="Asset Category" BGCOLOR="#FFFFFF" CLASS=SMALL NOWRAP><%=rsCalendar.fields("Category").value%></TD>               
      			<TD ALIGN="CENTER" TITLE="Oracle or Generic Item Number" BGCOLOR="#FFFFFF" CLASS=SMALL><%=rsCalendar.fields("Item_Number").value%></TD>    
      			<TD ALIGN="CENTER" TITLE="Revision" BGCOLOR="#FFFFFF" CLASS=SMALL><%=rsCalendar.fields("Revision_Code").value%></TD>            
            <%

            ' Language
            if UCase(rsCalendar("Language")) = "ENG" then
              response.write "<TD TITLE=""Language"" BGCOLOR=""#FFFFFF"" ALIGN=""CENTER"" CLASS=Small>"
            else  
              response.write "<TD TITLE=""Language"" BGCOLOR=""#CFCFCF"" ALIGN=""CENTER"" CLASS=Small>"
            end if
        
            response.write UCase(rsCalendar("Language"))
            %>
            </TD>
      		</TR>
          <%

          Asset_ID_Old = rsCalendar("ID")
        
        end if

  	  	rsCalendar.MoveNext

      loop
	  
    	rsCalendar.close
    	set rsCalendar = nothing
       
      %>
      <SCRIPT language="JavaScript">
      <!--
        document.getElementById("Standby").style.visibility = "hidden";
      //-->
      </SCRIPT>
      <%
      
    end if
    response.write "</TABLE>"
    Call Table_End
	   Response.Write "</Div>"
	 else
				Response.Write "<Div name=""AssetDiv"" id=""AssetDiv"">" 
				Response.Write "</Div>"
  end if
  Response.Write "<input type=hidden name=""rowfound"" value=""" & rowfound & """>"
  response.write "</FORM>"
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
	response.write "<TABLE CELLPADDING=10 ID=""Table2""><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
	response.write strmessage & "<br><br>"
	if Request("Result") <> "True" then
			response.write "Unable to retrieve database records.<br><br>"
			response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
  end if		
	response.write "</TD></TR></TABLE>" & vbCrLf
	Call Nav_Border_End
	response.write "</FORM>" & vbCrLf
	response.write "</DIV>"
	response.write "</BODY>"
	response.write "</HTML>"
		on error goto 0
	if Request("Result") <> "True" then
		response.end
	end if
end sub	
%>

<SCRIPT LANGUAGE=JAVASCRIPT>
<!--
var sid =  <%=Site_ID%>
var relate =  '<%=Associate%>'                               
  
  function  ChangeProducts(letter,strselectBox1,strselectBox2) {
    var iproductcount;
    var strproduct;
    var i;
    var ValidChars = "0123456789.";

				var selectBox1 = eval('document.<%=FormName%>.' + strselectBox1);
				var selectBox2   = eval('document.<%=FormName%>.' + strselectBox2);
				if (strselectBox1=="PCat_LinkedProducts")
				{
						document.getElementById("AssetDiv").style.visibility = "hidden";
				}
    for(i=(selectBox2.options.length-1);i>=0;i--) {
      selectBox2.options.remove(i);
    }
				if (document.<%=FormName%>.displayed.value== true) {
					if (strselectBox1=="PCat_LinkedProducts")	{
						for (iproductcount=0;iproductcount<document.<%=FormName%>.PNLCat_Category.options.length;iproductcount++) {
	  					if (document.<%=FormName%>.PNLCat_Category.options[iproductcount].value == letter) {
								document.<%=FormName%>.PNLCat_Category.selectedIndex = iproductcount;
							}
						}
					}
				}
    
    if (letter=="1") {
      for (iproductcount=0;iproductcount<selectBox1.options.length;iproductcount++) {   
        strproduct=selectBox1.options[iproductcount].text.toUpperCase();
        if (ValidChars.indexOf(strproduct.substring(0,1)) != -1) {
          optnew = document.createElement("OPTION") ;
          optnew.text=selectBox1.options[iproductcount].text;
          optnew.value=selectBox1.options[iproductcount].value;
          selectBox2.options.add(optnew);  
        }
      }    
    } 	
    else {
      for (iproductcount=0;iproductcount<selectBox1.options.length;iproductcount++) {
        strproduct=selectBox1.options[iproductcount].text.toUpperCase();
        if (strproduct.substring(0,1)==letter.toUpperCase()) {
          optnew = document.createElement("OPTION") ;
          optnew.text=selectBox1.options[iproductcount].text;
          optnew.value=selectBox1.options[iproductcount].value;
          selectBox2.options.add(optnew);  
        }
      } 
    }
  }
    
  function RetrieveAssets(productId) {
  	document.body.style.cursor="Wait";
   document.getElementById("Standby").style.visibility = "visible";
  	document.<%=FormName%>.operation.value="A";
  	document.<%=FormName%>.action="SW-PCAT_FNET_PROD_ASSET_REL.asp?Site_id=" + sid + "&Associate=" + relate;
  	document.<%=FormName%>.submit();
  } 

  function setRelationship() {
   var blnchecked=false;
  	var i;
  	var strMessage="";
   
    if (document.<%=FormName%>.rowfound.value == "True")
    {
  					if (document.<%=FormName%>.AssetId.length)
  					{
  						for (i = 0; i < document.<%=FormName%>.AssetId.length; i++)	
  						{
  									if (document.<%=FormName%>.AssetId[i].checked == true) {
  										blnchecked=true;
  							}
  						}
  					}
  					else
  					{
  								if(document.<%=FormName%>.AssetId.checked== true)
  								{
  										blnchecked=true;
  								}
  					}
  		}
  	
    if (blnchecked==false) {
  		strMessage="Please select Assets to associate with a Product.";
  	}
  
  	if(document.<%=FormName%>.PCat_NLProducts.selectedIndex ==-1)	{
  		strMessage=strMessage + "\nPlease select a Product to associate checked Assets to.";
  	}
  	
    if (strMessage !="")
    {
  		alert(strMessage);
  		return false;
  	}
   document.getElementById("Processing").style.visibility = "visible"; 
  	document.<%=FormName%>.operation.value="S";
  	document.<%=FormName%>.action="SW-PCAT_FNET_PROD_ASSET_REL.asp?Site_id=" + sid + "&Associate=" + relate + "&Letter=" + document.<%=FormName%>.PCat_NLProducts.options[document.<%=FormName%>.PCat_NLProducts.selectedIndex].text;
  	document.<%=FormName%>.submit();
  }
  function Redirect(pId)
  {
  
	document.<%=FormName%>.action="/sw-administrator/Calendar_Edit.asp?ID=" + pId + "&Site_id=" + sid + "&operation=A&PLCat_Category=" + document.<%=FormName%>.PLCat_Category.options[document.<%=FormName%>.PLCat_Category.selectedIndex].value + "&PCat_LProducts=" + document.<%=FormName%>.PCat_LProducts.options[document.<%=FormName%>.PCat_LProducts.selectedIndex].value;
	document.<%=FormName%>.submit();
	}
  function ChangeStatus(productId)
  {var i;
  var firsttime=false;
					if (document.<%=FormName%>.rowfound.value == "True")
					{
  			for (i = 0; i < document.<%=FormName%>.AssetId.length; i++)	
  			{
  				if (document.<%=FormName%>.AssetId[i].value == productId)
  				 {
  								if (firsttime== true)
  								{
  									document.<%=FormName%>.AssetId[i].checked= !(document.<%=FormName%>.AssetId[i].checked);
  								}
  								firsttime=true;
  				  }
  			}
  		}
  }
//-->
</SCRIPT>
