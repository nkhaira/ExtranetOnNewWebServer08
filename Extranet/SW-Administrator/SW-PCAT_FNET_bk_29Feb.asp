<!--#include virtual="/sw-administrator/SW-PCAT_FNET_IISSERVER.asp"-->
<%
'on error resume next
PcatFlag=false
optiontext =""
'' PCAT Interface
'' Called by Calendar_Edit_Add.asp or Calendar_Edit_Update.asp depending on PCAT_System value.
if IsNumeric(Calendar_ID) then
    SQL = "SELECT Item_Number, Count(*) AS Counts FROM dbo.Calendar WHERE (Item_Number = '" & CStr(rs("Item_Number") & "") & "') GROUP BY Item_Number"
    Set rsAllItems = Server.CreateObject("ADODB.Recordset")
    rsAllItems.Open SQL, conn, 3, 3

    if not rsAllItems.EOF then
	    if(isnumeric(rsAllItems("Counts")) and CInt(rsAllItems("Counts")) > 1) then
		    SQL = "SELECT * FROM dbo.Calendar WHERE (Item_Number = '" & CStr(rs("Item_Number")) & "') ORDER  BY Revision_Code DESC, UDate DESC"
		    Set rsDuplicateItems = Server.CreateObject("ADODB.Recordset")
		    rsDuplicateItems.Open SQL, conn, 3, 3
    		
		    if not rsDuplicateItems.EOF then
			    rsDuplicateItems.MoveFirst
			    if (CStr(rsDuplicateItems("ID")) <> CStr(Calendar_ID)) then
                    optiontext = " checked disabled "
                    showpcat=false
                else
                    optiontext = ""
			    end if
		    end if
		    rsDuplicateItems.close
		    set rsDuplicateItems = nothing		
	    end if
    end if

    rsAllItems.close
    set rsAllItems = nothing
end if
'' Check whether to display the Product catalog screen or not.
'' 
if not IsNumeric(Calendar_ID) then
    showpcat=true
else
    if  clng(rs("PID"))=-1 then
        showpcat=false
    else 
        showpcat=true  
    end if
end if        

with response
    ' Build PCAT Relationship Fields  
    'Changed by zensar on 13-01-2006 for pcat-asset relationship
    '***************************************************************
    '**************************************************************** 
    'Code for getting the products from the Fnet
    set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
    'Pass the actual url here
    call Products.open("POST",striisserverpath,0,0,0)
    call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")

    if IsObject(rs) then
        strPID=rs("PID")
    else
        strPID=""
    end if
    'Response.Write striisserverpath
    'Response.End
    call Products.send("operation=P&assetpid=" & strPID & "&SiteID=" & Site_ID)
    '.Write "sendwa"
    strProducts=Products.responseXML.XML
    '.Write "<script language=javascript>alert('" & strPID & "');</script>"    
    '.Write "test123"
    '.Write strProducts
    '.Write err.Description
    '.End
    set objxml=Server.CreateObject("msxml2.domdocument")
    call objxml.loadxml(strProducts)
    
    .write "<TR height=0>" & vbCrLf
    .write "<TD height=0 coslpan=3 CLASS=Medium>" & vbCrLf
    
    
    if not IsNumeric(Calendar_ID) then
      'When posted to Calendar_Admin.asp, code generates new PID value before posting other PCAT dependant values to PCAT DB
      PID_Value = 0
      response.write "<INPUT TYPE=""hidden"" NAME=""PCat"" VALUE=""" & PID_Value & """>" & vbCrLf
      Response.Write "<INPUT TYPE=""hidden"" NAME=""pidDelete"" VALUE=""" & PID_Value & """>" & vbCrLf
      Response.Write "<input type=hidden name=oldLanguage value=''>"
      Response.Write "<input type=hidden name=oldItemNumber value=''>"
    else
          response.write "<INPUT TYPE=""hidden"" NAME=""PCat"" VALUE=""" & rs("PID") & """>" & vbCrLf  
          Response.Write "<input type=hidden name=oldLanguage value=" & rs.fields("language").value & ">"
          Response.Write "<input type=hidden name=oldItemNumber value=" & rs.fields("Item_Number").value & "" & ">"
          'Added by zensar on 26-05-2006.
	      'This hidden variable will be used to retain the PID in case if the user checks the 
	      '"Do not set relationship checkbox.This PID will be passed to Productengine soap"
	      PcatValidateSql="select Id,PID from calendar where id=" & rs.fields("clone").value  & _
	      " and site_id= " & CInt(Site_ID)
	      set rsValidate=conn.execute(PcatValidateSql)
	      if not(rsvalidate.eof or rsvalidate.bof) then
            if clng(rs.fields("clone").value) > 0 then
                Response.Write "<INPUT TYPE=""hidden""  NAME=""pidDelete"" VALUE=""" & rsValidate("PID") & """>" & vbCrLf
                PcatFlag=true
	        else
	            Response.Write "<INPUT TYPE=""hidden""  NAME=""pidDelete"" VALUE=""" & rs("PID") & """>" & vbCrLf
	        end if 
	      else
	            Response.Write "<INPUT TYPE=""hidden""  NAME=""pidDelete"" VALUE=""" & rs("PID") & """>" & vbCrLf
	      end if   
	  '>>>>>>>>>
      '>>>>>>>>>>>>>>>>>>>>>>>
    end if
    
    response.write "<INPUT TYPE=""HIDDEN"" NAME=""Exclude"" VALUE=""false"">" & vbCrLf  
    
    'Modified by zensar for Pcat-Asset Relationship on 21-01-2006 
    set objcol=objxml.selectsingleNode("Info")
    if not(objcol is nothing) then
        set objcolProducts=objcol.firstChild
        set objcol=objcolProducts.firstChild
        if (objcol is nothing) then
            Response.Clear
            ShowRetrieveError(objcolProducts.nodevalue)
        end if
        strletters=""
      
        .write "<SELECT  style=""display:none;""  size=""0"" name=""PCat_AllProducts"" CLASS=Medium>" & vbCrLf
        for icol=0 to objcol.childnodes.length-1
            set objinfo=objcol.childnodes(icol)
            set objinfo1=objinfo.firstchild
            set objinfo2=objinfo.lastchild
            
            if isnumeric(left(objinfo2.text,1))=false then
                if instr(1,strletters,left(objinfo2.text,1)) <=0 then
                    strletters=strletters & left(objinfo2.text,1) 
                end if 
            else
               blnnumbers=true       
            end if
            Response.Write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,30) & "</OPTION>" & vbCrLf  
        next 
        .write "			</SELECT>" & vbCrLf
       
        if err.number<>0 then
            ShowRetrieveError(err.Description)
        end if
    end if  
    .write "</TD>" & vbCrLf
    .Write "</tr>"  & vbCrLf
'***********************************************************************************************************  
    'Exclude from Pcat
    ' Field Title
    .write "<TR>" & vbCrLf
    .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    '.write Translate("Do not relate to Product Catalog",Login_Language,conn) & ":"
    .write Translate("Do not show this item on the web",Login_Language,conn) & ":"
    .write "</TD>" & vbCrLf

    'Required Icon or Space
    .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
    .write "</TD>"
    '.Write rs("PID")
    .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"
    if (not IsNumeric(Calendar_ID)) or optiontext = " checked disabled " then
          'When posted to Calendar_Admin.asp, code generates new PID value before posting other 
          'PCAT dependant values to PCAT DB
          'New record.
          response.write "<INPUT name=""PcatRelationNone"" title=""-1"" onclick=""EnableDisableControls()"" TYPE = CHECKBOX VALUE=""-1""" & optiontext & ">"       
          showpcat = true
    else
          if clng(rs.fields("PID").value) = -1 then
            'If the relationship with Pcat is not set earlier and if the record is cloned then
            'dont allow to edit the checkbox again.
				'''Validation for clones.
				if PcatFlag=true then   
					if clng(rsvalidate.fields("PID").value) <= 0 then
						response.write "<INPUT onclick=""EnableDisableControls()"" disabled name=""PcatRelationNone"" TYPE = CHECKBOX checked VALUE=""" & rs("PID") &  """" & optiontext & ">" 	
					else
						response.write "<INPUT onclick=""EnableDisableControls()"" name=""PcatRelationNone"" TYPE = CHECKBOX checked VALUE=""" & rs("PID") &  """" & optiontext & ">" 
					end if
			    else				
				    response.write "<INPUT onclick=""EnableDisableControls()"" name=""PcatRelationNone"" TYPE = CHECKBOX checked VALUE=""" & rs("PID") &  """" & optiontext & ">" 
			    end if				
                showpcat=false
          else
             response.write "<INPUT onclick=""EnableDisableControls()"" name=""PcatRelationNone"" title=""" & rs("PID") & """ TYPE = CHECKBOX VALUE=""" & rs("PID") &  """" & optiontext & ">" 
             showpcat = true
          end if 
          set rsValidate =nothing 
    end if
    
    .Write "</TD>"
    .write "</TR>"
    'Category
    'Field Title
    .write "<TR>" & vbCrLf
    .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    .write Translate("Sort Products",Login_Language,conn) & ":"
    .write "</TD>" & vbCrLf

    ' Required Icon or Space
    .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
    .write      "&nbsp;" ' Or required Icon
    .write "</TD>"

    ' Field
    .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"
    stralphabets="ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
    strfound=""
    for icnt=1 to len(stralphabets)
        if instr(1,strletters,mid(stralphabets,icnt,1))>0 then
            strfound=strfound & mid(stralphabets,icnt,1)
        end if
    next

    if showpcat=true then  
        .write "<SELECT name=""PCat_Category"" CLASS=Medium onchange=""changeproducts(this.options[this.selectedIndex].value)"">" & vbCrLf
    else
        .write "<SELECT name=""PCat_Category"" disabled CLASS=Medium onchange=""changeproducts(this.options[this.selectedIndex].value)"">" & vbCrLf    
    end if
    
    .write "<OPTION Value=""0"" SELECTED>" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
    if blnnumbers=true then
        .write "<OPTION Value=""1"" >0 - 9</OPTION>" & vbCrLf
    end if
    for icnt=1 to len(strfound)
        .write "<OPTION " & strselected & "Value=""" & mid(strfound,icnt,1) & """ >" & mid(strfound,icnt,1) & "</OPTION>" & vbCrLf
    next
    .write "</SELECT>" & vbCrLf
    .write "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF="""" onclick=""Category_Window=window.open('/sw-administrator/SW-PCAT_FNET_PRODUCT_LIST.asp?Site_ID=" & Site_ID & "&Language=" & Login_Language &  "','Category_Window','status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no');Category_Window.focus();return false;"" CLASS=Medium><SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;" & Translate("Matrix",Login_language,conn) & "&nbsp;&nbsp;</SPAN></A>"                
    .write "</TD>" & vbCrLf
    .write "</TR>" & vbCrLf

    'Field Title
    .write "<TR>" & vbCrLf
    .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    .write Translate("PCat Products",Login_Language,conn) & ":<br /><br />"
    .Write Translate("Hint: Select which products on the web where you want this asset to be displayed. (The asset will appear in the Knowledge and information tab for each product you select.)",Login_Language,conn)
    .write "</TD>" & vbCrLf

    'Required Icon or Space
    .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
    .write      "&nbsp;" ' Or required Icon
    .write "</TD>"

    'Field
    .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"

    .write "<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0>" & vbCrLf
    .write "<TR>" & vbCrLf
    
    .write "<TD CLASS=Medium WIDTH=""48%"">" & vbCrlf
    .write "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"" height=""80"" WIDTH=""100%"" CLASS=NavLeftHighlight1>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD CLASS=Medium>"
    .write        Translate("Available Products",Login_Language,conn)
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD>" & vbCrLf
    if showpcat=true then  
        .write "<SELECT  size=""5"" multiple name=""PCat_AProducts"" CLASS=Medium>" & vbCrLf
    else
        .write "<SELECT  size=""5"" disabled multiple name=""PCat_AProducts"" CLASS=Medium>" & vbCrLf
    end if
    .write "			</SELECT>" & vbCrLf
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>	
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "<TD width=""2%"">" & vbCrLf
    .write "<TABLE height=""61"" cellSpacing=""1"" cellPadding=""1"" border=""0"">" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "  	<TD CLASS=Medium ALIGN=CENTER>" & vbCrLf
    .write "      &nbsp;<INPUT type=""button"" value="">"" CLASS=NavLeftHighlight1 name=""btnAproducts"" onclick=""AddRemoveOptions('PCat_AProducts','PCat_SProducts')"">&nbsp;" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "	<TR>" & vbCrLf
    .write "		<TD CLASS=Medium ALIGN=CENTER>" & vbCrLf
    'modified by zensar on 20 - 01 -2005
    .write "      &nbsp;<INPUT type=""button"" value=""<"" CLASS=NavLeftHighlight1 name=""btnRproducts"" onclick=""RemoveOption('PCat_SProducts')"">&nbsp;" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "<TD WIDTH=""48%"">" & vbCrLf
    .write "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"" height=""80"" "" WIDTH=100%"" CLASS=NavLeftHighlight1>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "		<TD CLASS=Medium>" & vbCrLf
    .write        Translate("Selected Products",Login_Language,conn) & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "	<TR>" & vbCrLf
    .write "    <TD CLASS=Medium>" & vbCrLf
    
    %>
    <%'Code modified by zensar on 23-01-2006 for asset-pcat relationship
    if showpcat=true then
         .write " <SELECT LANGUAGE=""JavaScript"" multiple size=""5"" NAME=""PCat_SProducts"" CLASS=""Medium"">"
	else
         .write " <SELECT LANGUAGE=""JavaScript"" disabled multiple size=""5"" NAME=""PCat_SProducts"" CLASS=""Medium"">"
    end if
    set objcol=objxml.selectsingleNode("Info")
    if not(objcol is nothing) then
			set objcolProducts=objcol.firstchild
			set objcol=objcolProducts.lastchild
			'Response.Write objcol.childnodes.length & "1222"
			
			if objcol.nodename <> "AllProducts" then
                for icol=0 to objcol.childnodes.length-1
                    set objinfo=objcol.childnodes(icol)
                    set objinfo1=objinfo.firstchild
                    set objinfo2=objinfo.lastchild
                    Response.Write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,60) & "</OPTION>" & vbCrLf  
                next
			end if
	end if
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>%>
    </SELECT>
    <%.write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf

    .write "</TR>" & vbCrLf
    
    if(Site_ID <> 82) then
        ''if (Request.QueryString("ID") <> "" and Request.QueryString("ID") <> "add") then
            ''.Write "<tr><TD>&nbsp;</TD></tr><tr><TD>Click here to enter multilingual data :<input type=button CLASS='NavLeftHighlight1' onclick=""Win=window.open('/sw-administrator/SW-PCAT_FNET_SAVE_ProductMultilingual.asp?ID=" & strPID & "&SiteId=" & Site_ID & "','Locale_Window','status=no,height=300,width=450,scrollbars=yes,resizable=yes,toolbar=yes,links=no,left=300,top=300');Win.focus();return false;"" value='Multilingual' /></TD></tr>"
        ''end if
    end if
    
    .write "</TABLE>" & vbCrLf & vbCrLf

    .write "</TD>" & vbCrLf
    .write "</TR>" & vbCrLf

    if IsObject(rs) then
    .write "<INPUT TYPE=""hidden"" NAME=""opt"" VALUE=""U"">" & vbCrLf  
    else
    .write "<INPUT TYPE=""hidden"" NAME=""opt"" VALUE=""A"">" & vbCrLf  
    end if 
    'end if

 if(Site_ID = 82) then
    .write "<TR>" & vbCrLf
    .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    .write Translate("Industry",Login_Language,conn) & ":"
    .write "</TD>" & vbCrLf

    ' Required Icon or Space
    .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
    .write      "&nbsp;" ' Or required Icon
    .write "</TD>"

    ' Field
    .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"
    set objcol=objxml.selectsingleNode("Info")
    if not(objcol is nothing) then
        set objcolProducts=objcol.lastchild
        set objcol=objcolProducts.firstChild
        set objcolselected=objcolProducts.lastChild
        'response.Write  objcolProducts.childnodes.length
        if (objcol is nothing) then
            response.write "Code for error handling goes here."
        else
            for icol=0 to objcol.childnodes.length-1
                set objinfo=objcol.childnodes(icol)
                set objinfo1=objinfo.firstchild
                set objinfo2=objinfo.lastchild
                strselected=""
                if cint(objcolProducts.childnodes.length)<>1 then
                    if not(objcolselected is nothing) then
                        for iselected=0 to objcolselected.childnodes.length-1
                            set objselectedinfo=objcolselected.childnodes(iselected)    
                            set objselectedinfo1= objselectedinfo.firstchild
                            if trim(objselectedinfo1.text)=trim(objinfo1.text) then
                                strselected="checked"
                                'Response.Write iselected & " loop <br>"
                            end if
                        next
                    end if
                end if    
                if showpcat=true then
                    .Write "<INPUT name=""IndustryCode"" type=checkbox " & strselected & " value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,30) & "<BR>" & vbCrLf  
                else
                    .Write "<INPUT name=""IndustryCode"" disabled type=checkbox " & strselected & " value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,30) & "<BR>" & vbCrLf  
                end if
            next
        end if     
    end if
  end if
  if (Site_ID <> 82) then
    .write "<TR>" & vbCrLf
    .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
    .write Translate("Web Categories",Login_Language,conn) & ":<br /><br />"
    .Write Translate("Hint: Select which application note categories where you want this asset to be displayed. (The asset will appear under Support/Application Notes.)", Login_Language,conn)
    .write "</TD>" & vbCrLf

    'Required Icon or Space
    .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER>" 
    .write      "&nbsp;" 
    .write "</TD>"

    .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=MEDIUM>"

    .write "<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 >" & vbCrLf
    .write "<TR>" & vbCrLf
    
    .write "<TD CLASS=Medium WIDTH=""48%"" >" & vbCrlf
    .write "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"" height=""80""  CLASS=NavLeftHighlight1>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD CLASS=Medium>"
    .write        Translate("Available Categories",Login_Language,conn)
    .write "    </TD>" & vbCrLf
    .write "  </TR>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "    <TD >" & vbCrLf
     'RI#850 gpd
     .write "    <div style=""overflow-y:scroll;overflow-x:scroll;height:100px; width:260px;overflow: -moz-scrollbars-horizontal;"" >" & vbCrLf
      
  
    'if showpcat=true then  
    '    .write "<SELECT  size=""5"" multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
    'else
    '   .write "<SELECT  size=""5"" disabled multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
    'end if

      set objcol=objxml.selectsingleNode("Info") 

    if not(objcol is nothing) then
        set objcolProducts=objcol.lastchild
        set objcol=objcolProducts.firstChild
        set objcolselected=objcolProducts.lastChild
        if (objcol is nothing) then
           'RI#850 gpd
           if showpcat=true then  
            .write "<SELECT  size=""5"" multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
           else
            .write "<SELECT  size=""5"" disabled multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
           end if     
          'RI#850 gpd  
            response.write "Code for error handling goes here."
        else     
            'RI#850 gpd
             if showpcat=true then          
               .write "<SELECT  size="&objcol.childnodes.length &"  multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
             else      
               .write "<SELECT   size="&objcol.childnodes.length &" disabled multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
             end if   
             'RI#850 gpd
            for icol=0 to objcol.childnodes.length-1           
                set objinfo=objcol.childnodes(icol)
                set objinfo1=objinfo.firstchild
                set objinfo2=objinfo.lastchild
                strselected=""
                if cint(objcolProducts.childnodes.length)<>1 then
                    if not(objcolselected is nothing) then
                        for iselected=0 to objcolselected.childnodes.length-1
                            set objselectedinfo=objcolselected.childnodes(iselected)    
                            set objselectedinfo1= objselectedinfo.firstchild
                            if trim(objselectedinfo1.text)=trim(objinfo1.text) then
                                strselected="checked"
                            end if
                        next
                    end if
                end if   
                 'RI#850 gpd
                .Write "<OPTION value=""" & objinfo1.text & """>" & objinfo2.text & "</OPTION>" & vbCrLf
                 ' .Write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,30) & "</OPTION>" & vbCrLf
            next
        end if     
    else
          'RI#850 gpd
           if showpcat=true then  
            .write "<SELECT  size=""5"" multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
           else
            .write "<SELECT  size=""5"" disabled multiple name=""PCat_AWebCats"" CLASS=Medium>" & vbCrLf
           end if     
          'RI#850 gpd 
    end if
    .write "			</SELECT>" & vbCrLf
    'RI#850 gpd
    .write "    </DIV>" &  vbCrLf
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>	
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
   
    .write "</TD>" & vbCrLf
    .write "<TD width=""2%"">" & vbCrLf
    .write "<TABLE height=""61"" cellSpacing=""1"" cellPadding=""1"" border=""0"">" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "  	<TD CLASS=Medium ALIGN=CENTER>" & vbCrLf
    .write "      &nbsp;<INPUT type=""button"" value="">"" CLASS=NavLeftHighlight1 name=""btnAWebCats"" onclick=""AddRemoveOptions('PCat_AWebCats','PCat_SWebCats')"">&nbsp;" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "	<TR>" & vbCrLf
    .write "		<TD CLASS=Medium ALIGN=CENTER>" & vbCrLf

    .write "      &nbsp;<INPUT type=""button"" value=""<"" CLASS=NavLeftHighlight1 name=""btnRWebCats"" onclick=""RemoveOption('PCat_SWebCats')"">&nbsp;" & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "<TD WIDTH=""48%"">" & vbCrLf
    .write "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"" height=""80"" "" WIDTH=100%"" CLASS=NavLeftHighlight1>" & vbCrLf
    .write "  <TR>" & vbCrLf
    .write "		<TD CLASS=Medium>" & vbCrLf
    .write        Translate("Selected Categories",Login_Language,conn) & vbCrLf
    .write "    </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "	<TR>" & vbCrLf
    .write "    <TD CLASS=Medium>" & vbCrLf

    %>
    <% 
    'RI#850 gpd
     .write "    <div style=""overflow-y:scroll;overflow-x:scroll;height:100px; width:200px;overflow: -moz-scrollbars-horizontal;"">" & vbCrLf
     
    if showpcat=true then
       .write " <SELECT LANGUAGE=""JavaScript"" multiple size=""5"" NAME=""PCat_SWebCats"" CLASS=""Medium"">"
	   else
       .write " <SELECT LANGUAGE=""JavaScript"" disabled multiple size=""5"" NAME=""PCat_SWebCats"" CLASS=""Medium"">"
    end if

    set objcol=objxml.selectsingleNode("Info")
    if not(objcol is nothing) then
        set objcolProducts=objcol.lastchild
        set objcol=objcolProducts.firstChild
        set objcolselected=objcolProducts.lastChild
        if (objcol is nothing) then
            response.write "Code for error handling goes here."
        else
            for icol=0 to objcol.childnodes.length-1
                set objinfo=objcol.childnodes(icol)
                set objinfo1=objinfo.firstchild
                set objinfo2=objinfo.lastchild
                strselected=""
                if cint(objcolProducts.childnodes.length)<>1 then
                    if not(objcolselected is nothing) then
                        for iselected=0 to objcolselected.childnodes.length-1
                            set objselectedinfo=objcolselected.childnodes(iselected)    
                            set objselectedinfo1= objselectedinfo.firstchild
                            if trim(objselectedinfo1.text)=trim(objinfo1.text) then
                                strselected="checked"
                                .Write "<OPTION value=""" & objinfo1.text & """>" & mid(objinfo2.text,1,30) & "</OPTION>" & vbCrLf
                            end if
                        next
                    end if
                end if    
            next
        end if     
    end if
    
    ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  %>
    </SELECT>
    <% 
    .write "  </div>" & vbCrLf
    .write "  </TD>" & vbCrLf
    .write "	</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf

    .write "</TR>" & vbCrLf
    .write "</TABLE>" & vbCrLf & vbCrLf
    .write "</TD>" & vbCrLf
    .write "</TR>" & vbCrLf
  end if  
%>
    <SCRIPT LANGUAGE="JAVASCRIPT" >
    function  changeproducts(letter)
    {
        var iproductcount;
        var strproduct;
        var i;
        var ValidChars = "0123456789.";

        //alert(FormName.PCat_AllProducts.length);
        for(i=(FormName.PCat_AProducts.options.length-1);i>=0;i--) {
	        FormName.PCat_AProducts.options.remove(i);
        }
        //alert(letter);
        if (letter=="1")
        {
            for (iproductcount=0;iproductcount<FormName.PCat_AllProducts.length;iproductcount++)
            {   
                strproduct=FormName.PCat_AllProducts.options[iproductcount].text.toUpperCase();
                //alert(ValidChars.indexOf(strproduct.substring(0,1));
                if (ValidChars.indexOf(strproduct.substring(0,1)) != -1) 
                {
                    optnew = document.createElement("OPTION") ;
                    optnew.text=FormName.PCat_AllProducts.options[iproductcount].text;
		            optnew.value=FormName.PCat_AllProducts.options[iproductcount].value;
		            FormName.PCat_AProducts.options.add(optnew);  
                }
            }    
        } 	
        else
        {
            for (iproductcount=0;iproductcount<FormName.PCat_AllProducts.length;iproductcount++)
            {   strproduct=FormName.PCat_AllProducts.options[iproductcount].text.toUpperCase();
                if (strproduct.substring(0,1)==letter.toUpperCase())
                {
                    optnew = document.createElement("OPTION") ;
                    optnew.text=FormName.PCat_AllProducts.options[iproductcount].text;
		            optnew.value=FormName.PCat_AllProducts.options[iproductcount].value;
		            FormName.PCat_AProducts.options.add(optnew);  
                }
            } 
        }
    }
    
	function checkifexists(strTo, productid)
    {   var iproductcount;
        var objto   = eval('document.<%=FormName%>.' + strTo);
        //if (FormName.PCat_SLocales.length==null)
        //{
            for (iproductcount=0;iproductcount<objto.options.length;iproductcount++)
            {   
                if (objto.options[iproductcount].value==productid)
                {
                    return false;
                }
            } 
        //}
    }
    
    function EnableDisableControls()
    {
        var i;
        var catcnt;
        FormName.PCat_AProducts.disabled = (!FormName.PCat_AProducts.disabled);
        FormName.PCat_SProducts.disabled = (!FormName.PCat_SProducts.disabled);
        if(FormName.PCat_AWebCats != null)
        {  FormName.PCat_AWebCats.disabled = (!FormName.PCat_AWebCats.disabled);  }
        if(FormName.PCat_SWebCats != null)
        {  FormName.PCat_SWebCats.disabled = (!FormName.PCat_SWebCats.disabled);  }
        
        if (FormName.PCat_SWebCats != null && FormName.PCat_SWebCats.disabled==true)
        {
            for(i=FormName.PCat_SWebCats.length-1;i>=0;i--)
            {
                FormName.PCat_SWebCats.options.remove(i);
            }
        }
        
        FormName.PCat_Category.disabled  = (!FormName.PCat_Category.disabled);
        if (FormName.PCat_Category.disabled==false)
			{
				if (FormName.pidDelete.value!=-1)
				{
					FormName.PCat.value=FormName.pidDelete.value;
				}
				else
				{
				    FormName.PCat.value=0;
				}
				/*if (FormName.ID.value!=FormName.Clone.value)
				{
				   for (catcnt=0;catcnt<FormName.Category_Change.length;catcnt++)
                   {   
                        if (FormName.Category_Change.options[catcnt].value==FormName.Category_ID.value)
                        {
                            FormName.Category_Change.options[catcnt].selected=true;  
                        }
                   }
				   FormName.Category_Change.disabled=true; 
				}*/
			}
		else	
			{
				FormName.PCat.value=-1;
				//FormName.Category_Change.disabled=false;
			}
        
        if (FormName.PCat_Category.disabled  ==true)
        {
            FormName.PCat_Category.Exclude=true;
        }
        for(i=0; i<FormName.elements.length; i++)
        {
            if(FormName.elements[i].name=="IndustryCode")
            {
                FormName.elements[i].disabled=(!FormName.elements[i].disabled);
                if (FormName.elements[i].disabled==true)
                {
                    FormName.elements[i].checked=false;
                }
            }
        }
        
        if (FormName.PCat_SProducts.disabled==true)
        {
            for(i=FormName.PCat_SProducts.length-1;i>=0;i--)
            {
                FormName.PCat_SProducts.options.remove(i);
            }
        }
    }
    </SCRIPT>
    <%.write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
end with
if err.number <> 0 then
    Response.Clear
    ShowRetrieveError(err.Description)
end if
sub ShowRetrieveError(strmessage)
    BackURL= "/sw-administrator/default.asp?Site_ID=" & Site_ID 
	response.write "<HTML>" & vbCrLf
	response.write "<HEAD>" & vbCrLf
	response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
	response.write "<TITLE>Error</TITLE>" & vbCrLf
	response.write "</HEAD>"
	response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
	response.write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """>" & vbCrLf
	response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """>" & vbCrLf
	response.write "<DIV ALIGN=CENTER>"
	Call Nav_Border_Begin
	response.write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
	Response.Write strmessage & "<br><br>"
	Response.Write "Unable to retrive the product records.<br><br>"
	response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
	response.write "</TD></TR></TABLE>" & vbCrLf
	Call Nav_Border_End
	response.write "</FORM>" & vbCrLf
	response.write "</DIV>"
	response.write "</BODY>"
	response.write "</HTML>"
	'response.flush
	on error goto 0
	Response.End
end sub	
%>


