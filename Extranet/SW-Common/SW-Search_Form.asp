<%
          with response
          
            if SearchDB = 0 or 999 then SearchDB = 1
            
            .write "<FORM NAME=""Search"" ACTION=""" & BackURL & """ METHOD=""POST"">" & vbCrLf         
            .write "<INPUT TYPE=""Hidden"" NAME=""Language"" VALUE=""" & Login_Language & """>"
        	  .write "<INPUT TYPE=""HIDDEN"" NAME=""BackURL"" VALUE=""" & BackURL & """>" & vbCrLf
            
            Call Nav_Border_Begin
            
            .write "<TABLE WIDTH=""100%"" BORDER=0 BORDERCOLOR=""GRAY"" CELLPADDING=4 CELLSPACING=0 ALIGN=CENTER>" & vbCrLf
            .write "  <TR>" & vbCrLf            
            .write "    <TD BGCOLOR=""" & Contrast & """ VALIGN=Top NOWRAP>"
            .write "      <FONT CLASS=SmallBold>&nbsp;" & Translate("Select Database to Search",Login_Language,conn) & ":"
            .write "    </TD>" & vbCrLf
            .write "    <TD BGCOLOR=""White"" ALIGN=LEFT VALIGN=MIDDLE CLASS=Medium>" & vbCrLf
            
            SQLNavigation = "SELECT Button_Enable FROM Navigation WHERE Site_ID=" & Site_ID & " AND Button=9003 AND Button_Enable=" & CInt(True)
            Set rsNavigation = Server.CreateObject("ADODB.Recordset")
            rsNavigation.Open SQLNavigation, conn, 3, 3
            if not rsNavigation.EOF then
              .write "      <INPUT TYPE=""RADIO"" NAME=""CINN"" VALUE=""1"" "
                            if SearchDB = 0 or SearchDB = 1 or SearchDB = 999 then .write " CHECKED"
              .write "      >&nbsp;&nbsp;<SPAN CLASS=Small>" &  Translate(Site_Description,Login_Language,conn) & " - " & Translate("General Site Search",Login_Language,conn) & "</SPAN><BR>" & vbCrLf
            end if
            rsNavigation.close
            set rsNavigation = nothing
            
            if Login_Access >= 4 and Login_Region <> 2 and Site_ID <> 17 and Site_ID <> 11 and Site <> 25 then
              .write "      <INPUT TYPE=""RADIO"" NAME=""CINN"" VALUE=""2"" "
              if SearchDB = 2 then .write " CHECKED"
              if Shopping_Cart = CInt(false) then .write " DISABLED"
              .write "      >&nbsp;&nbsp;<SPAN CLASS=Small>" &  Translate("Literature",Login_Language,conn) & "</SPAN>"
              'if Shopping_Cart = CInt(false) then
              '  .write "      &nbsp;&nbsp;<SPAN CLASS=Smallest>(" & Translate("Currently under development",Login_Language,conn) & ")</SPAN>"
              'end if
              .write "      <BR>" & vbCrLf
            end if
              
'            .write "      <INPUT DISABLED TYPE=""RADIO"" NAME=""CINN"" VALUE=""3"" "
'                            if SearchDB = 3 then .write " CHECKED"
'            .write        ">&nbsp;&nbsp;<SPAN CLASS=Small>" &  Translate("Manuals",Login_Language,conn) & "</SPAN>"
'            .write "      &nbsp;&nbsp;<SPAN CLASS=Smallest>(" & Translate("Currently under development",Login_Language,conn) & ")</SPAN>"
'            .write "      <BR>" & vbCrLf
             .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.open('http://www.fluke.com/products/manuals.asp?AppGrp_ID=24','_blank')"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Manuals",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf

            'Special Redirects
            select case Site_ID
              case 11  ' Met-Support-Gold Site
                .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='/met-support-gold/Default.asp?Site_ID=11&NS=False&CID=9003&SCID=0&PCID=0&CIN=22&CINN=146'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Met/Cal Procedures",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf
                .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='/met-support-gold/Default.asp?Site_ID=11&NS=False&CID=9003&SCID=0&PCID=0&CIN=27&CINN=200'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Portocal II Procedures",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf
                .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='/met-support-gold/Default.asp?Site_ID=11&NS=False&CID=9003&SCID=0&PCID=0&CIN=33&CINN=253'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("User Contributed Procedures",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf
              case 17  ' Service-Center Site
                '.write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='/sw-common/SvcIndex_Form.asp'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Service Index",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf
                .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='/sw-common/msg_procedure_form.asp?lv=gold'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Met/Cal Procedures",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf
                .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='/sw-common/portocal_procedure_form.asp'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Portocal II Procedures",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf                
            end select
            
            ' Order Inquiry
            SQLOrder = "SELECT Button_Enable, Button_URL FROM Navigation WHERE Button=9007 AND Button_Enable=" & CInt(True) & " AND Site_ID=" & Site_ID
            Set rsOrder = Server.CreateObject("ADODB.Recordset")
            rsOrder.Open SQLOrder, conn, 3, 3
            if not rsOrder.EOF then
              .write "<INPUT TYPE=""RADIO"" VALUE=""999"" NAME=""CINN"" ONCLICK=""window.location.href='" & rsOrder("Button_URL") & "'"">&nbsp;&nbsp;<SPAN CLASS=Small>" & Translate("Order Inquiry",Login_Language,conn) & "</SPAN></OPTION><BR>" & vbCrLf
            end if
            rsOrder.Close
            set rsOrder = nothing

            .write "     </TD>" & vbCrLf        
            .write "   </TR>" & vbCrLf

            .write "  <TR>" & vbCrLf
            .write "    <TD BGCOLOR=""" & Contrast & """ VALIGN=""MIDDLE"" NOWRAP>"
            .write "      <SPAN CLASS=SmallBold>&nbsp;" & Translate("Enter Keywords for Search",Login_Language,conn) & ":&nbsp;<SPAN>"
            .write "    </TD>" & vbCrLf
            .write "    <TD BGCOLOR=""White"" ALIGN=LEFT VALIGN=MIDDLE CLASS=Medium WIDTH=""75%"">" & vbCrLf
            .write "      &nbsp;&nbsp;<INPUT CLASS=Small TYPE=""Text"" NAME=""KeySearch"" SIZE=""35"" VALUE=""" & Replace(KeySearch,","," ") & """ MAXLENGTH=""255"" VALUE="""">"
            .write "      &nbsp;&nbsp;<SELECT CLASS=Small NAME=""BolSearch"">" & vbCrLf
            .write "      <OPTION CLASS=Small VALUE=""0"""
                          if BolSearch = 0 then response.write " SELECTED"
            .write "      >" & Translate("All Words",Login_Language,conn) & "</OPTION>" & vbCrLf
            .write "      <OPTION CLASS=Small VALUE=""1"""
                          if BolSearch = 1 then response.write " SELECTED"            
            .write "      >" & Translate("Any Word",Login_Language,conn) & "</OPTION>" & vbCrLf
            .write "      <OPTION CLASS=Small VALUE=""2"""
                          if BolSearch = 2 then response.write " SELECTED"            
            .write "      >" & Translate("Exact Phrase",Login_Language,conn) & "</OPTION>" & vbCrLf
            .write "      </SELECT>" & vbCrLf
            
            .write "      &nbsp;&nbsp;<SPAN CLASS=SMALL>" & Translate("Sort by",Login_Language,conn) & ":</SPAN> "
            .write "      <SELECT NAME=""SortBy"" CLASS=Small>" & vbCrLf
            .write "      <OPTION CLASS=SMALL VALUE=""0"""          
                          if SortBy = 0 then response.write " SELECTED"
            .write "      >" & Translate("Product",Login_Language,conn) & "</OPTION>" & vbCrLf                      
            .write "      <OPTION CLASS=SMALL VALUE=""1"""          
                          if SortBy = 1 then response.write " SELECTED"
            .write "      >" & Translate("Category",Login_Language,conn) & "</OPTION>" & vbCrLf                      
            .write "      <OPTION CLASS=SMALL VALUE=""2"""          
                          if SortBy = 2 then response.write " SELECTED"
            .write "      >" & Translate("Date",Login_Language,conn) & "</OPTION>" & vbCrLf                      
            .write "      </SELECT>&nbsp;" & vbCrLf
            
            .write "      &nbsp;&nbsp;<INPUT CLASS=NavLeftHighlight1 TYPE=""Submit"" NAME=""Go"" VALUE="" " & Translate("GO",Login_Language,conn) & " "" onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
            .write "    </TD>" & vbCrLf
            .write "  </TR>" & vbCrLf
            .write "</TABLE>" & vbCrLf
            
            Call Nav_Border_End
            
            .write "</FORM>" & vbCrLf            

            .write "<TABLE WIDTH=""100%"" BORDER=0 BORDERCOLOR=""GRAY"" CELLPADDING=4 CELLSPACING=0 ALIGN=CENTER>" & vbCrLf
            .write "   <TR>" & vbCrLf
            .write "     <TD COLSPAN=2 CLASS=Small>" & vbCrLf
            .write "       <P><BR><BR>" & vbCrLf
            .write "       <SPAN CLASS=SmallBold>" & Translate("Tips for Obtaining the Best Search Results",Login_Language,conn) & "</SPAN><P>" & vbCrLf
            .write "       <OL>" & vbCrLf
            .write "       <LI>" & Translate("Use more than one keyword, and then check for correct spelling.",Login_Language,conn) & "<P>" & vbCrLf
            .write "       <SPAN CLASS=SmallBold><UL><LI>" & Translate("Good examples:",Login_Language,conn) & "</SPAN><BR>" & vbCrLf 
            .write "          Catalog Accessories Clip" & "<BR>" & vbCrLf
            .write "          Application Note DMM" & "</LI><P>" & vbCrLf
            .write "       <SPAN CLASS=SmallBold><LI>" & Translate("Bad examples:",Login_Language,conn) & "</SPAN><BR>" & vbCrLf 
            .write "          Do you have a example of how to use aligator clips? " & Translate("(too specific; too many keywords; may return no results)",Login_Language,conn) & "<BR>" & vbCrLf
            .write "          accessory " & Translate("(too general; will return too many results)",Login_Language,conn) & "</LI></UL><P>" & vbCrLf
            .write "       <LI>" & Translate("Find and Use the Commonly-Used Terms.",Login_Language,conn) & "</LI><BR>" & vbCrLf
            .write "       <LI>" & Translate("When you are reading the search results, look for commonly-used terms, and then use them in your search.",Login_Language,conn) & "</LI><BR>" & vbCrLf
            .write "       <LI>" & Translate("Avoid the Use of Exact Text.",Login_Language,conn) & "<BR>" & vbCrLf
            .write "       <LI>" & Translate("Avoid the plural forms of the keyword such as 'Accessories'. Using the singular form 'Accessory' or with wildcard, 'Accessor%' will return more results.",Login_Language,conn) & "<BR>" & vbCrLf
            .write "       <LI>" & Translate("When you type the exact text that is provided in the product's Help or the exact text from an error message, you may not receive any results. Instead, use a few of the keywords, not the exact text.",Login_Language,conn) & "</LI><BR>" & vbCrLf
            .write "       <LI>" & Translate("Wildcards can be used within the keywords. A '%' (percent)is used to denote a string of zero or more characters. An '_' (underscore) represents a single character.",Login_Language,conn) & "</LI>" & vbCrLf
            .write "       <LI>" & Translate("With the exception of keywords containing wildcards, keywords found will appear in red text to help you see what your kewword matches are as stand alone or embedded matches.",Login_Language,conn) & "</LI>" & vbCrLf
            .write "       <LI>" & Translate("To begin a new search or modify your old search, click on the [Search] navigation button on the left side of your screen..",Login_Language,conn) & "</LI>" & vbCrLf            
            .write "       </OL>" & vbCrLf
            .write "      </TD>" & vbCrLf            
            .write "    </TR>" & vbCrLf
            .write "  </TABLE>" & vbCrLf             

          end with
%>