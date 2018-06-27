<%

  Call WhatsNew_Filter

  SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.Open SQL, conn, 3, 3

  WhatsNewFlag = False
  
  do while not rsCategory.EOF

    SQL  = "SELECT Calendar.*, " &_
              "Literature_Items_US.Item AS Lit_Item, " &_            
              "Literature_Items_US.COST_CENTER AS Cost_Center, " &_
              "Literature_Items_US.STATUS AS Lit_Status, " &_
              "Literature_Items_US.STATUS_NAME AS Lit_Status_Name, " &_                              
              "Literature_Items_US.[ACTION] AS Lit_Action, " &_
              "Literature_Items_US.Revision AS Lit_Revision, " &_              
              "Literature_Items_US.Inventory_Rule AS Lit_Inventory_Rule, " &_              
              "Literature_Items_US.SMALL_LIMIT AS Lit_Small_Limit, " &_
              "Literature_Items_US.LARGE_LIMIT AS Lit_Large_Limit, " &_
              "Literature_Items_US.LITERATURE_TYPE AS Lit_Type, " &_
              "Literature_Items_US.LIT_CODE as Lit_Code, " &_
              "Literature_Items_US.PSize as Lit_PSize, " &_
              "Literature_Items_US.COLORS as Lit_Colors, " &_            
              "Literature_Items_US.UOM AS Lit_UoM, " &_
              "Literature_Items_US.[Print] AS Lit_Print, " &_                                        
              "Literature_Items_US.POD AS Lit_POD, " &_
              "Literature_Items_US.CD AS Lit_CD, " &_
              "Literature_Items_US.Display AS Lit_Display, " &_
              "Literature_Items_US.Video_NTSC AS Lit_Video_NTSC, " &_
              "Literature_Items_US.Video_PAL AS Lit_Video_PAL, " &_
              "Literature_Items_US.END_USER AS Lit_End_User, " &_
              "Literature_Items_US.CUSTOMER_ORDER AS Lit_C_Order, " &_
              "Literature_Items_US.INTERNAL_ORDER AS Lit_I_Order, " &_
              "Literature_Items_US.EFULFILLMENT AS Lit_Description, " &_
              "Literature_Items_US.[LANGUAGE] AS Lit_Language " &_
          "FROM dbo.Calendar LEFT OUTER JOIN " &_
              "dbo.Literature_Items_US ON dbo.Calendar.Revision_Code = dbo.Literature_Items_US.REVISION AND " &_
              "dbo.Calendar.Item_Number = dbo.Literature_Items_US.ITEM LEFT OUTER JOIN " &_
              "dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code AND dbo.Calendar.Site_ID = dbo.Calendar_Category.Site_ID " &_
              "WHERE Calendar.Site_ID=" & CInt(Site_ID) &_
              "   AND Calendar.Category_ID=" & rsCategory("ID")    
    
    ' Build LIKE Clause for Sub-Group Memberships

    if (Access_Level <= 2 or Access_Level = 6) then
      SQL = SQL & " AND (Calendar.SubGroups LIKE '%all%'"
      for i = 0 to UserSubGroups_Max
        SQL = SQL & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
      next
      SQL = SQL & ")"        
    end if   
    
    ' Determine if Active or Pending

    if (Access_Level <= 2 or Access_Level = 6) then
      SQL = SQL & " AND Calendar.Status=1"
      SQL = SQL & " AND ((Calendar.LDate<='" & Date & "' AND Calendar.LDate>='" & DateAdd("d", (Show_Days * -1), Date) & "')"
      SQL = SQL & " OR"
      SQL = SQL & " ( Calendar.BDate<='" & Date & "' AND Calendar.BDate>='" & DateAdd("d", (Show_Days * -1), Date) & "'))"
      
    else ' Show all for Admin
      SQL = SQL & " AND (Calendar.Status=0 OR Calendar.Status=1)"
      SQL = SQL & " AND ((Calendar.LDate<='" & Date & "' AND Calendar.LDate>='" & DateAdd("d", (Show_Days * -1), Date) & "')"
      SQL = SQL & " OR"
      SQL = SQL & " ( Calendar.BDate<='" & Date & "' AND Calendar.BDate>='" & DateAdd("d", (Show_Days * -1), Date) & "'))"
      
    end if
      
    ' Restricted Countries
    
    if (Access_Level <= 2 or Access_Level = 6) then
      SQL = SQL & " AND (Country = 'none'" &_
        " OR (Country LIKE '%0%' AND Country NOT LIKE '%" & Login_Country & "%')" &_
        " OR (Country NOT LIKE '%0%' AND Country LIKE '%" & Login_Country & "%'))"
    end if

    ' Filter to English or Preferred Language for Users

    if (Access_Level = 0 or Access_Level = 6)then
      if LCase(Login_Language) <> "eng" then
        SQL = SQL & " AND (Calendar.Language='eng' OR Calendar.Language='" & Login_Language & "')"
      else  
        SQL = SQL & " AND Calendar.Language='eng'"
      end if
    end if

    ' Sort Order
    
    select case SortBy
      case 1          ' Sub-Category
        SQL = SQL & " ORDER BY Calendar.Sub_Category, Calendar.Product, Calendar.BDATE DESC"          
      case 2          ' Date
        SQL = SQL & " ORDER BY Calendar.BDATE DESC"
      case else       ' Product
        SQL = SQL & " ORDER BY Calendar.Product, Calendar.Sub_Category, Calendar.BDATE DESC"
    end select             

'    response.write SQL & "<P>"
    
    Set rsCalendar = Server.CreateObject("ADODB.Recordset")
    rsCalendar.Open SQL, conn, 3, 3

    if rsCalendar.EOF then
      TableOn = false
    else
      Product  = ""
      Category = ""            
      response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
      TableOn = true
      WhatsNewFlag = True
    end if
    
    Record_Number = 0

    Do while not rsCalendar.EOF

     ' Separator
     
      if Record_Number = 0 then
        response.write "  <TR>" & vbCrLf
        response.write "    <TD COLSPAN=6 HEIGHT=32 VALIGN=MIDDLE CLASS=Normal>"          
        response.write "    <FONT CLASS=Heading3Red>" & Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</FONT>"
        response.write "    </TD>" & vbCrLf            
        response.write "  </TR>" & vbCrLf      
      else
        response.write "  <TR>" & vbCrLf
        response.write "    <TD COLSPAN=6 HEIGHT=16 VALIGN=MIDDLE CLASS=Normal>"          
        response.write "    &nbsp;"
        response.write "    </TD>" & vbCrLf            
        response.write "  </TR>" & vbCrLf
      end if  
    
      Call Update_Fields          
      Call Display_Category_Item
  
      rsCalendar.MoveNext

    loop
                   
    rsCalendar.close
    set rsCalendar = nothing

    if TableOn then
      response.write "</TABLE><BR><BR>" & vbCrLf
    end if

    rsCategory.MoveNext
    
  loop
  
  if not WhatsNewFlag then
    response.write Translate("There have been no New Items Posted to this Site within the past",Login_Language, conn) & " " & Show_Days & " " & Translate("days",Login_Language,conn) & "."
  end if  

  rsCategory.close
  set rsCategory = nothing

' --------------------------------------------------------------------------------------

sub WhatsNew_Filter()
   
  response.write "<FORM NAME=""Whats_New"">"
  response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
  response.write "  <TR>" & vbCrLf  
 
  ' Filter Menu
  
  response.write "    <TD Width=""50%"" HEIGHT=16 CLASS=Small ALIGN=LEFT VALIGN=TOP NOWRAP>"

  Call Nav_Border_Begin
  
  response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Items Posted within the Past",Login_Language,conn) & "&nbsp;" & vbCrLf
  response.write "      <SELECT NAME=""Show_Days"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&Show_Detail=" & Show_Detail & "&SortBy=" & SortBy & "&Show_Days='+this.options[this.selectedIndex].value"">" & vbCrLf                        

  response.write "        <OPTION VALUE=""7"""          
  if Show_Days = 7 then response.write " SELECTED"
  response.write ">7</OPTION>" & vbCrLf                      

  response.write "        <OPTION VALUE=""14"""          
  if Show_Days = 14 then response.write " SELECTED"
  response.write ">14</OPTION>" & vbCrLf                                                                      

  response.write "        <OPTION VALUE=""30"""          
  if Show_Days = 30 then response.write " SELECTED"
  response.write ">30</OPTION>" & vbCrLf

  response.write "        <OPTION VALUE=""45"""          
  if Show_Days = 45 then response.write " SELECTED"
  response.write ">45</OPTION>" & vbCrLf
 
  response.write "        <OPTION VALUE=""60"""          
  if Show_Days = 60 then response.write " SELECTED"
  response.write ">60</OPTION>" & vbCrLf

  response.write "      </SELECT>" & vbCrLf
 
  response.write "&nbsp;" & Translate("Days",Login_Language,conn)
  response.write "</SPAN>"

  Call Nav_Border_End

  response.write "    </TD>" & vbCrLf
  
  ' Sort By
  
  response.write "    <TD COLSPAN=2 Width=""25%"" HEIGHT=16 CLASS=SMALL ALIGN=RIGHT VALIGN=TOP NOWRAP>"

  Call Nav_Border_Begin
  
  response.write "<SPAN CLASS=SmallBoldGold>"
  response.write Translate("Sort",Login_Language,conn) & ":&nbsp;" & vbCrLf
  response.write "      <SELECT NAME=""SortBy"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&Show_Detail=" & Show_Detail & "&Show_Days=" & Show_Days & "&SortBy='+this.options[this.selectedIndex].value"">" & vbCrLf
 
  response.write "<OPTION VALUE=""0"""          
  if SortBy = 0 then response.write " SELECTED"
  response.write ">" & Translate("Product",Login_Language,conn) & "</OPTION>" & vbCrLf                      
 
  response.write "<OPTION VALUE=""1"""          
  if SortBy = 1 then response.write " SELECTED"
  response.write ">" & Translate("Category",Login_Language,conn) & "</OPTION>" & vbCrLf                      
  
  response.write "<OPTION VALUE=""2"""          
  if SortBy = 2 then response.write " SELECTED"
  response.write ">" & Translate("Date",Login_Language,conn) & "</OPTION>" & vbCrLf                      
  
  response.write "</SELECT>" & vbCrLf
  response.write "</SPAN>"
  
  Call Nav_Border_End
  
  response.write "    </TD>" & vbCrLf

  ' Show Active Tite or Active Detail
  
  response.write "    <TD WIDTH=""25%"" HEIGHT=16 CLASS=Small ALIGN=RIGHT VALIGN=TOP NOWRAP>"
  
  Call Nav_Border_Begin
    
  response.write "      <SPAN CLASS=SmallBoldGold>" & Translate("View",Login_Language,conn) & ":&nbsp;" & vbCrLf
  response.write "      <SELECT NAME=""Show_Detail"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&SortBy=" & SortBy & "&Show_Days=" & Show_Days & "&Show_Detail='+this.options[this.selectedIndex].value"">" & vbCrLf                      
 
  response.write "        <OPTION VALUE=""-1"""          
  if Show_Detail = -1 then response.write " SELECTED"
  response.write ">" & Translate("Active",Login_Language,conn) & " + " & Translate("Detail",Login_Language,conn) & "</OPTION>" & vbCrLf                      
 
  response.write "        <OPTION VALUE=""1"""          
  if Show_Detail = 1 then response.write " SELECTED"
  response.write ">" & Translate("Active",Login_Language,conn) & " + " & Translate("Title",Login_Language,conn) & "</OPTION>" & vbCrLf                                                                      
 
  response.write "      </SELECT>" & vbCrLf
  response.write "      </SPAN>"
  
  Call Nav_Border_End

  response.write "    </TD>" & vbCrLf    

  response.write "  </TR>" & vbCrLf  
  response.write "</TABLE>" & vbCrLf
  response.write "</FORM>"
  response.write "<BR>"

end sub

%>