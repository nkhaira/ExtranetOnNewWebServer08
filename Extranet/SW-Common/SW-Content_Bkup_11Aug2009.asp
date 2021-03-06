<%
if CID = 9004 then  ' Search

        
else                ' All Other
        
  ' Build Library Category Title
  
  SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=-1" & " AND Calendar_Category.Code=" & CIN

  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.Open SQL, conn, 3, 3

  KeyFlag = false
  if not rsCategory.EOF then
    KeyFlag = true
    response.write "<FONT CLASS=Heading3>" & Translate(rsCategory("Title"),Login_Language,conn) & "</FONT>"
    if not isblank(rsCategory("Description")) then
      response.write "<BR><BR><FONT CLASS=Normal>" & Translate(rsCategory("Description"),Login_Language,conn) & "</FONT>"
    end if
    response.write "<BR><BR>"

    Title_View = CInt(rsCategory("Title_View"))   ' Appends Sub-Category to Title

  end if

  rsCategory.close
  set rsCategory = nothing
          
end if  
  
' --------------------------------------------------------------------------------------

if CID = 9004 and CINN = 1 Then    ' Search Partner Portal
  
  if bolSearch <> 2 then      ' 2 = Exact Phrase
    KeySearch = replace(replace(Trim(keySearch)," ",","),"'","")
    KeySearch = replace(replace(replace(KeySearch,",,,,",","),",,,",","),",,",",")
  end if  
         
  SQL     = "SELECT Calendar.*, " &_
            "Literature_Items_US.Item AS Lit_Item, " &_        
            "Literature_Items_US.COST_CENTER AS Cost_Center, " &_
            "Literature_Items_US.STATUS AS Lit_Status, " &_
            "Literature_Items_US.STATUS_NAME AS Lit_Status_Name, " &_            
            "Literature_Items_US.[ACTION] AS Lit_Action, " &_
            "Literature_Items_US.Revision AS Lit_Revision, " &_
            "Literature_Items_US.ACTIVE_FLAG AS Lit_Active_Flag, " &_            
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
            "Literature_Items_US.[LANGUAGE] AS Lit_Language, " &_
            "(Select isnull(count(*),0) from Asset_accesscodes where assetid = calendar.id) as CheckCode " &_
            "FROM   Calendar " &_
            "   LEFT OUTER JOIN " &_
                    "Literature_Items_US ON Calendar.Item_Number = Literature_Items_US.ITEM AND Calendar.Revision_Code = Literature_Items_US.Revision " &_
            "   LEFT OUTER JOIN " &_
                    "Calendar_Category ON Calendar.Code = Calendar_Category.Code AND Calendar.Site_ID = Calendar_Category.Site_ID " &_         
            "WHERE (Calendar.Site_ID = " & Site_ID & ") "
 
  if instr(1,KeySearch,",") > 0 then
    strSearch     = Split(KeySearch,",")
  else
    ReDim strSearch(0)
    strSearch(0)  = Trim(KeySearch)
  end if
     
  ItemSearch = False
  for x = 0 to ubound(strSearch)
    if Len(Trim(strSearch(x))) = 7 and isnumeric(strSearch(x)) then
      ItemSearch = True
      exit for
    end if
  next
  sqlTemp = "       convert(varchar(50),   COALESCE(Calendar_Category.Category, '')) + " &_
            "       convert(varchar(255),  COALESCE(Calendar.Sub_Category, '')) + " &_
            "       convert(varchar(255),  COALESCE(Calendar.Product, '')) + " &_
            "       convert(varchar(255),  COALESCE(Calendar.Title, '')) + " &_
            "       convert(varchar(10),   COALESCE(Calendar.Item_Number, '')) + " &_                    
            "       convert(varchar(10),   COALESCE(Literature_Items_US.Item, '')) + " &_                    
            "       convert(varchar(150),  COALESCE(Literature_Items_US.Literature_Type, '')) + " &_
            "       convert(varchar(2000), COALESCE(Literature_Items_US.Description, '')) + " &_                    
            "       convert(varchar(2000), COALESCE(Literature_Items_US.Efulfillment, '')) + " &_                    
            "       convert(varchar(2000), COALESCE(Calendar.Description, '')) + " &_
            "       'Asset ID: ' + convert(varchar(14),   COALESCE(Calendar.ID, '')) "
  if ItemSearch = True then          
    sqlTemp = sqlTemp & " + convert(varchar(50),   COALESCE(Calendar.Item_Number, '')) "
'    sqlTemp = sqlTemp & " + convert(varchar(50),   COALESCE(Calendar.Item_Number, ''))  "    
  end if  

  for x = 0 to ubound(strSearch)
    select case x
      case 0
        SQL = SQL & " AND (("
      case else
        if BolSearch = 1 then
          SQL = SQL & " OR ("
        else
          SQL = SQL & " AND ("
        end if                              
    end select
    SQL = SQL & sqlTemp & " LIKE '%" & strSearch(x) & "%') "
  next
  SQL = SQL & ") "

elseif CID = 9004 and CINN = 2 Then    ' Search Literature
          
  if bolSearch <> 2 then      ' 2 = Exact Phrase
    KeySearch = replace(replace(Trim(keySearch)," ",","),"'","")
    KeySearch = replace(replace(replace(KeySearch,",,,,",","),",,,",","),",,",",")
  end if  
  
        SQL = "SELECT Calendar.*, " &_
            "Literature_Items_US.Item AS Lit_Item, " &_
            "Literature_Items_US.COST_CENTER AS Cost_Center, " &_
            "Literature_Items_US.STATUS AS Lit_Status, " &_
            "Literature_Items_US.STATUS_NAME AS Lit_Status_Name, " &_                
            "Literature_Items_US.[ACTION] AS Lit_Action, " &_
            "Literature_Items_US.Revision AS Lit_Revision, " &_
            "Literature_Items_US.ACTIVE_FLAG AS Lit_Active_Flag, " &_                        
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
            "FROM   dbo.Lit_Cost_Center " &_
            "LEFT OUTER JOIN dbo.Literature_Items_US ON dbo.Lit_Cost_Center.Cost_Center = dbo.Literature_Items_US.COST_CENTER " &_
            "LEFT OUTER JOIN dbo.Calendar " &_
            "LEFT OUTER JOIN dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code " &_
            "AND dbo.Calendar.Site_ID = dbo.Calendar_Category.Site_ID ON dbo.Literature_Items_US.ITEM = dbo.Calendar.Item_Number " &_
    "WHERE  (Lit_Cost_Center.Site_ID = " & Site_ID & ") " &_
            "AND Literature_Items_US.Status = 'Active' " &_
            "AND (Literature_Items_US.Action = 'Complete' " &_
            "OR (Literature_Items_US.Action = 'Incomplete' AND Literature_Items_US.Revision <> 'A')) "

  if Login_Type_Code <> 5  then ' Fluke
    SQL = SQL & "AND Literature_Items_US.Customer_Order = " & CInt(True) & " "
  end if  
 
  if instr(1,KeySearch,",") > 0 then
    strSearch     = Split(KeySearch,",")
  else
    ReDim strSearch(0)
    strSearch(0)  = Trim(KeySearch)
  end if
     
  ItemSearch = False
  for x = 0 to ubound(strSearch)
    if Len(Trim(strSearch(x))) = 7 and isnumeric(strSearch(x)) then
      ItemSearch = True
      exit for
    end if
  next
  sqlTemp = "       convert(varchar(50),   COALESCE(Calendar_Category.Category, '')) + " &_
            "       convert(varchar(255),  COALESCE(Calendar.Sub_Category, '')) + " &_
            "       convert(varchar(255),  COALESCE(Calendar.Product, '')) + " &_
            "       convert(varchar(255),  COALESCE(Calendar.Title, '')) + " &_
            "       convert(varchar(10),   COALESCE(Calendar.Item_Number, '')) + " &_                    
            "       convert(varchar(10),   COALESCE(Literature_Items_US.Item, '')) + " &_                    
            "       convert(varchar(150),  COALESCE(Literature_Items_US.Literature_Type, '')) + " &_
            "       convert(varchar(2000), COALESCE(Literature_Items_US.Description, '')) + " &_
            "       convert(varchar(2000), COALESCE(Literature_Items_US.Efulfillment, '')) + " &_                    
            "       convert(varchar(2000), COALESCE(Calendar.Description, '')) "
  if ItemSearch = True then          
    sqlTemp = sqlTemp & " + convert(varchar(50),   COALESCE(Calendar.Item_Number, '')) "
  end if  
            
  for x = 0 to ubound(strSearch)
    select case x
      case 0
        SQL = SQL & " AND (("
      case else
        if BolSearch = 1 then
          SQL = SQL & " OR ("
        else
          SQL = SQL & " AND ("
        end if                              
    end select
    SQL = SQL & sqlTemp & " LIKE '%" & strSearch(x) & "%') "
  next
  SQL = SQL & ") "
  
' --------------------------------------------------------------------------------------

elseif CID <> 9004 and CINN > 0 then    ' Not Search and Category has been selected
            
  if CInt(Shopping_Cart) = CInt(True) then
    ' Add Literature Shopping Cart Data to Recordset for non-European accounts
    SQL = "SELECT Calendar.*, " &_
            "Literature_Items_US.Item AS Lit_Item, " &_
            "Literature_Items_US.COST_CENTER AS Cost_Center, " &_
            "Literature_Items_US.STATUS AS Lit_Status, " &_
            "Literature_Items_US.STATUS_NAME AS Lit_Status_Name, " &_                        
            "Literature_Items_US.[ACTION] AS Lit_Action, " &_
            "Literature_Items_US.Revision AS Lit_Revision, " &_
            "Literature_Items_US.ACTIVE_FLAG AS Lit_Active_Flag, " &_                        
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
            "Literature_Items_US.[LANGUAGE] AS Lit_Language, " &_
            "(Select isnull(count(*),0) from Asset_accesscodes where assetid = calendar.id) as CheckCode " &_
            "FROM Literature_Items_US " &_
            "RIGHT OUTER JOIN " &_
            "      dbo.Calendar ON dbo.Literature_Items_US.REVISION = dbo.Calendar.Revision_Code AND " &_
            "      dbo.Literature_Items_US.ITEM = dbo.Calendar.Item_Number " &_
            "WHERE Calendar.Site_ID=" & CInt(Site_ID) &" AND Calendar.Category_ID=" & CInt(CINN)
  else

   Bypass_Active_Flag = true  
   SQL = "SELECT Calendar.*, NULL as Lit_PSize, NULL AS Lit_Colors, Null as Lit_Code, Item_Number AS Lit_Description, Item_Number AS Lit_Item, " &_
   "(Select isnull(count(*),0) from Asset_accesscodes where assetid = calendar.id) as CheckCode FROM Calendar WHERE Calendar.Site_ID=" & CInt(Site_ID) &" AND Calendar.Category_ID=" & CInt(CINN)

  end if  

end if

' --------------------------------------------------------------------------------------        

if CID = 9004 or CINN > 0 then  ' Search or Category Selected
  
  if CID <> 9004 or (CID = 9004 and CINN <> 2) then
  
    ' Build LIKE Clause for Sub-Group Memberships

    if (Access_Level <= 2 or Access_Level = 6) then
      SQL = SQL & " AND (Calendar.SubGroups LIKE '%all%'"        
      for i = 0 to UserSubGroups_Max
        SQL = SQL & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
      next
      SQL = SQL & ")"        
    end if   
    
    ' Determine if Active or Archive
    
    if (Access_Level <= 2 or Access_Level = 6) then
      if abs(Show_Detail) = 0 or abs(Show_Detail) = 1 then
        SQL = SQL & " AND Calendar.Status=1 AND ((Calendar.LDate<='" & Date & "' AND Calendar.XDAYS=0) OR (Calendar.LDate<='" & Date & "' AND Calendar.XDate>'" & Date & "'))"
      elseif abs(Show_Detail) = 2 then  ' Archive
        SQL = SQL & " AND (Calendar.Status=2 OR (Calendar.XDays=0 AND '" & Date & "'>Calendar.XDate))"
      end if
    else ' Show all for Admin
      if abs(Show_Detail) = 1 then
        SQL = SQL & " AND (Calendar.Status=0 OR Calendar.Status=1) AND (Calendar.XDAYS=0 OR Calendar.XDate>'" & Date & "')"            
      else            
        SQL = SQL & " AND (Calendar.Status=" & abs(Show_Detail) & " OR (Calendar.XDAYS>0 AND '" & Date & "'>Calendar.XDate))"
      end if
    end if

   ' Restricted Countries

    if (Access_Level <= 2 or Access_Level = 6) then
      SQL = SQL & " AND (Calendar.Country = 'none'" &_
                  " OR (Calendar.Country LIKE '%0%' AND Calendar.Country NOT LIKE '%" & Login_Country & "%')" &_
                  " OR (Calendar.Country NOT LIKE '%0%' AND Calendar.Country LIKE '%" & Login_Country & "%'))"
    end if
    
    ' Filter for CIN Special Groupings
    
    if CIN < 8000 or CIN > 8999 then
      ' Individual or Individual + Product Introduction or Campaign
      'SQL = SQL & " AND (Calendar.Content_Group=0 or Calendar.Content_Group=1 or Calendar.Content_Group=3)"
    end if
    
  end if   
    
  ' Filter to English or Preferred Language for Users

  if (Access_Level = 0 or Access_Level = 6) and isblank(Language_Filter) then
    if LCase(Login_Language) <> "eng" then
      if CID <> 9004 then
        SQL = SQL & " AND (Calendar.Language='eng' OR Calendar.Language='" & Login_Language & "')"
      elseif CID = 9004 then
        SQL = SQL & " AND ((Calendar.Language='eng' OR Calendar.Language='" & Login_Language & "')"
        SQL = SQL & " OR (Literature_Items_US.[LANGUAGE]='eng' OR Literature_Items_US.[LANGUAGE]='" & Login_Language & "'))"                
      end if  
    else  
      SQL = SQL & " AND Calendar.Language='eng'"
    end if
  elseif not isblank(Language_Filter) then   
    if CID <> 9004 then
      SQL = SQL & " AND (Calendar.Language='" & Language_Filter & "')"
    elseif CID = 9004 then
      SQL = SQL & " AND ((Calendar.Language='" & Language_Filter & "')"
      SQL = SQL & " OR (Literature_Items_US.[LANGUAGE]='" & Language_Filter & "'))"                
    end if  
  end if
  
  if Bypass_Active_Flag = false then
    SQL = SQL & " AND (Literature_Items_US.ACTIVE_FLAG IS NULL OR Literature_Items_US.ACTIVE_FLAG = -1) "  
  end if  
    
  ' Sort Order

  if CID <> 9004 or (CID = 9004 and CINN <> 2) then  
    select case SortBy
      case 1          ' Sub-Category
        SQL = SQL & " ORDER BY Calendar.Sub_Category, Calendar.Product, Calendar.Title"
      case 2          ' Date
        SQL = SQL & " ORDER BY Calendar.BDATE DESC"
      case else       ' Product
        SQL = SQL & " ORDER BY Calendar.Product, Calendar.Sub_Category, Calendar.Title"
    end select
  elseif CID = 9004 and CINN = 2 then  
    select case SortBy
      case 1          ' Sub-Category
        SQL = SQL & " ORDER BY Literature_Items_US.Literature_Type, Calendar.Product, Calendar.Title, Literature_Items_US.Efulfillment"
      case 2          ' Date
        SQL = SQL & " ORDER BY Calendar.BDATE DESC, Literature_Items_US.QA_Update_Date"
      case 3          ' Item Number
        SQL = SQL & " ORDER BY Calendar.Item_Number, Calendar.Revision_Code"
      case else       ' Product
        SQL = SQL & " ORDER BY Calendar.Product, Calendar.Sub_Category, Calendar.Title, Literature_Items_US.Efulfillment, Literature_Items_US.Literature_Type"
    end select
  end if        

  if instr(1,KeySearch,"debug") > 0 then
    response.write SQL & "<P>"
  end if  

  Set rsCalendar = Server.CreateObject("ADODB.Recordset")
  rsCalendar.Open SQL, conn, 3, 3
response.write SQL
  ' No Active Records
  
  if rsCalendar.EOF then
   
    if CID = 9004 then
      if abs(Show_Detail) = 2 then
        if isblank(Language_Filter) then
          response.write Translate("There are no Archive items found based on your Search Key Words.",Login_Language,conn)
        else
          response.write Translate("There are no Archive items found based on your Search Key Words and your &quot;Filter by Language&quot; setting.",Login_Language,conn)
        end if  
      else  
        if isblank(Language_Filter) then
          response.write Translate("There are no Active items found based on your Search Key Words.",Login_Language,conn)              
        else
          response.write Translate("There are no Active items found based on your Search Key Words and your &quot;Filter by Language&quot; setting.",Login_Language,conn)
        end if  
      end if  
    else  
      if abs(Show_Detail) = 2 then
        if isblank(Language_Filter) then
          response.write Translate("There are no Archive items found for this Category.",Login_Language,conn)
        else  
          response.write Translate("There are no Archive items found for this Category based on your &quot;Filter by Language&quot; setting.",Login_Language,conn)
        end if
      else  
        if isblank(Language_Filter) then      
          response.write Translate("There are no Active items found for this Category.",Login_Language,conn)              
        else  
          response.write Translate("There are no Active items found for this Category based on your &quot;Filter by Language&quot; setting.",Login_Language,conn)
        end if
      end if  
    end if
    
    response.write "<P>" & Translate("Click on the [Search] button again to modify your search criteria.",Login_Language,conn)                          

    rsCalendar.close
    set rsCalendar = nothing

    ' Try to provide link to Archive if no Active Content
    if abs(Show_Detail) <> 2 then

      SQL = Replace(SQL,"Calendar.Status=0","Calendar.Status=2")
      SQL = Replace(SQL,"Calendar.Status=1","Calendar.Status=2")
      Set rsCalendar = Server.CreateObject("ADODB.Recordset")         
      rsCalendar.Open SQL, conn, 3, 3
      if not rsCalendar.EOF then
        response.write "<BR><BR><A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=0" & "&CIN=" & CIN & "&CINN=" & CINN & "&SortBy=" & SortBy & "&Show_Detail=-2"">" & Translate("Archive",Login_Language,conn) & "</A>" & vbCrLf
      end if

      rsCalendar.close
      set rsCalendar = nothing

    end if

  ' --------------------------------------------------------------------------------------            
  ' Records
  ' --------------------------------------------------------------------------------------

  else

    Product       = ""
    Category      = ""

    Record_Count = 0
    Old_ID       = 0

    do while not rsCalendar.EOF
      if Old_ID <> rsCalendar("ID") then
        Record_Count = Record_Count + 1
      end if
      rsCalendar.MoveNext
    loop
    rsCalendar.MoveFirst
    
'   Record_Count  = rsCalendar.RecordCount
    Record_Pages  = Record_Count \ Record_Limit
    if Record_Count mod Record_Limit > 0 then Record_Pages = Record_Pages + 1

    Page_QS = "Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&CID=" & CID &  "&SCID=" & SCID & "&CIN=" & CIN & "&CINN=" & CINN & "&SortBy=" & SortBy & "&KeySearch=" & KeySearch & "&BolSearch=" & BolSearch & "&Language_Filter=" & Language_Filter
    xPCID = 1
    Record_Number = 0
    
    ' Quick Find --

    response.write "<FORM NAME=""QuickFind"">" & vbCrLf
    response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
    response.write "<TR>" & vbCrLf

    ' Find Document

    response.write "<TD COLSPAN=7 HEIGHT=16 CLASS=SMALL VALIGN=TOP>" & vbCrLf
    
    Call Nav_Border_Begin
     
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Find",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf
    response.write "<SELECT STYLE=""font-size:8.5pt; font-weight:Normal; color:Black; background:White; text-decoration:none;font-family:'Lucida Console, Courier'"" LANGUAGE=""JavaScript"" ONCHANGE=""if(this.options[this.selectedIndex].value != '') { window.location.href=''+this.options[this.selectedIndex].value } else { alert('" & Translate("You have selected a Category Title and not a Document",Login_Language,conn) & "') }"">" & vbCrLf
    response.write "<OPTION CLASS=Small VALUE=""#_Top"">-- " & Translate("Select Document",Login_Language,conn) & " --</OPTION>" & vbCrLf

    Do while not rsCalendar.EOF

      if Old_ID <> rsCalendar("ID") then

        Call Update_Fields
  
        ' Product or Product Series Title and Sub-Category
  
  '      if Product <> Field_Data(xProduct) and (SortBy = 0 or SortBy = 1) then
  '        if SortBy = 1 and Category <> Field_Data(xSub_Category) then
  '          response.write "<OPTION Class=Region4NavMedium VALUE="""">" & "</OPTION>" & vbCrLf                
  '          response.write "<OPTION Class=Region4NavMedium VALUE="""">" & Field_Data(xSub_Category) & "</OPTION>" & vbCrLf
  '        end if
  '        response.write "<OPTION Class=PRODUCT VALUE="""">" & Field_Data(xProduct) & "</OPTION>" & vbCrLf
  '      end if
        
        if SortBy = 1 and Category <> Field_Data(xSub_Category) then
          response.write "<OPTION Class=Region4NavMedium VALUE="""">" & "</OPTION>" & vbCrLf                
          response.write "<OPTION Class=Region4NavMedium VALUE="""">" & Field_Data(xSub_Category) & "</OPTION>" & vbCrLf
        end if
        if Product <> Field_Data(xProduct) and (SortBy = 0 or SortBy = 1) then
          response.write "<OPTION Class=PRODUCT VALUE="""">" & Field_Data(xProduct) & "</OPTION>" & vbCrLf
        end if
  
        Product  = Field_Data(xProduct)
        Category = Field_Data(xSub_Category)
        
        ' Record
        
        Field_Data(xTitle) = Trim(Field_Data(xTitle))
        
        if not isblank(Field_Data(xTitle)) then
          if Len(Field_Data(xTitle)) > 36 then Field_Data(xTitle) = mid(Field_Data(xTitle),1,36)
        elseif CID = 9004 then
          if Len(rsCalendar("Lit_Description")) > 36 then
            Field_Data(xTitle) = ProperCase(mid(rsCalendar("Lit_Description"),1,36))
          else
            Field_Data(xTitle) = ProperCase(rsCalendar("Lit_Description"))
          end if
        end if
        
        if len(Field_Data(xTitle)) <= 36 then
          Field_Data(xTitle) = Field_Data(xTitle) & Mid("                                         ",1,(36 - Len(Field_Data(xTitle))))
        end if  
        
        if isblank(Field_Data(xItem_Number)) and not isblank(rsCalendar("Lit_Item")) then
        
          Field_Flag(xItem_Number) = True
          Field_Data(xItem_Number) = rsCalendar("Lit_Item")
          Field_Flag(xRevision_Code)    = True
          Field_Data(xRevision_Code)    = rsCalendar("Lit_Revision")
          
          SQLLanguage = "SELECT dbo.Language.Description " &_
                        "FROM   dbo.Literature_Items_US INNER JOIN " &_
                        "       dbo.Language ON dbo.Literature_Items_US.[LANGUAGE] = dbo.Language.Code " &_
                        "WHERE dbo.Literature_Items_US.Item=" & rsCalendar("Lit_Item") & " AND dbo.Language.enable=-1"
  
          'response.write sqllanguage & "<P>"
                        
          Set rsLanguage = Server.CreateObject("ADODB.Recordset")         
          rsLanguage.Open SQLLanguage, conn, 3, 3
          if not rsLanguage.EOF then
            Field_Flag(xLanguage)    = True
            Field_Data(xLanguage)    = rsLanguage("Description")
          else
            Field_Flag(xLanguage)    = True
            Field_Data(xLanguage)    = "---"
          end if
          rsLanguage.close
          set rsLanguage = nothing  
          
        end if        
        
        if CInt(Field_Flag(xItem_Number)) = CInt(true) then
          Field_Data(xTitle) = Field_Data(xTitle) & " | "
          if len(Field_Data(xItem_Number)) = 6 then Field_Data(xTitle) = Field_Data(xTitle) & " "
          Field_Data(xTitle) = Field_Data(xTitle) & UCase(Field_Data(xItem_Number))
          if CInt(Field_Flag(xRevision_Code)) = CInt(true) then
            Field_Data(xTitle) = Field_Data(xTitle) & " " & UCase(Field_Data(xRevision_Code))
          end if
        end if
  
        if CInt(Field_Flag(xItem_Number)) = CInt(False) and not isblank(Field_Data(xLanguage)) then
          Field_Data(xTitle) = Field_Data(xTitle) & " |           | " & UCase(Field_Data(xLanguage))
        elseif not isblank(Field_Data(xLanguage)) then 
          Field_Data(xTitle) = Field_Data(xTitle) & " | " & UCase(Field_Data(xLanguage))      
        end if
        
        If Record_Number >= Record_Limit then
          xPCID = xPCID + 1
          Record_Number = 1
        else
          Record_Number = Record_Number + 1
        end if
        
        Field_Data(xTitle) = Replace(Field_Data(xTitle)," ","&nbsp;")
	response.write "Checkcode is " & rsCalendar("CheckCode")
        if rsCalendar("CheckCode").value = 0 then
            response.write "<OPTION VALUE=""default.asp?" & Page_QS & "&PCID=" & xPCID & "#" & Field_Data(xID) & """>&nbsp;"  & Field_Data(xTitle) & "</OPTION>" & vbCrLf 
        else
           set rsCheckPriceList = conn.execute("exec PriceList_CheckAccessCode " & rsCalendar("ID") & ",'" & _
           Session("PriceListCode") & "'")
           if trim(rsCheckPriceList.fields(0).value) ="True" then
              response.write "<OPTION VALUE=""default.asp?" & Page_QS & "&PCID=" & xPCID & "#" & Field_Data(xID) & """>&nbsp;"  & Field_Data(xTitle) & "</OPTION>" & vbCrLf 
           end if
           set rsCheckPriceList = nothing
        end if        
      end if
      
      Old_ID = rsCalendar("ID")
        
      rsCalendar.MoveNext
      
    loop

    response.write "        </SELECT>" & vbCrLf
    
    ' Sort By
    
    response.write "&nbsp;&nbsp;&nbsp;"
    
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Sort",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf         
    response.write "<SELECT NAME=""SortBy"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & "&KeySearch=" & KeySearch & "&BolSearch=" & BolSearch & "&Show_Detail=" & Show_Detail & "&Language_Filter=" & Language_Filter & "&SortBy='+this.options[this.selectedIndex].value"">" & vbCrLf

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
    
    Call Nav_Border_End
    
    response.write "      </TD>" & vbCrLf
    response.write "      </TR>" & vbCrLf    

    ' Page Navigation

    response.write "<TR>"  & vbCrLf
    response.write "<TD COLSPAN=7 HEIGHT=16 CLASS=Medium BGCOLOR=WHITE>"  & vbCrLf
    
    Call RS_Page_Navigation
    
    response.write "</TD>"  & vbCrLf
    response.write "</TR>"  & vbCrLf

    ' View
    
    response.write "<TR>"  & vbCrLf
    response.write "<TD COLSPAN=7 HEIGHT=16 CLASS=Medium BGCOLOR=WHITE>"  & vbCrLf

    Call Nav_Border_Begin

    if Access_Level >= 4 then

      response.write "<SPAN CLASS=SmallBoldGold>" & Translate("View",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf
      response.write "<SELECT NAME=""Show_Detail"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=0" & "&CIN=" & CIN & "&CINN=" & CINN & "&KeySearch=" & KeySearch & "&BolSearch=" & BolSearch & "&SortBy=" & SortBy & "&Language_Filter=" & Language_Filter & "&Show_Detail='+this.options[this.selectedIndex].value"">" & vbCrLf                      
  
      response.write "<OPTION VALUE=""-1"""          
      if Show_Detail = -1 then response.write " SELECTED"
      response.write ">" & Translate("Active",Login_Language,conn) & " + " & Translate("Detail",Login_Language,conn) & "</OPTION>" & vbCrLf                      
  
      response.write "<OPTION VALUE=""1"""          
      if Show_Detail = 1 then response.write " SELECTED"
      response.write ">" & Translate("Active",Login_Language,conn) & " + " & Translate("Title",Login_Language,conn) & "</OPTION>" & vbCrLf                      
                                                      
      response.write "<OPTION VALUE=""-2"""          
      if Show_Detail = -2 then response.write " SELECTED"
      response.write ">" & Translate("Archive",Login_Language,conn) & " + " & Translate("Detail",Login_Language,conn) & "</OPTION>" & vbCrLf
  
      response.write "<OPTION VALUE=""2"""          
      if Show_Detail = 2 then response.write " SELECTED"
      response.write ">" & Translate("Archive",Login_Language,conn) & " + " & Translate("Title",Login_Language,conn) & "</OPTION>" & vbCrLf
  
      response.write "</SELECT>&nbsp;&nbsp;" & vbCrLf

    end if
    
    ' Document Language Filter
    
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Filter by Language",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf
    response.write "<SELECT NAME=""Language_Filter"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='" & HomeURL & "?Site_ID=" & Site_ID & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & SCID & "&PCID=0" & "&CIN=" & CIN & "&CINN=" & CINN & "&KeySearch=" & KeySearch & "&BolSearch=" & BolSearch & "&SortBy=" & SortBy & "&Show_Detail=" & Show_Detail & "&Language_Filter='+this.options[this.selectedIndex].value"">" & vbCrLf                            

  '  SQLLanguageFilter = "SELECT DISTINCT dbo.Literature_Items_US.[LANGUAGE], dbo.[Language].Code, dbo.[Language].Sort, dbo.[Language].Description " &_
   '                     "FROM            dbo.Literature_Items_US INNER JOIN " &_
    '                    "                dbo.[Language] ON dbo.Literature_Items_US.[LANGUAGE] = dbo.[Language].Code AND dbo.[Language].enable=-1" &_
     '                   "ORDER BY dbo.[Language].Sort"
     
       SQLLanguageFilter = "SELECT DISTINCT dbo.Literature_Items_US.[LANGUAGE], dbo.[Language].Code, dbo.[Language].Sort, dbo.[Language].Description " &_
                       "FROM        dbo.[Language]  LEFT  JOIN " &_
                        "                 dbo.Literature_Items_US  ON dbo.Literature_Items_US.[LANGUAGE] = dbo.[Language].Code Where dbo.[Language].enable=-1" &_
                        "ORDER BY dbo.[Language].Sort"
                   
    Set rsLanguageFilter = Server.CreateObject("ADODB.Recordset")         
    rsLanguageFilter.Open SQLLanguageFilter, conn, 3, 3
                        
    response.write "<OPTION CLASS=""RegionXNavSmall"" VALUE="""">" & Translate("None",Login_Language,conn) & "</OPTION>" & vbCrLf
      
    do while not rsLanguageFilter.EOF
  
      response.write "<OPTION CLASS=""Region5NavSmall"" VALUE=""" & rsLanguageFilter("Code") & """"
      if LCase(Language_Filter) = LCase(rsLanguageFilter("Code")) then
        response.write " SELECTED"
      end if  
      response.write ">" & Translate(rsLanguageFilter("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
      
      rsLanguageFilter.MoveNext
      
    loop  
  
    response.write "</SELECT>&nbsp;&nbsp;" & vbCrLf         
    Call Nav_Border_End

    rsLanguageFilter.Close
    'Response.Write SQLLanguageFilter
    set rsLanguageFilter = nothing
    set SQLLanguageFilter = nothing
    
    response.write "</TD>"  & vbCrLf
    response.write "</TR>"  & vbCrLf
 
    ' List Records
    
		  rsCalendar.MoveFirst
    if Record_Limit * (PCID - 1) > 0 then
   	rsCalendar.Move (Record_Limit * (PCID - 1))
    end if  

    Record_Number = 1
    
    Old_ID = 0
    'Dim rsCheckPriceList
    blnPriceList = false 
    do while not rsCalendar.EOF and Record_Number <= Record_Limit
      
      if Old_ID <> rsCalendar("ID") then
        blnPriceList = false
        Call Update_Fields
  
        ' Anchor
        response.write "<TR>"  & vbCrLf
        response.write "<TD COLSPAN=7 HEIGHT=16 CLASS=NORMAL VALIGN=MIDDLE BGCOLOR=WHITE><A NAME=""" & Field_Data(xID) & """></A>&nbsp;</TD>" & vbCrLf
        response.write "</TR>" & vbCrLf
        'response.write  Session("PriceListCode")
        if rsCalendar("CheckCode").value =0 then
            Call Display_Category_Item
        else
           blnPriceList = true
           set rsCheckPriceList = conn.execute("exec PriceList_CheckAccessCode " & rsCalendar("ID") & ",'" & _
           Session("PriceListCode") & "'")
           if trim(rsCheckPriceList.fields(0).value) ="True" then
              Call Display_Category_Item
           end if
           set rsCheckPriceList = nothing
        end if
      end if  
      
      Old_ID = rsCalendar("ID")
      
      rsCalendar.MoveNext
    loop
    rsCalendar.close
    set rsCalendar = nothing
    
    ' Page Navigation

    response.write "<TR>"  & vbCrLf
    response.write "<TD COLSPAN=7 CLASS=Medium BGCOLOR=WHITE><BR>"  & vbCrLf
    Call RS_Page_Navigation
    response.write "<BR>&nbsp;"
    response.write "</TD>"  & vbCrLf
    response.write "</TR>"  & vbCrLf

    response.write "</TABLE>" & vbCrLf
    response.write "</FORM>"  & vbCrLf

  end if
          
  response.write "<BR><BR>" & vbCrLf
  
end if
        
' --------------------------------------------------------------------------------------        
%>