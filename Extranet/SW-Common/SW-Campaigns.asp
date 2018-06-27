<%

if CIN = 8000 or CIN = 8001 then
  SQL = "SELECT Code, Title FROM Calendar_Category WHERE Site_ID=" & Site_ID & " AND Code=" & CIN
  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.Open SQL, conn, 3, 3
  if not rsCategory.EOF then
    Campaign_Title = rsCategory("Title")
  else
    Campaign_Title = ""
  end if
  rsCategory.close
  set rsCategory = nothing
end if  

if not isblank(Campaign_Title) then
  response.write "<SPAN CLASS=Heading3>" & Translate(Campaign_Title,Login_Language,conn) & "</SPAN>"
  response.write "<BR><BR>"
end if

Title_View = True  ' Append SubCategory to Title
  
' Splash Header and Footer

SQL = "SELECT Calendar.ID, Calendar.Splash_Header, Calendar.Splash_Footer, Calendar.Thumbnail, Calendar.Item_Number_2, Calendar.Category_ID FROM Calendar WHERE Calendar.ID=" & SCID_Campaign
Set rsHeader = Server.CreateObject("ADODB.Recordset")
rsHeader.Open SQL, conn, 3, 3

if not rsHeader.EOF then
  if not isblank(rsHeader("Splash_Header")) then
    Splash_Header = rsHeader("Splash_Header")
  else
    Splash_Header = ""  
  end if  
  if not isblank(rsHeader("Item_Number_2")) then
    if isnumeric(rsHeader("Item_Number_2")) then
      Intro_Letter = rsHeader("Item_Number_2")
    else
      Intro_Letter = 0
    end if    
  else
    Intro_Letter = 0
  end if
  Category_PIC = rsHeader("Category_ID")
  if not isblank(rsHeader("Splash_Footer")) then
    Splash_Footer = rsHeader("Splash_Footer")
  else
    Splash_Footer = ""  
  end if
  if not isblank(rsHeader("Thumbnail")) then
    Splash_Thumbnail = rsHeader("Thumbnail")
  else
    Splash_Thumbnail = ""  
  end if
  Category_PIC = rsHeader("Category_ID")
end if

rsHeader.close
set rsHeader = nothing  

' Categories 

SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
Set rsCategory = Server.CreateObject("ADODB.Recordset")
rsCategory.Open SQL, conn, 3, 3

WhatsNewFlag = False

' Display Splash Header

if not isblank(Splash_Header) then
  response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=4 BGCOLOR=""#F3F3F3"">" & vbCrLf
  response.write "<TR>"
  if not isblank(Splash_Thumbnail) then
    response.write "<TD CLASS=Small WIDTH=""80"" VALIGN=TOP><IMG SRC=""/" & Site_Code & "/" & Splash_Thumbnail & """ BORDER=1 WIDTH=80></TD>"
  end if
  response.write "<TD CLASS=Small VALIGN=TOP>" & RestoreQuote(Splash_Header) & "</TD>"
  response.write "</TR>"
  response.write "</TABLE>"
  response.write "<BR><BR>"
end if

' Display Introduction Letter

if Intro_Letter > 0 then

  SQL = "SELECT Calendar.*, Item_Number AS Lit_Description, Item_Number AS Lit_Item FROM Calendar WHERE Calendar.ID=" & Intro_Letter
  Set rsCalendar = Server.CreateObject("ADODB.Recordset")
  rsCalendar.Open SQL, conn, 3, 3

  if not rsCalendar.EOF then           
  
    Record_Number = 0

    response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
    
    Call Update_Fields
    Call Display_Category_Item
    
    response.write "</TABLE>"
    response.write "<BR><BR>"

  end if
  
  rsCalendar.close
  set rsCalendar = nothing
  
end if  

' For Each Category

do while not rsCategory.EOF

' SQL     = "SELECT Calendar.*, Item_Number AS Lit_Description, Item_Number AS Lit_Item " & _
'           "FROM Calendar " &_
'           "WHERE Calendar.Site_ID=" & CInt(Site_ID) &_
'           " AND Calendar.Category_ID=" & rsCategory("ID") &_
'           " AND Calendar.Campaign=" & SCID_Campaign

  SQL     = "SELECT Calendar.*, " &_
                "Literature_Items_US.Item AS Lit_Item, " &_          
                "Literature_Items_US.STATUS AS Lit_Status, " &_
                "Literature_Items_US.STATUS_Name AS Lit_Status_Name, " &_
                "Literature_Items_US.[ACTION] AS Lit_Action, " &_                                
                "Literature_Items_US.[Print] AS Lit_Print, " &_                                        
                "Literature_Items_US.POD AS Lit_POD, " &_
                "Literature_Items_US.CD AS Lit_CD, " &_
                "Literature_Items_US.Display AS Lit_Display, " &_
                "Literature_Items_US.Video_NTSC AS Lit_Video_NTSC, " &_
                "Literature_Items_US.Video_PAL AS Lit_Video_PAL, " &_
                "Literature_Items_US.CUSTOMER_ORDER AS Lit_C_Order, " &_
                "Literature_Items_US.INTERNAL_ORDER AS Lit_I_Order, " &_
                "Literature_Items_US.EFULFILLMENT AS Lit_Description, " &_
                "Literature_Items_US.COST_CENTER AS Cost_Center, " &_
                "Literature_Items_US.LITERATURE_TYPE AS Lit_Type, " &_
                "Literature_Items_US.LIT_CODE as Lit_Code, " &_
                "Literature_Items_US.PSize as Lit_PSize, " &_
                "Literature_Items_US.COLORS as Lit_Colors, " &_            
                "Literature_Items_US.LARGE_LIMIT AS Lit_Large_Limit, " &_
                "Literature_Items_US.SMALL_LIMIT AS Lit_Small_Limit, " &_
                "Literature_Items_US.END_USER AS Lit_End_User, " &_
                "Literature_Items_US.UOM AS Lit_UOM, " &_
                "Literature_Items_US.Revision AS Lit_Revision, " &_                            
                "Literature_Items_US.Inventory_Rule AS Lit_Inventory_Rule, " &_                
                "Literature_Items_US.[LANGUAGE] AS Lit_Language " &_
            "FROM   Calendar " &_
            "   LEFT OUTER JOIN " &_
                    "Literature_Items_US ON Calendar.Item_Number = Literature_Items_US.ITEM AND dbo.Literature_Items_US.REVISION = dbo.Calendar.Revision_Code " &_
            "   LEFT OUTER JOIN " &_
                    "Calendar_Category ON Calendar.Code = Calendar_Category.Code AND Calendar.Site_ID = Calendar_Category.Site_ID " &_         
            "WHERE Calendar.Site_ID=" & CInt(Site_ID) &_
            "   AND Calendar.Category_ID=" & rsCategory("ID") &_
            "   AND Calendar.Campaign=" & SCID_Campaign

  if Intro_Letter > 0 then          
    SQL = SQL & " AND Calendar.ID<>" & Intro_Letter
  end if      

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

  if Bypass_Active_Flag = false then
    SQL = SQL & " AND (Literature_Items_US.ACTIVE_FLAG IS NULL OR Literature_Items_US.ACTIVE_FLAG = -1) "  
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

'response.write SQL & "<P>"  
  
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

' Display Splash Footer

if not isblank(Splash_Footer) then
  response.write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=2 CELLSPACING=4 BGCOLOR=""#F3F3F3"">" & vbCrLf
  response.write "<TR>"
  response.write "<TD CLASS=Small VALIGN=TOP>" & RestoreQuote(Splash_Footer) & "</TD>"
  response.write "</TR>"
  response.write "</TABLE>"
  response.write "<BR><BR>"
end if


rsCategory.close
set rsCategory = nothing

SCID_Campaign = 0
  
%>