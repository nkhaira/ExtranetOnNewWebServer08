<%

' --------------------------------------------------------------------------------------
' Name  : Site Wide Headlines (Include File)
' Author: K. D. Whitlock
' Date  : 11/01/2000
' --------------------------------------------------------------------------------------

  SQL = "SELECT Calendar_Category.* FROM Calendar_Category WHERE Calendar_Category.Site_ID=" & CInt(Site_ID) & " AND Calendar_Category.Enabled=" & CInt(True) & " ORDER BY Calendar_Category.Sort, Calendar_Category.Title"
  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.Open SQL, conn, 3, 3
  
  Have_Headlines = false
  blnShowRecord = false
  do while not rsCategory.EOF

    Headline_Counter = 0          ' Maximum Displayed records per Category
    
    SQL = "SELECT Calendar.*, Item_Number AS Lit_Description,(Select isnull(count(*),0) from Asset_accesscodes where assetid = calendar.id) as CheckCode FROM Calendar WHERE Calendar.Site_ID=" & CInt(Site_ID) &" AND Calendar.Category_ID=" & rsCategory("ID")
    


    ' Build LIKE Clause for Sub-Group Memberships

    if (Access_Level <= 2 or Access_Level = 6) then
      SQL = SQL & " AND (Calendar.SubGroups LIKE '%all%'"
      for i = 0 to UserSubGroups_Max
        SQL = SQL & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
      next
      SQL = SQL & ")"        
    end if   
    
    ' Determine if Current - By Pre Announce Date and then again on Begin Date

      SQL = SQL & " AND Calendar.Status=1"
      SQL = SQL & " AND"
      SQL = SQL & " ((Calendar.LDate<='" & Date & "' AND Calendar.LDate>='" & DateAdd("d", (Show_Days * -1), Date) & "')"
      SQL = SQL & " OR"
      SQL = SQL & " ( Calendar.BDate<='" & Date & "' AND Calendar.BDate>='" & DateAdd("d", (Show_Days * -1), Date) & "'))"
      
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

    SQL = SQL & " ORDER BY Calendar.Sub_Category, Calendar.Product, Calendar.BDATE DESC"              


    Set rsCalendar = Server.CreateObject("ADODB.Recordset")
    rsCalendar.Open SQL, conn, 3, 3
    
    if TableOn = False then
      
      TableOn = True
      
      HeadLine_Number    = 0
      ScrollerWidth      = 440
      ScrollerHeight     = 240
      ScrollerForeGround = "Black"
      ScrollerBackGround = "White"
      ScrollerBorder     = "White"
      ScrollerFont       = "Arial, Helvetica"
      ScrollerFontSize   = "2"
      
      ' Set up Container for Scroller
      with response
      	.write vbCrLf
        Call Nav_Border_Begin
        .write "<TABLE WIDTH=""100%"" CELLPADDING=2>" & vbCrLf
        .write "<TR>" & vbCrLf
        .write "<TD BGCOLOR=""Black"" WIDTH=""100%"" CLASS=MediumBoldGold>" & vbCrLf
if site_id=46 then
  .write Translate("This weeks's news, articles, tools and events",Login_Language,conn) & vbCrLf
else
  .write Translate("This Weeks's News, Articles, Tools and Events",Login_Language,conn) & vbCrLf
end if
      
        .write "</TD>" & vbCrLf
        .write "</TR>" & vbCrLf
        .write "<TD CLASS=Medium VALIGN=TOP>" & vbCrLf
      end with
      
      ' Set up Scroller Includes and Pre-Sets
      with response
        .write "<SCRIPT LANGUAGE=""JavaScript"" src=""/include/SW-dHTML-lib.js""></SCRIPT>" & vbCrLf
        .write "<SCRIPT LANGUAGE=""JavaScript"" src=""/include/SW-Scroller.js""></SCRIPT>" & vbCrLf
        .write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
        
        .write "var myScroller1 = new Scroller(0, 0, " & ScrollerWidth & ", " & ScrollerHeight & ", 1, 0);" & vbCrLf
        .write "myScroller1.setColors(""" & ScrollerForeGround & """, """ & ScrollerBackGround & """, """ & ScrollerBorder & """);" & vbCrLf
        .write "myScroller1.setFont(""" & ScrollerFont & """, " & ScrollerFontSize & ");" & vbCrLf
        
        ' Scroller Function
        .write "function RunScroller() {" & vbCrLf
        .write "  var layer;" & vbCrLf
        .write "  var xHorizontal, yVertical;" & vbCrLf

        ' Locate Placeholder layer so we can use it to position the Scroller.

        .write "  layer       = getLayer(""placeholder"");" & vbCrLf
        .write "  xHorizontal = getPageLeft(layer);" & vbCrLf
        .write "  yVertical   = getPageTop(layer);" & vbCrLf

        ' Create the Scroller and position it.

        .write "  myScroller1.create();" & vbCrLf
        .write "  myScroller1.hide();" & vbCrLf
        .write "  myScroller1.moveTo(xHorizontal, yVertical);" & vbCrLf
        .write "  myScroller1.setzIndex(100);" & vbCrLf
        .write "  myScroller1.show();" & vbCrLf
        .write "}" & vbCrLf
      end with

    end if

    Record_Number = 0 
        
    do while not rsCalendar.EOF and Headline_Counter <= 2
     
     if rsCalendar("CheckCode").value =0 then
             blnShowRecord = true
     else
           set rsCheckPriceList = conn.execute("exec PriceList_CheckAccessCode " & rsCalendar("ID") & ",'" & _
           Session("PriceListCode") & "'")
           if trim(rsCheckPriceList.fields(0).value) ="True" then
                 blnShowRecord = false
           end if
           set rsCheckPriceList = nothing
      end if
      if     (CDate(rsCalendar("BDate")) <> CDate(rsCalendar("EDate")) and CDate(rsCalendar("EDate")) < Date) then
      elseif (rsCalendar("XDays") > 0 and CDate(rsCalendar("EDate")) < Date) then

      else
        if blnShowRecord = true then
                 Have_Headlines = true
                 
                 for i = 0 to Field_Max
                   if isnumeric(rsCalendar(Field_Name(i))) or isdate(rsCalendar(Field_Name(i))) or not isblank(Trim(rsCalendar(Field_Name(i)))) then
                     Field_Data(i) = Trim(rsCalendar(Field_Name(i)))
                     Field_Flag(i) = True
                   else
                     Field_Data(i) = ""
                     Field_Flag(i) = False
                   end if
                 next

                 ' Begin AddItem Container
                 
                 response.write "myScroller1.addItem('"
                 
                 ' Title
                 if Field_Flag(xTitle) = True then
                  ' Category
                   if Record_Number = 0 then
                     with response
                       .write "<TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=0 CELLSPACING=0>"
                       .write "<TR>"
                       .write "<TD VALIGN=MIDDLE>"            
                       .write "<FONT CLASS=Heading3Red>" & Translate(RestoreQuote(rsCategory("Title")),Login_Language,conn) & "</FONT>"
                       .write "</TD>"
                       .write "<TD VALIGN=TOP>"
                     end with
                     if isblank(Logo) then
                       response.write "<IMG SRC=""/images/flukelogo.gif"" WIDTH=80 ALIGN=RIGHT VSPACE=10 HSPACE=10>"
                     else
                       response.write "<IMG SRC=""" & Logo & """ WIDTH=80 ALIGN=RIGHT VSPACE=10 HSPACE=10>"
                     end if
                     with response
                       .write "</TD>"
                       .write "</TR>"
                       .write "</TABLE>"
                     end with  
                   end if  

                   with response
                     .write "<A HREF="""   & HomeURL
                     .write "?Site_ID="    & Site_ID
                     .write "&Language="   & Login_Language
                     .write "&NS="         & Top_Navigation
                     .write "&CID="        & "9001"

                     .write "&SCID="       & "0"     'SCID
                     .write "&PCID="       & "2"     'PCID
                     .write "&CIN="        & "0"     'rsCalendar("Code")
                     .write "&CINN="       & "0"     'rsCalendar("Category_ID")
                     .write "&Show_Days="  & "14"
                     .write "#"            & rsCalendar("ID")
                     .write """>"

                     .write "<FONT CLASS=NormalBold><FONT COLOR=""Blue"">"
                     .write KillCrLf(Field_Data(xTitle))
                     .write "</FONT>"

                     .write "</FONT>"
                     .write "</A>"
                     .write "<BR>"
                     .write "<IMG SRC=""/images/1X1TRANS.GIF""  WIDTH=""100%"" HEIGHT=8>"            
                   end with
                 end if
                         
                 ' Product
                 if Field_Flag(xProduct) = True then
                   with response
                     .write "<FONT CLASS=MediumBold>"
                     .write Translate("Product",Login_Language,conn) & ": " & KillCrLf(Field_Data(xProduct))
                     .write "</FONT>"
                     .write "<BR>"
                   end with
                 end if

                 ' Date
                 if Field_Flag(xBDATE) = True and Field_Flag(xEDate) = True then
                   response.write "<FONT CLASS=Small>"
                   response.write Day(Field_Data(xBDATE)) & " " & Translate(MonthName(Month(Field_Data(xBDATE))),Login_Language,conn) & " " & Year(Field_Data(xBDATE))
                   if Field_Data(xBDate) <> Field_Data(xEDate) then
                     response.write " - "
                     response.write Day(Field_Data(xEDATE)) & " " & Translate(MonthName(Month(Field_Data(xEDATE))),Login_Language,conn) & " " & Year(Field_Data(xEDATE))
                   end if
                   response.write "</FONT>"
                   with response
                     .write "<IMG SRC=""/images/1X1LINE.GIF""  WIDTH=""100%"" HEIGHT=1>"
                     .write "<IMG SRC=""/images/1X1TRANS.GIF""  WIDTH=""100%"" HEIGHT=4>"
                   end with
                 end if
                   
                 ' Location
                 if Field_Flag(xLocation) = True then
                   with response
                     .write "<FONT CLASS=Small>"
          	         .write Translate("Location",Login_Language,conn) & ": " & KillCrLf(Field_Data(xLocation))
                     .write "</FONT>"
                     .write "<BR>"
                   end with
                 end if
           
                 ' Description
                 
                 with response
                   .write "<TABLE WIDTH=""100%"" BORDER=0 COLPADDING=2 COLSPACING=0>"
                   .write "<TR>"
                 end with
                     
                 ' Thumbnail
                                 
                 if Field_Flag(xThumbnail) = True then
                   with response
                     .write "<TD CLASS=Small VALIGN=TOP>"
                     .write "<IMG SRC=""" & Field_Data(xThumbnail) & """ WIDTH=60 BORDER=1>"
                     .write "</TD>"
                   end with            
                 end if
                   
                 ' Description / Language Cell
                 
                 response.write "<TD CLASS=Small>"
                 
                 if Field_Flag(xDescription) = True then
                   response.write KillCrLf(Field_Data(xDescription))
                   response.write "<BR>"
                 end if
                 
                 ' Language
                 
                 if Field_Flag(xLanguage) = True then
                   SQL = "SELECT * FROM Language WHERE Language.Code='" & Field_Data(xLanguage) & "'"
                   Set rsLanguage = Server.CreateObject("ADODB.Recordset")
                   rsLanguage.Open SQL, conn, 3, 3    
           
                   with response
                     .write "<FONT COLOR=""Gray"">"
                     .write Translate("Language",Login_Language,conn) & ": "
                     .write Translate(rsLanguage("Description"),Login_Language,conn)
                     .write "</FONT>"
                   end with
           
                   rsLanguage.close
                   Set rsLanguage = nothing          
                 end if
                       
                 with response
                   .write "</TD>"
                   .write "</TR>"
                   .write "</TABLE>"
                 end with
                 
                 response.write "');" & vbCrLf & vbCrLf
                 
                 Record_Number = Record_Number + 1
           
               end if

               Headline_Counter = Headline_Counter + 1
      end if  'Show  Record End
      rsCalendar.MoveNext

      
    loop

    rsCalendar.close
    set rsCalendar = nothing

    Record_Number = 0
    
    rsCategory.MoveNext
    
  loop

  rsCategory.close
  set rsCategory = nothing

' --------------------------------------------------------------------------------------
' Begin AddItem Container for no New Items
' --------------------------------------------------------------------------------------

  with response
    .write "myScroller1.addItem('"
    if Have_Headlines = false then      
      .write "<B>" & Translate("No new items have been added to this site within the past " & Show_Days & "-days.",Login_Language,conn) & "</B><P>"
    end if
    .write "<B>" & Translate("Tip",Login_Language,conn) & "</B>: " & Translate("You can click on ""What&rsquo;s New"" and select a longer time period from 14 up to 60-days to quickly view items that have been added to this site within those time periods.",Login_Language,conn)
    .write "');" & vbCrLf & vbCrLf
  end with

  Headline_Counter = Headline_Counter + 1

' --------------------------------------------------------------------------------------

  if TableOn = True then
    with response
      .write vbCrLf
      .write "window.onload=RunScroller" & vbCrLf
      .write "</SCRIPT>" & vbCrLf

      .write "<DIV ID=""placeholder"" STYLE=""position:relative; width:" & ScrollerWidth & "px; height:" & ScrollerHeight & "px;"">&nbsp;</DIV>" & vbCrLf

      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf & vbCrLf
                Call Nav_Border_End
    end with  
    
  end if

%>
