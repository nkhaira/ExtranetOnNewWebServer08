<%
' --------------------------------------------------------------------------------------
' Individual Record Data
' --------------------------------------------------------------------------------------

' Release Task  :   688
' Updated by    :   Amol Jagtap
' Description   :   To change all the CInt value conversion to CLng to avoid Interger range violation in future
    
sub Display_Category_Item
    
  ' Begin Record Listing
          
  Record_Number = Record_Number + 1                                            

  ' Product or Product Series Title and Sub-Category
  
  if Product <> Field_Data(xProduct) or Category <> Field_Data(xSub_Category) then   
       
    response.write "<TR>" & vbCrLf
    response.write "<TD HEIGHT=16 COLSPAN=4 CLASS=PRODUCT>&nbsp;"
    response.write Field_Data(xProduct)
    response.write "</TD>" & vbCrLf
    if Product <> Field_Data(xProduct) then
      Category = ""
    end if  
    if Category <> Field_Data(xSub_Category) then
      response.write "<TD HEIGHT=16 COLSPAN=3 CLASS=PRODUCT ALIGN=RIGHT>" & Field_Data(xSub_Category) & "&nbsp;</TD>" & vbCrLf
    else
      response.write "<TD HEIGHT=16 COLSPAN=3 CLASS=PRODUCT>&nbsp;</TD>" & vbCrLf
    end if                                        
    response.write "</TR>" & vbCrLf
  end if
  
  Product  = Field_Data(xProduct)
  Category = Field_Data(xSub_Category)                                    

  response.write "<TR>" & vbCrLf          
  response.write "<TD HEIGHT=1 COLSPAN=7><IMG SRC=""/images/1X1LINE.GIF"" WIDTH=""100%"" HEIGHT=1></TD>"
  response.write "</TR>" & vbCrLf                                  
    
  ' Title
                                                               
  response.write "<TR>" & vbCrLf
  response.write "<TD COLSPAN=4 WIDTH=""50%"" CLASS=NORMALBOLD VALIGN=TOP>" & "<A NAME=""" & Field_Data(xID) & """></A>"
  if not isblank(Field_Data(xTitle)) then
    if not isblank(KeySearch) then
      TempString = Field_Data(xTitle)
      for x = 0 to UBound(strSearch)
        TempString = Highlight_Keyword(TempString,strSearch(x), "#FF0000")
      next
      response.write TempString  
    else  
      response.write Field_Data(xTitle)
    end if
    if CLng(Field_Flag(xSub_Category)) = CLng(True) and CLng(Title_View) = CLng(True) and (CIN < 8000 or CIN > 8999) then
      response.write " - " & Translate(Field_Data(xSub_Category),Login_Language,conn)
    end if
  elseif CID = 9004 then
    if not isblank(rsCalendar("Lit_Description")) then
      if not isblank(KeySearch) then
        TempString = ProperCase(rsCalendar("Lit_Description"))
        for x = 0 to UBound(strSearch)
          TempString = Highlight_Keyword(TempString,strSearch(x), "#FF0000")
        next
        response.write TempString  
      else
        response.write ProperCase(rsCalendar("Lit_Description"))
      end if  
    else
      response.write "&nbsp;"  
    end if
  else
    response.write "&nbsp;"  
  end if
  
  response.write "</TD>" & vbCrLf
  
  ' Date
  
  response.write "<TD COLSPAN=3 WIDTH=""25%"" CLASS=SmallBold VALIGN=TOP ALIGN=RIGHT>"
    
  if CLng(Field_Flag(xBDATE)) = CLng(True) and CLng(Field_Flag(xEDate)) = CLng(True) then
    response.write        Day(Field_Data(xBDATE)) & " " & Translate(MonthName(Month(Field_Data(xBDATE))),Login_Language,conn) & " " & Year(Field_Data(xBDATE))
    if Field_Data(xBDate) <> Field_Data(xEDate) then
      response.write " - "
      response.write Day(Field_Data(xEDATE)) & " " & Translate(MonthName(Month(Field_Data(xEDATE))),Login_Language,conn) & " " & Year(Field_Data(xEDATE))
    end if
      
    if CLng(Field_Flag(xLDate)) = CLng(True) and (Access_Level = 4 or Access_Level >= 8) then
      if DateDiff("d",Field_Data(xLDate),Date) < 0 then
        response.write "<BR><FONT COLOR=""#009900"">" & Translate("Go Live Date",Login_Language,conn) & ": " & Day(Field_Data(xLDATE)) & " " & Translate(MonthName(Month(Field_Data(xLDATE))),Login_Language,conn) & " " & Year(Field_Data(xLDATE)) & "</FONT>"
      elseif CLng(Field_Flag(xUDATE)) = CLng(True) then
        response.write "<BR><FONT CLASS=Small>" & Translate("Aging",Login_Language,conn) & ": " & DateDiff("d",Field_Data(xUDate),Date) & " "
        if (DateDiff("d",Field_Data(xUDate),Date) > 1  or DateDiff("d",Field_Data(xUDate),Date) < -1) then
         response.write Translate("days",Login_Language,conn) 
        else
         response.write Translate("day",Login_Language,conn) 
        end if 
        response.write "</FONT>"
      end if
    end if  
  else
    response.write "&nbsp;"
  end if    
  response.write "</TD>" & vbCrLf            
  response.write "</TR>" & vbCrLf

  ' Location
  
  if CLng(Field_Flag(xLocation)) = CLng(True) then
    response.write "<TR>" & vbCrLf            
    response.write "<TD COLSPAN=7 CLASS=Small VALIGN=TOP>" & Field_Data(xLocation) & "</TD>" & vbCrLf
    response.write "<TR>" & vbCrLf                          
  end if
    
  ' Thumbnail / Description Container

  if Show_Detail < 0 _
    and ((CLng(Field_Flag(xThumbnail))  = CLng(True)  _
    or  CLng(Field_Flag(xDescription))  = CLng(True)  _
    or  CLng(Field_Flag(xInstructions)) = CLng(True)  _
    or  CLng(Field_Flag(xItem_Number))  = CLng(True)  _
    or (CLng(Field_Flag(xItem_Number))  = CLng(True) and CLng(Field_Flag(xItem_Number_Show)) = CLng(True) and CLng(Field_Data(xItem_Number_Show)) = CLng(True)) _
    or (CLng(Field_Flag(xItem_Number))  = CLng(False) and not isblank(rsCalendar("Lit_Item"))) _
    or  CLng(Field_Flag(xFile_Name))    = CLng(True)  _    
    or  CLng(Field_Flag(xConfidential)) = CLng(True)  _    
    or  CLng(Field_Flag(xPEData))       = CLng(True))) then

    response.write "<TR>" & vbCrLf
    response.write "<TD COLSPAN=7 CLASS= Normal VALIGN=TOP>" & vbCrLf
    response.write "<TABLE WIDTH=""100%"" CELLPADDING=4 CELLSPACING=2 BORDER=0 BGCOLOR=""#F3F3F3"">" & vbCrLf
    response.write "<TR>" & vbCrLf
    
    'Thumbnail          
    
    response.write "<TD WIDTH=""1%"" ALIGN=CENTER VALIGN=MIDDLE CLASS=NORMAL>" & vbCrLf
    if CLng(Field_Flag(xThumbnail)) = CLng(True) and CLng(Show_Thumbnail) = CLng(True) then
      if CLng(Field_Flag(xFile_Name)) = CLng(True) then
        select case UCase(Mid(Field_Data(xFile_Name),Instr(1,Field_Data(xFile_Name), ".")))
          case ".EXE", ".ZIP", ".MDB", ".BAT"
          case else
            response.write "<A HREF=""javascript:void(0);"" TITLE=""View File"
            if CLng(Field_Flag(xFile_Name_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xFile_Name_Size) / 1024)),0) & " KBytes"
            response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),0,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
        end select                                                                    
      end if                                                        
      response.write "<IMG SRC=""http://" & Request("SERVER_NAME") & "/" & Site_Code & "/" & Field_Data(xThumbnail) & """ WIDTH=80 BORDER=1 COLOR=Black>"
      if CLng(Field_Flag(xFile_Name)) = CLng(True) then
        response.write "</A>" & vbCrLf
      end if                                          
    elseif Access_Level > 0 then
        Write_Thumbnail_Icon
        'response.write "<IMG SRC=""/images/Blank_Doc.jpg"" WIDTH=""80"" BORDER=1 VSPACE=0 COLOR=""#C0C0C0"">" & vbCrLf
    else  
      response.write "<IMG SRC=""/images/Blank.gif"" HEIGHT= ""1"" WIDTH=""80"" BORDER=0 VSPACE=0>" & vbCrLf
    end if  
    response.write "</TD>" & vbCrLf

    ' Description
    response.write "<TD WIDTH=""99%"" CLASS=Small VALIGN=TOP>" & vbCrLf
    response.write "<SPAN CLASS=SmallBold>" & Translate("Description",Login_Language,conn) & "</SPAN><BR>"

    if CLng(Field_Flag(xTitle)) = CLng(True) and CLng(Field_Flag(xDescription)) = CLng(False) then
      Field_Data(xDescription) = Field_Data(xTitle)
      Field_Flag(xDescription) = CLng(True)
    end if

    if CLng(Field_Flag(xDescription)) = CLng(True) then
      if not isblank(KeySearch) then
        TempString = Field_Data(xDescription)
        for x = 0 to UBound(strSearch)
          TempString = Highlight_Keyword(TempString,strSearch(x), "#FF0000")
        next
        response.write TempString  
      else  
        response.write Field_Data(xDescription)
      end if  
      response.write "<P>" & vbCrLf
    elseif CID = 9004 then
      if not isblank(rsCalendar("Lit_Description")) then
        if not isblank(KeySearch) then
          TempString = ProperCase(rsCalendar("Lit_Description"))
          for x = 0 to UBound(strSearch)
            TempString = Highlight_Keyword(TempString,strSearch(x), "#FF0000")
          next
          response.write TempString  
        else
          response.write ProperCase(rsCalendar("Lit_Description"))
        end if  
        response.write "<P>" & vbCrLf
      end if
    end if
    
    ' Special Instructions
    if CLng(Field_Flag(xInstructions)) = CLng(True) then
      response.write "<FONT CLASS=SmallBold>" & Translate("Special Instructions",Login_Language,conn) & "</FONT><BR>" & Field_Data(xInstructions) & "<BR><BR>"
    end if  

    ' Item / Reference Number
    
    if (CLng(Field_Flag(xItem_Number)) = CLng(True) and CLng(Field_Flag(xItem_Number_Show)) = CLng(True) and CLng(Field_Data(xItem_Number_Show)) = CLng(True)) then
      response.write "<FONT CLASS=SmallBold>" & Translate("Literature Code",Login_Language,conn) & "</FONT><BR>"
      if not isblank(KeySearch) then
        TempString = Field_Data(xItem_Number)
        if ItemSearch = True then  
          for x = 0 to UBound(strSearch)
            TempString = Highlight_Keyword(TempString,strSearch(x), "#FF0000")
          next
        end if  
        response.write TempString  
      else  
        response.write Field_Data(xItem_Number)
      end if
      if CLng(Field_Flag(xRevision_Code)) = CLng(True) then
        response.write " " & Field_Data(xRevision_Code)
      end if
      
      if Access_Level > 0 and not isblank(rsCalendar("Lit_Code")) then
        response.write "&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR=""#666666"">(" & rsCalendar("Lit_Code") & ")</FONT>"
      end if    

    elseif CID = 9004 then
      if not isblank(rsCalendar("Lit_Item")) then    
        response.write "<FONT CLASS=SmallBold>" & Translate("Literature Code",Login_Language,conn) & "</FONT><BR>"
        if not isblank(KeySearch) then
          TempString = rsCalendar("Lit_Item")
          if ItemSearch = True then  
            for x = 0 to UBound(strSearch)
              TempString = Highlight_Keyword(TempString,strSearch(x), "#FF0000")
            next
          end if  
          response.write TempString  
        else  
          response.write rsCalendar("Lit_Item")
        end if
        if not isblank(rsCalendar("Lit_Revision")) then
          response.write " " & rsCalendar("Lit_Revision")
        end if
        
        if Access_Level > 0 and not isblank(rsCalendar("Lit_Code")) then
          response.write "&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR=""#666666"">(" & rsCalendar("Lit_Code") & ")</FONT>"
        end if    
          
      end if        
    end if
    
    if Access_Level > 0 and not isblank(rsCalendar("Lit_Colors")) then
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR=""#666666"">" & ProperCase(rsCalendar("Lit_Colors")) & "</FONT>&nbsp;&nbsp;"
    end if

    if Access_Level > 0 and not isblank(rsCalendar("Lit_PSize")) then
        TempString = LCase(rsCalendar("Lit_PSize"))
        TempString = Replace(TempString,"x"," x ")        
        TempString = Replace(TempString," by "," x ")
        TempString = Replace(TempString,"""","in ")
        TempString = Replace(TempString,"'","ft ")
        TempString = Replace(TempString,"?","")
        TempString = Replace(TempString,"tbd.","")
        TempString = Replace(TempString,"1/2"," 1/2 ")
        TempString = Replace(TempString,"1/3"," 1/3 ")
        TempString = Replace(TempString,"1/4"," 1/4 ")
        TempString = Replace(TempString,"1/8"," 1/8 ")
        TempString = Replace(TempString,"   "," ")
        TempString = Replace(TempString,"  "," ")
        response.write "<FONT COLOR=""#666666""> " & Replace(Replace(Replace(ProperCase(TempString),"In","in"),"Ft","ft"),"X","x") & "</FONT>&nbsp;&nbsp;"
    end if  
    
    ' File Format
    if CLng(Field_Flag(xFile_Name)) = CLng(True) then
      SQL="SELECT * FROM Asset_Type WHERE Asset_Type.File_Extension='" & UCase(mid(Field_Data(xFile_Name),Instr(1,Field_Data(xFile_Name),".")+1)) & "'"
      Set rsAsset_Type = Server.CreateObject("ADODB.Recordset")
      rsAsset_Type.Open SQL, conn, 3, 3
      if not rsAsset_Type.EOF then
        response.write "<P><FONT CLASS=SmallBold>" & Translate("File Format",Login_Language,conn) & "</FONT><BR>"  & Translate(rsAsset_Type("File_Type"),Login_Language,conn)
      end if
      rsAsset_Type.close
      set rsAsset_Type = nothing
      
      ' Page Count
      if rsCalendar("File_Page_Count") > 0 then
        response.write "&nbsp&nbsp&nbsp&nbsp;<FONT CLASS=Small>" & Translate("Page Count",Login_Language,conn) & "</FONT>:&nbsp"  & rsCalendar("File_Page_Count")
      else
        response.write "<P>"  
      end if
        
    end if

    ' Confidential
    if CLng(Field_Flag(xConfidential)) = CLng(True) then
      if CLng(Field_Data(xConfidential)) = CLng(True) then
        response.write "<P><FONT CLASS=SmallBoldRed>" & Translate("Confidential Information - Not for Public Release",Login_Language,conn) & "</FONT>"
      end if
    end if
    
    ' Embargoed Information
    if CLng(Field_Flag(xPEDate)) = CLng(True) then
      if CDate(Date()) < CDate(Field_Data(xPEDate)) then    
        response.write "<P><FONT CLASS=SmallBoldRed>" & Translate("Embargoed Information - Not for Public Release until",Login_Language,conn) & ": "  & Day(Field_Data(xPEDATE)) & " " & Translate(MonthName(Month(Field_Data(xPEDATE))),Login_Language,conn) & ", " & Year(Field_Data(xPEDATE)) & "</FONT><BR>"
      end if
    end if
    
    ' Exclude / Include Countries
    if CLng(Field_Flag(xCountry)) = CLng(True) and (Access_Level = 2 or Access_Level = 4 or Access_Level >= 8) then
      if instr(1,Field_Data(xCountry),"0,") > 0 then
        response.write "<P><FONT COLOR=""#666666""><B>" & Translate("Exclude from Countries",Login_Language,conn) & "</B>: "  & Replace(Field_Data(xCountry),"0, ","") & "</FONT>"
      elseif LCase(Field_Data(xCountry)) <> "none" then
        response.write "<P><FONT COLOR=""#666666""><B>" & Translate("Limit to Countries",Login_Language,conn) & "</B>: "  & Field_Data(xCountry) & "</FONT>"      
      end if
    end if


    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
  
  end if
                
  Call RecordNavBar

end sub

' --------------------------------------------------------------------------------------
' Individual Record Navigation Bar
' --------------------------------------------------------------------------------------

sub RecordNavBar

  response.write "<TR>" & vbCrLf           
  response.write "<TD HEIGHT=11 VALIGN=MIDDLE WIDTH=""14%"" CLASS=NavRecordLight>"

  ' New Icon
  
  if CLng(Field_Flag(xLDate)) = CLng(True) then
    if DateDiff("d", Date, DateAdd("d",7,Field_Data(xLDate))) > 0 then
      response.write "&nbsp;&nbsp;<IMG SRC=""/images/new.gif"" ALT=""New Item within the last 7 Days"" ALIGN=ABSMIDDLE WIDTH=30 HEIGHT=14>&nbsp;&nbsp;&nbsp;&nbsp;"
    else
      response.write "&nbsp;&nbsp;<IMG SRC=""/images/1x1trans.gif"" ALIGN=ABSMIDDLE WIDTH=30 HEIGHT=14>&nbsp;&nbsp;&nbsp;&nbsp;"      
    end if
  end if  
    
  ' View Text
  
  if CLng(Field_Flag(xLink)) = CLng(True) or CLng(Field_Flag(xFile_Name)) = CLng(True) then
    select case LCase(Mid(Field_Data(xFile_Name),Instr(Field_Data(xFile_Name), ".") + 1 ))
      case "zip", "exe", "mdb", "bat"
        response.write "&nbsp;"
      case "iso"
        response.write Translate("Download",Login_Language,conn) & ":&nbsp;&nbsp;"
      case else
        response.write Translate("View",Login_Language,conn) & ":&nbsp;&nbsp;"
    end select  
  end if
    
  ' View Link
  if CLng(Field_Flag(xFile_Name)) = CLng(False) and CLng(Field_Flag(xLink)) = CLng(True) then
    if CLng(Field_Data(xLink_PopUp_Disabled)) = CLng(True) then     
      Link_Name    = Replace(Trim(Field_Data(xLink)),"https://support.fluke.com","https://" & request.ServerVariables("SERVER_NAME"))
      Link_Name    = Replace(Link_Name,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))                
      response.write "<A HREF=""" & Link_Name & """ TITLE=""View URL"">"
      response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
      response.write "</A>"      
    else
      response.write "<A HREF=""javascript:void(0);"" TITLE=""View URL"
      response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),7,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
      response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
      response.write"</A>"
    end if
  end if

  ' View File or Download ISO
  if CLng(Field_Flag(xFile_Name)) = CLng(True) then
    select case LCase(Mid(Field_Data(xFile_Name),Instr(Field_Data(xFile_Name), ".") + 1 ))
      case "zip", "exe", "mdb", "bat"
        response.write "&nbsp;"
      case "iso"        ' Download File - ISO File Only
        response.write "<A HREF=""javascript:void(0);"" TITLE=""Download ISO Disk Image File"
        if CLng(Field_Flag(xFile_Name_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xFile_Name_Size) / 1024)),0) & " KBytes"
        response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),0,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
        Icon_Type = 0
        Call Write_Icon
        response.write "</A>"
      case else
        response.write "<A HREF=""javascript:void(0);"" TITLE=""View File"
        if CLng(Field_Flag(xFile_Name_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xFile_Name_Size) / 1024)),0) & " KBytes"
        response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),0,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
        Icon_Type = 0
        Call Write_Icon
        response.write "</A>"
    end select
  end if
    
  response.write "</TD>" & vbCrLf          

  ' Download
  response.write "<TD HEIGHT=11 VALIGN=MIDDLE WIDTH=""15%"" CLASS=NavRecordLight>"

  ' Download Title
  if CLng(Field_Flag(xFile_Name)) = CLng(True) or CLng(Field_Flag(xArchive_Name)) = CLng(True) then
    response.write "&nbsp;&nbsp;" & Translate("Download",Login_Language,conn) & ":&nbsp;&nbsp;"
  else
    response.write "&nbsp;"          
  end if

  ' Download Check Archive First and Skip Regular File (Since View handles this)
  if CLng(Field_Flag(xArchive_Name)) = CLng(True) then

    response.write "<A HREF=""javascript:void(0);"" TITLE=""Download ZIP File"
    if CLng(Field_Flag(xArchive_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xArchive_Size) / 1024)),0) & " KBytes"
    response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),1,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
    Icon_Type = 1
    Call Write_Icon
    response.write "</A>"

  ' Download File - No Archive file
  elseif CLng(Field_Flag(xFile_Name)) = CLng(True) then
    response.write "<A HREF=""javascript:void(0);"" TITLE=""Download File"
    if CLng(Field_Flag(xFile_Name_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xFile_Name_Size) / 1024)),0) & " KBytes"
    response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),1,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
    Icon_Type = 0
    Call Write_Icon
    response.write "</A>"

  end if
 
  response.write "</TD>" & vbCrLf          

  ' Send File as Email Attachment  if CLng(Field_Flag(xFile_Name)) = CLng(True) then
  response.write "<TD HEIGHT=11 VALIGN=MIDDLE WIDTH=""15%"" CLASS=NavRecordLight>"

    if (CLng(Field_Flag(xFile_Name)) = CLng(True) and instr(1,LCase(Field_Data(xFile_Name)),".exe") = 0) and (CLng(Field_Flag(xFile_Name)) = CLng(True) and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0) or (CLng(Field_Flag(xArchive_Name)) = CLng(True) and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0) then
  
    response.write "&nbsp;&nbsp;" & Translate("Send",Login_Language,conn) & ":&nbsp;&nbsp;"

    if instr(1,LCase(Field_Data(xFile_Name)),".exe") = 0 and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0 then
      ' Non Zip Version
      response.write "<A HREF=""javascript:void(0);"" TITLE=""Email File to Yourself"
      if CLng(Field_Flag(xFile_Name_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xFile_Name_Size) / 1024)),0) & " KBytes"
      response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),6,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
      Icon_Type = 0
      Call Write_Icon
      response.write "</A>"
    end if
    
    ' Zip Version
    if CLng(Field_Flag(xArchive_Name)) = CLng(True) and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0 then    
      response.write "&nbsp;&nbsp;&nbsp;<A HREF=""javascript:void(0);"" TITLE=""Email ZIP File to Yourself"
      if CLng(Field_Flag(xArchive_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xArchive_Size) / 1024)),0) & " KBytes"
      response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),2,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
      Icon_Type = 1
      Call Write_Icon
      response.write "</A>"
    end if  

  else
    response.write "&nbsp;"          
  end if
  
  response.write "</TD>" & vbCrLf          

  ' Link, PI/C or Email Asset to Associate
  
  response.write "<TD HEIGHT=11 VALIGN=MIDDLE WIDTH=""14%"" CLASS=NavRecordLight>"   

  if CLng(Field_Flag(xLink)) = CLng(True) then     

    response.write "&nbsp;&nbsp;" & Translate("Link",Login_Language,conn) & ":&nbsp;&nbsp;"

    if CLng(Field_Data(xLink_PopUp_Disabled)) = CLng(True) then  
      Link_Name    = Replace(Trim(Field_Data(xLink)),"https://support.fluke.com","https://" & request.ServerVariables("SERVER_NAME"))
      Link_Name    = Replace(Link_Name,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))    
      response.write "<A HREF=""" & Link_Name & """ TITLE=""View URL"">"
      response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
      response.write "</A>"      
    else
      response.write "<A HREF=""javascript:void(0);"" TITLE=""View URL"
      response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),7,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
      response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
      response.write"</A>"
    end if
    
  else
    ' Release Task  :   688
    ' Updated by    :   Amol Jagtap
    ' Description   :   To convert Field_Data(xCampaign) to Long from Int
    if (CID = 9003 and CIN >= 8000 and CIN <= 8999 and CLng(Field_Data(xCampaign)) = 0) then

      response.write "&nbsp;&nbsp;" & Translate("More Info",Login_Language,conn) & ":&nbsp;&nbsp;"   
      response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&CID=" & CID & "&SCID=" & Field_Data(xID) & "&PCID=" & PCID & "&CIN=" & CIN & "&CINN=" & CINN & """>"
      response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
      response.write"</A>"

    elseif CID = 9001 or CID = 9003 or CID = 9004 then
    
      if (Field_Data(xCode) >= 8000 and Field_Data(xCode) <= 8999 and CLng(Field_Data(xCampaign)) = 0) then     
        response.write "&nbsp;&nbsp;" & Translate("More Info",Login_Language,conn) & ":&nbsp;&nbsp;"   
        response.write "<A HREF=""" & HomeURL & "?Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&CID=9003" & "&SCID=" & Field_Data(xID) & "&PCID=1" & "&CIN=" & Field_Data(xCode) & "&CINN=" & Field_Data(xCategory_ID) & """>"
        response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
        response.write"</A>"
      
      ' Email Document to Associate
      elseif CLng(Field_Flag(xFile_Name)) = CLng(True) then
      
        ' Do not display Send It if Confidential or Public Release Date has not been met
        
        SendIt_Show = true
        if CLng(Field_Data(xConfidential)) = CLng(True) then
          SendIt_Show = false
        elseif not isblank(Field_Data(xPEDate)) then
          if CDate(Date()) < CDate(Field_Data(xPEDate)) then
            SendIt_Show = false
          end if
        end if
        
        if (SendIt_Show = true and instr(1,LCase(Field_Data(xFile_Name)),".exe") = 0) and (SendIt_Show = true and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0) or ((SendIt_Show = true and CLng(Field_Flag(xArchive_Name)) = CLng(True)) and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0) then
  
          response.write "&nbsp;&nbsp;" & Translate("Email",Login_Language,conn) & ":&nbsp;&nbsp;"
      
          if instr(1,LCase(Field_Data(xFile_Name)),".exe") = 0 and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0 then
            ' Non Zip Version
            response.write "<A HREF=""javascript:void(0);"" TITLE=""Send File to an Associate"
            if CLng(Field_Flag(xFile_Name_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xFile_Name_Size) / 1024)),0) & " KBytes"
            response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),14,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
            Icon_Type = 0
            Call Write_Icon
            response.write "</A>"
          end if
      
          ' Zip Version
          if CLng(Field_Flag(xArchive_Name)) = CLng(True)  and instr(1,LCase(Field_Data(xFile_Name)),".iso") = 0 then    
            response.write "&nbsp;&nbsp;&nbsp;<A HREF=""javascript:void(0);"" TITLE=""Send ZIP File to an Associate"
            if CLng(Field_Flag(xArchive_Size)) = CLng(True) then response.write " | Size: " & FormatNumber(CDbl(CDbl(Field_Data(xArchive_Size) / 1024)),0) & " KBytes"
            response.write """ onclick=""openit('" & xLocator(Site_ID,Login_ID,Field_Data(xID),13,Language_ID,Session("Session_ID"),CID,SCID,PCID,CIN,CINN) & "','Vertical');return false;"">"
            Icon_Type = 1
            Call Write_Icon
            response.write "</A>"
          end if
          
        else
          response.write "&nbsp;"  
        end if  
      
      else
        response.write "&nbsp;"  
      end if  
    
    else
      response.write "&nbsp;"  
    end if

  end if

  response.write "</TD>" & vbCrLf

  ' Shopping Cart
  response.write "<TD HEIGHT=11 VALIGN=MIDDLE WIDTH=""14%"" CLASS=NavRecordLight>"   

  if isblank(Field_Data(xCode)) then Field_Data(xCode) = 0
  if isblank(Field_Data(xSubGroups)) then Field_Data(xSubGroups) = ""
  
  if isblank(Login_Region) then Login_Region = 1
      
  if CLng(Shopping_Cart) = CLng(True) _
     and (Field_Data(xCode) < 8000 or Field_Data(xCode) > 8999) _
     and instr(1,Field_Data(xSubGroups),"shpcrt") = 0 then
      Lit_Inventory_Rule = ""

    if not isblank(rsCalendar("Lit_Inventory_Rule")) then
      Lit_Inventory_Rule = LCase(Replace(rsCalendar("Lit_Inventory_Rule")," ","_"))
    end if  

    if   (((LCase(rsCalendar("Lit_Status")) = "active" and LCase(rsCalendar("Lit_Action")) = "complete" and LCase(rsCalendar("Lit_Status_Name")) = "final loaded") _
       or  (LCase(rsCalendar("Lit_Status")) = "active" and LCase(rsCalendar("Lit_Action")) = "complete" and LCase(rsCalendar("Lit_Status_Name")) = "reprint") _
       or  (LCase(rsCalendar("Lit_Status")) = "active" and LCase(rsCalendar("Lit_Action")) = "n/a" and LCase(rsCalendar("Lit_Status_Name")) = "final loaded") _
       or  (LCase(rsCalendar("Lit_Status")) = "active" and LCase(rsCalendar("Lit_Action")) = "n/a" and LCase(rsCalendar("Lit_Status_Name")) = "reprint"))) then

      response.write "&nbsp;&nbsp;" & Translate("Order",Login_Language,conn) & ":&nbsp;&nbsp;"
  
      Add_Cart_Item = "/sw-common/sw-shopping_cart_lit.asp"
      Add_Cart_Item = Add_Cart_Item & "?Language=" & Login_Language
      Add_Cart_Item = Add_Cart_Item & "&Action=Add&Cart_ID=" & Field_Data(xID)
      Add_Cart_Item = Add_Cart_Item & "&Lit_ID=" & rsCalendar("Lit_Item")
      Add_Cart_Item = Add_Cart_Item & "&Cart_Type="
      
      'Modified by Zensar on 28th Nov 2009 for changing the quantity to 50 from 2500 after discussion with Dennis Sims.      
      if IsNull(rsCalendar("Lit_Large_Limit")) then 
        Lit_Large_Limit = 0        
      else
        if CDBL(rsCalendar("Lit_Large_Limit")) >= 999999 then Lit_Large_Limit = 50 else Lit_Large_Limit = CDBL(rsCalendar("Lit_Large_Limit"))
      end if
      
      'Modified by Zensar on 9/18, now small limit is not getting used any more.Commenting out the below code.
      'if IsNull(rsCalendar("Lit_Small_Limit")) then
      '  Lit_Small_Limit = 0        
      'else
      '  if CDBL(rsCalendar("Lit_Small_Limit")) >= 999999 then Lit_Small_Limit =  500 else Lit_Small_Limit = CDBL(rsCalendar("Lit_Small_Limit"))
      'end if
      '***************************************
      select case Login_Region
        case 1        ' US
          Add_Cart_Item = Add_Cart_Item & "lit_us"
          if Login_Type_Code = 5 then
            Add_Cart_Item = Add_Cart_Item & "&Max_Limit=" & Lit_Large_Limit
          else
             'Modified by Zensar on 9/18, now small limit is not getting used any more.Commenting out the below code.
             ' There will be no distinction made between the Fluke Entity and Non Fluke Entity users.
             'Add_Cart_Item = Add_Cart_Item & "&Max_Limit=" & Lit_Small_Limit
             '***********************************
             Add_Cart_Item = Add_Cart_Item & "&Max_Limit=" & Lit_Large_Limit
          end if  
        case 2        ' Europe
          Add_Cart_Item = Add_Cart_Item & "lit_eu"
          Add_Cart_Item = Add_Cart_Item & "&Max_Limit=" & Lit_Large_Limit
        case 3        ' Intercon
          Add_Cart_Item = Add_Cart_Item & "lit_us"
          Add_Cart_Item = Add_Cart_Item & "&Max_Limit=" & Lit_Large_Limit
      end select
      Add_Cart_Item = Add_Cart_Item & "#Cart_List"
      Add_Cart_Item = Replace(Add_Cart_Item,"https://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))
      Add_Cart_Item = Replace(Add_Cart_Item,"http://support.fluke.com","http://" & request.ServerVariables("SERVER_NAME"))
  
      response.write "<A HREF=""javascript:void(0);"" TITLE=""Add Item Number to Shopping Cart"
      response.write """ LANGUAGE=""JavaScript"" onclick=""Shopping_Cart = window.open('" & Add_Cart_Item & "','Shopping_Cart','fullscreen=0,toolbar=0,status=0,menubar=0,scrollbars=1,resizable=1,directories=0,location=0'); Shopping_Cart.window.blur(); return false;"">"
      response.write "<IMG SRC=""/images/Button-Cart.gif"" BORDER=0 WIDTH=16 VSPACE=0 ALIGN=ABSMIDDLE>"
      response.write"</A>"
    else
      response.write "&nbsp;"
    end if    
  else
    response.write "&nbsp;"
  end if

  response.write "</TD>" & vbCrLf

  ' Language
        
  response.write "<TD HEIGHT=11 VALIGN=MIDDLE WIDTH=""21%"" CLASS=NavRecordLight>"
  response.write "&nbsp;&nbsp;" & Translate("Language",Login_Language,conn) & ":&nbsp;&nbsp;"
  
  if CLng(Field_Flag(xLanguage)) = CLng(True) then
    SQL = "SELECT * FROM Language WHERE Language.Code='" & Field_Data(xLanguage) & "' AND enable=-1"
    Set rsLanguage = Server.CreateObject("ADODB.Recordset")
    rsLanguage.Open SQL, conn, 3, 3    
    response.write Translate(rsLanguage("Description"),Login_Language,conn)
    rsLanguage.close
    Set rsLanguage = nothing
  elseif CID = 9004 then
  
    SQLLanguage = "SELECT dbo.Language.Description " &_
                  "FROM   dbo.Literature_Items_US INNER JOIN " &_
                  "       dbo.Language ON dbo.Literature_Items_US.[LANGUAGE] = dbo.Language.Code " &_
                  "WHERE dbo.Literature_Items_US.Item=" & rsCalendar("Lit_Item") & " AND dbo.Language.enable=-1"

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

    response.write Translate(Field_Data(xLanguage),Login_Language,conn)
  else
    response.write "&nbsp;"
  end if
  response.write "</TD>" & vbCrLf  

  ' Status and ID Number

  response.write "<TD HEIGHT=11 VALIGN=MIDDLE ALIGN=CENTER WIDTH=""7%"" BGCOLOR="

  if CLng(Field_Flag(xStatus)) = CLng(True) then  
    Select Case Field_Data(xStatus)
      case 1      ' Live
        if (Access_Level <= 2) then
          response.write """Silver"""
        else
          response.write """#00CC00"""
        end if  
      case 2      ' Archive
        response.write """#AAAAFF"""            
      case else   ' Pending
        response.write """Yellow"""
    end select
  
    response.write " CLASS=Small>"
  end if  

  if CLng(Field_Flag(xID)) = CLng(True) then  
    
    ' Release Task  :   668
    ' Updated by    :   Amol Jagtap
    ' Description   :   To convert Field_Data(xID) to Long from Int
    if (Access_Level = 4 or Access_Level >= 8) and CLng(Field_Data(xID)) > 0 then
      Edit_Content = "/sw-administrator/Calendar_Edit.asp?ID=" & Field_Data(xID) & "&Site_ID=" & Site_ID & "&Logon_User=" & Session("Logon_User")
      response.write "<A HREF=""javascript:void(0);"" onclick=""var MyPop1 = window.open('" & Edit_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes,directories=yes,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""" & Translate("Click to Edit this Item",Login_Language,conn) & """>" & Field_Data(xID) & "</A>"    
    elseif CLng(Field_Data(xID)) > 0 then       
      response.write Field_Data(xID)
    else
      response.write "&nbsp;"  
    end if
  else
    response.write "&nbsp;"
  end if
        
  response.write "</TD>" & vbCrLf
  response.write "</TR>" & vbCrLf    

end sub

' --------------------------------------------------------------------------------------
' Update Record Set Fields
' --------------------------------------------------------------------------------------

sub Update_Fields

  for i = 0 to Field_Max
    
    myField = Trim(rsCalendar(Field_Name(i)))
    
    if not isblank(myField) then
      Field_Data(i) = RestoreQuote(myField)
      Field_Flag(i) = CLng(True)
    else
      Field_Data(i) = ""
      Field_Flag(i) = CLng(False)
    end if
  next
  
  if Field_Flag(xCampaign) = false then
    Field_Data(xCampaign) = 0
  end if
  if Field_Flag(xCode) = false then
    Field_Data(xCode) = 0
  end if
  

end sub

' --------------------------------------------------------------------------------------
' Record Set Page Navigation
' --------------------------------------------------------------------------------------

Sub RS_Page_Navigation

  Page_QS = "Site_ID=" & Site_ID & "&Language=" & Login_Language & "&NS=" & Top_Navigation & "&CID=" & CID &  "&SCID=" & SCID & "&CIN=" & CIN & "&CINN=" & CINN & "&SortBy=" & SortBy & "&Show_Detail=" & Show_Detail & "&KeySearch=" & KeySearch & "&BolSearch=" & BolSearch & "&Language_Filter=" & Language_Filter
	if PCID = 0 then PCID = 1

  ltEnabled = 0
  
  if Record_Pages > 1 then

    Call Nav_Border_Begin
    
    response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Page", Login_Language, conn) & ": &nbsp;</SPAN>"

  	if PCID = 1 then
  		Call RS_Page_Numbers
    		response.write "<A HREF=""Default.asp?" & Page_QS & "&PCID=" & PCID + 1 & """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "&nbsp;&gt;&gt;&nbsp;</A>"
        response.write "&nbsp;&nbsp;"
  	else
  		if PCID = Record_Pages then
        ltEnabled = 1
  			response.write "<A HREF=""Default.asp?" & Page_QS & "&PCID=" & PCID - 1 & """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
        response.write "&nbsp;&lt;&lt;&nbsp</A>&nbsp;&nbsp;"
    		Call RS_Page_Numbers
  		else
        ltEnabled = 1
  			response.write "<A HREF=""Default.asp?" & Page_QS & "&PCID=" & PCID - 1 &  """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
        response.write "&nbsp;&lt;&lt;&nbsp;</A>&nbsp;&nbsp;"
    		Call RS_Page_Numbers
  			response.write "<A HREF=""Default.asp?" & Page_QS & "&PCID=" & PCID + 1 &  """ CLASS=NAVLEFTHIGHLIGHT1 TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
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
	  	response.write "<A HREF=""Default.asp?" & Page_QS & "&PCID=" & i & """ CLASS=NAVLEFTHIGHLIGHT1>"
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
			response.write "<A HREF=""Default.asp?" & Page_QS & "&PCID=" & i & """ CLASS=NavTopHighLight onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='NavTopHighLight'"">"
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

' --------------------------------------------------------------------------------------

sub Write_Icon

  if Icon_Type = 0 then
    Icon_Extension = UCase(mid(Field_Data(xFile_Name),Instr(1,Field_Data(xFile_Name),".")+1))
  else
    Icon_Extension = UCase(mid(Field_Data(xArchive_Name),Instr(1,Field_Data(xArchive_Name),".")+1))
  end if

  SQLIcon = "SELECT * FROM Asset_Type WHERE File_Extension='" & Icon_Extension & "'"
  Set rsIcon = Server.CreateObject("ADODB.Recordset")
  rsIcon.Open SQLIcon, conn, 3, 3
  
  if not rsIcon.EOF then
    response.write "<IMG SRC=""" & rsIcon("Icon_File") & """ BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"
  else
    response.write "<IMG SRC=""/images/Button-URL.gif"" BORDER=0 width=12 VSPACE=0 ALIGN=ABSMIDDLE>"      
  end if

  rsIcon.close
  set rsIcon  = nothing
  set SQLIcon = nothing
      
end sub

sub Write_Thumbnail_Icon

  Icon_Extension = UCase(mid(Field_Data(xFile_Name),Instr(1,Field_Data(xFile_Name),".")+1))

  SQLIcon = "SELECT * FROM Asset_Type WHERE File_Extension='" & Icon_Extension & "'"
  Set rsIcon = Server.CreateObject("ADODB.Recordset")
  rsIcon.Open SQLIcon, conn, 3, 3
  
  if not rsIcon.EOF then
    response.write "<IMG SRC=""" & rsIcon("Icon_File") & """ BORDER=0  VSPACE=0 ALIGN=ABSMIDDLE>"
  else
    response.write "<IMG SRC=""/images/Blank_Doc.jpg"" BORDER=0  width=80 VSPACE=0 ALIGN=ABSMIDDLE>"      
  end if

  rsIcon.close
  set rsIcon  = nothing
  set SQLIcon = nothing
      
end sub

' --------------------------------------------------------------------------------------
'Modified by zensar on 25-04-2009 for RI 506 - Price List changes
'Added variable strPath
function xLocator(Site_ID,Login_ID,Asset_ID,Method,Language_ID,Session_ID,CID,SCID,PCID,CIN,CINN)
  strPath =""
  if isblank(Login_ID) then
    LID = 1
  else
    LID = Login_ID
  end if

  if blnPriceList = true then
     strPath = "/SW-Common/SW-Find_ItPriceList.asp?SW-Locator="
  else
     strPath = "/SW-Common/SW-Find_It.asp?SW-Locator="
  end if

  xLocator = strPath       & _ 
             CStr(Site_ID)                                 & "O" & _
             CStr(LID)                                     & "O" & _
             CStr(Asset_ID)                                & "O" & _
             CStr(Method)                                  & "O" & _
             CStr(Encode_Key(Site_ID, Login_ID, Asset_ID)) & "O" & _
             CStr(CLng(Date))                              & "O" & _
             CStr(Language_ID)                             & "O" & _
             CStr(Session_ID)                              & "O" & _
             CStr(CID)                                     & "O" & _
             CStr(SCID)                                    & "O" & _
             CStr(PCID)                                    & "O" & _
             CStr(CIN)                                     & "O" & _
             CStr(CINN)

  xLocator = xLocator & "&Style=" & CStr(Site_ID)
  
  if Method = 13 or Method = 14 then
    xLocator = xLocator & "#" & "SEND_DOCUMENT"
  end if
  
end function

' --------------------------------------------------------------------------------------
' Floating View Shopping Cart Button
' --------------------------------------------------------------------------------------

if CLng(Session("Cart_Active")) = CLng(True) and CLng(Shopping_Cart) = CLng(True) and instr(1,LCase(request.servervariables("SCRIPT_NAME")),"/sw-administrator") = 0 then
  %>

  <SCRIPT TYPE="text/javascript">
    if (!document.layers) {
      document.write('<DIV ID="divStayTopRight" STYLE="position:absolute">');
    }
  </SCRIPT>
  
  <LAYER ID="divStayTopRight">
  <FORM NAME="Cart">
    <%Call Nav_Border_Begin%>
    <INPUT TYPE=BUTTON NAME="View_Cart" CLASS=NAVLEFTHIGHLIGHT1 VALUE="<%=Translate("Shopping Cart",Login_Language,conn)%>" ONCLICK="Shopping_Cart = window.open('/sw-common/sw-shopping_cart_lit.asp?Language=<%=Login_Language%>&Action=NoOp','Shopping_Cart','fullscreen=0,toolbar=0,status=0,menubar=0,scrollbars=1,resizable=1,directories=0,location=1'); return false;" TITLE="Click to View Shopping Cart">
    <%Call Nav_Border_End%>
  </FORM>
  </LAYER>
  
  <SCRIPT TYPE="text/javascript">
    var verticalpos = "fromtop"
    
    if (!document.layers) {
      document.write('</DIV>');
    }
    
    function xFloatTopDiv() {
    
      var startX;
      var startY;
      if (document.all) {
        startX = document.body.clientWidth;
      }
      else if (document.layers || document.getElementById) {
          startX = self.innerWidth; 
      }
    
      //startX = startX - 142;
      startX = startX - 160      
    //	var startX = screen.innerWidth - 190,
    	startY = 100;
    	var ns = (navigator.appName.indexOf("Netscape") != -1);
    	var d = document;
    	function ml(id)
    	{
    		var el = d.getElementById?d.getElementById(id) : d.all ? d.all[id] : d.layers[id];
    		if(d.layers)el.style = el;
    		el.sP = function(x,y){this.style.left=x; this.style.top=y;};
    		el.x  = startX;
    		if (verticalpos == "frombottom")
    		el.y  = startY;
    		else{
    		el.y  = ns ? pageYOffset + innerHeight : document.body.scrollTop + document.body.clientHeight;
    		el.y -= startY;
    		}
    		return el;
    	}
    	window.stayTopRight = function(){
    		if (verticalpos == "fromtop"){
    		var pY = ns ? pageYOffset : document.body.scrollTop;
    		ftlObj.y += (pY + startY - ftlObj.y)/8;
    		}
    		else{
    		var pY = ns ? pageYOffset + innerHeight : document.body.scrollTop + document.body.clientHeight;
    		ftlObj.y += (pY - startY - ftlObj.y)/8;
    		}
    		ftlObj.sP(ftlObj.x, ftlObj.y);
    		setTimeout("stayTopRight()", 10);
    	}
    	ftlObj = ml("divStayTopRight");
    	stayTopRight();
    }
    xFloatTopDiv();
  </SCRIPT>

<%
end if

' --------------------------------------------------------------------------------------
%>