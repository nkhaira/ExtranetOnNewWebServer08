<%

  with response
    .write "<B>" & Translate("Browser Versions Supported", Login_Language, conn) & "</B>" & vbCrLf
    .write "<BR><BR>" & vbCrLf     
    .write Translate("This site has been optimized to use the Microsoft Internet Explorer, version", Login_Language, conn)%> 5.0+. <%response.write Translate("If you would like to upgrade your browser, to use the advanced features of this site, see the Site Utilities section under the Library Navigation button.", Login_Language, conn) & vbCrLf
    %>
    <script language="JavaScript1.2">
    <!--
    var correctwidth  = 1024
    var correctheight = 768
    if (screen.width < correctwidth||screen.height < correctheight) {
    var correctmessage = "<B><%response.write Translate("Browser Screen Resolution",Login_Language,conn)%></B><BR><BR>"
    correctmessage = correctmessage + "<%response.write Translate("This site is best viewed with your screen resolution set to: ",Login_Language,conn)%>" + correctwidth + " x "+correctheight+". " + "<%response.write Translate("Your current screen resolution is set to: ",Login_Language,conn)%>" + screen.width + " x " + screen.height + ". " + "<%response.write Translate("If possible, please change your screen resolution.",Login_Language,conn)%>"
    document.write("<BR><BR>")
    document.write(correctmessage)
    }
    //-->
    </script>
    <%
  end with  

' Key Table Begin

  Number_of_Columns = 3

  with response
    .write "<BR><BR><BR>" & vbCrLf
    .write "<B>" & Translate("Content and Event Navigation Icons", Login_Language, conn) & "</B>" & vbCrLf
    .write "<BR><BR>" & vbCrLf
    .write Translate("For each content item or event, depending on the type of item, listed you have the ability to view, download, send the document to your email address as an attachment, or link to a different web page for more information. The following table represents the icons used to identify each type of file format. If you hold your mouse pointer over the icon, addition information will be displayed, related to the type of view, file size in KBytes and approximated download time depending on your connection speed to the internet that you selected in your account profile.", Login_Language, conn) & vbCrLf
    .write "<BR><BR><BR>" & vbCrLf

    Call Nav_Border_Begin

    .write "<TABLE WIDTH=""100%"" BGCOLOR=""#F3F3F3"" CELLSPACING=0 CELLPADDING=2>" & vbCrLf
  end with
  
  response.write "<TR>" & vbCrLf
  for i = 1 to Number_of_Columns
    with response
      .write "<TD WIDTH=""4%""  CLASS=PRODUCT VALIGN=MIDDLE ALIGN=CENTER>" & Translate("Icon", Login_Language, conn) & "</TD>" & vbCrLf
      .write "<TD WIDTH=""8%""  CLASS=PRODUCT VALIGN=MIDDLE ALIGN=CENTER>" & Translate("Extension", Login_Language, conn) & "</TD>" & vbCrLf
      .write "<TD WIDTH=""20%"" CLASS=PRODUCT VALIGN=MIDDLE>" & Translate("Description", Login_Language, conn) & "</TD>" & vbCrLf
    end with
    if i <> Number_of_Columns then response.write "<TD CLASS=PRODUCT VALIGN=MIDDLE WIDTH=""1%""></TD>" & vbCrLf
  next
  response.write "</TR>" 

  SQL =       "SELECT Asset_Type.* "
  SQL = SQL & "FROM Asset_Type "
  SQL = SQL & "WHERE Asset_Type.Enabled=" & CInt(True) & " "
  SQL = SQL & "ORDER BY Asset_Type.File_Type"

  Set rsAsset = Server.CreateObject("ADODB.Recordset")
  rsAsset.Open SQL, conn, 3, 3
  
  if not rsAsset.EOF then
  
    Asset_Count  = rsAsset.RecordCount
    Asset_Filler = Asset_Count
    
    do while (Asset_Filler / Number_of_Columns) > int(Asset_Filler / Number_of_Columns)
      Asset_Filler = Asset_Filler + 1
    loop  
    
    ReDim Asset_Value(2,Asset_Filler)
    
    for i = 1 to Asset_Count
      Asset_Value(0,i) = "<IMG SRC=""" & rsAsset("Icon_File") & """ WIDTH=16 BORDER=0></TD>"
      Asset_Value(1,i) = "<B>" & UCase(rsAsset("File_Extension")) & "</B>"
      Asset_Value(2,i) = Translate(rsAsset("File_Type"),Login_Language,conn)
      rsAsset.MoveNext
    next
    
    for i = Asset_Count + 1 to Asset_Filler
      Asset_Value(0,i) = "&nbsp;"
      Asset_Value(1,i) = "&nbsp;"
      Asset_Value(2,i) = "&nbsp;"
      if not rsAsset.EOF then
        rsAsset.MoveNext
      end if  
    next  
  
    for j = 1 to Int(Asset_Filler / Number_of_Columns)
    
      response.write "<TR>" & vbCrLf
  
      with response
        .write "<TD CLASS=Small VALIGN=MIDDLE ALIGN=CENTER>"  & Asset_Value(0,j) & "</TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE ALIGN=CENTER>" & Asset_Value(1,j) & "</TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE>" & Asset_Value(2,j) & "</TD>" & vbCrLf
        .write "<TD CLASS=PRODUCT VALIGN=MIDDLE></TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE ALIGN=CENTER>"  & Asset_Value(0,j+(1*(Asset_Filler / Number_of_Columns))) & "</TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE ALIGN=CENTER>" & Asset_Value(1,j+(1*(Asset_Filler / Number_of_Columns))) & "</TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE>" & Asset_Value(2,j+(1*(Asset_Filler / Number_of_Columns))) & "</TD>" & vbCrLf
        .write "<TD CLASS=PRODUCT VALIGN=MIDDLE></TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE ALIGN=CENTER>"  & Asset_Value(0,j+(2*(Asset_Filler / Number_of_Columns))) & "</TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE ALIGN=CENTER>" & Asset_Value(1,j+(2*(Asset_Filler / Number_of_Columns))) & "</TD>" & vbCrLf
        .write "<TD CLASS=Small VALIGN=MIDDLE>" & Asset_Value(2,j+(2*(Asset_Filler / Number_of_Columns))) & "</TD>" & vbCrLf
      end with
        
      response.write "</TR>" & vbCrLf
      
    next
    
    with response
      .write "</TR>" & vbCrLf
      .write "</TABLE>" & vbCrLf
      
      Call Nav_Border_End
      
      .write "<P><BR>"

    end with
    
  end if  

%>