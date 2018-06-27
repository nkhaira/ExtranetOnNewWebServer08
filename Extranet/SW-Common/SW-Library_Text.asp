<%
response.write "<BR>"
response.write Translate("The " & Site_Description & " digital library provides a wealth of information at your fingertips in support of these products.",Login_Language,conn)
response.write "<P>"
response.write Translate("Specific information can be found by selecting one of the categories in the left navigation menu.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("The following is a brief help section that describes how the inner navigation functions for these category page listings:",Login_language,conn) & "<P>"

response.write "<FORM NAME=""Library_Help"">"
response.write "<TABLE WIDTH=""100%"" CELLPADDING=6 CELLSPACING=0>"

' Quick Find
response.write "<TR>"
response.write "<TD Class=Small>"
Call Nav_Border_Begin
response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Find",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf
response.write "<SELECT CLASS=Small>" & vbCrLf
response.write "<OPTION CLASS=Small VALUE="""">-- " & Translate("Select Document",Login_Language,conn) & " --</OPTION>" & vbCrLf
response.write "        </SELECT>" & vbCrLf
Call Nav_Border_End
response.write "</TD>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Find",Login_Language,conn) & "</B> - "
response.write Translate("Quick Find is a list of all items contained within this category organized alphabetically by subcategory, then by the title of each library item.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("Use the dropdown list to scroll down to find your library item of interest.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("Once a library item has been selected, your browser will refresh to the page containing that library item.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"

' Sort By
response.write "<TR>"
response.write "<TD Class=Small>"
Call Nav_Border_Begin
response.write "<SPAN CLASS=SmallBoldGold>" & Translate("Sort",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf         
response.write "<SELECT NAME=""SortBy"" CLASS=Small>" & vbCrLf
response.write "<OPTION>" & Translate("Product",Login_Language,conn) & "</OPTION>" & vbCrLf                      
response.write "<OPTION>" & Translate("Category",Login_Language,conn) & "</OPTION>" & vbCrLf                      
response.write "<OPTION>" & Translate("Date",Login_Language,conn) & "</OPTION>" & vbCrLf                      
response.write "</SELECT>" & vbCrLf
Call Nav_Border_End
response.write "</TD>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Sort",Login_Language,conn) & "</B> - "
response.write Translate("Allows you to re-sort the listing by Category/Subcategory (default), in alphabetical order by the product or product series that the item references, or by descending date in which the item was added to the library.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("To quickly find the most recent items that have been added to this category select, Sort: Date, or visit the What's New section of this site for a complete listing.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"

' Detail or Title
response.write "<TR>"
response.write "<TD Class=Small>"
Call Nav_Border_Begin
response.write "<SPAN CLASS=SmallBoldGold>" & Translate("View",Login_Language,conn) & ":&nbsp;</SPAN>" & vbCrLf
response.write "<SELECT Class=Small>" & vbCrLf                      
response.write "<OPTION>" & Translate("Active",Login_Language,conn) & " + " & Translate("Detail",Login_Language,conn) & "</OPTION>" & vbCrLf                      
response.write "<OPTION>" & Translate("Active",Login_Language,conn) & " + " & Translate("Title",Login_Language,conn) & "</OPTION>" & vbCrLf                      
response.write "</SELECT>" & vbCrLf
Call Nav_Border_End
response.write "<TD Class=Small>"
response.write "<B>" & Translate("View",Login_Language,conn) & "</B> - "
response.write Translate("Allows you to see all the detail information about the item or restrict the listing to only the title of the item.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"

' Headerliner
response.write "<TR>"
response.write "<TD Class=NormalBold COLSPAN=2><P>&nbsp;<BR>"
if Site_ID = 24 then
  response.write Translate("Available anytime, you can use the Educator's Portal to view, download, email to yourself, send to a colleague or student, PDF document files of product information, application notes and other classroom materials.",Login_Language,conn)
else
  response.write Translate("Available anytime, you can use the Partner Portal to view, download, email to yourself, send to a colleague or customer, PDF document files of product information, application notes, etc.",Login_Language,conn)
end if  
if CInt(Shopping_Cart) = CInt(True) then
  response.write "<P>" & Translate("In addition, you can order printed copies of literature items, by simply adding the item to your shopping cart.",Login_Language,conn) & "&nbsp;&nbsp;" & Translate("Just look for the following icon.",Login_Language,conn) 
  response.write "&nbsp;&nbsp;<IMG SRC=""/images/Button-Cart.gif"" ALT=""Add to Shopping Cart"" BORDER=0 WIDTH=16 HEIGHT=12 VSPACE=0 ALIGN=ABSMIDDLE>&nbsp;&nbsp;" & Translate("Literature items in reasonable quantities are at no charge, unless otherwise noted.",Login_Language,conn)
end if  
response.write "</TD>"
response.write "</TR>"

' Item Link Bar
response.write "<TR>"
response.write "<TD Class=Medium COLSPAN=2><P>&nbsp;<BR>"
response.write Translate("At the bottom of each item, there is a link bar that will allow you to View / Download / Email or go to a web page on a different site.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("Moving your mouse over the icon will display the physical size of the file in KBytes so that you can estimate download time.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("These links have some special functionality that is described below:",Login_Language,conn) & "<P>"
response.write "</TD>"
response.write "</TR>"

' View
response.write "<TR>"
response.write "<TD Class=Small>"
response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
response.write Translate("View",Login_Language,conn) & ":&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-PDF.gif"" ALT=""View File \ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "</TD></TR></TABLE>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("View",Login_Language,conn) & "</B> - "
response.write Translate("Retrieves the item from the site and attempts to display it on your browser, provided you have the required plug-in for viewing this item type.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("If you do not have the required plug-in, the browser will display the Save As... file dialog box to allow you to save this item to your computer for off-line viewing.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"

' Download
response.write "<TR>"
response.write "<TD Class=Small>"
response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
response.write Translate("Download",Login_Language,conn) & ":&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-PDF.gif"" ALT=""Download File \ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-ZIP.gif"" ALT=""Download ZIP File \ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "</TD></TR></TABLE>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Download",Login_Language,conn) & "</B> - "
response.write Translate("Retrieves the item from the site allows you to select whether or not you want to Open (display it on your browser, provided you have the required plug-in for viewing), or Save As... (opens the file dialog box to allow you to save this item to your computer) for off-line viewing.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("In addition to the first icon which represents the item's native format, a second icon may appear to denote that this item is also available in ZIP/Archive format.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("The ZIP/Archive format should be selected, if your company's firewall prohibits or blocks access to the file in its native format.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"

' Send
response.write "<TR>"
response.write "<TD Class=Small>"
response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
response.write Translate("Send",Login_Language,conn) & ":&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-PDF.gif"" ALT=""Email File to Yourself \ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-ZIP.gif"" ALT=""Email Zip File to Yourself \ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "</TD></TR></TABLE>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Send",Login_Language,conn) & "</B> - "
response.write Translate("Retrieves the item from the site send the file as an attachment to your email address.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("The Email option frees you from waiting to download the item if you have a slow internet connection.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("In addition to the first icon which represents the item's native format, a second icon may appear to denote that this item is also available in ZIP/Archive format.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("The ZIP/Archive format should be selected, if your company's firewall prohibits or blocks access to the file in its native format.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"

' Email
response.write "<TR>"
response.write "<TD Class=Small>"
response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
response.write Translate("Email",Login_Language,conn) & ":&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-PDF.gif"" ALT=""Email File to an Associate \ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-ZIP.gif"" ALT=""Email ZIP File to an Associate\ Size: ### KBytes"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "</TD></TR></TABLE>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Email",Login_Language,conn) & "</B> - "
response.write Translate("Retrieves the item from the site send the file as an attachment or a link to the file to an email address you specify.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("The Email option frees you from waiting to download the item if you have a slow internet connection.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("In addition to the first icon which represents the item's native format, a second icon may appear to denote that this item is also available in ZIP/Archive format.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("The ZIP/Archive format should be selected, if your company's firewall prohibits or blocks access to the file in its native format.",Login_Language,conn)
response.write Translate("This is a great tool for sending product information documents or other information directly to your customers or other associates.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"


' Link
response.write "<TR>"
response.write "<TD Class=Small>"
response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
response.write Translate("Link",Login_Language,conn) & ":&nbsp;&nbsp;"
response.write "<IMG SRC=""/images/Button-URL.gif"" ALT=""View URL"" BORDER=0 WIDTH=11 VSPACE=0 ALIGN=ABSMIDDLE>"
response.write "</TD></TR></TABLE>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Link",Login_Language,conn) & "</B> - "
response.write Translate("A Link icon denotes that the item you wish to view is located on another page of this website, or is located on another external website.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write "</TD>"
response.write "</TR>"

' Language
response.write "<TR>"
response.write "<TD Class=Small>"
response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
response.write Translate("Language",Login_Language,conn) & ":&nbsp;&nbsp;" & Translate("English",Login_Language,conn)
response.write "</TD></TR></TABLE>"
response.write "<TD Class=Small>"
response.write "<B>" & Translate("Language",Login_Language,conn) & "</B> - "
response.write Translate("Denotes that language that the item is written in.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("Items are always displayed in their English form and if you have choosen to view the site in an alternate &quot;Preferred Language&quot; then items matching your &quot;Preferred Language&quot; setting will also be displayed if available.",Login_Language,conn)
response.write "</TD>"
response.write "</TR>"


if CInt(Shopping_Cart) = CInt(True) then

  ' Shopping Cart
  response.write "<TR>"
  response.write "<TD Class=Small>"
  response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0><TR><TD CLASS=SmallBoldWhite WIDTH=""100%"">"
  response.write Translate("Shopping Cart",Login_Language,conn) & ":&nbsp;&nbsp;"
  response.write "<IMG SRC=""/images/Button-Cart.gif"" ALT=""Add to Shopping Cart"" BORDER=0 WIDTH=16 HEIGHT=12 VSPACE=0 ALIGN=ABSMIDDLE>"
  response.write "</TD></TR></TABLE>"
  response.write "<TD Class=Small>"
  response.write "<B>" & Translate("Shopping Cart",Login_Language,conn) & "</B> - "
  response.write Translate("A Shopping Cart icon denotes that the item is available for ordering in quantity through the Literature Ordering System.",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write Translate("Once you have 1 or more items in your Shopping Cart, you can access your Shopping Cart at any time by clicking on the [Shopping Cart] button located on the right side of your screen.",Login_Language,conn)
  response.write "</TD>"
  response.write "</TR>"
  
end if

if Access_Level = 4 or Access_Level >= 8 then

  ' Shopping Cart
  response.write "<TR>"
  response.write "<TD Class=Small>"
  
  response.write "<TABLE WIDTH=""100%"" BGCOLOR=""Silver"" BORDER=0 CELLSPACING=0 CELLPADDING=0><TR>"

  ' Review
  response.write "<TD CLASS=Small ALIGN=CENTER  WIDTH=""33%"" BGCOLOR="""
  if Access_Level = 0 then
    response.write "Silver"
  else
    response.write "Yellow"
  end if
  response.write """><U>1234</U>"
  response.write "</TD>"


  ' Live
  response.write "<TD CLASS=Small ALIGN=CENTER  WIDTH=""33%"" BGCOLOR="""
  if Access_Level = 0 then
    response.write "Silver"
  else
    response.write "#00CC00"
  end if
  response.write """><U>1234</U>"
  response.write "</TD>"

  ' Archive
  response.write "<TD CLASS=Small ALIGN=CENTER WIDTH=""33%"" BGCOLOR="""
  if Access_Level = 0 then
    response.write "Silver"
  else
    response.write "#AAAAFF"
  end if
  response.write """><U>1234</U>"
  response.write "</TD>"

  response.write "</TR>"
  response.write "</TABLE>"
  response.write "</TD>"
  
  response.write "<TD Class=Small>"
  response.write "<B>" & Translate("Administration",Login_Language,conn) & "</B> - "
  response.write Translate("Access to the item's Content Container is accessible by clicking on the underlined item ID number.",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write Translate("This will bring up a pop-up edit window for this item.",Login_Language,conn) & "&nbsp;&nbsp;"
  response.write Translate("You can quickly determine the status of an item by noting the background color.  Yellow = Review, Green = Live and Lavendar = Archive.",Login_Language,conn)
  response.write "</TD>"
  response.write "</TR>"
  
end if

' Help
response.write "<TR>"
response.write "<TD Class=Medium COLSPAN=2><P>&nbsp;<BR>"
response.write Translate("For additional help with this site or site features or if you would like to view a listing of file/icon types, select the [Help] navigation button on the left side navigation menu.",Login_Language,conn) & "&nbsp;&nbsp;"
response.write Translate("Browser Helper Utilities or links to Browser Plug-Ins can be found in the [Site Utility] section of the [Library].",Login_Language,conn) & "<P>"
response.write "</TD>"
response.write "</TR>"

response.write "</TABLE>"
response.write "</FORM>"












%>
