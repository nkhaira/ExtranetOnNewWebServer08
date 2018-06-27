<%
  Spacer = False
  response.write "<SELECT NAME=""SC2"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Account_List.asp?Roopy=Yes&Site_ID=" & Site_ID & "&SC0=1" & "&SC1=0" & "&SC3=0" & "&SC4=" & SC4 & "&SC9=" & SC9 & "&SCA=" & SCA & "&Page=1" & "&SC2=" & "'+this.options[this.selectedIndex].value"">" & vbCrLf

  response.write "<OPTION CLASS=Small>" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
  
  SQL = "SELECT UserData.NewFlag, UserData.Site_ID FROM UserData WHERE UserData.NewFlag <>" & CInt(False) & " AND UserData.Site_ID=" & Site_ID
  Set rsNewFlag = Server.CreateObject("ADODB.Recordset")
  rsNewFlag.Open SQL, conn, 3, 3
  
  if not rsNewFlag.EOF then
    response.write "<OPTION CLASS=Region4NavSmall VALUE=""0"""
    if SC2 = 0 then response.write " SELECTED"
    response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Pending Approval",Login_Language,conn) & " - " & Translate("New",Login_Language,conn) & "</OPTION>" & vbCrLf
    Spacer = True
  end if

  rsNewFlag.close
  set rsNewFlag = nothing

  SQL = "SELECT UserData.ExpirationDate, UserData.Site_ID FROM UserData WHERE UserData.ExpirationDate<='" & Date & "' AND UserData.Site_ID=" & Site_ID
  Set rsExpirationDate = Server.CreateObject("ADODB.Recordset")
  rsExpirationDate.Open SQL, conn, 3, 3

  if not rsExpirationDate.EOF then
    response.write "<OPTION CLASS=Region4NavSmall VALUE=""7"""
    if SC2 = 7 then response.write " SELECTED"
    response.write ">" & Translate("Account",Login_Language,conn) & " - " &  Translate("Expired",Login_Language,conn) & "</OPTION>" & vbCrLf
    Spacer = True
  end if
  
  rsExpirationDate.close
  set rsExpirationDate = nothing

  SQL = "SELECT UserData.Reg_Approval_Date, UserData.Site_ID FROM UserData WHERE UserData.Reg_Approval_Date='" & Date & "' AND UserData.Site_ID=" & Site_ID
  Set rsRAD = Server.CreateObject("ADODB.Recordset")
  rsRad.Open SQL, conn, 3, 3

  if not rsRAD.EOF then
    response.write "<OPTION CLASS=Region1NavSmall VALUE=""1"""
    if SC2 = 1 then response.write " SELECTED"
    response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Updated Today",Login_Language,conn) & "</OPTION>" & vbCrLf
    Spacer = True
  end if
  
  rsRAD.close
  set rsRAD = nothing

  if Spacer = true then
    response.write "<OPTION CLASS=Small></OPTION>" & vbCrLf
  end if
  
  response.write "<OPTION CLASS=Region5NavSmall VALUE=""15"""
  if SC2 = 15 then response.write " SELECTED"
  response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("All",Login_Language,conn) & vbCrLf

  response.write "<OPTION CLASS=Region5NavSmall VALUE=""2"""
  if SC2 = 2 then response.write " SELECTED"
  response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Fluke",Login_Language,conn) & " - " & Translate("Excluded",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=Region5NavSmall VALUE=""3"""
  if SC2 = 3 then response.write " SELECTED"
  response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Fluke",Login_Language,conn) & " - " & Translate("Only",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=Region1NavSmall VALUE=""4"""
  if SC2 = 4 then response.write " SELECTED"
  response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Region 1 - US Only",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=Region2NavSmall VALUE=""5"""
  if SC2 = 5 then response.write " SELECTED"
  response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Region 2 - Europe Only",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=Region3NavSmall VALUE=""6"""
  if SC2 = 6 then response.write " SELECTED"
  response.write ">" & Translate("Account",Login_Language,conn) & " - " & Translate("Region 3 - Intercon Only",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=Small></OPTION>" & vbCrLf

  response.write "<OPTION CLASS=NavLeftHighlight1 VALUE=""90"""
  if SC2 = 90 then response.write " SELECTED"
  response.write ">"  & Translate("Account",Login_Language,conn) & " - " & Translate("Search",Login_Language,conn) & "</OPTION>" & vbCrLf

  if SC2 = 91 then
    response.write "<OPTION CLASS=NavLeftHighlight1 VALUE=""91"""
    if SC2 = 91 then response.write " SELECTED"
    response.write ">"  & Translate("Account",Login_Language,conn) & " - " & Translate("Search Results",Login_Language,conn) & "</OPTION>" & vbCrLf
  end if

  response.write "<OPTION CLASS=Small></OPTION>" & vbCrLf
  
  response.write "<OPTION CLASS=Region4NavSmall VALUE=""99"""
  if SC2 = 99 then response.write " SELECTED"
  response.write ">" & Translate("Add New Account Profile",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=Small></OPTION>" & vbCrLf
  
  response.write "<OPTION CLASS=Region2NavSmall VALUE=""42"""
  if SC2 = 42 then response.write " SELECTED"
  response.write ">" & Translate("Branch Edit",Login_Language,conn) & "</OPTION>" & vbCrLf
  
  response.write "<OPTION CLASS=Region2NavSmall VALUE=""43"""
  if SC2 = 43 then response.write " SELECTED"
  response.write ">" & Translate("Edit Discounts",Login_Language,conn) & "</OPTION>" & vbCrLf
  

  response.write "<OPTION CLASS=Small></OPTION>" & vbCrLf
 
  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""40"""
  if SC2 = 40 then response.write " SELECTED"
  response.write ">" & Translate("Order Inquiry Administrator",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""41"""
  if SC2 = 41 then response.write " SELECTED"
  response.write ">" & Translate("Order Inquiry Search",Login_Language,conn) & "</OPTION>" & vbCrLf
  
  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""8"""
  if SC2 = 8 then response.write " SELECTED"
  response.write ">" & Translate("Account Manager",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""20"""
  if SC2 = 20 then response.write " SELECTED"
  response.write ">" & Translate("Forum Moderator",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""9"""
  if SC2 = 9 then response.write " SELECTED"
  response.write ">" & Translate("Content Submitter",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""10"""
  if SC2 = 10 then response.write " SELECTED"
  response.write ">" & Translate("Content Administrator",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""11"""
  if SC2 = 11 then response.write " SELECTED"
  response.write ">" & Translate("Content Administrator - Matrix",Login_Language,conn) & "</OPTION>" & vbCrLf
  
  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""12"""
  if SC2 = 12 then response.write " SELECTED"
  response.write ">" & Translate("Account Administrator",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""13"""
  if SC2 = 13 then response.write " SELECTED"
  response.write ">" & Translate("Account Administrator - Matrix",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "<OPTION CLASS=RegionXNavSmall VALUE=""14"""
  if SC2 = 14 then response.write " SELECTED"
  response.write ">" & Translate("Site Administrator",Login_Language,conn) & "</OPTION>" & vbCrLf

  response.write "</SELECT>" & vbCrLf
  
  if SC2 = 0 then
  
    response.write "&nbsp;&nbsp;" & Translate("Filter",Login_Language,conn) & ": " & vbCrLf
    response.write "<SELECT NAME=""SC9"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Account_List.asp?Roopy=Yes&Site_ID=" & Site_ID & "&SC0=1" & "&SC1=0" & "&SC3=0" & "&SC4=" & SC4 & "&Page=1" & "&SC2=" & SC2 & "&SC9=" & "'+this.options[this.selectedIndex].value"">" & vbCrLf     

    response.write "<OPTION CLASS=Small VALUE=""0"""
    if SC9 = 0 then response.write " SELECTED"
    response.write ">" & Translate("All Regions",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Region1NavSmall VALUE=""1"""
    if SC9 = 1 then response.write " SELECTED"
    response.write ">" & Translate("United States",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Region2NavSmall VALUE=""2"""
    if SC9 = 2 then response.write " SELECTED"
    response.write ">" & Translate("Europe",Login_Language,conn) & "</OPTION>" & vbCrLf

    response.write "<OPTION CLASS=Region3NavSmall VALUE=""3"""
    if SC9 = 3 then response.write " SELECTED"
    response.write ">" & Translate("Intercon",Login_Language,conn) & "</OPTION>" & vbCrLf
  
    response.write "</SELECT>" & vbCrLf

    response.write "&nbsp;&nbsp;" '& Translate("Country",Login_Language,conn) & ": " & vbCrLf
    
    ' Return Countries that have pending accounts for site
        
    SQLCountry = "SELECT DISTINCT dbo.Country.Abbrev, dbo.Country.Name, dbo.Country.Region, dbo.Country.Enable " &_
                 "FROM            dbo.Country RIGHT OUTER JOIN " &_
                 "                dbo.UserData ON dbo.Country.Abbrev = dbo.UserData.Business_Country " &_
                 "WHERE          (dbo.Country.Enable = - 1) AND (dbo.UserData.NewFlag = - 1) AND (dbo.UserData.Site_ID = " & Site_ID & ") " &_
                 "ORDER BY dbo.Country.Name"
                 
    Set rsCountry = Server.CreateObject("ADODB.Recordset")
    rsCountry.Open SQLCountry, conn, 3, 3

    response.write "<SELECT NAME=""SCA"" CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Account_List.asp?Roopy=Yes&Site_ID=" & Site_ID & "&SC0=1" & "&SC1=0" & "&SC3=0" & "&SC4=" & SC4 & "&Page=1" & "&SC2=" & SC2 & "&SC9=" & SC9 & "&SCA=" & "'+this.options[this.selectedIndex].value"" TITLE=""Countries that have Pending Accounts"">" & vbCrLf         

    response.write "<OPTION CLASS=Small VALUE="""""
    if SCA = "" then response.write " SELECTED"
    response.write ">" & Translate("All Countries",Login_Language,conn) & "</OPTION>" & vbCrLf

    do while not rsCountry.EOF
      response.write "<OPTION "
      select case rsCountry("Region")
        case 1
          response.write "CLASS=Region1NavSmall "
        case 2
          response.write "CLASS=Region2NavSmall "
        case 3
          response.write "CLASS=Region3NavSmall "
        case else
          response.write "CLASS=RegionXNavSmall "
      end select
      response.write "VALUE=""" & rsCountry("Abbrev") & """ "
      if SCA = rsCountry("Abbrev") then
        response.write "Selected "
      end if
      response.write ">" & rsCountry("Name") & "</OPTION>" & vbCrLf
      rsCountry.MoveNext
    loop
    response.write "</SELECT>" & vbCrLf
    
    rsCountry.close
    set rsCount = nothing
          
  end if      
%>
