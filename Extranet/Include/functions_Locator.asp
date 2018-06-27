<%
  ' --------------------------------------------------------------------------------------
  ' Used by /common/SW-Find_It.asp and /register/Logon.asp
  ' --------------------------------------------------------------------------------------

  xSite_ID          = 0 ' Required  (Integer - Site, ID)
  xAccount_ID       = 1 ' Required  (Integer - UserData, ID)
  xAsset_ID         = 2 ' Required  (Integer - Calendar, ID)
  xMethod           = 3 ' Required  (Integer - See Key Above
  xKey              = 4 ' Required  (Integer - Calculated Encoded/Decode Key)
  xExpiration_Date  = 5 ' Required  if method > 2 (Integer - Serial Date of Notification + 5 days.
  xLanguage         = 6 ' Optional  (Integer - Language, ID) (Default = 0 / 'eng' if not supplied)
  xSession_ID       = 7 ' Optional  (Integer - Server Variable)
  xCID              = 8 ' Optional  (Integer - Primary   Navigation Button)
  xSCID             = 9 ' Optional  (Integer - Secondary Navigation Button)
  xPCID             =10 ' Optional  (Integer - Paged     Navigation Sequence)
  xCIN              =11 ' Optional  (Integer - Content Index Number)
  xCINN             =12 ' Optional  (Integer - Content Index Sequence Number)
  xRegion           =13 ' Optional  (Integer - Sales Region)
  xCountry          =14 ' Optional  (VarChar - User Country)
  ' --------------------------------------------------------------------------------------

  Dim qString
  Dim qString_Max
  
  qString     = Split(UCase(Locator),"O")       '<---- Character UCase("O") not numeric 0
  qString_Max = Ubound(qString)
  
  Dim Parameter_Max
  Parameter_Max = 14
  Dim Parameter(14)
  Dim Parameter_Key(14)
  
  Parameter_Key(0) = "Site ID"
  Parameter_Key(1) = "Account ID"
  Parameter_Key(2) = "Asset ID"
  Parameter_Key(3) = "Method"    
  Parameter_Key(4) = "Key"
  Parameter_Key(5) = "Expiration Date"
  Parameter_Key(6) = "Language"
  Parameter_Key(7) = "Session"
  Parameter_Key(8) = "CID"
  Parameter_Key(9) = "SCID"
  Parameter_Key(10) = "PCID"
  Parameter_Key(11) = "CIN"
  Parameter_Key(12) = "CINN"
  Parameter_Key(13) = "Region"
  Parameter_Key(14) = "Country"
  
  for i = 0 to Parameter_Max - 1
    if i <= qString_Max then
      if isnumeric(qString(i)) or i = xCountry then
        Parameter(i) = qString(i)
      else
        Parameter(i) = 0
        ErrString = ErrString + "<LI>Missing or Invalid Asset Locator Parameter: " & CStr(i) & "</LI>"
      end if
    else
      select case i
        case xMethod
          Parameter(xMethod)   = 0                           ' Default to Subscription View
        case xLanguage
          Parameter(xLanguage) = 0                           ' Default to English
        case xSession_ID, xCID, xSCID, xPCID, xCIN, xCINN    ' Default
          Parameter(i)         = 0
        case xCountry
          Parameter(xCountry) = "US"
        case else
          ErrString = ErrString & "<LI>Asset Parameter Out of Range</LI>"  
      end select    
    end if
  next
  
  Parameter(Parameter_Max) = 0
    
  ' If Expiration SerialDate is not supplied, then default to todays date to allow On-Line views to work if Session_ID > 0.
  
  'if Parameter(xExpiration_Date) = 0 and Parameter(xSession_ID) > 0 then
   '  Parameter(xExpiration_Date) = CLng(Date)
  'elseif isnumeric(Parameter(xExpiration_Date)) then
   ' Parameter(xExpiration_Date)  = CLng(Parameter(xExpiration_Date))          Convert Serial Date to Date
  'end if
  Parameter(xExpiration_Date) = CLng(Date)
    
%>