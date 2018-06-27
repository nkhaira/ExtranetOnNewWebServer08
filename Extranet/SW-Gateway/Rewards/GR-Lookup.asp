<!-- #include virtual="/Connections/connection_Global_FormData.asp" -->
<!-- #include virtual="/include/functions_string.asp" -->

<%

' Comma separated list of 2-digit country ISO codes participating in the program

sCountryList = "US,CA"
sCountryList = Replace(sCountryList," ","")

if instr(1,sCountryList,",") > 0 then
  aCountry    = Split(sCountryList,",")
  iCountryMax = UBound(aCountry)
else
  Redim aCountry(0)
  aCountry    = CountryList
  iCountryMax = 0
end if

ErrorMsg = ""

if isblank(request("accountid")) or not isnumeric(request("accountid")) then
  ErrorMsg = ErrorMsg & "AccountID: Invalid format or Missing "
end if
if UCase(request("partner")) <> "FLUKE_USA" then
  ErrorMsg = ErrorMsg & "Partner: Invalid Value "
end if

if not isblank(request("accountid")) and isnumeric(request("accountid")) and Ucase(request("partner")) = "FLUKE_USA" then

  Call Connect_GlobalFormData()

  sSQL =  "SELECT FirstName, LastName, BusinessCountry FROM dbo.con_Core " &_
          "WHERE (PersonID=" & request("accountid") & ")"

  'response.write sSQL & "<P>"
  
  Set rsCore = Server.CreateObject("ADODB.Recordset")
  rsCore.Open sSQL, dbConnGlobalFormData, 3, 3
  
  if not rsCore.EOF then
  
    bCountryOK = false
    for x = 0 to iCountryMax
      if UCase(rsCore("BusinessCountry")) = UCase(aCountry(x)) then
        bCountryOK = true
        exit for
      end if
    next

    if bCountryOK = true then
      PersonID = request("accountid")
      FirstName = rsCore("FirstName")
      LastName  = rsCore("LastName")
      ErrorMsg  = ""
    else
      ErrorMsg = ErrorMsg & "Country: Account Country not valid for reward program "
    end if
  else
    ErrorMsg = "AccountID: Not Found "
  end if
  
  dbConnGlobalFormData.Close
  set dbConnGlobalFormData = nothing
  
  Call Disconnect_GlobalFormData()
  
end if

'sXML = "" & vbCrLf &_

sXML = "<?xml version=""1.0"" ?>" & vbCrLf &_
       "<customer>" & vbCrLf &_
       "  <partner-client>" & request("partner") & "</partner-client>" & vbCrLf &_
       "  <partner-client-participant>" & vbCrLf &_
       "    <user-name>" & PersonID & "</user-name>" & vbCrLf &_
       "    <password>" & PersonID & "</password>" & vbCrLf &_             
       "    <first-name>" & FirstName & "</first-name>" & vbCrLf &_
       "    <last-name>" & LastName & "</last-name>" & vbCrLf &_             
       "  </partner-client-participant>" & vbCrLf &_
       "  <partner-error>" & ErrorMsg & "</partner-error>" & vbCrLf & _
       "</customer>" & vbCrLf

response.write sXML

%>