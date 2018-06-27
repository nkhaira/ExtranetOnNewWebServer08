<%
' --------------------------------------------------------------------------------------
'
' Author: Kelly Whitlock
'
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

Site_ID = 3

Moot = GetNextGenericNumber(Site_ID)

response.write Moot
  
function GetNextGenericNumber(Site_ID)

  sSite_ID = mid("00",1,2 - Len(Trim(CStr(Site_ID)))) & Trim(CStr(Site_ID))

  Start_Number = "9" & sSite_ID & "0000" 
  End_Number   = "9" & sSite_ID & "9999"
  
  SQL = "SELECT TOP 1 L.Item_Number + 1 AS Start " &_
        "FROM dbo.Calendar L LEFT OUTER JOIN " &_
        "     dbo.Calendar R ON L.Item_Number + 1 = R.Item_Number " &_
        "WHERE (L.Item_Number >= " & Start_Number & ") AND (R.Item_Number IS NULL) AND (L.Item_Number <= " & End_Number & ") " &_
        "ORDER BY L.Item_Number"

  Set rsGeneric = Server.CreateObject("ADODB.Recordset")
  rsGeneric.Open SQL, conn, 3, 3
  
  if not rsGeneric.EOF then
    Item_Number = rsGeneric("Start")
  else
    Item_Number = -1
  end if
  
  rsGeneric.close
  set rsGeneric = nothing
  
  GetNextGenericNumber = Item_Number
          
end function











Call Disconnect_SiteWide
%>