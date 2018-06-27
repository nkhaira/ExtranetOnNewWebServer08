<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/include/adovbs.inc"-->
<%

Call Connect_Sitewide

SQL_UserData = "SELECT * FROM UserData Where ID=110"
  Set rsUserData = Server.CreateObject("ADODB.Recordset")
  rsUserData.Open SQL_UserData, conn, 3, 3

int_rsUserDataCount = rsUserData.Fields.Count

response.write int_rsUserDataCount & "<P>"

response.write "<TABLE><TR>"
response.write "<TD>Index</TD><TD>Field Name</TD><TD>Type</TD><TD>Size</TD><TD>SQL</TD>"
response.write "</TR>"

index = 0
for each field In rsUserData.Fields
   response.write "<TR>"
   response.write "<TD>" & index & "</TD>"   
   response.write "<TD>" & field.Name & "</TD>"
   response.write "<TD>" & field.Type & "</TD>"
   response.write "<TD>" & field.DefinedSize & "</TD>"
'   response.write "<TD>" & field.Value & "</TD>"
   response.write "<TD>"
   response.write Construct_SQLKeyValueSet(field.Name, field.Type, field.Value, True)
   response.write "</TD>"
   response.write "</TR>"
   
   index = index + 1
Next 

response.write "</TABLE>"

Call Disconnect_Sitewide

function Construct_SQLKeyValueSet(myFieldName, myFieldType, myFieldData, NullFlag)

  Dim tempName, tempData, tempType

  tempName = myFieldName
  tempData = myFieldData
  tempType = myFieldType
  
  
  select case myFieldType
    ' Number
    case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, _
         adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt, adSingle, _
         adDouble, adCurrency, adDecimal, adNumeric, adBoolean, _
         adBinary, adVarBinary, adLongVarBinary, adVarNumeric
         
         if isnull(myFieldData) and NullFlag = True then
           tempData = "NULL"
         end if  

    ' String
    case adVariant, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, _
         adLongVarWChar, adChapter

         if (myFieldData = "" or isnull(myFieldData)) and NullFlag = True then
           tempData = "NULL"
         else 
           tempData = "'" & Replace(tempData,"'","''") & "'"
         end if  
    
    ' Date     
    case adFileTime, adDBFileTime, adDate, adDBDate, adDBTime, adDBTimeStamp, _
         adBSTR

         if (myFieldData = "" or isnull(myFieldData)) and NullFlag = True then
           tempData = "NULL"
         else
           tempData = "'" & Replace(tempData,"'","''") & "'"
         end if  
             
    ' Undefined     
    case adPropVariant, adError, adUserDefined, adIDispatch, adIUnknown, _
         adGUID, adEmpty
         
         tempData = "Invalid Data Type"
         
  end select
  
  Construct_SQLKeyValueSet = tempName & "=" & tempData
  
end function

%> 