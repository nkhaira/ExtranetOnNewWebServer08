<%
if not AFOVBS_LOADED then
  %>
  <!--#Include Virtual="/include/Adovbs.inc" -->
  <%
end if  

' --------------------------------------------------------------------------------------

function Get_New_Record_ID (myTable, myField, myValue, connection)

  ' Creates new Record, then Returns ID Number for UPDATE
 
  set objRS = Server.CreateObject("ADODB.Recordset")

  with objRS
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
  	.ActiveConnection = connection
  end with

  objRS.Open myTable,,,,adCmdTable
  objRS.AddNew
  objRS(myField) = myValue
  objRS.Update
  objRS.Close
  objRS.open "select @@identity", connection, 3,3
  my_ID = objRS(0)
  set objRS = Nothing

  Get_New_Record_ID = my_ID
  
end function

' --------------------------------------------------------------------------------------

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