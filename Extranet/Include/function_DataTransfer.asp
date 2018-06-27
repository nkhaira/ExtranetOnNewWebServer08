<%

' Used by Register_Admin.asp, SW-Profile_Admin.asp, Account_Admin.asp

sub Send_2_CMS

  SQL_CMS = "SELECT CMS_System.* FROM CMS_System WHERE CMS_System.Site_ID=" & Site_ID & " AND CMS_System.Region_ID=" & User_Region & " AND CMS_System.Enabled <> 0"

  Set rsCMS = Server.CreateObject("ADODB.Recordset")
  rsCMS.Open SQL_CMS, conn, 3, 3

  if not rsCMS.EOF then   ' Active Posting Information

    ' Clean up and Filter Query String

    if len(strPost_QueryString) = 0 then
      strPost_QueryString = "NoKey='NoValue'"
    end if
    
    'response.write "A|" & strPost_QueryString & "|<P>"
    
    ' Trim SQL Prefix "UPDATE **** SET " if supplied
    strPost_TrimPos = Instr(UCase(Mid(strPost_QueryString,1,instr(strPost_QueryString,"=")-1))," SET ")
    if strPost_TrimPos > 0 then
      strPost_QueryString = Mid(strPost_QueryString,strPost_TrimPos + 5)
    end if  

    ' Trim SQL Suffix " WHERE " if supplied
    if InStrRev(strPost_QueryString, " WHERE ") then   
      strPost_QueryString = Trim(Mid(strPost_QueryString,1,InStrRev(strPost_QueryString, " WHERE ")))
    end if

    ' Remove NVarchar Delimiter for Strings
    strPost_QueryString = Replace(strPost_QueryString,"=N'","='")
    strPost_QueryString = Replace(strPost_QueryString,"=n'","='")

    strPost_QueryString = replace(strPost_QueryString, "=''", "=NULL")    
    strPost_QueryString = replace(strPost_QueryString,"&acute;","''")
    strPost_QueryString = replace(strPost_QueryString,"&quot;","""")  

    ' Add Action Dependant Fields 

    strPost_QueryString  = strPost_QueryString & ",Action=" & Action
    strPost_QueryString  = strPost_QueryString & ",Account_ID=" & New_Account_ID
    'strPost_QueryString  = strPost_QueryString & ",ID=" & New_Account_ID    
    
    'response.write "B|" & strPost_QueryString & "|<P>"
     
    ' Bar "|" is used as an array delemeter, Substitute for comma before sending using Replace()
    strPost_QueryString = Replace(URLEncode(Trim(strPost_QueryString), "=", "'", ","),"|",", ")
    
    ' Sort Array so values from different forms at least arrive to the receiver in the same order
    str_sort_Post_QueryString = Replace(strPost_QueryString,"&",",")
    
    'response.write "C|" & strPost_QueryString & "|<P>"
    
    str_sort_Post_QueryString = Split(str_sort_Post_QueryString,",")
    call SortArray(str_sort_Post_QueryString)
    strPost_QueryString = ""
    for x = 0 to UBOUND(str_sort_Post_QueryString)
      strPost_QueryString = strPost_QueryString & "&" & str_sort_Post_QueryString(x)
    next    

    strPost_QueryString = mid(strPost_QueryString,2)
    
    'response.write "D|" & strPost_QueryString & "|<P>"
    'response.write Replace(strPost_QueryString,"&","&<BR>") & "<P>"
    'response.flush
    'response.end
    
    ' HTTP_Post Setup Parameters

    Dim strKeyValueDelimiter, strPairDelimeter, strReferrerFile, strResponse, bResponse

    strKeyValueDelimiter = "="
    strPairDelimiter     = "&"
    strReferrerFile      = request.ServerVariables("Script_Name")
    strResponse          = ""

    bResponse = PostDataToServer(strPost_QueryString, strKeyValueDelimiter, strPairDelimiter, Site_ID, User_Region, strReferrerFile, strResponse)

    if (CInt(bResponse) = CInt(True) or LCase(bResponse) = "true") and Instr(1,strResponse,"200 OK") > 0 then
      if Send_2_CMS_Debug = True then
        ErrorMessage = "(1) CMS_System Post Successful:<BR><BR>"
        ErrorMessage = ErrorMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR><BR>"
        ErrorMessage = ErrorMessage & strResponse & "<BR>"
        response.flush
      end if
    elseif (CInt(bResponse) = CInt(True) or LCase(bResponse) = "true") and Instr(1, strResponse,"100 Continue") > 0 then  
      if Send_2_CMS_Debug = True then
        ErrorMessage = "(1) CMS_System Post Successful:<BR><BR>"
        ErrorMessage = ErrorMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR><BR>"
        ErrorMessage = ErrorMessage & strResponse & "<BR>"
        response.flush
      end if
    elseif CInt(bResponse) = CInt(False) or LCase(bResponse) = "false" then
      ErrorMessage = "(2) CMS_System Post Error (" & bResponse & "):<BR><BR>"
      ErrorMessage = ErrorMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR><BR>"
      ErrorMessage = ErrorMessage & strResponse & "<BR>"
      response.flush
    else
      ErrorMessage = "(3) CMS_System Post Error (" & bResponse & "):<BR><BR>"
      ErrorMessage = ErrorMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR><BR>"
      ErrorMessage = ErrorMessage & strResponse & "<BR>"
      response.flush
    end if
    
  else
  
    '  Not active Posting Information Error Handler
    
  end if
  
  rsCMS.close
  set rsCMS = nothing  
  
end sub

function SortArray(arrArray)

  Dim row, j
  Dim StartingKeyValue, StartingCseValue, NewKeyValue, NewCseValue, swap_pos

  for row = 0 To UBound(arrArray) - 1
    'Take a snapshot of the first element
    'in the array because if there is a 
    'smaller value elsewhere in the array 
    'we'll need to do a swap.
    StartingKeyValue = UCase(arrArray(row))
    StartingCseValue = arrArray(row)
    NewKeyValue      = UCase(arrArray(row))
    NewCseValue      = arrArray(row)
    swap_pos         = row
	    	
    for j = row + 1 to UBound(arrArray)
      'Start inner loop.
      if UCase(arrArray(j)) < NewKeyValue then
        'This is now the lowest number - remember it's position.
        swap_pos = j
        NewKeyValue = UCase(arrArray(j))
        NewCseValue = arrArray(j)
      end if
    next
	    
    if swap_pos <> row then
      'If we get here then we are about to do a swap within the array.		
      arrArray(swap_pos) = StartingCseValue
      arrArray(row) = NewCseValue
    end if	
  next
  
  SortArray = arrArray
  
end function

%>