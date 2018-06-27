<%
' --------------------------------------------------------------------------------------
' Receiving Script Template
' Kelly Whitlock
' 09/17/2002
' --------------------------------------------------------------------------------------

Dim Verbose
Dim ErrorMesg, ErrorCode

Verbose   = true  ' Debug Mode
ErrorMesg = ""
ErrorCode = 0

if request.form("Action") <> "" then
  Action = request.form("Action")
elseif request("Action") <> "" then
  Action = request("Action")
else
  Action = ""
end if  

' --------------------------------------------------------------------------------------

with response

  .write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf
  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf
  .write "	<TITLE>DCG Receiving Script from Fluke Partner Portal</TITLE>" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY>" & vbCrLf
  .write "<?xml version=""1.0""?>"
  .write vbCrLf
  
  ' --------------------------------------------------------------------------------------
  ' MAIN Select Action and perform method
  ' --------------------------------------------------------------------------------------
  
  select case LCase(Action)
  
    case "order_post"
  
      ' --------------------------------------------------------------------------------------
      ' Data is received from the Fluke Externet Site as a HTTP POST
      ' Map response.form("") elements to DCG DB fields and set POST_OK flag below.
      ' Set Verbose = True above to enumerate during development and debug
      ' --------------------------------------------------------------------------------------
      ' DCG code begin
      ' --------------------------------------------------------------------------------------
  
      ' Decode Account Info Example
      Account_Info = Split(request.form("SCUSID"),":")  ' FlukePP + SCUSID + Ship To ID
      Fluke_PP = Account_Info(0)
      SCUSID   = Account_Info(1)
      Ship_Who = Account_Info(2)
      
      select case Ship_Who
        case 0    ' Order by and Ship to are same person
                  ' If DCG account has not been established, use Oxxxxx Order By profile for account creation, ship to profile is for order by person
        case else ' Order by and Ship to are a different person
                  ' If DCG account has not been established, use Oxxxxx Order By profile for account creation, ship to profile is not the order by person
      end select  
  
      ' Decode Item Numbers and Quantity Example
      Item_Number = Split(request.form("SITSNA"),",")
      Item_Number_Count = UBound(Item_Number)
      
      for x = 0 to Item_Number_Count step 2
        SIT = Item_Number(x)
        SNA = Item_Number(x+1)
      next  
  
      ' ...
  
      Post_OK = True  ' Determined whether post was successful 'true' or not 'false' and set flag
      ErrorMesg = ""  ' If Error, add text error message
      ErrorCode = 0   ' If Error, add numeric error code

      ' --------------------------------------------------------------------------------------
      ' DCG code end
      ' --------------------------------------------------------------------------------------      
        
      if Post_OK then
        .write "<DATASET>"   & "<BR>" & vbCrLf
        .write "<ACTION>"    & Action & "</ACTION>"    & "<BR>" & vbCrLf
        .write "<SCUSID>"    & "[SCUSID]"             & "</SCUSID>"    & "<BR>" & vbCrLf
        .write "<SORDNH>"    & Session.SessionID      & "</SORDNH>"    & "<BR>" & vbCrLf
        .write "</DATASET>"                                            & "<BR>" & vbCrLf
      end if   
            
    case "order_status"   ' Send Order Status
  
      ' --------------------------------------------------------------------------------------
      ' Data is received from the Fluke Externet Site as a HTTP POST
      ' Map response.form("") elements to DCG DB fields and set POST_OK flag below.
      ' Set Verbose = True above to enumerate during development and debug
      ' --------------------------------------------------------------------------------------
      ' DCG Code Goes in Here
      ' --------------------------------------------------------------------------------------

      Post_OK = True  ' Determined whether post was successful or not and set flag
      ErrorMesg = ""  ' If Error, add text error message
      ErrorCode = 0   ' If Error, add numeric error code

      if Post_OK then      
        .write "<DATASET>" & "<BR>" & vbCrLf
        .write "<ACTION>"  & Action & "</ACTION>" & "<BR>" & vbCrLf
        .write "<SORDNH>"  & request.form("SORDNH") & "</SORDNH>" & "<BR>" & vbCrLf
        .write "<SRDSTS>"  & "10"  & "</SRDSTS>" & "<BR>" & vbCrLf
        .write "<SPPKG>"   & "1Z9425850305046012"   & "</SPPKG>"  & "<BR>" & vbCrLf
        .write "<SDTSP>"   & Date() & "</SDTSP>" & "<BR>" & vbCrLf
        .write "<SPCAR>"   & "UPS"   & "</SPCAR>"  & "<BR>" & vbCrLf
        .write "</DATASET>"
      end if  
    
    case "item_status"    ' Send Item Number On-Hand Quantity
  
      ' --------------------------------------------------------------------------------------
      ' Data is received from the Fluke Externet Site as a HTTP POST
      ' Map response.form("") elements to DCG DB fields and set POST_OK flag below.
      ' Set Verbose = True above to enumerate during development and debug
      ' --------------------------------------------------------------------------------------
      ' DCG Code Goes in Here
      ' --------------------------------------------------------------------------------------

      Post_OK = True  ' Determined whether post was successful or not and set flag
      ErrorMesg = ""  ' If Error, add text error message
      ErrorCode = 0   ' If Error, add numeric error code
      
      if Post_OK then      
        .write "<DATASET>"                                         & "<BR>" & vbCrLf
        .write "<ACTION>"   & Action & "</ACTION>" & "<BR>" & vbCrLf
        .write "<SIT>"      & "[SIT]"                & "</SIT>"    & "<BR>" & vbCrLf
        .write "<SITQTY>"   & "[SITQTY]"             & "</SITQTY>" & "<BR>" & vbCrLf
        .write "</DATASET>"                                        & "<BR>" & vbCrLf
      end if  
    
    case else ' Invalid Action
      
      Post_OK = True
      .write "<DATASET>"                                            & "<BR>" & vbCrLf
      .write "<ACTION>" & Action & "</ACTION>"    & "<BR>" & vbCrLf
      .write "<ERRORMESG>" & "Invalid Action"      & "</ERRORMESG>" & "<BR>" & vbCrLf
      .write "<ERRORCODE>" & "1"                   & "</ERRORCODE>" & "<BR>" & vbCrLf
      .write "</DATASET>"                                           & "<BR>" & vbCrLf
        
  end select

  ' --------------------------------------------------------------------------------------
  ' Error Messages
  ' --------------------------------------------------------------------------------------        
  
  if not Post_OK then
    .write "<DATASET>"                                             & "<BR>" & vbCrLf
    .write "<ACTION>"    & Action & "</ACTION>"    & "<BR>" & vbCrLf
    .write "<ERRORMESG>" & ErrorMesg              & "</ERRORMESG>" & "<BR>" & vbCrLf
    .write "<ERRORCODE>" & ErrorCode              & "</ERRORCODE>" & "<BR>" & vbCrLf         
    .write "</DATASET>"                                            & "<BR>" & vbCrLf
  end if  
  
  ' --------------------------------------------------------------------------------------
  ' Debug
  ' --------------------------------------------------------------------------------------        
  
  if Verbose then ' if Verbose is true, enumerate response.form object for development and debug
    .write "<P>"
    .write "<TABLE BGCOLOR=BLACK COLPADDING=0 COLSPACING=0 BORDER=0>"
    .write "<TR>"
    .write "<TD>"
    
    .write "<TABLE  COLPADDING=0 COLSPACING=0 Border=0>"
  
    .write "<TR>"
    .write "<TD BGCOLOR=Silver><FONT FACE=Arial SIZE=2><B>Key</B></FONT></TD>"
    .write "<TD BGCOLOR=Silver><FONT FACE=Arial SIZE=2><B>Value</B></FONT></TD>"
    .write "<TD BGCOLOR=Silver ALIGN=CENTER><FONT FACE=Arial SIZE=2><B>Array</B></FONT></TD>"  
    .write "</TR>"
    
    ' for each item in request.querystring
  
    for each item in request.form
      .write "<TR>"
      .write "<TD BGCOLOR=WHITE><FONT FACE=Arial SIZE=2>" & item & "</FONT></TD>"
      .write "<TD BGCOLOR=WHITE><FONT FACE=Arial SIZE=2>"
      if request(item) = "" then
        .write "&nbsp;"
      else
        .write request.form(item)
      end if
      .write "</FONT></TD>"
      .write "<TD BGCOLOR=WHITE><FONT FACE=Arial SIZE=2>"
      
      Check4Array = Split(request(item),",")
      if UBound(Check4Array) > 0 then
        .write "Yes"
        else
        .write "&nbsp;"
      end if
      .write "</FONT></TD>"
      .write "</TR>"
    next
    .write "</TABLE>"
  
    .write "</TD>"
    .write "</TR>"
    .write "</TABLE>"
              
  end if  

  ' --------------------------------------------------------------------------------------
  ' End of Main
  ' --------------------------------------------------------------------------------------        
  
  .write vbCrLf
  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf
  
end with
%>
