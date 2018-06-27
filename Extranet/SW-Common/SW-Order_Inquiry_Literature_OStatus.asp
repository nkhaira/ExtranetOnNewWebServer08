<%
'--------------------------------------------------------------------------------------

function DCG_Status(Action_Sequence, Reference_Number, conn)

  ' Order Status: Reference Number = Order_Number
  ' Item  Status: Reference Number = Item_Number

  Dim strRemoteHostName, strRemoteHostTargetFile, strRemoteHostURL, strLocalHostReferrerFile
  Dim strResponse, bResponse

  Script_Debug   = False
  Transfer_Debug = False

  %>
  <!--#include virtual="/connections/Connection_Literature_Order_System.asp"-->
  <%

  strLocalHostReferrerFile = request.ServerVariables("SCRIPT_Name")
  strResponse              = ""
  
  select case LCase(Action_Sequence)
    case "order_status"
      strPost_QueryString      = "Action=" & Action_Sequence & "&SORDNH=" & Replace(Reference_Number,"FLUKECO","")
    case "item_status"
      strPost_QueryString      = "Action=" & Action_Sequence & "&SIT=" & Reference_Number
      ' Sample:  http://www.hkmdm.com/fluke/response.asp?Action=Item_Status&SIT=1547191
  end select  
  
  Dim HTTPRequest
  
  Set HTTPRequest = Server.CreateObject("Msxml2.XMLHTTP.3.0") 
  
  if not Script_Debug then
    HTTPRequest.Open "POST", strRemoteHostURL, False 
    HTTPRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    on error resume next          
    HTTPRequest.Send strPost_QueryString
    
    if err.Number <> 0 then
      strResponse = err.Description & "<BR>Order not received by remote server."
      bResponse   = err.Number            
    else  
      strResponse = HTTPRequest.responseText
      bResponse = HTTPRequest.status
    end if
    
    on error goto 0  
    
  else
    response.write "URL: " & strRemoteHostURL & "<P>"
    response.write "Form Data: " & Replace(strPost_QueryString,"&","<BR>&") & "<P>"
    response.flush
    response.end        
  end if  
  
  if bResponse = 200 then
  
    if Transfer_Debug then
      PostErrMessage = PostErrMessage & "----------------<P>"
      PostErrMessage = PostErrMessage & "<B>DCG Post Successful</B><BR>"
      PostErrMessage = PostErrMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR>"
      PostErrMessage = PostErrMessage & "<FONT COLOR=BLUE>" & strResponse & "</FONT><P>"
    end if
    
    ' Decode Response and Extract Data
  
    if instr(1,LCase(strResponse),"<dataset>") > 0 then
      xmlText = Mid(strResponse,instr(1,LCase(strResponse),"<dataset>")+9)
    else  
      response.write "<SPAN CLASS=SMALLBOLDRED>Error: No &lt;dataset&gt; Opening Tag Found - " & Reference_Number & "</SPAN><BR>"
    end if  
    if instr(1,LCase(strResponse),"</dataset>") > 0 then
      xmlText = Mid(xmlText,1,instr(1,LCase(xmlText),"</dataset>") - 1)
    else  
      response.write "<SPAN CLASS=SMALLBOLDRED>Error: No &lt;dataset&gt; Ending Tag Found - " & Reference_Number & "</SPAN><BR>"
    end if  
    
    xmlTags = "ACTION,ERRORMESG,ERRORCODE,SCUSID,SORDNH,SRDSTS,SPPKG,SPCAR,SDTSP,SITQTY,SIT"
    xmlTag  = Split(xmlTags,",")
    xmlData = Split(xmlTags,",")
    xmlTagCount = UBound(xmlTag)
  
    for TagCount = 0 to xmlTagCount
      if instr(1,LCase(xmlText),"<" & LCase(xmlTag(TagCount)) & ">") > 0 then
        xmlTemp = xmlText
        xmlTemp = Mid(xmlTemp,instr(1,LCase(xmlTemp),"<" & LCase(xmlTag(TagCount)) & ">") + Len(xmlTag(TagCount)) + 2 )
        xmlTemp = Mid(xmlTemp,1,instr(1,LCase(xmlTemp),"</" & LCase(xmlTag(TagCount)) & ">") - 1)
        if len(xmlTemp) > 0 then
          xmlData(TagCount) = Trim(xmlTemp)
        else
          xmlData(TagCount) = ""
        end if    
      else
        xmlData(TagCount) = ""
      end if
    next
    
    select case LCase(Action_Sequence)
    
      case "order_status"
      
        Order_Number     = GetXMLData(xmlTag, xmlData, "SORDNH")
        Order_Status     = GetXMLData(xmlTag, xmlData, "SRDSTS")
        Ship_Carrier     = GetXMLData(xmlTag, xmlData, "SPCAR")
        Ship_Tracking_No = GetXMLData(xmlTag, xmlData, "SPPKG")
        Ship_Date        = GetXMLData(xmlTag, xmlData, "SDTSP")
        Error_Text       = GetXMLData(xmlTag, xmlData, "ERRORMESG")
        Error_Number     = GetXMLData(xmlTag, xmlData, "ERRORCODE")        
        
        if Transfer_Debug = True then
          PostErrMessage = PostErrMessage & "Action: " & Action_Sequence & "<BR>"
          PostErrMessage = PostErrMessage & "Order Number: " & Order_Number & "<BR>"
          PostErrMessage = PostErrMessage & "Order Status: " & Order_Status & "<BR>"
          PostErrMessage = PostErrMessage & "Ship Carrier: " & Ship_Carrier & "<BR>"
          PostErrMessage = PostErrMessage & "Ship Tracking No: " & Ship_Tracking_No & "<BR>"
          PostErrMessage = PostErrMessage & "Ship Date: " & Ship_Date & "<BR>"
          PostErrMessage = PostErrMessage & "Error Text: " & Error_Text & "<BR>"
          PostErrMessage = PostErrMessage & "Error Number: " & Error_Number & "<P>"
        end if  

        if not isblank(Ship_Carrier) then

          SQLStatus = "SELECT * FROM Shopping_Cart_Ship_Tracking WHERE Name='" & Ship_Carrier & "'"
          Set rsStatus = Server.CreateObject("ADODB.Recordset")
          rsStatus.Open SQLStatus, conn, 3, 3
        
          if rsStatus.EOF then
            rsStatus.close
            set rsStatus = nothing
            Ship_Carrier = CInt(Get_New_Record_ID("Shopping_Cart_Ship_Tracking", "Name", Ship_Carrier, conn))
          else
            Ship_Carrier = rsStatus("ID")
            rsStatus.close
            set rsStatus = nothing
          end if
        end if  
            
        SQLStatus =             "UPDATE Shopping_Cart_Lit SET "
        SQLStatus = SQLStatus & "Order_Status="
        if isblank(Order_status) then Order_Status = 0
        SQLStatus = SQLStatus & Order_Status
        SQLStatus = SQLStatus & ", Ship_Carrier="
        if isblank(Ship_Carrier) or Ship_Carrier = "0" then Ship_Carrier = "0"
        SQLStatus = SQLStatus & Ship_Carrier
        SQLStatus = SQLStatus & ", Ship_Tracking_No="
        if isblank(Ship_Tracking_No) then Ship_Tracking_No = "NULL" else Ship_Tracking_No = "'" & Ship_Tracking_No & "'"
        SQLStatus = SQLStatus & Ship_Tracking_No
        SQLStatus = SQLStatus & ", Order_Ship_Date="
        if isblank(Ship_Date) then Ship_Date = "NULL" else Ship_Date = "'" & Ship_Date & "'"       
        SQLStatus = SQLStatus & Ship_Date & " "
        SQLStatus = SQLStatus & "WHERE ORDER_Number='"
        if instr(Order_Number,"FLUKECO0") = 0 then Order_Number = "FLUKECO0" & Right(Order_Number,7) 
        SQLStatus = SQLStatus & Order_Number & "'"
        conn.execute SQLStatus
        
        DCG_Status = Order_Status
  
      case "item_status"

        Onhand_Qty = GetXMLData(xmlTag, xmlData, "SITQTY")
        DCG_Status = Onhand_Qty
        
      case else
      
        SQLStatus = "UPDATE Shopping_Cart_Lit SET "
        SQLStatus = SQLStatus & "Order_Status="
        if isblank(Order_status) then Order_Status = 100
        SQLStatus = SQLStatus & Order_Status
        SQLStatus = SQLStatus & ", Ship_Carrier="
        if isblank(Ship_Carrier) or Ship_Carrier = "0" then Ship_Carrier = "0"
        SQLStatus = SQLStatus & Ship_Carrier
        SQLStatus = SQLStatus & ", Ship_Tracking_No="
        if isblank(Ship_Tracking_No) then Ship_Tracking_No = "NULL" else Ship_Tracking_No = "'" & Ship_Tracking_No & "'"
        SQLStatus = SQLStatus & Ship_Tracking_No
        SQLStatus = SQLStatus & ", Order_Ship_Date="
        if isblank(Ship_Date) then Ship_Date = "NULL" else Ship_Date = "'" & Ship_Date & "'"       
        SQLStatus = SQLStatus & Ship_Date & " "
        SQLStatus = SQLStatus & "WHERE ORDER_Number='"
        if instr(Order_Number,"FLUKECO0") = 0 then Order_Number = "FLUKECO0" & Right(Order_Number,7) 
        SQLStatus = SQLStatus & Order_Number & "'"
        conn.execute SQLStatus
        
        Error_Status = 100
        DCG_Status = Error_Status

    end select

  else    

    if Transfer_Debug then
      PostErrMessage = PostErrMessage & "----------------<P>"
      PostErrMessage = PostErrMessage & "<FONT COLOR=RED><B>DCG Post Failure:</B></FONT><BR>"
      PostErrMessage = PostErrMessage & "Calling Script: " & request.ServerVariables("Script_Name") & "<BR>"
      PostErrMessage = PostErrMessage & "<FONT COLOR=BLUE>" & strResponse & "</FONT><P>"
    end if

  end if
  
  if Transfer_Debug then
    response.write PostErrMessage
  end if
  
  
end function

'--------------------------------------------------------------------------------------

function GetXMLData(xmlTag, xmlData, Pointer)

  temp = ""
  for v = 0 to UBound(xmlTag)
    if LCase(xmlTag(v)) = LCase(Pointer) then
      temp = xmlData(v)
      exit for
    end if
  next

  GetXMLData = temp
  
end function
  
'--------------------------------------------------------------------------------------
%>                               
