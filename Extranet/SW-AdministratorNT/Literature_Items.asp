<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     K. D. Whitlock
' Date:       06/1/2000
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------

Dim Site_ID
Dim Record_Count, Record_Pages, Record_Limit, PCID, Page_QS
Dim Lan_Type, Lit_Type, Sort_By, Show_Detail, Search

if not isblank(request.querystring("Site_ID")) then
  Site_ID        = request.querystring("Site_ID")
else
  Site_ID = 3
end if

if not isblank(request.querystring("PCID")) then
  PCID = CInt(request.querystring("PCID"))
else
  PCID = 0
end if
    
'Login_Language = "eng"

Call Connect_SiteWide

SQL = "SELECT Site.* FROM Site WHERE Site.ID=" & Site_ID
Set rsSite = Server.CreateObject("ADODB.Recordset")
rsSite.Open SQL, conn, 3, 3

Site_Code        = rsSite("Site_Code")     
Screen_Title     = rsSite("Site_Description") & " - " & Screen_TitleX
Bar_Title        = rsSite("Site_Description") & "<BR><FONT CLASS=SmallBoldGold>" & "Literature Items" & "</FONT>"
Navigation       = false
Top_Navigation   = false
Content_Width    = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->
<%

rsSite.close
set rsSite=nothing

' --------------------------------------------------------------------------------------

Dim Field_Max
Field_Max = 53

Dim Field_Name(53)
Dim Field_Data(53)
Dim Field_Pointer

' Field Name Pointers

xITEM               = 0
xREVISION           = 1
xLITERATURE_TYPE    = 2  
xLANGUAGE           = 3

xCOST_CENTER        = 4  
xPRODUCT_FAMILY     = 5
xMODEL_GROUP        = 6
xMARKET_MODEL       = 7
xDESCRIPTION        = 8
xMARCOM_ITEM_DESC   = 9
xEFULFILLMENT       = 10

xORIGIN_CODE        = 11
xREPLACEMENT        = 12

xEND_USER           = 13
xPDF                = 14
xPOD                = 15
xPRINT              = 16
xWEB                = 17
xFAX                = 18
xCD                 = 19
xDISPLAY            = 20
xDISKS              = 21
xURL                = 22
xVIDEO_NTSC         = 23
xVIDEO_PAL          = 24
xSUPPORT            = 25

xCUSTOMER_ORDER     = 26
xINTERNAL_ORDER     = 27
xPRICED             = 28

xUOM                = 29
xLARGE_LIMIT        = 30
xSMALL_LIMIT        = 31
xMINIMUM_ORD_QTY    = 32
xMAXIMUM_ORD_QTY    = 33
xMIN_MINMAX_QTY     = 34
xMAX_MINMAX_QTY     = 35
xONHAND_QTY         = 36
xTOTAL_QUANTITY     = 37
xREORDER_QTY        = 38

xSTATUS             = 39
xACTION             = 40
xRELEASE_DATE       = 41
xQA_UPDATE_DATE     = 42
xREVIEW_DATE        = 43
xITEM_UPDATE_DATE   = 44
xEXPIRE_DATE        = 45
        
xOBSOLETE_RULE      = 46
xINVENTORY_RULE     = 47  

xPRINTER            = 48
xPRINTER_PHONE      = 49
xDELIVERY_DATE      = 50
xVENDOR_COMMENTS    = 51

xPRODUCTION_MANAGER = 52
xMARCOM_MANAGER     = 53

Call Get_Field_Names

' --------------------------------------------------------------------------------------
' Decode and Setup Filters
' --------------------------------------------------------------------------------------

' Literature Type
if not isblank(request.querystring("Lit_Type")) then
  Lit_Type = request.querystring("Lit_Type")
else
  Lit_Type = ""
end if    

' Language Type
if not isblank(request.querystring("Lan_Type")) then
  Lan_Type = request.querystring("Lan_Type")
else
  Lan_Type = ""
end if    

' Search
if not isblank(request.querystring("Search")) then
  Search = request.querystring("Search")
else
  Search = ""
end if    

' Literature Type

SQL = "SELECT DISTINCT Literature_Type FROM Literature_Items_US ORDER BY Literature_Type"
Set rsTypes = Server.CreateObject("ADODB.Recordset")
rsTypes.Open SQL, conn, 3, 3

if not rsTypes.EOF then
  response.write "<SPAN CLASS=SmallBold>" & Translate("Literature Type",Login_Language,conn) & ":</SPAN> "
  response.write "<SELECT CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Literature_Items.asp?Site_ID=" & Site_ID & "&Sort_By=" & Sort_By & "&Lan_Type=" & Lan_Type & "&PCID=0" & "&Search=" & Search & "&Lit_Type='" & "+this.options[this.selectedIndex].value"" NAME=""Lit_Type"">" & vbCrLf
  response.write "<OPTION CLASS=Small Value="""">" & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf
  do while not rsTypes.EOF
    response.write "<OPTION Class=Small VALUE=""" & rsTypes("Literature_Type") & """ "
    if LCase(Lit_Type) = LCase(rsTypes("Literature_Type")) then response.write " SELECTED"
    response.write ">" & Translate(ProperCase(rsTypes("Literature_Type")),Login_Language,conn) & "</OPTION>" & vbCrLf
    rsTypes.MoveNext
  loop
  response.write "</SELECT>" & vbCrLf
end if    

rsTypes.close
set rsTypes = nothing

' Language

SQL = "SELECT DISTINCT Language FROM Literature_Items_US ORDER BY Language"
Set rsLanguage = Server.CreateObject("ADODB.Recordset")
rsLanguage.Open SQL, conn, 3, 3

if not rsLanguage.EOF then
  response.write "&nbsp;&nbsp;"
  response.write "<SPAN CLASS=SmallBold>" & Translate("Language",Login_Language,conn) & ":</SPAN> " & vbCrLf
  response.write "<SELECT CLASS=Small LANGUAGE=""JavaScript"" ONCHANGE=""window.location.href='Literature_Items.asp?Site_ID=" & Site_ID & "&Sort_By=" & Sort_By & "&Lit_Type=" & Lit_Type & "&PCID=0" & "&Search=" & Search & "&Lan_Type='" & "+this.options[this.selectedIndex].value"" NAME=""Lan_Type"">" & vbCrLf
  response.write "<OPTION CLASS=Small Value="""">" & Translate("All",Login_Language,conn) & "</OPTION>" & vbCrLf
  do while not rsLanguage.EOF
    response.write "<OPTION Class=Small VALUE=""" & rsLanguage("Language") & """ "
    if LCase(Lan_Type) = LCase(rsLanguage("Language")) then response.write " SELECTED"
    response.write ">" & Translate(ProperCase(rsLanguage("Language")),Login_Language,conn) & "</OPTION>" & vbCrLf
    rsLanguage.MoveNext
  loop
  response.write "</SELECT>" & vbCrLf
end if

rsLanguage.close
set rsLanguage = nothing

' Search Criteria
response.write "&nbsp;&nbsp;"
response.write "<SPAN CLASS=SmallBold>" & Translate("Keyword",Login_Language,conn) & ":</SPAN> " & vbCrLf
response.write "<INPUT CLASS=Small TYPE=""TEXT"" VALUE=""" & Search & """ LANGUAGE=""JavaScript"" ONBLUR=""window.location.href='Literature_Items.asp?Site_ID=" & Site_ID & "&Sort_By=" & Sort_By & "&Lit_Type=" & Lit_Type & "&PCID=0" & "&Lan_Type=" & Lan_Type & "&Search='" & "+this.value"" NAME=""Search"">"
'if isblank(Search) then
  response.write "&nbsp;<INPUT TYPE=""BUTTON"" CLASS=NavLeftHighlight1 VALUE="" " & Translate("Search",Login_Language,conn) & " "">" & vbCrLf
'end if  

if not isblank(request.querystring("Sort_By")) and isnumeric(request.querystring("Sort_By"))then
  Sort_By = CInt(request.querystring("Sort_By"))
else
  Sort_By = xLiterature_Type
end if

SQL = "SELECT * FROM Literature_Items_US Site WHERE Status='Active' "

if not isblank(Lit_Type) then
  SQL = SQL & " AND Literature_Type='" & Lit_Type & "' "
end if  

if not isblank(Lan_Type) then
  SQL = SQL & " AND (Language='" & Lan_Type & "' OR Language='Multiple') "
end if

if not isblank(Search) then
  SQL = SQL & " AND ("
  SQL = SQL & Field_Name(xProduct_Family) & " LIKE '%" & Search & "%' OR "
  SQL = SQL & Field_Name(xModel_Group)    & " LIKE '%" & Search & "%' OR "
  SQL = SQL & Field_Name(xMarket_Model)   & " LIKE '%" & Search & "%' OR "  
  SQL = SQL & Field_Name(xeFulfillment)   & " LIKE '%" & Search & "%') "  
end if

'response.write "<BR>" & SQL & "<BR>"

SQL = SQL & "ORDER BY " & Field_Name(Sort_By) & ", Item"
Set rsItems = Server.CreateObject("ADODB.Recordset")
rsItems.Open SQL, conn, 3, 3

TableOn = false

if not rsItems.EOF then

  Record_Limit  = 20
  Record_Count  = rsItems.RecordCount
  Record_Pages  = Record_Count \ Record_Limit
  if Record_Count mod Record_Limit > 0 then Record_Pages = Record_Pages + 1

  response.write "<BR><SPAN CLASS=SmallBold>" & Translate("Items Found",Login_Language,conn) & ": " & rsItems.RecordCount & "</SPAN><BR>"

  response.write "<TABLE WIDTH=""100%"" ALIGN=CENTER BORDER=0>"
  response.write "<TR>"
  response.write "<TD WIDTH=""100%"">"
  Call RS_Page_Navigation
  response.write "</TD>"
  response.write "</TR>"
  response.write "</TABLE>"

  TableOn = True

  %>
  <TABLE WIDTH="100%" BORDER="1" CELLPADDING=0 CELLSPACING=0 BORDERCOLOR="#666666" BGCOLOR="#666666">
    <TR>
      <TD>
        <TABLE CELLPADDING=4 CELLSPACING=1 BORDER=0  WIDTH="100%">
          <TR>
            <TD BGCOLOR="Red" ALIGN="CENTER" CLASS=SmallBoldWhite>Item</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Rev</TD>
            <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Type</TD>              
            <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Language</TD>
            <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Family</TD>
            <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Group</TD>
            <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Model</TD>                       
            <TD BGCOLOR="#000000" ALIGN="LEFT" CLASS=SmallBoldGold>Description</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>U<BR>S<BR>R</TD>              
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>P<BR>D<BR>F</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>P<BR>O<BR>D</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>P<BR>R<BR>T</TD>            
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>W<BR>E<BR>B</TD>            
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>F<BR>A<BR>X</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>C<BR>D<BR>&nbsp;</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>D<BR>S<BR>P</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>D<BR>S<BR>K</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>U<BR>R<BR>L</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>V<BR>N<BR>T</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>V<BR>P<BR>L</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>S<BR>U<BR>P</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>C<BR>O<BR>R</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>I<BR>O<BR>R</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>$<BR>$<BR>$</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>Unit</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>SLMT</TD>
            <TD BGCOLOR="#000000" ALIGN="CENTER" CLASS=SmallBoldGold>LLMT</TD>
            <TD BGCOLOR="#000000" ALIGN="RIGHT" CLASS=SmallBoldGold>Release</TD>
      		</TR>
  <%

	rsItems.MoveFirst
  if Record_Limit * (PCID - 1) > 0 then
  	rsItems.Move (Record_Limit * (PCID - 1))
  end if  
  
  Record_Number = 1
  Toggle = false

  do while not rsItems.EOF and Record_Number <= Record_Limit

    Call Get_Field_Data ()

    response.write "<TR>"

    for Field_Pointer = 0 to Field_Max
      select case Field_Pointer
        case xItem, xRevision, xLiterature_Type, xLanguage, _
             xPRODUCT_FAMILY, xMODEL_GROUP, xMARKET_MODEL, xEFulfillment, _
             xEnd_User, xPDF, xPOD, xPrint, xWeb, xFax, xCD, xDisplay, xDisks, xURL, xVideo_NTSC, xVideo_PAL, xSupport, _
             xCustomer_Order, xInternal_Order, xPriced, _
             xUOM, _
             xSmall_Limit, xLarge_Limit, _
             xRelease_Date

          Call Format_Data()
        
      end select
    next

    response.write "</TR>"

    if Gray_Toggle then Gray_Toggle = false else GrayToggle = True
      
    rsItems.MoveNext

    Record_Number = Record_Number + 1
  
  loop

  if TableOn then
    response.write "</TABLE>" & vbCrLf
    response.write "</TD>" & vbCrLf
    response.write "</TR>" & vbCrLf
    response.write "</TABLE>" & vbCrLf & vbCrLf
  end if

else
  response.write "<BR><SPAN CLASS=SmallBold>" & Translate("Items Found",Login_Language,conn) & ": " & rsItems.RecordCount & "</SPAN><BR>"  
end if  

rsItems.close
set rsItems = nothing
  

' --------------------------------------------------------------------------------------

Call Disconnect_SiteWide

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"--> 
<%

' --------------------------------------------------------------------------------------
' Functions and Subroutines
' --------------------------------------------------------------------------------------

sub Format_Data()

  select case Field_Pointer

    ' Boolean Fields

    case xEnd_User, xPDF, xPOD, xPrint, xWeb, xFax, xCD, xDisplay, _
         xDisks, xURL, xVideo_NTSC, xVideo_PAL, xSupport, _
         xCustomer_Order, xInternal_Order, xPriced

      response.write "<TD CLASS=Small ALIGN=CENTER BGCOLOR=""LightGrey"">"
      if not isblank(Field_Data(Field_Pointer)) then
        if isnumeric(Field_Data(Field_Pointer)) then
          if CInt(Field_Data(Field_Pointer)) = CInt(True) then
            response.write "Y"
          else
            response.write "&nbsp;"  
          end if  
        else
          response.write "&nbsp;"
        end if
      else    
        response.write "&nbsp;"
      end if
      response.write "</TD>"  
    
    ' Numeric Fields

    case xLarge_Limit, xSmall_Limit, xMinimum_Ord_Qty, xMaximum_Ord_Qty, _
         xMin_MinMax_Qty, xMax_MinMax_Qty, xOnHand_Qty, xTotal_Quantity, xReOrder_Qty

      response.write "<TD CLASS=Small ALIGN=RIGHT BGCOLOR=""White"">"
      if not isblank(Field_Data(Field_Pointer)) then
        if isnumeric(Field_Data(Field_Pointer)) then
          response.write FormatNumber(Field_Data(Field_Pointer),0)
        else
          response.write "&nbsp;"
        end if  
      else    
        response.write "&nbsp;"
      end if
      response.write "</TD>"  
    
    ' Date Fields

    case xRelease_Date, xQA_Update_Date, xReview_Date, xItem_Update_Date, xExpire_Date
         
      response.write "<TD CLASS=Small ALIGN=RIGHT BGCOLOR=""White"">"
      if not isblank(Field_Data(Field_Pointer)) then
        if isdate(Field_Data(Field_Pointer)) then
          response.write FormatDateTime(Field_Data(Field_Pointer))
        else
          response.write "&nbsp;"
        end if    
      else    
        response.write "&nbsp;"
      end if
      response.write "</TD>"  
    
    ' Text Fields
    
    case else
    
      response.write "<TD CLASS=Small ALIGN=LEFT BGCOLOR=""White"">"
      if not isblank(Field_Data(Field_Pointer)) then
        Keyword = Search
        response.write Highlight_Keyword(Field_Data(Field_Pointer),Keyword, "#FF0000")
      else    
        response.write "&nbsp;"
      end if
      response.write "</TD>"  

  end select
  
end sub

' --------------------------------------------------------------------------------------

sub Get_Field_Names()

  ' Field Locators
  
  Field_Name(xITEM)               = "ITEM"
  Field_Name(xREVISION)           = "REVISION"
  Field_Name(xLITERATURE_TYPE)    = "LITERATURE_TYPE"
  Field_Name(xLANGUAGE)           = "LANGUAGE"

  Field_Name(xCOST_CENTER)        = "COST_CENTER"
  Field_Name(xPRODUCT_FAMILY)     = "PRODUCT_FAMILY"
  Field_Name(xMODEL_GROUP)        = "MODEL_GROUP"
  Field_Name(xMARKET_MODEL)       = "MARKET_MODEL"
  Field_Name(xDESCRIPTION)        = "DESCRIPTION"
  Field_Name(xMARCOM_ITEM_DESC)    = "MARCOM_ITEM_DESC"
  Field_Name(xEFULFILLMENT)       = "EFULFILLMENT"
  
  Field_Name(xORIGIN_CODE)        = "ORIGIN_CODE"
  Field_Name(xREPLACEMENT)         = "REPLACEMENT"

  Field_Name(xEND_USER)           = "END_USER"
  Field_Name(xPDF)                = "PDF"
  Field_Name(xPOD)                = "POD"
  Field_Name(xPRINT)              = "PRINT"
  Field_Name(xWEB)                = "WEB"
  Field_Name(xFAX)                = "FAX"
  Field_Name(xCD)                 = "CD"
  Field_Name(xDISPLAY)            = "DISPLAY"
  Field_Name(xDISKS)              = "DISKS"
  Field_Name(xURL)                = "URL"
  Field_Name(xVIDEO_NTSC)         = "VIDEO_NTSC"
  Field_Name(xVIDEO_PAL)          = "VIDEO_PAL"
  Field_Name(xSUPPORT)            = "SUPPORT"

  Field_Name(xCUSTOMER_ORDER)     = "CUSTOMER_ORDER"
  Field_Name(xINTERNAL_ORDER)     = "INTERNAL_ORDER"
  Field_Name(xPRICED)             = "PRICED"

  Field_Name(xUOM)                = "UOM"
  Field_Name(xLARGE_LIMIT)        = "LARGE_LIMIT"
  Field_Name(xSMALL_LIMIT)        = "SMALL_LIMIT"
  Field_Name(xMINIMUM_ORD_QTY)    = "MINIMUM_ORD_QTY"
  Field_Name(xMAXIMUM_ORD_QTY)    = "MAXIMUM_ORD_QTY"
  Field_Name(xMIN_MINMAX_QTY)     = "MIN_MINMAX_QTY"
  Field_Name(xMAX_MINMAX_QTY)     = "MAX_MINMAX_QTY"
  Field_Name(xONHAND_QTY)         = "ONHAND_QTY"
  Field_Name(xTOTAL_QUANTITY)     = "TOTAL_QUANTITY"
  Field_Name(xREORDER_QTY)        = "REORDER_QTY"

  Field_Name(xSTATUS)             = "STATUS"
  Field_Name(xACTION)             = "ACTION"
  Field_Name(xRELEASE_DATE)       = "RELEASE_DATE"
  Field_Name(xQA_UPDATE_DATE)     = "QA_UPDATE_DATE"
  Field_Name(xREVIEW_DATE)        = "REVIEW_DATE"
  Field_Name(xITEM_UPDATE_DATE)   = "ITEM_UPDATE_DATE"
  Field_Name(xEXPIRE_DATE)        = "EXPIRE_DATE"

  Field_Name(xOBSOLETE_RULE)      = "OBSOLETE_RULE"
  Field_Name(xINVENTORY_RULE)     = "INVENTORY_RULE"

  Field_Name(xPRINTER)            = "PRINTER"
  Field_Name(xPRINTER_PHONE)      = "PRINTER_PHONE"
  Field_Name(xDELIVERY_DATE)      = "DELIVERY_DATE"
  Field_Name(xVENDOR_COMMENTS)    = "VENDOR_COMMENTS"

  Field_Name(xPRODUCTION_MANAGER) = "PRODUCTION_MANAGER"
  Field_Name(xMARCOM_MANAGER)     = "MARCOM_MANAGER"

  
end sub

' --------------------------------------------------------------------------------------

sub Get_Field_Data ()
  
  for Field_Pointer = 0 to Field_Max
    if isblank(rsItems(Field_Name(Field_Pointer))) then
      Field_Data(Field_Pointer) = ""
    else
      Field_Data(Field_Pointer) = CStr(rsItems(Field_Name(Field_Pointer)))
    end if  
  next
    
end sub

' --------------------------------------------------------------------------------------
' Record Set Page Navigation
' --------------------------------------------------------------------------------------

Sub RS_Page_Navigation

  Page_QS = "Site_ID=" & Site_ID & "&Language=" & Login_Language & "&Lan_Type=" & Lan_Type & "&Lit_Type=" & Lit_Type & "&Sort_By=" & Sort_By & "&Show_Detail=" & Show_Detail & "&Search=" & Search
	if PCID = 0 then PCID = 1

  if Record_Pages > 1 then

    response.write "<FONT CLASS=SmallBold>&nbsp;" & Translate("Page", Login_Language, conn) & ": &nbsp;&nbsp;</FONT>"

  	if PCID = 1 then
  		Call RS_Page_Numbers
    		response.write "<A HREF=""Literature_Items.asp?" & Page_QS & "&PCID=" & PCID + 1 & """ TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"
        response.write "&nbsp;&nbsp;"
  	else
  		if PCID = Record_Pages then
  			response.write "<A HREF=""Literature_Items.asp?" & Page_QS & "&PCID=" & PCID - 1 & """ TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;"
    		Call RS_Page_Numbers
  		else
  			response.write "<A HREF=""Literature_Items.asp?" & Page_QS & "&PCID=" & PCID - 1 &  """ TITLE=""" & Translate("Previous Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&lt;&lt;&nbsp;&nbsp;</FONT></A>&nbsp;&nbsp;"
    		Call RS_Page_Numbers
  			response.write "<A HREF=""Literature_Items.asp?" & Page_QS & "&PCID=" & PCID + 1 &  """ TITLE=""" & Translate("Next Page", Alt_Language, conn) & """>"
        response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;&nbsp;&gt;&gt;&nbsp;&nbsp;</FONT></A>"
  		end if
  		
  	end if

  end if

End Sub

' --------------------------------------------------------------------------------------
' Record Set Page Numbers
' --------------------------------------------------------------------------------------

Sub RS_Page_Numbers

  for i = 1 to Record_Pages
  	if i = PCID then
	  	response.write "<A HREF=""Literature_Items.asp?" & Page_QS & "&PCID=" & i & """>"
      response.write "<FONT CLASS=NAVLEFTHIGHLIGHT1>&nbsp;"
      if i < 10 then response.write "&nbsp;"
      response.write CStr(i) & "&nbsp;</FONT></A> &nbsp;"
  	else
			response.write "<A HREF=""Literature_Items.asp?" & Page_QS & "&PCID=" & i & """>"
      response.write  "<FONT CLASS=NavTopHighLight>&nbsp;"
      if i < 10 then response.write "&nbsp;"
      response.write CStr(i) & "&nbsp;</FONT></A> &nbsp;"
  	end if
  next

end sub

' --------------------------------------------------------------------------------------



%>
