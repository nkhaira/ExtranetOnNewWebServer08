<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Common Replacement Parts - Search</TITLE>
<LINK REL="stylesheet" TYPE="text/css" HREF="/include/FlukeStyle.css">
</HEAD>
<BODY>

<!-- Remove any code above this line -->

<!--#include virtual="/connections/connection_eStore.asp" -->

<%
Post_URL     = "/usen/support/crp/default.asp"
Post_URL     = "/sw-administrator/common_replacement_parts_search.asp"
Post_Method  = "POST"

Dim Model, Family, Family_Code, Family_Max

Model   = request("CRP_Model")

if request("Family") <> "" then
  Family  = UCase(request("Family"))
else
  Family  = ""                                         ' Comma separated Family codes or single Family
end if

if instr(1,Family,",") > 0 then                         ' List one or more family / models in DB
  Family_Code = Split(Replace(Family," ",""),",")
  Family_Max  = Ubound(Family_Code)
elseif Family <> "" then                               ' List one family / models in DB
  ReDim Family_Code(0)
  Family_Code(0) = ""
else
  Family_Max = -1                                       ' List all family / models in DB
end if

Country = LCase("us")

Call Connect_eStoreDatabase

with response
  .write "<TABLE BORDER=0 WIDTH = ""590""CELLPADDING=4 CELLSPACING=0>" & vbCrLf
  .write "  <TR>" & vbCrLf
  .write "    <TD VALIGN=TOP ALIGN=LEFT COLSPAN=4 CLASS=PlainText>" & vbCrLf
  .write "      <SPAN CLASS=BlueHeadline>Common Replacement Parts - Search</SPAN>" & vbCrLf
  .write "      <P></P>" & vbCrLf
  .write "      <P>To search for a common replacement part for your product, select your model number from the drop down selection box.</P>" & vbCrLf
  .write "      <P>If you can't find your product listed, you may for replacement parts by calling the Fluke Parts Department at:" & vbCrLf
  .write "      <UL>" & vbCrLf
  .write "        <LI>North America and Asia: 1.800.526.4731 or 1.425.347.6100</LI>" & vbCrLf
  .write "        <LI>Europe: (31 40) 2.678.200</LI>" & vbCrLf
  .write "      </UL></P>" & vbCrLf
  .write "    </TD>" & vbCrLf
  .write "  </TR>" & vbCrLf
  .write "</TBODY>" & vbCrLf
  .write "</TABLE>" & vbCrLf & vbCrLf
  
  FormName = "SelectModel"
  
  .write "<FORM NAME=""" & FormName & """ ACTION=""" & Post_URL & """ METHOD=""" & Post_Method & """>" & vbCrLf
  .write "<INPUT TYPE=""HIDDEN"" NAME=""Family"" VALUE=""" & Family & """>" & vbCrLf
  .write "<TABLE WIDTH=""590"" CELLPADDING=4 CELLSPACING=0 BORDER=0 BGCOLOR=""#FFCC00"">" & vbCrLf
  .write "  <TR>" & vbCrLf
  .write "    <TD VALIGN=""top"" ALIGN=""left"" CLASS=PlainText>"
  
  SQL = "SELECT DISTINCT Model " &_
        "FROM vcturbo_replaceable_parts_xref "
        
        if Family_Max >= 0 then
          SQL = SQL & "WHERE "
          for fc = 0 to Family_Max
            if fc > 0 and fc < Family_Max then
              SQL = SQL & " OR "
              SQL = SQL & "Family=N'" & Family_Code(fc) & "' "
            else
              SQL = SQL & "Family=N'" & Family_Code(fc) & "' "
            end if
          next
        end if

        SQL = SQL & "ORDER BY Model"

  Set rsModels = Server.CreateObject("ADODB.Recordset")
  rsModels.open SQL,eConn,3,1,1
  
  .write "<SPAN CLASS=SmallText><B>Model: </B></SPAN>"
  .write "<SELECT NAME=""CRP_Model"" CLASS=SmallText>" & vbCrLf
  
  .write "<OPTION VALUE="""">Select from list</OPTION>" & vbCrLf
  
  do while not rsModels.EOF
  
    .write "<OPTION VALUE=""" & rsModels("Model") & """"
    if LCase(Model) = LCase(rsModels("Model")) then
      .write " SELECTED"
    end if
    .write ">" & rsModels("Model") & "</OPTION>" & vbCrLf
    rsModels.MoveNext
  
  loop
  
  rsModels.close
  set rsModels = nothing
  
  .write "      </SELECT>" & vbCrLf

  .write "      <INPUT TYPE=""IMAGE"" SRC=""http://us.fluke.com/images/applications/wtb/go.gif"" BORDER=0 ALT=""Click GO to Search"">" & vbCrLf
  .write "   </TD>" & vbCrLf
  .write "	</TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf & vbCrLf
  .write "<P>" & vbCrLf
  
end with
  
' --------------------------------------------------------------------------------------  
' List Search Results
' --------------------------------------------------------------------------------------
  
with response

if request("CRP_Model") <> "" then

  SQL = "SELECT vcturbo_replaceable_parts_xref.Model AS Model, " &_
        "       vcturbo_replaceable_parts_xref.Brand AS Brand, " &_
        "       vcturbo_replaceable_parts_xref.Family AS Family, " &_
        "       vcturbo_replaceable_parts_xref.Part_Description AS Part_Description, " &_
        "       vcturbo_replaceable_parts_xref.Part_Exception AS Part_Exception, " &_
        "       vcturbo_replaceable_parts_xref.Serial_Range AS Serial_Range, " &_
        "       vcturbo_replaceable_parts_xref.Item_Number AS Item_Number, " &_
        "       vcturbo_product_family.list_price AS US_List, " &_
        "       vcturbo_product_family.PT2 AS Replaced_By, " &_
        "       vcturbo_product_family.C3 AS Status_Code " &_
        "FROM   vcturbo_replaceable_parts_xref LEFT OUTER JOIN " &_
        "       vcturbo_product_family ON vcturbo_replaceable_parts_xref.Item_Number = vcturbo_product_family.pfid " &_
        "WHERE (vcturbo_replaceable_parts_xref.Model=N'" & Model & "') " &_
        "ORDER BY vcturbo_replaceable_parts_xref.Part_Description"

  Set rsParts = Server.CreateObject("ADODB.Recordset")
  
  rsParts.open SQL,eConn,3,1,1
  
  if not rsParts.EOF then
  
    FormName = "ListParts"
  
    .write "<TABLE WIDTH=""590"" BGCOLOR=""#000000"" CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
    .write "<TR>"
    .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=PrdSpecHeadline>Model</TD>" & vbCrLf
    .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=PrdSpecHeadline>Part Description</TD>" & vbCrLf
    '.write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=PrdSpecHeadline>Part Exception</TD>" & vbCrLf
    .write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=PrdSpecHeadline>Serial Number Range</TD>" & vbCrLf
    .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=PrdSpecHeadline>Part Number</TD>" & vbCrLf
    .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=PrdSpecHeadline>Replaced By</TD>" & vbCrLf
    if Country = "us" then
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=PrdSpecHeadline>US List Price</TD>" & vbCrLf
    end if
    .write "</TR>" & vbCrLf
    
    Replaced_By_Flag = False
    
    BGColor = "#FFFFFF"
    
    do while not rsParts.EOF
    
      .write "<TR>" & vbCrLf
      
      ' Model
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SmallText NOWRAP>"
      .write rsParts("Model") & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Part Description
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SmallText>" 
      .write rsParts("Part_Description") & vbCrLf
      .write "</TD>" & vbCrLf

      ' Part Exception
      '.write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SmallText>" 
      '.write rsParts("Part_Exception") & vbCrLf
      '.write "</TD>" & vbCrLf
      
      ' Serial Number Range
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SmallText>" 
      .write rsParts("Serial_Range") & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Part Number
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""RIGHT"" CLASS=SMALLText>" 
      .write rsParts("Item_Number") & vbCrLf
      .write "</TD>" & vbCrLf

      ' Replaced By
      if isnumeric(rsParts("Replaced_By")) then Replaced_By_Flag = True
      .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SmallText>"
      .write "<FONT COLOR=""#FF0000"">" & vbCrLf
      .write rsParts("Replaced_By") & vbCrLf
      .write "</FONT>" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' US List
      if isnumeric(rsParts("US_List")) then Price = CDBL(rsParts("US_List")) / 100 else Price = 0
      if Country = "us" then
        .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SmallText>"
        .write FormatNumber(Price,2) & vbCrLf
        .write "</TD>" & vbCrLf
      end if
      
      .write "</TR>" & vbCrLf
      
      rsParts.MoveNext
      
    loop
    
    if Replaced_By_Flag = True then
    
      .write "<TR>" & vbCrLf
      .write "<TD COLSPAN=10 BGCOLOR=""#CCCCCC"" CLASS=SmallText>" & vbCrLf
      .write "One or more replacement parts above indicate a Replaced By part number as listed below." & vbCrLf
      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf

      rsParts.MoveFirst
      
      do while not rsParts.EOF
      
        if isnumeric(rsParts("Replaced_By")) then
        
          SQL = "SELECT  pfid AS Item_Number, short_description AS Oracle_Description, list_price AS US_List " &_
                "FROM vcturbo_product_family " &_
                "WHERE pfid = '" & rsParts("Replaced_By") & "'"
          Set rsParts_RB = Server.CreateObject("ADODB.Recordset")
          rsParts_RB.open SQL,eConn,3,1,1
          
          if not rsParts_RB.EOF then
  
            .write "<TR>" & vbCrLf
            .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left""  CLASS=SmallText NOWRAP>" & "&nbsp;" & "</TD>" & vbCrLf
            .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left""  CLASS=SmallText>" & "&nbsp;" & "</TD>" & vbCrLf
            '.write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left""  CLASS=SmallText>" & "&nbsp;" & "</TD>" & vbCrLf
            .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left""  CLASS=SmallText>" & "&nbsp;" & "</TD>" & vbCrLf
            .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SmallText><FONT COLOR=""#FF0000"">" & rsParts_RB("Item_Number") & "</FONT></TD>" & vbCrLf

            if Country = "us" then
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SmallText>" & "&nbsp;" & "</TD>" & vbCrLf
              if isnumeric(rsParts_RB("US_List")) then Price = CDBL(rsParts_RB("US_List")) / 100 else Price = 0
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SmallText>" & FormatNumber(Price,2) & "</TD>" & vbCrLf
            end if

            .write "</TR>" & vbCrLf

          end if
          
          rsParts_RB.close
          set rsParts_RB = nothing
          
        end if
        
        rsParts.MoveNext
        
      loop
      
    end if
    
    .write "</TABLE>" & vbCrLf & vbCrLf
    
  else
    .write "No Common Replacement Parts have been found for the model you have selected."
  end if
end if
  
end with

Call DisConnect_eStoreDatabase

%>
<!-- Remove any code below this line -->

</BODY>
</HTML>
   