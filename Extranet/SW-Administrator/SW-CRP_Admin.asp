<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   3/18/2005
' Edit DB - eStore.vcturbo_replaceable_parts_xref
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_DB.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_eStore.asp" -->
<%

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim DebugFlag
DebugFlag = false

if DebugFlag then
  response.write "<TABLE>"
  for each item in request.form
    response.write "<TR><TD>" & item & "</TD><TD>" & request.form(item) & "</TD></TR>"
  next
  response.write "</TABLE>"
end if

Admin_Access = 3 ' Remove when Security is in place

Dim Site_ID, Site_Code, FormName, FormNameNav, Script_Name

Script_Name = request.ServerVariables("SCRIPT_NAME")

FormName = ""
FormNameNav = ""

if not isblank(request("Site_ID")) then
	Site_ID = request("Site_ID")
elseif not isblank(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
	Site_ID = 3
end if

if not isblank(request("BackURL")) then
	Session("BackURL") = request("BackURL")
  BackURL = request("BackURL")
elseif not isblank(Session("Session")) then
  BackURL = Session("BackURL")
else
	BackURL = ""
end if

Dim Post_Method
Post_Method = "POST"

if request("Language") = "XON" then
  Session("ShowTranslation") = True
elseif request("Language")="XOF" then
  Session("ShowTranslation") = False
end if

Dim ID
if not isblank(request("ID")) then
	ID = LCase(request("ID"))
else
	ID = 0
end if

Dim model
if not isblank(request("Model")) then
	model = request("Model")
else
	model = ""
end if

Dim family
if not isblank(request("Family")) then
	family = request("Family")
else
	family = ""
end if

Dim brand
if not isblank(request("Brand")) then
	brand = request("Brand")
else
	brand = ""
end if

Dim sequence
if not isblank(request("Sequence")) then
	sequence = Trim(LCase(request("Sequence")))
  if sequence = "cancel" then
    sequence = "list parts"
    ID = 0
  end if
else
	sequence = "list parts"
  ID = 0
end if

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

' --------------------------------------------------------------------------------------
' Determine Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

if sequence <> "export to excel" then

  Screen_Title = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Common Replacement Parts Lookup - Administration Screen",Alt_Language,conn)
  Bar_Title = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=NormalBoldGold>" & Translate("Common Replacement Parts Lookup - Administration Screen",Login_Language,conn) & "</FONT>"
  Navigation = False
  Top_Navigation = False
  Content_Width = 95  ' Percent

  %>
  <!--#include virtual="/SW-Common/SW-Header.asp"-->
  <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
  <IFRAME STYLE="display:none;position:absolute;width:148;height:194;z-index=100" ID="CalFrame" MARGINHEIGHT=0 MARGINWIDTH=0 NORESIZE FRAMEBORDER=0 SCROLLING=NO SRC="/SW-Common/SW-Calendar_PopUp.asp"></IFRAME>
  <%

  response.write "<FONT CLASS=NormalBoldRed>"
  select case Admin_Access
    case 2
      response.write Translate("Content Submitter",Login_Language,conn)
    case 3
      response.write Translate("Gateway Application Administrator",Login_Language,conn)
    case 4
      response.write Translate("Content Administrator",Login_Language,conn)
    case 6
      response.write Translate("Account Administrator",Login_Language,conn)
    case 8
      response.write Translate("Site Administrator",Login_Language,conn)
    case 9
      response.write Translate("Domain Administrator",Login_Language,conn)
  end select
  response.write "</FONT><BR>" & Admin_FirstName & " " & Admin_LastName & "<BR>" & Admin_Company & "<BR><BR>"
  
end if

' --------------------------------------------------------------------------------------

if  Admin_Access = 2 or Admin_Access = 3 or Admin_Access = 8 or Admin_Access = 9 then

  if DebugFlag = true then

    Call Connect_eStoreDatabase
    
    SQL = "SELECT * FROM vcturbo_replaceable_parts_xref WHERE (Part_Description LIKE '%""%')"
    Set rsParts = Server.CreateObject("ADODB.Recordset")
    rsParts.open SQL,eConn,3,1,1
    
    do while not rsParts.EOF
    
      SQLU = "UPDATE vcturbo_replaceable_parts_xref SET Part_Description=N'" & Replace(rsParts("Part_Description"),"""","") & "' WHERE ID=" & rsParts("ID")
      eConn.execute SQLU
      
      rsParts.MoveNext
    loop
    
    rsParts.close
    set rsParts = nothing
    
    Call Disconnect_eStoreDatabase
    
  end if

' --------------------------------------------------------------------------------------
' Save
' --------------------------------------------------------------------------------------

  if sequence = "save" or sequence = "update" then
  
    Call Connect_eStoreDatabase
    
    if ID = "add" then
      New_ID = Get_New_Record_ID("vcturbo_replaceable_parts_xref", "Model", "0", eConn)
      ID = CInt(New_ID)
    end if
    
    if isnumeric(ID) then
    
      if not isblank(request("Part_Description")) then
        Part_Description = "Part_Description=N'" & Replace(Mid(replace(request("Part_Description"),chr(13) & chr(10),"<BR>"),1,1024),"'","''") & "'"
      else
        Part_Description = "Part_Description=NULL"
      end if
    
      if not isblank(request("Part_Exception")) then
        Part_Exception = "Part_Exception=N'" & Replace(Mid(replace(request("Part_Exception"),chr(13) & chr(10),"<BR>"),1,2048),"'","''") & "'"
      else
        Part_Exception = "Part_Exception=NULL"
      end if
      
      if not isblank(request("Serial_Range")) then
        Serial_Range = "Serial_Range=N'" & Replace(Mid(replace(request("Serial_Range"),chr(13) & chr(10),"<BR>"),1,1024),"'","''") & "'"
      else
        Serial_Range = "Serial_Range=NULL"
      end if
      
      if not isblank(request("Item_Number")) then
        Item_Number = "Item_Number=N'" & Replace(Mid(replace(request("Item_Number"),chr(13) & chr(10),"<BR>"),1,1024),"'","''") & "'"
      else
        Item_Number = "Item_Number=NULL"
      end if
      
      if not isblank(request("Brand")) then
        Brand = "Brand=N'" & Replace(Mid(replace(request("Brand"),chr(13) & chr(10),"<BR>"),1,1024),"'","''") & "'"
      else
        Brand = "Brand=NULL"
      end if
      
      if not isblank(request("Family")) then
        Family = "Family=N'" & Replace(Mid(replace(request("Family"),chr(13) & chr(10),"<BR>"),1,1024),"'","''") & "'"
      else
        Family = "Family=NULL"
      end if
    
      SQL = "UPDATE    vcturbo_replaceable_parts_xref " &_
            "SET       Model=N'"        & Model & "', " &_
            "      " & Part_Description & ", "  &_
            "      " & Part_Exception   & ", "  &_
            "      " & Serial_Range     & ", "  &_
            "      " & Item_Number      & ", "  &_
            "      " & Brand            & ", "  &_
            "      " & Family           & " "   &_
            "WHERE     ID="             & ID

      eConn.execute SQL
      
    elseif ID = "model" and not isblank(Model) and not isblank(request("ModelChange")) then
    
      SQL = "UPDATE    vcturbo_replaceable_parts_xref " &_
            "SET       Model=N'" & request("ModelChange")  & "' " &_    

            "WHERE     Model='" & Model & "'"
  
      eConn.execute SQL
      
      Model = request("ModelChange")
    
    end if  
    
    sequence = "list parts"
    ID = 0
    
    Call Disconnect_eStoreDatabase
    
  end if
  
' --------------------------------------------------------------------------------------
' Delete
' --------------------------------------------------------------------------------------  

  if sequence = "delete" then
  
    Call Connect_eStoreDatabase
    SQL = "DELETE FROM vcturbo_replaceable_parts_xref WHERE (ID=" & ID & ")"
    
    eConn.execute SQL
    
    Call Disconnect_eStoreDatabase
    
    ID = 0
    sequence = "list parts"
    
  end if
  
' --------------------------------------------------------------------------------------
' Add Part
' --------------------------------------------------------------------------------------  
 
  if sequence = "add part" then
  
    Call Connect_eStoreDatabase
  
    BGColor = "#FFFFFF"
    
    FormName = "AddPart"
    
    with response
      .write "<FORM ACTION=""" & Script_Name & """ NAME=""" & FormName & """ METHOD=""" & Post_Method & """ LANGUAGE=""JavaScript"" ONSUBMIT=""return(CheckRequiredFields(this.form));"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""add"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""BN"" VALUE="""">" & vbCrLf
    end with
   
    Call Table_Begin
    
    with response
      .write "<TABLE BGCOLOR=""#000000"" CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
      .write "<TR>"
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=SMALLBOLDGOLD>" & Translate("Model",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=SMALLBOLDGOLD>" & Translate("Part Description",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=SMALLBOLDGOLD>" & Translate("Comments",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SMALLBOLDGOLD>" & Translate("Serial Number Range",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD>" & Translate("Part Number",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD>" & Translate("Family Code",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD COLSPAN=2>" & Translate("Action",Login_Language,conn) & "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      
      .write "<TR>" & vbCrLf
      
      ' Model
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL NOWRAP>"
      .write "<A NAME=""EAdd""></A>"
      .write "<INPUT TYPE=""Text"" NAME=""Model"" VALUE=""" & Model & """ CLASS=SMALL SIZE=10 MAXLENGTH=50 LANGUAGE=""JavaScript"" ONCHANGE=""return(ResetFamily(this.form));"">" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Part Description
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL>"       
      .write "<TEXTAREA COLS=""25"" ROWS=""4"" CLASS=SMALL NAME=""Part_Description""></TEXTAREA>" & vbCrLf
      .write "</TD>" & vbCrLf

      ' Part Comments / Exception
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL>" 
      .write "<TEXTAREA COLS=""25"" ROWS=""4"" CLASS=SMALL NAME=""Part_Exception""></TEXTAREA>" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Serial Number Range
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL>" 
      .write "<TEXTAREA COLS=""25"" ROWS=""4"" CLASS=SMALL NAME=""Serial_Range""></TEXTAREA>" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Part Number
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""CENTER"" CLASS=SMALL>" 
      .write "<INPUT TYPE=TEXT CLASS=SMALL SIZE=10 MAXLENGTH=20 NAME=""Item_Number"" VALUE="""">" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Family Code
      
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""CENTER"" CLASS=SMALL>" & vbCrLf
      
      if isblank(Family) then
        SQL = "SELECT DISTINCT Family FROM vcturbo_replaceable_parts_xref WHERE Model='" & Model & "'"
        Set rsFamily = Server.CreateObject("ADODB.Recordset")
        rsFamily.open SQL,eConn,3,1,1
      
        if not rsFamily.EOF then
          Family = Trim(rsFamily("Family"))
        end if
        
        rsFamily.close
        set rsFamily = nothing
      end if
      
      SQL =     "SELECT DISTINCT Family " &_
                "FROM   vcturbo_product_family " &_
                "WHERE  (Dept_ID = 20) AND " &_
                "       (Family IS NOT NULL) AND " &_
                "       (Family <> '') AND " &_
                "       (Family NOT LIKE '%ACC%') AND " &_
                "       (Family <> 'CVAS') AND " &_
                "       (Family <> 'DISTRA') AND " &_
                "       (Family <> 'PARTS') AND " &_
                "       (Family <> 'SERVC') AND " &_
                "       (Model_Group NOT LIKE '%ACC%') AND " &_
                "       (Model_Group <> 'CALSW') " &_
                "ORDER BY Family"
                
      Set rsFamily = Server.CreateObject("ADODB.Recordset")
      rsFamily.open SQL,eConn,3,1,1                
      
      .write "<SELECT NAME=""Family"" CLASS=Small>" & vbCrLf
      .write "<OPTION VALUE="""">Select</OPTION>" & vbCrLf
      
      .write "<OPTION VALUE=""ACC"""
      if LCase(Family) = LCase("acc") then .write " SELECTED"
      .write ">ACC</OPTION>" & vbCrLf

      do while not rsFamily.EOF
      
        .write "<OPTION VALUE=""" & rsFamily("Family") & """"
        if LCase(Family) = LCase(Trim(rsFamily("Family"))) then .write " SELECTED"
        .write ">" & rsFamily("Family") & "</OPTION>" & vbCrLf
        
        rsFamily.MoveNext
      
      loop
      
      rsFamily.close
      set rsFamily = nothing
      
      .write "</SELECT>" & vbCrLf
      .write "</TD>" & vbCrLf

      ' Button Save
      .write "<TD BGCOLOR=""#00FF00"" ALIGN=""Center"" CLASS=SMALL>"        
      .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE=""Save"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonName('Save');"">" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Button Save / Cancel
      .write "<TD BGCOLOR=""#00FF00"" ALIGN=""Center"" CLASS=SMALL>"        
      .write "<INPUT TYPE=""SUBMIT"" NAME=""Sequence"" VALUE=""Cancel"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonName('Cancel');"">" & vbCrLf
      .write "</TD>" & vbCrLf          
      
      .write "</TR>" & vbCrLf
      
      .write "</TABLE>" & vbCrLf
      
    end with
      
    Call Table_End
    
    with response

      .write "</FORM>"
    
      .write Translate("Changing the model value will establish a new entry in the model select drop down.",Login_Language,conn) & " " & Translate("This feature is used to establishing a new common replacement parts listing for a new model.",Login_language,conn) & vbCrLf
     
    end with
    
    Call Disconnect_eStoreDatabase
      
  end if
  
' --------------------------------------------------------------------------------------
' Change Model
' --------------------------------------------------------------------------------------  

  if sequence = "change model" then
  
    Call Connect_eStoreDatabase
    
    BGColor = "#FFFFFF"
    
    FormName = "ChangeModel"
    
    with response
      .write "<FORM ACTION=""" & Script_Name & """ NAME=""" & FormName & """ METHOD=""" & Post_Method & """ LANGUAGE=""JavaScript"" ONSUBMIT=""return(CheckRequiredFields(this.form));"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""model"">" & vbCrLf
      .write "<INPUT TYPE=""HIDDEN"" NAME=""BN"" VALUE="""">" & vbCrLf
    end with
   
    Call Table_Begin
    
    with response
      .write "<TABLE BGCOLOR=""#000000"" CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
      .write "<TR>"
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=SMALLBOLDGOLD>" & Translate("Change Model From",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Left""   CLASS=SMALLBOLDGOLD>" & Translate("Change Model To",Login_Language,conn) & "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD COLSPAN=2>" & Translate("Action",Login_Language,conn) & "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      
      .write "<TR>" & vbCrLf
      
      ' Model
      
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL NOWRAP>"
    
      SQL = "SELECT DISTINCT Model " &_
      "FROM vcturbo_replaceable_parts_xref " &_
      "ORDER BY Model"

      Set rsModels = Server.CreateObject("ADODB.Recordset")
      rsModels.open SQL,eConn,3,1,1
      
      .write "<SPAN CLASS=SMALLBoldWHITE>" & Translate("Select Model",Login_Language,conn) & ": </SPAN>"
      .write "<SELECT NAME=""Model"" CLASS=SMALL>" & vbCrLf
      
      .write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
      
      do while not rsModels.EOF
      
        .write "<OPTION VALUE=""" & rsModels("Model") & """"
        if LCase(Model) = LCase(rsModels("Model")) then .write " SELECTED"
        .write ">" & rsModels("Model") & "</OPTION>" & vbCrLf
        rsModels.MoveNext
      
      loop
      
      rsModels.close
      set rsModels = nothing
      
      .write "</SELECT>" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Change Model To
      .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""CENTER"" CLASS=SMALL>" 
      .write "<INPUT TYPE=TEXT CLASS=SMALL SIZE=10 MAXLENGTH=20 NAME=""ModelChange"" VALUE="""">" & vbCrLf
      .write "</TD>" & vbCrLf

      ' Button Save
      .write "<TD BGCOLOR=""#00FF00"" ALIGN=""Center"" CLASS=SMALL>"        
      .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE=""Update"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonName('Update');"">" & vbCrLf
      .write "</TD>" & vbCrLf
      
      ' Button Save / Cancel
      .write "<TD BGCOLOR=""#00FF00"" ALIGN=""Center"" CLASS=SMALL>"        
      .write "<INPUT TYPE=""SUBMIT"" NAME=""Sequence"" VALUE=""Cancel"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonName('Cancel');"">" & vbCrLf
      .write "</TD>" & vbCrLf          
      
      .write "</TR>" & vbCrLf
      
      .write "</TABLE>" & vbCrLf
      
    end with
      
    Call Table_End
    
    with response
    
      .write "</FORM>"
    
      .write Translate("Changing the Model (to) name will change all occurances of the original Model name (from) in the database.",Login_Language,conn)
    
    end with
    
    Call Disconnect_eStoreDatabase
      
  end if

' --------------------------------------------------------------------------------------
' Export to Excel - .CSV File
' --------------------------------------------------------------------------------------

if sequence = "export to excel" then

  %>
  <SCRIPT LANGUAGE="JavaScript">
    <!--//
    alert("<%=Translate("View Data in Excel",Login_Language,conn)%>\r\n\n<%=Translate("Use the back button of your browser when done viewing / saving this data extract to return to the main menue",Login_Language,conn)%>");
    //-->
  </SCRIPT>
  <%

  Call Connect_eStoreDatabase
  
  Set fstemp = server.CreateObject("Scripting.FileSystemObject")
  Extract_Path = "/find-sales/download/Common_Replacement_Parts.csv"
  File_path = server.mappath(Extract_Path)
  Set filetemp = fstemp.CreateTextFile(File_Path, True)
  
  Column_Titles = """Model""," &_
                  """Part Description""," &_
                  """Comments""," &_
                  """Serial Number Range""," &_
                  """Part Number""," &_
                  """Replaced By""," &_
                  """US List Price""," &_
                  """Family""," &_
                  """Oracle Description"""
                  
  filetemp.writeLine(Column_Titles)
                  

  SQL = "SELECT     vcturbo_replaceable_parts_xref.ID AS ID, vcturbo_replaceable_parts_xref.Model AS Model, " &_
        "           vcturbo_replaceable_parts_xref.Brand AS Brand, vcturbo_replaceable_parts_xref.Family AS Family, " &_
        "           vcturbo_replaceable_parts_xref.Part_Description AS Part_Description, vcturbo_replaceable_parts_xref.Part_Exception AS Part_Exception, " &_
        "           vcturbo_replaceable_parts_xref.Serial_Range AS Serial_Range, vcturbo_replaceable_parts_xref.Item_Number AS Item_Number, " &_
        "           vcturbo_product_family.short_description AS Oracle_Description, vcturbo_product_family.list_price AS US_List, vcturbo_product_family.PT2 AS Replaced_By, " &_
        "           vcturbo_product_family.C3 AS Status_Code " &_
        "FROM       vcturbo_replaceable_parts_xref LEFT OUTER JOIN " &_
        "           vcturbo_product_family ON vcturbo_replaceable_parts_xref.Item_Number = vcturbo_product_family.pfid " &_
        "ORDER BY vcturbo_replaceable_parts_xref.Model, vcturbo_replaceable_parts_xref.Part_Description"
        
    Set rsParts = Server.CreateObject("ADODB.Recordset")
    
    rsParts.open SQL,eConn,3,1,1
    
    old_model = LCase(rsParts("Model"))
    
    do while not rsParts.EOF
    
      if old_model <> LCase(rsParts("Model")) then
      
        Part_Info = ""
        old_model = LCase(rsParts("Model"))
        
        filetemp.writeLine(Part_Info)
      
      end if
      
      if isnumeric(rsParts("US_List")) then Price = FormatNumber(CDBL(rsParts("US_List")) / 100,2) else Price = "0.00"
        
        Part_Info = """" & ConvertBR2CHR(rsParts("Model"))              & """," &_
                    """" & ConvertBR2CHR(rsParts("Part_Description"))   & """," &_
                    """" & ConvertBR2CHR(rsParts("Part_Exception"))     & """," &_
                    """" & ConvertBR2CHR(rsParts("Serial_Range"))       & """," &_
                    """" & ConvertBR2CHR(rsParts("Item_Number"))        & """," &_
                    """" & ConvertBR2CHR(rsParts("Replaced_By"))        & """," &_
                    """" & ConvertBR2CHR(Price)                         & """," &_
                    """" & ConvertBR2CHR(rsParts("Family"))             & """," &_
                    """" & ConvertBR2CHR(rsParts("Oracle_Description")) & """"
        filetemp.writeLine(Part_Info)
        
        rsParts.MoveNext
      
    loop
    
    rsParts.close
    set rsParts = nothing
    
  Call Disconnect_eStoreDatabase
  Call Connect_SiteWide
  
  filetemp.close
  set filetemp = nothing
  set fstemp = nothing
  
  response.redirect Extract_Path
  
end if

' --------------------------------------------------------------------------------------
' Edit
' --------------------------------------------------------------------------------------  

	if sequence = "list parts" or sequence = "edit" then
    
    Call Connect_eStoreDatabase
    
    FormNameNav = "SelectModel"
    
    with response
    .write "<FORM ACTION=""" & Script_Name & """ NAME=""" & FormNameNav & """ METHOD=""" & Post_Method & """ LANGUAGE=""JavaScript"" ONSUBMIT=""return(CheckRequiredFieldsNav(this.form));"">"   & vbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE="""   & Site_ID & """>" & vbCrLf
    .write "<INPUT TYPE=""HIDDEN"" NAME=""BN"" VALUE="""">" & vbCrLf    
    
    Call Table_Begin
    
    SQL = "SELECT DISTINCT Model " &_
          "FROM vcturbo_replaceable_parts_xref " &_
          "ORDER BY Model"
          
    Set rsModels = Server.CreateObject("ADODB.Recordset")
    rsModels.open SQL,eConn,3,1,1
    
    .write "<SPAN CLASS=SMALLBoldWHITE>" & Translate("Model",Login_Language,conn) & ": </SPAN>"
    .write "<SELECT NAME=""Model"" CLASS=SMALL>" & vbCrLf
    
    .write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
    
    do while not rsModels.EOF
    
      .write "<OPTION VALUE=""" & rsModels("Model") & """"
      if LCase(Model) = LCase(rsModels("Model")) then
        .write " SELECTED"
      end if
      .write ">" & rsModels("Model") & "</OPTION>" & vbCrLf
      rsModels.MoveNext
    
    loop
    
    .write "</SELECT>" & vbCrLf
    
    .write "&nbsp;&nbsp;"
    .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE="" List Parts "" CLASS=NavleftHighlight1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonNameNav('List Parts');"">" & vbCrLf
    .write "&nbsp;&nbsp;&nbsp;&nbsp;"
    .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE=""Add Part"" CLASS=NavLeftHighlight1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonNameNav('Add Part');"">" & vbCrLf
    .write "&nbsp;&nbsp;&nbsp;&nbsp;<SPAN CLASS=SmallBoldWhite>" & Translate("Tools",Login_Language,conn) & ": </SPAN>" & vbCrLf
    .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE=""Change Model"" CLASS=NavLeftHighlight1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonNameNav('Change Model');"">" & vbCrLf
    .write "&nbsp;&nbsp;&nbsp;&nbsp;"
    .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE=""Export to Excel"" CLASS=NavLeftHighlight1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonNameNav('Export to Excel');"">" & vbCrLf
    
    if not isblank(BackURL) then
      .write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      .write "<INPUT TYPE=BUTTON CLASS=NAVLEFTHIGHLIGHT1 NAME=BACKURL VALUE="" " & Translate("Home",Login_Language,conn) & " "" LANGUAGE=""JavaScript"" ONCLICK=""window.location.href='" & BackURL & "';"">" & vbCrLf
    else
      .write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      .write "<INPUT TYPE=BUTTON CLASS=NAVLEFTHIGHLIGHT1 NAME=BACKURL VALUE="" " & Translate("Exit",Login_Language,conn) & " "" LANGUAGE=""JavaScript"" ONCLICK=""window.opener.focus(); window.self.close();"">" & vbCrLf
    end if

    
    end with
    
    rsModels.close
    set rsModels = nothing
    
    Call Table_End
    
    response.write "</FORM>"
    
    if sequence = "list parts" or ID <> 0 then
      response.write Translate("Listing is alphebetically ordered by the Part Description field.",Logon_Language,conn) & "<P>" & vbCrLf
    end if

    SQL = "SELECT DISTINCT " &_
          "       vcturbo_replaceable_parts_xref.ID AS ID, vcturbo_replaceable_parts_xref.Model AS Model, " &_
          "       vcturbo_replaceable_parts_xref.Brand AS Brand, vcturbo_replaceable_parts_xref.Family AS Family, " &_
          "       vcturbo_replaceable_parts_xref.Part_Description AS Part_Description, vcturbo_replaceable_parts_xref.Part_Exception AS Part_Exception, " &_
          "       vcturbo_replaceable_parts_xref.Serial_Range AS Serial_Range, vcturbo_replaceable_parts_xref.Item_Number AS Item_Number, " &_
          "       vcturbo_product_family.short_description AS Oracle_Description, vcturbo_product_family.list_price AS US_List, vcturbo_product_family.PT2 AS Replaced_By, " &_
          "       vcturbo_product_family.C3 AS Status_Code " &_
          "FROM   vcturbo_replaceable_parts_xref LEFT OUTER JOIN " &_
          "       vcturbo_product_family ON vcturbo_replaceable_parts_xref.Item_Number = vcturbo_product_family.pfid " &_
          "WHERE (vcturbo_replaceable_parts_xref.Model=N'" & Model & "') " &_
          "ORDER BY vcturbo_replaceable_parts_xref.Part_Description"

    Set rsParts = Server.CreateObject("ADODB.Recordset")
    
    rsParts.open SQL,eConn,3,1,1
    
    if not rsParts.EOF then
    
      FormName = "EditModel"
    
      with response
        .write "<FORM ACTION=""" & Script_Name & """ NAME=""" & FormName & """ METHOD=""" & Post_Method & """ LANGUAGE=""JavaScript"" ONSUBMIT=""return(CheckRequiredFields(this.form));"">" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Brand"" VALUE="""   & Brand & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Model"" VALUE="""   & Model & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""Family"" VALUE="""  & Family & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""ID"" VALUE=""" & ID & """>" & vbCrLf
        .write "<INPUT TYPE=""HIDDEN"" NAME=""BN"" VALUE="""">" & vbCrLf
      end with
    
      Call Table_Begin
      
      with response
        .write "<TABLE BGCOLOR=""#000000"" CELLPADDING=2 CELLSPACING=1 BORDER=0  WIDTH=""100%"">" & vbCrLf
        .write "<TR>"
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Left"" CLASS=SMALLBOLDGOLD>"   & Translate("Model",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Left"" CLASS=SMALLBOLDGOLD>"   & Translate("Part Description",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Left"" CLASS=SMALLBOLDGOLD>"   & Translate("Comments",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""CENTER"" CLASS=SMALLBOLDGOLD>" & Translate("Serial Number Range",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD>" & Translate("Part Number",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD>" & Translate("Replaced By",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD>" & Translate("US List Price",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Left"" CLASS=SMALLBOLDGOLD>"   & Translate("Oracle Description",Login_Language,conn) & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""#000000"" ALIGN=""Center"" CLASS=SMALLBOLDGOLD COLSPAN=2>" & Translate("Action",Login_Language,conn) & "</TD>" & vbCrLf
        .write "</TR>" & vbCrLf
      
      Replaced_By_Flag = False
      
      do while not rsParts.EOF
      
        if CInt(rsParts("ID")) = CInt(ID) and sequence = "edit" then
          BGColor   = "#ffff66"
          BGColor_R = "#FFFFFF"
          BGColor_G = "#FFFFFF"
        else
          BGColor   = "#FFFFFF"
          BGColor_R = "#FF0000"
          BGColor_G = "#00FF00"
        end if
      
        .write "<TR>" & vbCrLf
        
        ' Model
        .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL NOWRAP>"
        .write "<A NAME=""E" & rsParts("ID") & """></A>"
        .write rsParts("Model") & "</TD>" & vbCrLf
        .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL>" 
        
        ' Part Description
        if CInt(rsParts("ID")) = CInt(ID) and sequence = "edit" then
          .write "<TEXTAREA COLS=""25"" ROWS=""4"" CLASS=SMALL NAME=""Part_Description"">"
          if instr(1,rsParts("Part_Description"),"<BR>") > 0 then
            .write replace(rsParts("Part_Description"),"<BR>",chr(13) & chr(10))
          else
            .write rsParts("Part_Description")
          end if
          .write "</TEXTAREA>" & vbCrLf
        else
          .write rsParts("Part_Description")
        end if
        .write "</TD>" & vbCrLf

        ' Part Comments / Exception
        .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL>" 
        if CInt(rsParts("ID")) = CInt(ID) and sequence = "edit" then
          .write "<TEXTAREA COLS=""25"" ROWS=""4"" CLASS=SMALL NAME=""Part_Exception"">"
          if instr(1,rsParts("Part_Exception"),"<BR>") > 0 then
            .write replace(rsParts("Part_Exception"),"<BR>",chr(13) & chr(10))
          else
            .write rsParts("Part_Exception")
          end if
          .write "</TEXTAREA>" & vbCrLf
        else
          .write rsParts("Part_Exception")
        end if
        .write "</TD>" & vbCrLf
        
        ' Serial Number Range
        .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""Left"" CLASS=SMALL>" 
        if CInt(rsParts("ID")) = CInt(ID) and sequence = "edit" then
          .write "<TEXTAREA COLS=""25"" ROWS=""4"" CLASS=SMALL NAME=""Serial_Range"">"
          if instr(1,rsParts("Serial_Range"),"<BR>") > 0 then
            .write replace(rsParts("Serial_Range"),"<BR>",chr(13) & chr(10))
          else
            .write rsParts("Serial_Range")
          end if
          .write "</TEXTAREA>" & vbCrLf
        else
          .write rsParts("Serial_Range")
        end if
        .write "</TD>" & vbCrLf
        
        ' Part Number
        .write "<TD BGCOLOR=""" & BGColor & """ ALIGN=""RIGHT"" CLASS=SMALL>" 
        if CInt(rsParts("ID")) = CInt(ID) and sequence = "edit" then
          .write "<INPUT TYPE=TEXT CLASS=SMALL SIZE=10 MAXLENGTH=20 NAME=""Item_Number"" VALUE=""" & rsParts("Item_Number") & """>" & vbCrLf
        else
          .write rsParts("Item_Number")
        end if
        .write "</TD>" & vbCrLf

        ' Replaced By
        if isnumeric(rsParts("Replaced_By")) then Replaced_By_Flag = True
        .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SMALLBoldRed>" & rsParts("Replaced_By") & "</TD>" & vbCrLf
        
        ' US List
        if isnumeric(rsParts("US_List")) then Price = CDBL(rsParts("US_List")) / 100 else Price = 0
        .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SMALL>" & FormatNumber(Price,2) & "</TD>" & vbCrLf
        
        ' Oracle Description
        if instr(1,rsParts("Oracle_Description"),",") > 0 then Oracle_Description = Replace(Replace(Replace(rsParts("Oracle_Description"), " ," ,","), ", ", ","), ",", ", ") else Oracle_Description = rsParts("Oracle_Description")
        .write "<TD BGCOLOR=""#CCCCCC"" ALIGN=""Left"" CLASS=SMALL>" & ProperCase(Oracle_Description) & "</TD>" & vbCrLf
        
        ' Button Edit / Delete
        if sequence = "list parts" then
          .write "<TD BGCOLOR=""" & BGColor_G & """ ALIGN=""Center"" CLASS=SMALL>"
          .write "<INPUT TYPE=""BUTTON"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""window.location.href='?Site_ID=" & Site_ID & "&Model=" & rsParts("Model") & "&Sequence=edit&ID=" & rsParts("ID") & "#E" & rsParts("ID") & "';"" VALUE="" " & Translate("Edit",Login_Language,conn) & " ""></TD>" & vbCrLf
        elseif CInt(rsParts("ID")) = CInt(ID) then
          .write "<TD BGCOLOR=""#00FF00"" ALIGN=""Center"" CLASS=SMALL>"        
          .write "<INPUT TYPE = ""SUBMIT"" NAME=""Sequence"" VALUE=""Save"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonName('Save');"">" & vbCrLf
        else
          .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Center"" CLASS=SMALL>"
          .write "&nbsp;"
        end if
        .write "</TD>" & vbCrLf
        
        ' Button Save / Cancel
        if sequence = "list parts" then
          .write "<TD BGCOLOR=""" & BGColor_R & """ ALIGN=""Center"" CLASS=SMALL>"        
          .write "<INPUT TYPE=""BUTTON"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""window.location.href='?Site_ID=" & Site_ID & "&Model=" & rsParts("Model") & "&Sequence=delete&ID=" & rsParts("ID") & "';"" VALUE=""" & Translate("Delete",Login_Language,conn) & """></TD>" & vbCrLf
        elseif CInt(rsParts("ID")) = CInt(ID) then
          .write "<TD BGCOLOR=""#00FF00"" ALIGN=""Center"" CLASS=SMALL>"        
          .write "<INPUT TYPE=""SUBMIT"" NAME=""Sequence"" VALUE=""Cancel"" CLASS=NAVLEFTHIGHLIGHT1 LANGUAGE=""JavaScript"" ONCLICK=""SaveButtonName('Cancel');"">" & vbCrLf
        else
          .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Center"" CLASS=SMALL>"
          .write "&nbsp;"
        end if
        .write "</TD>" & vbCrLf          
        
        .write "</TR>" & vbCrLf
        
        rsParts.MoveNext
        
      loop
      
      if Replaced_By_Flag = True then
      
        .write "<TR><TD COLSPAN=10 BGCOLOR=""#CCCCCC"" CLASS=SMALL>" & Translate("One or more replacement parts above indicate a Replaced By Part Number as listed below.",Login_Language,conn) & "</SPAN></TD></TR>" & vbCrLf

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
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left"" CLASS=SMALL NOWRAP>" & "&nbsp;" & "</TD>" & vbCrLf
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left"" CLASS=SMALL>" & "&nbsp;" & "</TD>" & vbCrLf
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left"" CLASS=SMALL>" & "&nbsp;" & "</TD>" & vbCrLf
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Left"" CLASS=SMALL>" & "&nbsp;" & "</TD>" & vbCrLf
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SMALLBoldRed>" & rsParts_RB("Item_Number") & "</TD>" & vbCrLf
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SMALL>" & "&nbsp;" & "</TD>" & vbCrLf
              if isnumeric(rsParts_RB("US_List")) then Price = CDBL(rsParts_RB("US_List")) / 100 else Price = 0
              .write "<TD BGCOLOR=""#FFFFFF"" ALIGN=""Right"" CLASS=SMALL>" & FormatNumber(Price,2) & "</TD>" & vbCrLf
              if instr(1,rsParts_RB("Oracle_Description"),",") > 0 then Oracle_Description = Replace(Replace(Replace(rsParts_RB("Oracle_Description"), " ," ,","), ", ", ","), ",", ", ") else Oracle_Description = rsParts_RB("Oracle_Description")
              .write "<TD BGCOLOR=""#CCCCCC"" ALIGN=""Left"" CLASS=SMALL>" & ProperCase(Oracle_Description) & "</TD>" & vbCrLf        
              .write "<TD BGCOLOR=""#CCCCCC"" ALIGN=""Center"" CLASS=SMALL>" & "&nbsp;" & "</TD>" & vbCrLf
              .write "<TD BGCOLOR=""#CCCCCC"" ALIGN=""Center"" CLASS=SMALL>" & "&nbsp;" & "</TD>" & vbCrLf
              .write "</TR>" & vbCrLf

            end if
            
            rsParts_RB.close
            set rsParts_RB = nothing
            
          end if
          
          rsParts.MoveNext
          
        loop
        
      end if
      
      .write "</TABLE>" & vbCrLf
      
      end with
      
      Call Table_End
      
      with response
      
        .write "</FORM>"
        
      end with
      
      
    else
    
      
      if ID <> 0 then
        response.write "<SPAN CLASS=SmallBoldRed>" & Translate("No Common Replacement Parts were found for this Model.",Login_Language,conn) & "</SPAN>" & vbCrLf
      end if

    end if
    
    rsParts.close
    set rsParts = nothing
    Call Disconnect_eStoreDatabase
    
  end if

' --------------------------------------------------------------------------------------  
' Unauthorized
' --------------------------------------------------------------------------------------
  
else

  %>
  <B><%=Translate("You are not authorized to use this application.",Login_Language,conn)%></B>
  <BR><BR>
  <%Call Nav_Border_Begin%>
  <INPUT TYPE="BUTTON" Value=" <%=Translate("Main Menu",Login_Language,conn)%> " onclick="location.href='default.asp?Site_ID=<%=Site_ID%>'" CLASS=NavLeftHighlight1>
  <%Call Nav_Border_End%>
  <%

end if

%>
<!--#include virtual="/SW-Common/SW-Footer.asp"-->
<%

' --------------------------------------------------------------------------------------
' Functions
' --------------------------------------------------------------------------------------

function ConvertBR2CHR(myString)

  if instr(1,myString,"<BR>") > 0 then
    mystring = replace(myString,"<BR>",chr(13) & chr(10))
  elseif isblank(myString) then
    mystring = ""
  end if
  ConvertBR2CHR = replace(replace(myString,"""",""),"'","''")
  
end function

' --------------------------------------------------------------------------------------

function ConvertCHR2BR(myString)

  if instr(1,myString,chr(13) & chr(10)) > 0 then
    mystring = replace(myString,chr(13) & chr(10),"<BR>")
  elseif isblank(myString) then
    mystring = ""
  end if
  ConvertCHR2BR = replace(myString,"""","")
  
end function

' --------------------------------------------------------------------------------------

Call Disconnect_SiteWide
%>

<SCRIPT Language="JavaScript">
<!--//

<%
if isblank(FormName) then
  response.write "var myForm = document.bubbba" & vbCrLf
else
  response.write "var myForm = document." & FormName & vbCrLf
end if

if isblank(FormNameNav) then
  response.write "var myFormNav = document.bubbba" & vbCrLf
else
  response.write "var myFormNav = document." & FormNameNav & vbCrLf
end if
%>

var ErrorMsg = "";

function CheckRequiredFields() {

  ErrorMsg = "";
  
  if (myForm.BN.value != "Cancel") {
  
    if ("<%=FormName%>" == "ChangeModel" || "<%=FormName%>" == "SelectModel") {
      if (myForm.Model.options[0].selected == true) {
        if (myForm.BN.value == "List Parts" || myForm.BN.value == "Change Model") {
          ErrorMsg = ErrorMsg + "<%=Translate("Select a Model from the down list",Alt_Language,conn)%>\r\n"
        }  
      }
    }
    
    if ("<%=FormName%>" == "ChangeModel") {
      if (myForm.ModelChange.value == "") {
        ErrorMsg = ErrorMsg + "<%=Translate("Change Model To",Alt_Language,conn)%>\r\n"
      }
    }
    
    
    if ("<%=FormName%>" == "AddPart" || "<%=FormName%>" == "EditModel") {
      if (myForm.Model.value == "") {
        ErrorMsg = ErrorMsg + "<%=Translate("Model",Alt_Language,conn)%>\r\n"
      }   
      if (myForm.Part_Description.value == "") {
        ErrorMsg = ErrorMsg + "<%=Translate("Part Description",Alt_Language,conn)%>\r\n"
      }   
      if (myForm.Item_Number.value != "N/A") {
        if (myForm.Item_Number.value.length > 0 && myForm.Item_Number.value.length != 6 && myForm.Item_Number.value.length != 7) {
          ErrorMsg = ErrorMsg + "<%=Translate("Part Number must be a 6 or 7-digit numeric value",Alt_Language,conn)%>\r\n"      
        }
        if (! IsNumeric(myForm.Item_Number.value)) {
          ErrorMsg = ErrorMsg + "<%=Translate("Part Number must be a 6 or 7-digit numeric value",Alt_Language,conn)%>\r\n"      
        }
        if (myForm.Item_Number.value.length == 0) {
          ErrorMsg = ErrorMsg + "<%=Translate("Part Number (Oracle Item Number)",Alt_Language,conn)%>\r\n"      
        }
      }
      
      if ("<%=FormName%>" != "EditModel") {
        if (myForm.Family.options[0].selected == true) {
          ErrorMsg = ErrorMsg + "<%=Translate("Select a Family Code from the drop down list",Alt_Language,conn)%>\r\n"
        }
      }  
    }     
    
    if (ErrorMsg.length) {
      ErrorMsg = "<%=Translate("Please enter the missing information for following REQUIRED fields (or use N/A)",Alt_Language,conn)%>:\r\n\n" + ErrorMsg;
      alert (ErrorMsg);
      return (false);
    }
    else {
      return (true);
    }
  }
}


function CheckRequiredFieldsNav() {

  ErrorMsg = "";
  
  if ("<%=FormNameNav%>" == "ModelChange" || "<%=FormNameNav%>" == "SelectModel") {
    if (myFormNav.Model.options[0].selected == true) {
      if (myFormNav.BN.value == "List Parts" || myFormNav.BN.value == "Change Model") {
        ErrorMsg = ErrorMsg + "<%=Translate("Select a Model from the drop down list",Alt_Language,conn)%>\r\n"
      }  
    }
  }
    
  if (ErrorMsg.length) {
    alert (ErrorMsg);
    return (false);
  }
  else {
    return (true);
  }
}

function SaveButtonName(myButton) {
  myForm.BN.value = myButton;
}  
  
function SaveButtonNameNav(myButton) {
  myFormNav.BN.value = myButton;
}  

function ResetFamily() {
  myForm.Family.options[0].selected = true;
}    

function IsNumeric(sText) {
  var ValidChars = "0123456789";
  var IsNumber = true;
  var Char;
  
  if (sText == null) sText = "";
   
  for (i = 0; i < sText.length && IsNumber == true; i++) { 
    Char = sText.charAt(i); 
    if (ValidChars.indexOf(Char) == -1) {
      IsNumber = false;
    }
  }
  return IsNumber;
}

//-->
</SCRIPT>