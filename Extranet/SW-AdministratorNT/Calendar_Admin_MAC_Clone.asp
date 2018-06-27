<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_DB.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

if isblank(request("Sequence")) then
  Sequence = 0
else
  Sequence = request("Sequence")
end if

if isblank(request("Site_ID")) then
  Site_ID = 24
else
  Site_ID = request("Site_ID")
end if

if isblank(request("MSID")) then
  Sequence = 4
else
  MAC_Source_ID = request("MSID")
end if

if isblank(request("SBID")) then
  Submitted_By_ID = 0
else
  Submitted_By_ID = request("SBID")
end if

if isblank(request("MT")) then
  MAC_Title = ""
else
  MAC_Title = request("MT")
end if

if isblank(request("ML")) then
  MAC_Language = "eng"
else
  MAC_Language = request("ML")
end if

Call Connect_SiteWide

select case CInt(Sequence)

  case 0    ' Configure New MAC Container and assets
  
    Screen_Title   = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Master Asset Container - Clone",Alt_Language,conn)
    Bar_Title      = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=MediumBoldGold>" & Translate("Master Asset Container - Clone",Login_Language,conn) & "</FONT>"
    Navigation     = false
    Top_Navigation = false
    Content_Width  = 95  ' Percent
    %>
    <!--#include virtual="/SW-Common/SW-Header.asp"-->
    <!--#include virtual="/SW-Common/SW-Navigation.asp"-->
    <%
  
    FormName = "CloneMAC"
    
    with response

      .write Translate("This utility will clone an English Language Master Asset Container and all of its associated assets to a new language version of the same.",Login_Language,conn) & "&nbsp;&nbsp;"
      .write Translate("There are three required fields that need to be specified before you click on the [GO] button to create this language version clone; Title of new Master Asset Container (in local language), language, and new administrator of this language version who will responsible for the localization of text, asset file upload, and other asset configuration settings such as, Go Live Date (Begin Date), groups allowed to view these assets, country restrictions, etc.",Login_language,conn)
      .write "<P>" & vbCrLf
    
      .write "<FORM NAME=""" & FormName & """ ACTION=""/SW-Administrator/Calendar_Admin_MAC_Clone.asp"" METHOD=""GET"" LANGUAGE=""JavaScript"" onsubmit=""return CheckRequiredFields(this.form);"">" & vbCrLf
      .write "<INPUT Type=""Hidden"" NAME=""Sequence"" VALUE=""1"">" & vbCrLf
      .write "<INPUT Type=""Hidden"" NAME=""MSID"" VALUE=""" & MAC_Source_ID & """>" & vbCrLf
      .write "<INPUT Type=""Hidden"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>" & vbCrLf
      
    Call Nav_Border_Begin
    
      .write "<TABLE BGCOLOR=""#EEEEEE"">" & vbCrLf
      
      ' Source MAC ID
      .write "<TR>"
      .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
      .write Translate("Source",Login_Language,conn) & " - " & Translate("Master Asset Container",Login_Language,conn) & " - " & Translate("ID",Login_Language,conn) & ":"
      .write "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>" & vbCrLf
    	.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
      .write MAC_Source_ID
      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      
      ' Source MAC Title
      .write "<TR>"
      .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
      .write Translate("Source",Login_Language,conn) & " - " & Translate("Master Asset Container",Login_Language,conn) & " - " & Translate("Title",Login_Language,conn) & ":"
      .write "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>&nbsp;</TD>" & vbCrLf
    	.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
      
      SQL = "SELECT Title FROM Calendar WHERE ID=" & MAC_Source_ID
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open SQL, conn, 3, 3
      .write rs("Title")
      rs.close
      set rs = nothing
      
      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      
      .write "<TR><TD COLSPAN=3 BGCOLOR=""Gray"" CLASS=Medium></TD></TR>"      

      ' Clone MAC Title
      .write "<TR>"
      .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
      .write Translate("Clone",Login_Language,conn) & " - &nbsp;&nbsp;" & Translate("Master Asset Container",Login_Language,conn) & " - " & Translate("Title",Login_Language,conn) & ":"
      .write "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
      .write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"">"
      .write "</TD>" & vbCrLf
    	.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
      .write "<INPUT TYPE=""Text"" NAME=""MT"" SIZE=""50"" MAXLENGTH=""255"" VALUE="""" CLASS=Medium>" & vbCrLf
      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      
      ' Clone Language
      .write "<TR>"
      .write "<TD BGCOLOR=""#EEEEEE"" VALIGN=TOP CLASS=Medium>" & vbCrLf
      .write Translate("Clone",Login_Language,conn) & " - &nbsp;&nbsp;" & Translate("Language",Login_Language,conn) & ":"
      .write "</TD>" & vbCrLf
      .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
      .write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"">"
      .write "</TD>" & vbCrLf
    	.write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>" & vbCrLf
                              
      .write "<SELECT Name=""ML"" CLASS=Medium>" & vbCrLf

      SQL = "SELECT * FROM Language WHERE Oracle_Enable=-1 OR Enable=-1 ORDER BY Language.Sort"
      Set rsLanguage = Server.CreateObject("ADODB.Recordset")
      rsLanguage.Open SQL, conn, 3, 3
      
   	  .write "<OPTION CLASS=Medium VALUE="""">" & Translate("Select from List",Login_Language,conn) & "</OPTION>" & vbCrLf
                            
      do while not rsLanguage.EOF
     	  .write "<OPTION CLASS=Medium VALUE=""" & rsLanguage("Code") & """>" & Translate(rsLanguage("Description"),Login_Language,conn) & "</OPTION>" & vbCrLf
    	  rsLanguage.MoveNext 
      loop
      
      .write "<OPTION CLASS=Medium VALUE=""elo"">" & Translate("English (non-clone)",Login_Language,conn) & "</OPTION>" & vbCrLf
      
      rsLanguage.close
      set rsLanguage=nothing
      set SQL = nothing
      
      .write "</SELECT>" & vbCrLf
      .write "</TD>" & vbCrLf
      .write "</TR>" & vbCrLf
      
      ' Reassign Owner
      ' List all Content Admins

      SQL =       "SELECT UserData.* "
      SQL = SQL & "FROM UserData "
      SQL = SQL & "WHERE (UserData.Site_ID=" & Site_ID & " OR UserData.Site_ID=0)"
      SQL = SQL & "AND (UserData.SubGroups LIKE '%domain%' OR UserData.SubGroups LIKE '%administrator%' OR UserData.SubGroups LIKE '%content%' OR UserData.SubGroups LIKE '%submitter%') "
      SQL = SQL & "ORDER BY UserData.LastName, UserData.FirstName"

      Set rsSubmitters = Server.CreateObject("ADODB.Recordset")
      rsSubmitters.Open SQL, conn, 3, 3
  
      if not rsSubmitters.EOF then
      
  		  .write "<TR>"
        .write "<TD BGCOLOR=""#EEEEEE"" CLASS=Medium>"
        .write Translate("Clone",Login_Language,conn) & " - &nbsp;&nbsp;" & Translate("Reassign Owner to",Login_Language,conn) & ":"
        .write "</TD>"
        .write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER CLASS=Medium>"
        .write "<IMG SRC=""/images/required.gif"" Border=0 WIDTH=""10"" HEIGHT=""10"">"
        .write "</TD>"
	      .write "<TD BGCOLOR=""White"" ALIGN=LEFT CLASS=Medium>"
      
        .write "<SELECT NAME=""SBID"" CLASS=Medium>" & vbCrLf
        .write "<OPTION CLASS=Medium VALUE="""">(-) " & Translate("Unassigned",Login_Language,conn) & "</OPTION>"                  

        do while not rsSubmitters.EOF
          .write "<OPTION CLASS=Region" & rsSubmitters("Region") & "NavMedium VALUE=""" & rsSubmitters("ID") & """"
          if CLng(Submitted_By_Current) = CLng(rsSubmitters("ID")) then .write " SELECTED"
          .write ">"
          if instr(1,rsSubmitters("SubGroups"),"domain") > 0 then
            .write "(D) "
          elseif instr(1,rsSubmitters("SubGroups"),"administrator") > 0 then
            .write "(A) "
          elseif instr(1,rsSubmitters("SubGroups"),"content") > 0 then
            .write "(C) "
          elseif instr(1,rsSubmitters("SubGroups"),"submitter") > 0 then
            .write "(S) "
          else                    
            .write "(-) "
          end if  
          .write rsSubmitters("LastName") & " " & rsSubmitters("FirstName") & "</OPTION>" & vbCrLf
          rsSubmitters.MoveNext
        loop
        
        .write "</SELECT>" & vbCrLf
        
      end if
      
      rsSubmitters.close
      set rsSubmitters = nothing  
      
      .write "</TD>"
      .write "</TR>"
      
      ' Navigation
      .write "<TR>"
      .write "<TD COLSPAN=2 CLASS=Medium BGCOLOR=""#666666"">"
      .write "&nbsp;"
      .write "</TD>"

      .write "<TD CLASS=Medium BGCOLOR=""#666666"">"
      .write "<INPUT TYPE=""Submit"" NAME=""Submit"" VALUE="" " & Translate("Go",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
      .write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      .write "<INPUT TYPE=""Button"" NAME=""Cancel"" VALUE="" " & Translate("Cancel",Login_Language,conn) & " "" CLASS=Navlefthighlight1>"
      .write "</TD>"
      .write "</TR>"

      .write "</TABLE>" & vbCrLf
      
    Call Nav_Border_End
    
      .write "</FORM>" & vbCrLf

    end with
    
    %>
    <!--#include virtual="/SW-Common/SW-Footer.asp"-->
    <%

  case 1    ' Duplicate Assets
  
    ' Clone MAC Container
    
    SQL = "SELECT * FROM Calendar where ID=" & MAC_Source_ID
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn, 3, 3
    
    SQLClone = "UPDATE Calendar SET "
    
    MAC_Clone_ID = Get_New_Record_ID("Calendar", "Content_Group", 0, conn)
    
    for iCounter = 0 To rs.Fields.Count - 1
                    
      Set fld = rs.Fields(iCounter)
      'response.write "Field: " & iCounter & " " & fld.name & " " & fld.type & " " & fld.value & "<BR>"
    
      if UCase(fld.name) <> "ID" then
    
        SQLClone = SQLClone & fld.Name & "="
      
        select case CInt(fld.type)
          case 0                                    ' Null
            SQLClone = SQLClone & "NULL,"
          case 7,8,12,129,133,134,135,200,201,203   ' Char

            if UCase(fld.name) = "ITEM_NUMBER_2" then
              if not isblank(Trim(rs("ITEM_NUMBER"))) then
                SQLClone = SQLClone & "'" & UCase(rs("ITEM_NUMBER")) & " " & UCase(rs("REVISION_CODE")) & " " & UCASE(rs("LANGUAGE")) & "',"
              else  
                SQLClone = SQLClone & "NULL,"
              end if
            elseif isblank(Trim(rs(fld.name))) then
              SQLClone = SQLClone & "NULL,"
            elseif UCase(fld.name) = "ITEM_NUMBER" then
              SQLClone = SQLClone & "NULL,"
            elseif UCase(fld.name) = "REVISION_CODE" then
              SQLClone = SQLClone & "NULL,"
            elseif UCase(fld.name) = "FILE_NAME" then
              SQLClone = SQLClone & "NULL,"
            elseif UCase(fld.name) = "ARCHIVE_NAME" then
              SQLClone = SQLClone & "NULL,"
            elseif UCase(fld.name) = "FILE_NAME_POD" then
              SQLClone = SQLClone & "NULL,"
            elseif UCase(fld.name) = "LDATE" then
              SQLClone = SQLClone & "'" & Date() & "',"
            elseif UCase(fld.name) = "BDATE" then
              SQLClone = SQLClone & "'" & Date() & "',"
            elseif UCase(fld.name) = "EDATE" then
              SQLClone = SQLClone & "'" & Date() & "',"
            elseif UCase(fld.name) = "PDATE" then
              SQLClone = SQLClone & "'" & Now() & "',"
            elseif UCase(fld.name) = "XDATE" then
              SQLClone = SQLClone & "'" & Date() & "',"
            elseif UCase(fld.name) = "UDATE" then
              SQLClone = SQLClone & "'" & Now() & "',"
            elseif UCase(fld.name) = "SUBGROUPS" then
              SQLClone = SQLClone & "'" & replace(replace(replace(replace(replace(Trim(rs(fld.Name)),"'","''"),"view, ",""),"fedl, ",""),"shpcrt, ",""),"nomac","") & "',"
            elseif UCase(fld.name) = "LANGUAGE" then
              if LCase(MAC_Language) <> "elo" then
                SQLClone = SQLClone & "'" & MAC_Language & "',"
              else
                SQLClone = SQLClone & "'eng',"
              end if
            else
              SQLClone = SQLClone & "'" & replace(Trim(rs(fld.Name)),"'","''") & "',"
            end if
          case 130, 202                             'NChar, NVarChar
            if  UCase(fld.name) = "TITLE" then
              SQLClone = SQLClone & "'" & Replace(MAC_Title,"'","''") & "',"
            elseif isblank(Trim(rs(fld.name))) then
              SQLClone = SQLClone & "NULL,"
            else
              SQLClone = SQLClone & "N'" & replace(Trim(rs(fld.Name)),"'","''") & "',"
            end if
          case else                                 ' Numeric
            if UCase(fld.name) = "STATUS" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "CLONE" then
              if LCase(MAC_Language) <> "elo" then
                SQLClone = SQLClone & MAC_Source_ID & ","
              else  
                SQLClone = SQLClone & "0,"
              end if
            elseif UCase(fld.name) = "FILE_SIZE" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "ARCHIVE_SIZE" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "FILE_SIZE_POD" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "SUBSCRIPTION" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "LDAYS" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "XDAYS" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "ITEM_NUMBER_SHOW" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "SUBMITTED_BY" then
              SQLClone = SQLClone & Submitted_By_ID & ","
            elseif UCase(fld.name) = "APPROVED_BY" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "REVIEW_BY" then
              SQLClone = SQLClone & "0,"
            elseif UCase(fld.name) = "REVIEW_BY_GROUP" then
              SQLClone = SQLClone & "0,"
            elseif isblank(rs(fld.name)) then
              SQLClone = SQLClone & "NULL,"
            else
              SQLClone = SQLClone & replace(Trim(rs(fld.Name)),"'","''") & ","
            end if
        end select
        
      end if
    
    next
    
    Set fld = nothing
    
    rs.Close
    Set rs = Nothing
    
    SQLClone = Mid(SQLClone,1,len(SQLClone)-1)
    SQLClone = SQLClone & " WHERE ID=" & MAC_Clone_ID
    
    conn.execute SQLClone
    
    ' Now Create the Cloned Assets for the MAC Container
    
    SQLAsset = "SELECT * FROM Calendar where Campaign=" & MAC_Source_ID
    Set rsAsset = Server.CreateObject("ADODB.Recordset")
    rsAsset.Open SQLAsset, conn, 3, 3
    
    if not rsAsset.EOF then
    
      do while not rsAsset.EOF
      
        SQL = "SELECT * FROM Calendar where ID=" & rsAsset("ID")
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open SQL, conn, 3, 3
      
        SQLClone = "UPDATE Calendar SET "
        
        Clone_ID = Get_New_Record_ID("Calendar", "Content_Group", 0, conn)
        
        for iCounter = 0 To rs.Fields.Count - 1
                        
          Set fld = rs.Fields(iCounter)
        
          'response.write "Field: " & iCounter & " " & fld.name & " " & fld.type & " " & fld.value & "<BR>"
          
          if UCase(fld.name) <> "ID" then
        
            SQLClone = SQLClone & fld.Name & "="
          
            select case CInt(fld.type)
              case 0                                      ' Null
                SQLClone = SQLClone & "NULL,"
              case 7,8,12,129,133,134,135,200,201,203     ' Char VChar

                if UCase(fld.name) = "ITEM_NUMBER_2" then
                  if not isblank(Trim(rs("ITEM_NUMBER"))) then
                    SQLClone = SQLClone & "'" & UCase(rs("ITEM_NUMBER")) & " " & UCase(rs("REVISION_CODE")) & " " & UCASE(rs("LANGUAGE")) & "',"
                  else
                    SQLClone = SQLClone & "NULL,"
                  end if
                elseif isblank(Trim(rs(fld.name))) then
                  SQLClone = SQLClone & "NULL,"
                elseif UCase(fld.name) = "ITEM_NUMBER" then
                  SQLClone = SQLClone & "NULL,"
                elseif UCase(fld.name) = "REVISION_CODE" then
                  SQLClone = SQLClone & "NULL,"
                elseif UCase(fld.name) = "FILE_NAME" then
                  SQLClone = SQLClone & "NULL,"
                elseif UCase(fld.name) = "ARCHIVE_NAME" then
                  SQLClone = SQLClone & "NULL,"
                elseif UCase(fld.name) = "FILE_NAME_POD" then
                    SQLClone = SQLClone & "NULL,"
                elseif UCase(fld.name) = "LDATE" then
                  SQLClone = SQLClone & "'" & Date() & "',"
                elseif UCase(fld.name) = "BDATE" then
                  SQLClone = SQLClone & "'" & Date() & "',"
                elseif UCase(fld.name) = "EDATE" then
                  SQLClone = SQLClone & "'" & Date() & "',"
                elseif UCase(fld.name) = "PDATE" then
                  SQLClone = SQLClone & "'" & Now() & "',"
                elseif UCase(fld.name) = "XDATE" then
                  SQLClone = SQLClone & "'" & Date() & "',"
                elseif UCase(fld.name) = "UDATE" then
                  SQLClone = SQLClone & "'" & Now() & "',"
                elseif UCase(fld.name) = "LANGUAGE" then
                  if LCase(MAC_Language) <> "elo" then
                    SQLClone = SQLClone & "'" & MAC_Language & "',"
                  else
                    SQLClone = SQLClone & "'eng',"
                  end if
                elseif UCase(fld.name) = "SUBGROUPS" then
                  SQLClone = SQLClone & "'" & replace(replace(replace(replace(replace(Trim(rs(fld.Name)),"'","''"),"view, ",""),"fedl, ",""),"shpcrt, ",""),"nomac","") & "',"
                else
                  SQLClone = SQLClone & "'" & replace(Trim(rs(fld.Name)),"'","''") & "',"
                end if
              case 130, 202                             'NChar, NVarChar
                if isblank(Trim(rs(fld.name))) then
                  SQLClone = SQLClone & "NULL,"
                else
                  SQLClone = SQLClone & "N'" & replace(Trim(rs(fld.Name)),"'","''") & "',"
                end if
              case else                                 ' Numeric
                if     UCase(fld.name) = "STATUS" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "CLONE" then
                  if LCase(MAC_Language) <> "elo" then
                    SQLClone = SQLClone & MAC_Source_ID & ","
                  else  
                    SQLClone = SQLClone & "0,"
                  end if
                elseif UCase(fld.name) = "CAMPAIGN" then
                  SQLClone = SQLClone & MAC_Clone_ID & ","
                elseif UCase(fld.name) = "FILE_SIZE" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "ARCHIVE_SIZE" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "FILE_SIZE_POD" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "SUBSCRIPTION" then
                    SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "LDAYS" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "XDAYS" then
                    SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "ITEM_NUMBER_SHOW" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "SUBMITTED_BY" then
                  SQLClone = SQLClone & Submitted_By_ID & ","
                elseif UCase(fld.name) = "APPROVED_BY" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "REVIEW_BY" then
                  SQLClone = SQLClone & "0,"
                elseif UCase(fld.name) = "REVIEW_BY_GROUP" then
                  SQLClone = SQLClone & "0,"
                elseif isblank(Trim(rs(fld.name))) then
                  SQLClone = SQLClone & "NULL,"
                else
                  SQLClone = SQLClone & replace(Trim(rs(fld.Name)),"'","''") & ","
                end if
            end select
            
          end if
        
        next
        
        Set fld = nothing
        
        SQLClone = Mid(SQLClone,1,len(SQLClone)-1)
        
        SQLClone = SQLClone & " WHERE ID=" & Clone_ID
        
        response.write "<P>" & SQLClone & "<P>"

        conn.execute SQLClone
        
        rs.Close
        Set rs = Nothing
        
        rsAsset.MoveNext
      
      loop
      
    end if
    
    rsAsset.Close
    set rsAsset = nothing
    
    response.redirect "/SW-Administrator/Site_Utility.asp?ID=Site_Utility&Campaign=" & MAC_Clone_ID & "&Site_ID=" & Site_ID & "&Utility_ID=50&View=4"

  case else
  
    response.redirect "/SW_Administrator/"
    
end select

Call Disconnect_SiteWide

%>

<SCRIPT LANGUAGE="JAVASCRIPT">
<!--//

var CheckMsg = "";
var ErrorMsg = "";
var FormName = document.<%=FormName%>
var CheckFlg = false;
var CheckSts = false;

function CheckRequiredFields() {

  if (FormName.MT.value.length == 0) {
    FormName.MT.style.backgroundColor = "#FFB9B9";
    ErrorMsg = ErrorMsg + "<%=Translate("Clone - Master Asset Container - Title (in Local Language)",Alt_Language,conn)%>\r\n";        
  }

  
  for (var i = 0; i < FormName.ML.length; i++) {
    if (FormName.ML[i].selected) {
      if (FormName.ML.value == '') {
        FormName.ML.style.backgroundColor = "#FFB9B9";            
        ErrorMsg = ErrorMsg + "<%=Translate("Clone - Language",Alt_Language,conn)%>\r\n";
      }    
    }
  }

  for (var i = 0; i < FormName.SBID.length; i++) {
    if (FormName.SBID[i].selected) {
      if (FormName.SBID.value == '') {
        FormName.SBID.style.backgroundColor = "#FFB9B9";            
        ErrorMsg = ErrorMsg + "<%=Translate("Clone - Reassign Owner to",Alt_Language,conn)%>\r\n";
      }    
    }
  }


  if (ErrorMsg.length) {

    ErrorMsg = "<%=Translate("Please enter the missing information for following REQUIRED fields.",Alt_Language,conn)%>:\r\n\n" + ErrorMsg;
    alert (ErrorMsg);
    ErrorMsg = "";
    return (false);
  }
  else {
    ErrorMsg = "";  
    return (true);
  }  
  
}
//-->
</SCRIPT>