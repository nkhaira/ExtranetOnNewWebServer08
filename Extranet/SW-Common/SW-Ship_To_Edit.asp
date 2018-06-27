<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author: Kelly Whitlock
' Date:   5/1/2003
' --------------------------------------------------------------------------------------

response.buffer = true    ' Set in case there is a redirect

%>
  <!--#include virtual="/include/functions_string.asp"-->
  <!--#include virtual="/include/functions_file.asp"-->
  <!--#include virtual="/include/functions_date_formatting.asp"-->
  <!--#include virtual="/include/functions_translate.asp"-->
  <!--#include virtual="/SW-Common/Preferred_Language.asp"-->
  <!--#include virtual="/connections/connection_SiteWide.asp"-->
  <!--#include virtual="/connections/connection_FormData.asp"-->

  <STYLE>
  .ZeroValue  {font-size:8.5pt;font-weight:Bold;color:Black;background:#FFFF99;text-decoration:none;font-family:Arial,Verdana;}
  </STYLE>
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim Debug_Flag
Dim SValue

if request("Debug") = "on" then
  Session("Debug_Flag") = True
elseif request("Debug") = "off" then
  Session("Debug_Flag") = False
elseif isblank(Session("Debug_Flag")) then
  Session("Debug_Flag") = False
end if    

Debug_Flag = Session("Debug_Flag")

Dim Sequence, seqAdd, seqEdit, seqDelete

seqAdd    = 0
seqEdit   = 1
seqUpdate = 2
seqDelete = 3

Dim Caption(3)
Caption(0) = "Add New Drop Ship Address"
Caption(1) = "Edit Drop Ship Address"
Caption(2) = "Update Drop Ship Address"
Caption(3) = "Delete Drop Ship Address"

Dim Account_ID, Border_Toggle

Border_Toggle = 0

Dim Form_Name
Form_Name = "Not_Set"

Contrast = "#FFCC00"

if not isblank(request("Ship_To_ID")) then
  Account_ID = CInt(request("Ship_To_ID"))
else
  Account_ID = 0
end if

if not isblank(request("Site_ID")) then
  Site_ID = CInt(request("Site_ID"))
else
  Site_ID = 0
end if


if not isblank(request("Sequence")) then
  Sequence = CInt(request("Sequence"))
else
  Sequence = seqAdd
end if

with response

' --------------------------------------------------------------------------------------

select case Sequence

  ' --------------------------------------------------------------------------------------
  
  case seqDelete
  
'    SQL = "DELETE FROM Shopping_Cart_Ship_To WHERE ID=" & Account_ID & " AND NTLogin='" & Session("Logon_User") & "'"
    SQL = "UPDATE Shopping_Cart_Ship_To SET Disabled=" & CInt(True) & " WHERE ID=" & Account_ID
    conn.execute (SQL)
    
    %>
    <SCRIPT LANGUAGE="JavaScript">
    opener.location.reload();
    self.close();
    </SCRIPT>
    <%

  ' --------------------------------------------------------------------------------------  

  case seqAdd, seqEdit ' Add & Edit

    select case Sequence
      case seqAdd
        
        Form_Name = "Add"
        
      case seqEdit
        
        Form_Name = "Edit"
        
        SQL = "SELECT * FROM Shopping_Cart_Ship_To WHERE ID=" & Account_ID
        Set rsProfile = Server.CreateObject("ADODB.Recordset")
        rsProfile.Open SQL, conn, 3, 3
        
    end select
    
    response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf & vbCrLf
    response.write "<!-- Whitlock's SiteWide Content Management System SW-CMS 3.0 " & Now & " PST -->" & vbCrLf & vbCrLf
    response.write "<HTML>" & vbCrLf
    response.write "<HEAD>" & vbCrLf
    response.write "<TITLE>" & Translate(Caption(Sequence),Alt_Language,conn) & "</TITLE>" & vbCrLf
    response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
    response.write "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=utf-8"">"
    response.write "<META HTTP-EQUIV=""Expires"" CONTENT=""0"">" & vbCrLf
    response.write "<META HTTP-EQUIV=""PRAGMA"" CONTENT=""NO-CACHE"">" & vbCrLf
    response.write "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache"">" & vbCrLf
    response.write "</HEAD>" & vbCrLf

    response.write "<BODY BGCOLOR=""White"" TOPMARGIN=""0"" LEFTMARGIN=""0"" MARGINWIDTH=""0"" MARGINHEIGHT=""0"" "
    response.write "LINK =""#000000"" "
    response.write "VLINK=""#000000"" "
    response.write "ALINK=""#000000"""
    response.write ">"
   
    Call Table_Begin
    %>
    <FORM ACTION="SW-Ship_To_Edit.asp" NAME="<%=Form_Name%>" METHOD="POST" LANGUAGE="JavaScript" onsubmit="return(CheckRequiredFields(this.form));" onKeyUp="highlight(event)" onClick="highlight(event)">
    <INPUT TYPE=HIDDEN NAME="Sequence" VALUE="<%=seqUpdate%>">
    <INPUT TYPE=HIDDEN NAME="Ship_To_ID" VALUE="<%=Account_ID%>">    
    <INPUT TYPE=HIDDEN NAME="Action" VALUE="<%=Form_Name%>">    

    <TABLE CELLPADDING=4 CELLSPACING=0 BORDER=0 BGCOLOR="<%=Contrast%>" WIDTH="100%">
    
      <TR>
        <TD WIDTH="40%" CLASS=SMALLBOLD><%=Translate("First Name",Login_Language,conn)%> :</TD>
        <TD WIDTH="60%" CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="FirstName" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("FIRSTNAME")%>">
        </TD>
      </TR>
    
      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Last Name",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="LastName" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("LASTNAME")%>">
        </TD>
      </TR>
      
      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Email",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Email" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("EMAIL")%>">
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Phone",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Business_Phone" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Business_Phone")%>">
        <SPAN CLASS=SmallBold>&nbsp;x&nbsp;</SPAN>
        <INPUT CLASS=SMALL NAME="Business_Phone_Extension" SIZE="4" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Business_Phone_Extension")%>">
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Fax",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Business_Fax" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Business_Fax")%>">
        </TD>
      </TR>
      
      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Company",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Company" SIZE="20" MAXLENGTH="100" VALUE="<%if sequence = seqEdit then response.write rsProfile("COMPANY")%>">
        </TD>
      </TR>
              
      <TR>
        <TD CLASS=SMALLBOLD>
        <%=Translate("Shipping",Login_Language,conn)%><BR>
        <%=Translate("Address",Login_Language,conn)%> :
        </TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Shipping_Address" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Shipping_ADDRESS")%>"><BR>
        <INPUT CLASS=SMALL NAME="Shipping_Address_2" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Shipping_ADDRESS_2")%>">
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("City",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Shipping_City" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Shipping_CITY")%>">
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("State or Province",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
           <SELECT NAME="Shipping_State" CLASS=SMALL>
           <%
           .write "<OPTION VALUE="""">" & Translate("Select from list",Login_Language,conn) & "</OPTION>" & vbCrLf
           if sequence = seqEdit then
             sValue = rsProfile("Shipping_State")
           else
             sValue = ""
           end if    
           %>  
           <!--#include virtual="/include/core_states.inc"-->
          </SELECT>
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Postal Code",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <INPUT CLASS=SMALL NAME="Shipping_Postal_Code" SIZE="20" MAXLENGTH="50" VALUE="<%if sequence = seqEdit then response.write rsProfile("Shipping_POSTAL_CODE")%>">
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD><%=Translate("Country",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
          <%
          if sequence = seqEdit then
            sValue = rsProfile("Shipping_Country")
          else
            sValue = ""
          end if    

          Call Connect_FormDatabase
          Call DisplayCountryList("Shipping_Country",SValue,Translate("Select from list",Login_Language,conn),"Small")
          Call Disconnect_FormDatabase
          %>
        </TD>
      </TR>

      <TR>
        <TD CLASS=SMALLBOLD VALIGN=TOP><%=Translate("Shipping Instructions",Login_Language,conn)%> :</TD>
        <TD CLASS=SMALL>
        <TEXTAREA CLASS=SMALL NAME="Comment" COLS=45 ROWS=6><%if sequence = seqEdit then response.write rsProfile("Comment")%></TEXTAREA>
        </TD>
      </TR>


      <TR>
        <TD CLASS=SMALL BGCOLOR="#777777">&nbsp;</TD>
        <TD CLASS=SMALL BGCOLOR="#777777">
          <INPUT CLASS=NAVLEFTHIGHLIGHT1 TITLE="Update Drop Ship Address Information" TYPE="SUBMIT" VALUE="<%=Translate("Update",Login_Language,conn)%>">&nbsp;&nbsp;&nbsp;
          <INPUT CLASS=NavLeftHighlight1 TITLE="Delete Drop Ship Address From List" TYPE="BUTTON" VALUE="<%=Translate("Delete",Login_Language,conn)%>" Language="JavaScript" ONCLICK="location.href='/sw-common/SW-Ship_To_Edit.asp?Sequence=3&Ship_To_ID=<%=Account_ID%>';">
        </TD>
      </TR>
    </TABLE>
    <%
    Call Table_End

    if sequence = seqEdit then
      rsProfile.close
      set rsProfile = nothing
    end if  

    response.write "</FORM>" & vbCrLf
    response.write "</BODY>" & vbCrLf
    response.write "</HTML>" & vbCrLf

     
  ' --------------------------------------------------------------------------------------

  case seqUpdate
  
    ' Create New Account using Add/Update Method to get New_Account_ID
    
    if isblank(request("Ship_To_ID")) or request("Ship_To_ID") = 0 Then
      Account_ID = CInt(Get_New_Record_ID ("Shopping_Cart_Ship_To", "Disabled", 0, conn))
    else
      Account_ID = request("Ship_To_ID")
    end if    
    
    ' Update New Account with Registration Data

    SQL = "UPDATE Shopping_Cart_Ship_To SET " &_
          "NTLogin='" & Session("Logon_User") & "', " &_
          "FirstName='" & Replace(request("FirstName"),"'","''")  & "', " &_
          "LastName='"  & Replace(request("LastName"),"'","''")   & "', " &_
          "Company='"   & Replace(request("Company"),"'","''")    & "', " &_
          "Email='"   & Replace(request("Email"),"'","''")    & "', " &_
          "Business_Phone='"   & Replace(FormatPhone(request("Business_Phone")),"'","''")    & "', " &_                              
          "Business_Phone_Extension='"   & Replace(request("Business_Phone_Extension"),"'","''")    & "', " &_          
          "Business_Fax='"   & Replace(FormatPhone(request("Business_Fax")),"'","''")    & "', " &_                              
          "Shipping_Address='"   & Replace(request("Shipping_Address"),"'","''")    & "', " &_          
          "Shipping_Address_2='"   & Replace(request("Shipping_Address_2"),"'","''")    & "', " &_          
          "Shipping_City='"   & Replace(request("Shipping_City"),"'","''")    & "', " &_          
          "Shipping_State='"   & request("Shipping_State")    & "', " &_
          "Shipping_Postal_Code='"   & Replace(request("Shipping_Postal_Code"),"'","''")    & "', " &_
          "Shipping_Country='"   & request("Shipping_Country")    & "', " & _
          "Comment='"   & Replace(Mid(request("Comment"),1,255),"'","''") & "' " & _          
          "WHERE ID=" & Account_ID
' response.write sql & "<P>"         
    conn.execute (SQL)
    
    %>
    <SCRIPT LANGUAGE="JavaScript">
    opener.location.reload();
    self.close();
    </SCRIPT>
    <%

  case else            
  
end select



' --------------------------------------------------------------------------------------
' Debug
' --------------------------------------------------------------------------------------

if CInt(Session("Debug_Flag")) = CInt(True) then

  .write "<P>------------------------------------------------<BR>" & vbCrLf
  .write "Debug - Key=Value Pairs<BR>" & vbCrLf
  .write "------------------------------------------------<P>" & vbCrLf
  .write "<B>Sequence: " & Sequence & "<P>" & vbCrLf
  .write "<B>QueryString Object</B><BR>" & vbCrLf
  for each item in request.querystring
    .write item & "=" & request.querystring(item) & "<BR>" & vbCrLf
  next  
  .write "<P><B>Form Object</B><BR>" & vbCrLf
  for each item in request.form
    .write item & "=" & request.form(item) & "<BR>" & vbCrLf
  next  
  .write "<BR>------------------------------------------------<P>" & vbCrLf
end if  

end with

' End Content

response.write "<BR><BR>" & vbCrLf

' --------------------------------------------------------------------------------------

%>  
<!--#include virtual="/include/core_countries.inc"-->
<%

' --------------------------------------------------------------------------------------
' Functions and Subroutines
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
  objRS.open "select @@identity", conn, 3,3
  my_ID = objRS(0)
  set objRS = Nothing

  Get_New_Record_ID = my_ID
  
end function

'--------------------------------------------------------------------------------------

sub Table_Begin()
    response.write "<TABLE BORDER=""" & Border_Toggle & """ CELLPADDING=""0"" CELLSPACING=""0"" CLASS=TableBorder VSPACE=""0"" HSPACE=""0"">" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif""            BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_TR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLr
    response.write "  <TR>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD VALIGN=""top"">" & vbCrLf
end sub      

'--------------------------------------------------------------------------------------

sub Table_End()
    response.write "    </TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>" & vbCrLf
    response.write "  <TR>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BL_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "    <TD BACKGROUND=""/images/SideNav_BR_corner.gif""><IMG SRC=""/images/Spacer.gif"" BORDER=""0"" WIDTH=""8"" HEIGHT=""6"" ALT="""" VSPACE=""0"" HSPACE=""0""></TD>" & vbCrLf
    response.write "  </TR>"
    response.write "</TABLE>" & vbCrLf
end sub

'--------------------------------------------------------------------------------------

select case Sequence
  case seqAdd, seqEdit
    %>

    <SCRIPT LANGUAGE=JAVASCRIPT>
    <!--
    // --------------------------------------------------------------------------------------
    
    function CheckRequiredFields() {
    
      var df = document.<%=Form_Name%>;
      var ErrorMsg;
      var strVal;
      var strChk;
      ErrorMsg = "";
      strVal = "";
    
      var TestRadio;
      var ctr;
      var RadioChecked;
      var valid = "0123456789-";
      var temp;
      var LastField = "";
          
      var badchars = /[\[\]:;|=,+*?<>"\\\/]/;
      var D_BadChars = '\\ / [ ] : ; | = , + * ? < > \' "';              //"'
    
      if (df.FirstName.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("First Name",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "FirstName";}
        df.FirstName.style.backgroundColor = "#FFB9B9";
      }
    
      if (df.LastName.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Last Name",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "LastName";}
        df.LastName.style.backgroundColor = "#FFB9B9";
      }
    
      if (df.Email.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Email",Alt_Language,conn))%>\r\n";  
        if (LastField.length == 0) {LastField = "Email";}
        df.Email.style.backgroundColor = "#FFB9B9";
      }
      
      if (!df.Email.value.match(/^[\w]{1}[\w\.\-_]*@[\w]{1}[\w\-_\.]*\.[\w]{2,6}$/i)) { 
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Invalid Email Address",Alt_Language,conn))%>\r\n";  
        if (LastField.length == 0) {LastField = "Email";}
        df.Email.style.backgroundColor = "#FFB9B9";
      } 
    
      if (df.Company.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Company Name",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Company";}
        df.Company.style.backgroundColor = "#FFB9B9";
      }
      
      if (df.Business_Phone.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Phone",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Business_Phone";}
        df.Business_Phone.style.backgroundColor = "#FFB9B9";
      }

      if (df.Shipping_Address.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Address",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Shipping_Address";}
        df.Shipping_Address.style.backgroundColor = "#FFB9B9";
      }

      if (df.Shipping_City.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("City",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Shipping_City";}
        df.Shipping_City.style.backgroundColor = "#FFB9B9";
      }
      
      if (df.Shipping_State.value == "") {
        ErrorMsg = ErrorMsg + "<%=Translate("USA State or Canadian Province",Alt_Language,conn) & " " & Translate("or",Alt_Language,conn) & " " & Translate("N/A",Alt_Language,conn)%>\r\n";
        if (LastField.length == 0) {LastField = "Shipping_State";}
        df.Shipping_State.style.backgroundColor = "#FFB9B9";
      }
    
      if (df.Shipping_Country.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Country",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Shipping_Country";}
        df.Shipping_Country.style.backgroundColor = "#FFB9B9";
      }

      if (df.Shipping_Postal_Code.value == "") {
        ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code",Alt_Language,conn))%>\r\n";
        if (LastField.length == 0) {LastField = "Shipping_Postal_Code";}
        df.Shipping_Postal_Code.style.backgroundColor = "#FFB9B9";
      }  
      
      if (df.Shipping_Country.value == "US" && df.Shipping_Postal_Code.value != "") {
        strVal = df.Shipping_Postal_Code.value;
        if (strVal.length != 5 && strVal.length != 10) {
          ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code - 5 digit or 5 digit + 4",Alt_Language,conn))%>\r\n";
          if (LastField.length == 0) {LastField = "Shipping_Country";}
          df.Shipping_Postal_Code.style.backgroundColor = "#FFB9B9";
        }
        else {
          for (var i=0; i < strVal.length; i++) {
            strChk = "" + strVal.substring(i, i+1);
            if (valid.indexOf(strChk) == "-1") {
              ErrorMsg = ErrorMsg + "<%=ReplaceRSQuote(Translate("Postal Code - Invalid Characters",Alt_Language,conn))%>\r\n";
              if (LastField.length == 0) {LastField = "Shipping_Postal_Code";}
              df.Shipping_Postal_Code.style.backgroundColor = "#FFB9B9";
            }
          }
        }
      }  
      
      if (ErrorMsg != "") {
        ErrorMsg = "<%=KillQuote(Translate("Please enter the missing information for following REQUIRED fields (or use N/A (Not Applicable))",Alt_Language,conn))%>:\r\n\n" + ErrorMsg;
        if (LastField.length > 0) {
          df[LastField].focus();
        }  
        alert (ErrorMsg);
        return false;
      }
      else {
      	return true;
      } 
    }
    
    // --------------------------------------------------------------------------------------
    
    var highlightcolor="lightyellow"
    var ns6=document.getElementById&&!document.all
    var previous=''
    var eventobj
    
    //Regular expression to highlight only form elements
    var intended=/INPUT|TEXTAREA/
    
    //Function to check whether element clicked is form element
    
    function checkel(which){
      if (which.style&&intended.test(which.tagName)){
        if (ns6&&eventobj.nodeType==3)
          eventobj=eventobj.parentNode.parentNode
        return true;
      }
      else
      return false;
    }
    
    //Function to highlight form element
    
    function highlight(e){
      eventobj=ns6? e.target : event.srcElement
      if (previous!=''){
      if (checkel(previous))
        previous.style.backgroundColor=''
        previous=eventobj
        if (checkel(eventobj))
        eventobj.style.backgroundColor=highlightcolor
      }
      else{
      if (checkel(eventobj))
        eventobj.style.backgroundColor=highlightcolor
        previous=eventobj
      }
    }
    // --------------------------------------------------------------------------------------
    //-->
    </SCRIPT>
    
  <%  
  case else
end select  

' --------------------------------------------------------------------------------------

Call Disconnect_SiteWide
%>