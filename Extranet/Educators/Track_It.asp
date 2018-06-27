<%@ Language="VBScript" CODEPAGE="65001" %>
<%
' --------------------------------------------------------------------------------------
' Author: K. D. Whitlock
' Date:   2/28/2005
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%
Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Dim Site_ID
Site_ID = 24        ' Educators Portal

Dim debug_flag
debug_flag = false

Dim TotalCount,OneCount,MultiCount

Dim Post_Method
Post_Method = "Post"

If isblank(Request("Code")) then

  with response
  
    .write "<FORM NAME=""Code"" ACTION=""Track_IT.asp"" METHOD=""" & Post_Method & """> "
    .write "<TABLE BGCOLOR=""#FFFFFF"" CELLSPACING=""2"" CELLPADDING=""2"" BORDER=""1"" >"
    .write "<TR>"
    .write "    <TD Class=Small>Begin Date:</TD>"
    
    .write "    <TD Class=Small><Input type=text name=BDate Class=Small Value="""
    if not isblank(request("BDate")) then
      .write request("BDate")
    else
      .write "1/1/" & Year(Date)
    end if
    .write """></TD>"
    
    .write "</TR>"
    .write "<TR>"
    .write "    <TD Class=Small>End Date:</TD>"
    .write "    <TD Class=Small><Input type=text name=EDate Class=Small Value="""
    if not isblank(request("EDate")) then
      .write request("EDate")
    else
      .write "12/31/" & Year(Date)
    end if
    .write """></TD>"
    
    .write "</TR>"
  
    .write "<TR>"
    .write "    <TD Class=Small>Region:</TD>"
    .write "    <TD Class=Small>"
    .write "<SELECT NAME=""Region"" CLASS=Small>" & vbCrLf
    .write "<OPTION VALUE=""0"""
    if isblank(request("Region")) or request("Region") = "0" then .write " SELECTED"
    .write ">All</OPTION>" & vbCrLf
    .write "<OPTION VALUE=""1"""
    if request("Region") = "1" then .write " SELECTED"
    .write ">United States</OPTION>" & vbCrLf
    .write "<OPTION VALUE=""2"""
    if request("Region") = "2" then .write " SELECTED"
    .write ">Europe</OPTION>" & vbCrLf
    .write "<OPTION VALUE=""3"""
    if request("Region") = "3" then .write " SELECTED"
    .write ">Intercon</OPTION>" & vbCrLf
    .write "</SELECT>" & vbCrLf
    .write "</TD>"
    .write "</TR>"


    SQLCode = "SELECT X_Description, Code " &_
              "FROM dbo.SubGroups " &_
              "WHERE (Site_ID = " & Site_ID & ") AND (Enabled = - 1)" &_
              "ORDER BY Order_Num"
    Set rsCode = Server.CreateObject("ADODB.Recordset")
    rsCode.Open SQLCode, conn, 3, 3
    
    FirstOne = True
    do while not rsCode.EOF
    
      .write "<TR>"
      .write "<TD Class=Small>"
      if FirstOne = True then
        .write "Select Groups:"
        FirstOne = False
      else
        .write "&nbsp;"
      end if
      .write "</TD>"
      .write "<TD Class=Small><INPUT TYPE=CHECKBOX NAME=CODE VALUE=""" & rsCode("Code") & """>&nbsp;&nbsp;" & rsCode("X_Description") & "</TD>"
      .write "</TR>"
      
      rsCode.MoveNext
      
    loop
  
    .write "<TR>"
    .write "    <TD Class=Small>&nbsp;</TD>"
    .write "    <TD Class=Small><INPUT TYPE=SUBMIT NAME=SUBMIT VALUE=""Calculate""></TD>"
    .write "</TR>"
  
    .write "</FORM>"
    
    rsCode.Close
    set rsCode = Nothing
  
  end with

else

  if instr(1,request("Code"),", ") > 0 then
    CodeValues = Split(request("Code"),",")
    CodeMax    = Ubound(CodeValues)
  else
    ReDim CodeValues(0)
    CodeValues(0) = request("Code")
    CodeMax = 0
  end if
  
  ' Build Filter
  
  SQL =   "SELECT DISTINCT " &_
          "       dbo.Activity.Account_ID, dbo.Activity.Session_ID, dbo.UserData.FirstName, dbo.UserData.LastName, dbo.UserData.Company, dbo.UserData.Business_Country " &_
          "FROM   dbo.Activity LEFT OUTER JOIN " &_
          "       dbo.UserData ON dbo.Activity.Account_ID = dbo.UserData.ID "
        
  SQLW =  "WHERE (dbo.Activity.View_Time >='" & Request("BDate") & "') AND (dbo.Activity.View_Time <= '" & Request("EDate") & "') AND (dbo.Activity.Site_ID = 24) AND "

  for x = 0 to CodeMax
    if x = 0 then SQLW = SQLW & "("
    SQLW = SQLW & "dbo.UserData.SubGroups LIKE '%" & Trim(CodeValues(x)) & "%' "
    if x <> CodeMax then SQLW = SQLW & " OR "
  next

  SQL = SQL & SQLW & ")"
  
  if not isblank(request("Region")) and request("Region") <> "0" then
    SQL = SQL & " AND dbo.UserData.Region=" & request("Region") & " "
  end if
  
  SQL = SQL & " ORDER BY dbo.UserData.Business_Country, dbo.Activity.Account_ID"

  if debug_flag = true then
    response.write SQL & "<P>"
  end if
  
  SQLT =  "SELECT DISTINCT dbo.Activity.Account_ID " &_
          "FROM            dbo.Activity LEFT OUTER JOIN " &_
          "                dbo.UserData ON dbo.Activity.Account_ID = dbo.UserData.ID "
  SQLT = SQLT & SQLW & ")"
  
  if not isblank(request("Region")) and request("Region") <> "0" then
    SQLT = SQLT & " AND dbo.UserData.Region=" & request("Region") & " "
  end if

  SQLT = SQLT & " ORDER BY dbo.Activity.Account_ID"

  TotalCount      = 0
  OneCount        = 0
  MulTiCount      = 0
  Account_Last    = 0
  Account_Current = 0
  Account_Next    = 0

  Set rsCount = Server.CreateObject("ADODB.Recordset")
  rsCount.Open SQL, conn, 3, 3
   
  do while not rsCount.EOF
  
    TotalCount = TotalCount + 1
    
    Account_Current = rsCount("Account_ID")    
    
    if CLng(Account_Current) <> CLng(Account_Last) then
    
      rsCount.MoveNext

      if not rsCount.EOF then

        Account_Next = rsCount("Account_ID")
        
        if CLNG(Account_Current) = CLng(Account_Next) and CLng(Account_Current) <> CLng(Account_Last) then
          MultiCount = MultiCount + 1
          if debug_flag = true then
            response.write "M: " & Account_Current & " " & MultiCount
          end if
        elseif CLng(Account_Current) <> CLng(Account_Next) then
          OneCount = OneCount + 1
          if debug_flag = true then
            response.write "O: " & Account_Current & " " & OneCount
          end if
        else
          if debug_flag = true then
            response.write "X: " & Account_Current
          end if
        end if

        if debug_flag = true then
          response.write " L:" & Account_Last & " C:" & Account_Current & " N:" & Account_Next & "<BR>"
        end if  
      end if
      
      rsCount.MovePrevious

    else
    
      if debug_flag = true then
        response.write "S: " & Account_Current
        response.write " L:" & Account_Last & " C:" & Account_Current & " N:" & Account_Next & "<BR>"
      end if  
      
    end if
   
    rsCount.MoveNext
    
    Account_Last = Account_Current    
    
  loop

  rsCount.Close
  set rsCount = nothing
  
  Set rsCount = Server.CreateObject("ADODB.Recordset")
  rsCount.Open SQLT, conn, 3, 3
  
  if not rsCount.EOF then
    Account_Total = rsCount.RecordCount
  else
    Account_Total = 0
  end if
  
  rsCount.Close
  set rsCount = nothing
  
  with response
    .write "Educators Portal on Support.Fluke.com<P>"
    .write "Date Range " & request("BDate") & " - " & request("EDate") & "<BR>"
    .write "Region "
    select case request("Region")
      case 1
        .write "United States"
      case 2
        .write "Europe"
      case 3
        .write "Intercon"
      case else
        .write "All"
    end select
    .write "<P>"
    
    .write "Total Unique Accounts: " & Account_Total & "<BR>"
    .write "Total Logons for Groups Selected: " & TotalCount & "<BR>"
    .write "Total Logons = 1: " & OneCount & "<BR>"  
    .write "Total Logons > 1: " & MultiCount & "<P>"    
  
    .write "<FORM ACTION=""Track_IT.asp"" ACTION=""" & Post_Method & """>"
    .write "<INPUT TYPE=""HIDDEN"" NAME=Region VALUE=""" & request("Region") & """>"
    .write "<INPUT TYPE=""HIDDEN"" NAME=BDate VALUE=""" & request("BDate") & """>"    
    .write "<INPUT TYPE=""HIDDEN"" NAME=EDate VALUE=""" & request("EDate") & """>"    
    .write "<INPUT TYPE=SUBMIT VALUE=""New Tally"">"
    .write"</FORM>"
    
  end with
  
  
  
end if


Call Disconnect_SiteWide
%>