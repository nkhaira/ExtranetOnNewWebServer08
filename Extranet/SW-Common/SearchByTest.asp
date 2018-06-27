<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Site_ID = 3
Login_Language = "eng"
Login_Country  = "us"
Call Connect_SiteWide
keySearch = "clamp"
dim UserSubGroups(1)
UserSubGroups(0) = "uss"
UserSubGroups(1) = "usm"
UserSubGroups_Max = 1

sqlTemp   = "       convert(varchar(50),   COALESCE(dbo.Calendar_Category.Category, '')) + " &_
            "       convert(varchar(255),  COALESCE(dbo.Calendar.Sub_Category, '')) + " &_
            "       convert(varchar(255),  COALESCE(dbo.Calendar.Product, '')) + " &_
            "       convert(varchar(255),  COALESCE(dbo.Calendar.Title, '')) + " &_
            "       convert(varchar(2000), COALESCE(dbo.Calendar.Description, '')) + " &_
            "       convert(varchar(50),   COALESCE(dbo.Calendar.Item_Number, '')) "

sqlSearch = "SELECT dbo.Calendar.* " &_
            "FROM   dbo.Calendar LEFT OUTER JOIN " &_
            "       dbo.Calendar_Category ON dbo.Calendar.Code = dbo.Calendar_Category.Code AND " &_
            "       dbo.Calendar.Site_ID = dbo.Calendar_Category.Site_ID " &_
            "WHERE (dbo.Calendar.Site_ID = " & Site_ID & ") "

            if instr(1,keySearch," ") > 0 then
              strSearch     = Split(keySearch," ")
              cntSearch     = Ubound(strSearch)
            else
              Dim strSearch(0)
              strSearch(0)  = Trim(keySearch)
              cntSearch     = 0
            end if  

            for x = 0 to cntSearch
              select case x
                case 0
                  sqlSearch = sqlSearch & " AND (("
                case else
                  sqlSearch = sqlSearch & " OR ("
              end select
              sqlSearch = sqlSearch & sqlTemp & " LIKE '%" & strSearch(x) & "%') "
            next
            sqlSearch = SqlSearch & ") "
            
'----
            if (Access_Level <= 2 or Access_Level = 6) then
              sqlSearch = sqlSearch & " AND (Calendar.SubGroups LIKE '%all%'"        
              for i = 0 to UserSubGroups_Max
                sqlSearch = sqlSearch & " OR Calendar.SubGroups LIKE '%" & UserSubGroups(i) & "%'"            
              next
              sqlSearch = sqlSearch & ")"        
            end if   
            
            ' Determine if Active or Archive
            
            if (Access_Level <= 2 or Access_Level = 6) then
              if abs(Show_Detail) = 0 or abs(Show_Detail) = 1 then
                sqlSearch = sqlSearch & " AND Calendar.Status=1 AND ((Calendar.LDate<='" & Date & "' AND Calendar.XDAYS=0) OR (Calendar.LDate<='" & Date & "' AND Calendar.XDate>'" & Date & "'))"
              elseif abs(Show_Detail) = 2 then  ' Archive
                sqlSearch = sqlSearch & " AND (Calendar.Status=2 OR (Calendar.XDays=0 AND '" & Date & "'>Calendar.XDate))"
              end if
            else ' Show all for Admin
              if abs(Show_Detail) = 1 then
                sqlSearch = sqlSearch & " AND (Calendar.Status=0 OR Calendar.Status=1) AND (Calendar.XDAYS=0 OR Calendar.XDate>'" & Date & "')"            
              else            
                sqlSearch = sqlSearch & " AND (Calendar.Status=" & abs(Show_Detail) & " OR (Calendar.XDAYS>0 AND '" & Date & "'>Calendar.XDate))"
              end if
            end if
    
           ' Restricted Countries
        
            if (Access_Level <= 2 or Access_Level = 6) then
              sqlSearch = sqlSearch & " AND (Country = 'none'" &_
                    " OR (Country LIKE '%0%' AND Country NOT LIKE '%" & Login_Country & "%')" &_
                    " OR (Country NOT LIKE '%0%' AND Country LIKE '%" & Login_Country & "%'))"
            end if 
  
            ' Filter to English or Preferred Language for Users
  
            if (Access_Level = 0 or Access_Level = 6)then
              if LCase(Login_Language) <> "eng" then
                sqlSearch = sqlSearch & " AND (Calendar.Language='eng' OR Calendar.Language='" & Login_Language & "')"
              else  
                sqlSearch = sqlSearch & " AND Calendar.Language='eng'"
              end if
            end if
            
            ' Filter for CIN Special Groupings
            
            if CIN < 8000 or CIN > 8999 then
              ' Individual or Individual + Product Introduction or Campaign
              sqlSearch = sqlSearch & " AND (Calendar.Content_Group=0 or Calendar.Content_Group=1 or Calendar.Content_Group=3)"
            end if
            
            sqlSearch = sqlSearch & " ORDER BY dbo.Calendar_Category.Title, dbo.Calendar.Sub_Category, dbo.Calendar.Product"

Set rsCalendar = Server.CreateObject("ADODB.Recordset")
rsCalendar.Open sqlSearch, conn, 3, 3

do while not rsCalendar.EOF
  response.write rsSearch("ID") & "<BR>"
  rsSearch.MoveNext
loop

rsCalendar.close
set rsCalendar = nothing
set sqlSearch  = nothing
            
Call Disconnect_SiteWide
%>