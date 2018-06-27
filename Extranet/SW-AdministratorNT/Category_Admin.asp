<%@ Language="VBScript" CODEPAGE="65001" %>
<%
'
' Author: K. Whitlock
'
 
' Functions
%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

'for each item in request.querystring
'  response.write item & ": " & request.querystring(item) & "<BR>"
'next

' --------------------------------------------------------------------------------------  

' Update Record

SQL = "UPDATE Calendar_Category SET "

if not isblank(replacequote(request("Title"))) then
  SQL = SQL & "Calendar_Category.Title=" & "N'" & replacequote(request("Title")) & "'"
else  
  SQL = SQL & "Calendar_Category.Title=NULL"
end if

if not isblank(replacequote(request("Preset_Sub_Category"))) then
  SQL = SQL & ",Calendar_Category.Preset_Sub_Category=" & "N'" & replacequote(request("Preset_Sub_Category")) & "'"
else  
  SQL = SQL & ",Calendar_Category.Preset_Sub_Category=NULL"
end if

if not isblank(replacequote(request("Description"))) then  
  SQL = SQL & ",Calendar_Category.Description=" & "N'" & replacequote(request("Description")) & "'"
else  
  SQL = SQL & ",Calendar_Category.Description=NULL"
end if
  
if isnumeric(request("Sort")) then
  SQL = SQL & ",Calendar_Category.Sort=" & replacequote(request("Sort"))
else
  SQL = SQL & ",Calendar_Category.Sort=999"
end if
  
if isnumeric(request("SortBy")) then
  SQL = SQL & ",Calendar_Category.SortBy=" & request("SortBy")
else
  SQL = SQL & ",Calendar_Category.SortBy=999"
end if

if request("Separator") = "on" then           
  SQL = SQL & ",Calendar_Category.Separator=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Separator=" & CInt(False)
end if

if request("PID") = "on" then           
  SQL = SQL & ",Calendar_Category.PID=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.PID=" & CInt(False)
end if

if request("Site_View") = "on" then           
  SQL = SQL & ",Calendar_Category.Site_View=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Site_View=" & CInt(False)
end if

if request("Sub_Category") = "on" then           
  SQL = SQL & ",Calendar_Category.Sub_Category=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Sub_Category=" & CInt(False)
end if

if request("Product_Series") = "on" then           
  SQL = SQL & ",Calendar_Category.Product_Series=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Product_Series=" & CInt(False)
end if

if request("Submission_Approve") = "on" then           
  SQL = SQL & ",Calendar_Category.Submission_Approve=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Submission_Approve=" & CInt(False)
end if 

if request("Reassign_Owner") = "on" then           
  SQL = SQL & ",Calendar_Category.Reassign_Owner=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Reassign_Owner=" & CInt(False)
end if

if request("Country_Restrictions") = "on" then           
  SQL = SQL & ",Calendar_Category.Country_Restrictions=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Country_Restrictions=" & CInt(False)
end if    


if request("Instructions") = "on" then           
  SQL = SQL & ",Calendar_Category.Instructions=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Instructions=" & CInt(False)
end if    

if request("Forum") = "on" then

  SQL = SQL & ",Calendar_Category.Calendar_View=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Item_Number=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Item_Number_2=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Location=" & CInt(False)
  SQL = SQL & ",Calendar_Category.File_Name=" & CInt(False)
  SQL = SQL & ",Calendar_Category.File_Name_POD=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Include=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Link=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Link_PopUp_Disabled=" & CInt(False)
  SQL = SQL & ",Calendar_Category.ImageStore=" & CInt(False)
  SQL = SQL & ",Calendar_Category.InformationStore=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Shopping_Cart=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Mark_Confidential=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Date_Basic=" & CInt(True)
  SQL = SQL & ",Calendar_Category.Date_PRD=" & CInt(False)
  SQL = SQL & ",Calendar_Category.Forum=" & CInt(True)  

else

  if request("Content_Group") = "on" then           
    SQL = SQL & ",Calendar_Category.Content_Group=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Content_Group=" & CInt(False)
  end if

  if request("Shopping_Cart") = "on" then           
    SQL = SQL & ",Calendar_Category.Shopping_Cart=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Shopping_Cart=" & CInt(False)
  end if

  if request("Mark_Confidential") = "on" then
    SQL = SQL & ",Calendar_Category.Mark_Confidential=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Mark_Confidential=" & CInt(False)
  end if

  if request("Date_Basic") = "on" then
    SQL = SQL & ",Calendar_Category.Date_Basic=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Date_Basic=" & CInt(False)
  end if

  if request("Date_PRD") = "on" then           
    SQL = SQL & ",Calendar_Category.Date_PRD=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Date_PRD=" & CInt(False)
  end if

  if request("Calendar_View") = "on" then           
    SQL = SQL & ",Calendar_Category.Calendar_View=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Calendar_View=" & CInt(False)
  end if

  if request("Item_Number") = "on" then           
    SQL = SQL & ",Calendar_Category.Item_Number=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Item_Number=" & CInt(False)
  end if    

  if request("Item_Number_2") = "on" then           
    SQL = SQL & ",Calendar_Category.Item_Number_2=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Item_Number_2=" & CInt(False)
  end if    

  if request("Location") = "on" then           
    SQL = SQL & ",Calendar_Category.Location=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Location=" & CInt(False)
  end if

  if request("File_Name") = "on" then           
    SQL = SQL & ",Calendar_Category.File_Name=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.File_Name=" & CInt(False)
  end if    
  
  if request("File_Name_POD") = "on" then           
    SQL = SQL & ",Calendar_Category.File_Name_POD=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.File_Name_POD=" & CInt(False)
  end if    

  if request("Include") = "on" then           
    SQL = SQL & ",Calendar_Category.Include=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Include=" & CInt(False)
  end if

  SQL = SQL & ",Calendar_Category.Forum=" & CInt(False)  

  if request("Link") = "on" then
    SQL = SQL & ",Calendar_Category.Link=" & CInt(True)
    if request("Link_PopUp_Disabled") = "on" then           
      SQL = SQL & ",Calendar_Category.Link_PopUp_Disabled=" & CInt(True)
    else
      SQL = SQL & ",Calendar_Category.Link_PopUp_Disabled=" & CInt(False)
    end if    
  else
    SQL = SQL & ",Calendar_Category.Link=" & CInt(False)
    SQL = SQL & ",Calendar_Category.Link_PopUp_Disabled=" & CInt(False)
  end if    

  if request("ImageStore") = "on" then           
    SQL = SQL & ",Calendar_Category.ImageStore=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.ImageStore=" & CInt(False)
  end if

  if request("Preset_EEF") = "on" then           
    SQL = SQL & ",Calendar_Category.Preset_EEF=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Preset_EEF=" & CInt(False)
  end if

  if request("Preset_FDL") = "on" then           
    SQL = SQL & ",Calendar_Category.Preset_FDL=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.Preset_FDL=" & CInt(False)
  end if

  if request("InformationStore") = "on" then           
    SQL = SQL & ",Calendar_Category.InformationStore=" & CInt(True)
  else
    SQL = SQL & ",Calendar_Category.InformationStore=" & CInt(False)
  end if

end if

if request("Thumbnail") = "on" then           
  SQL = SQL & ",Calendar_Category.Thumbnail=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Thumbnail=" & CInt(False)
end if    

if request("Subscription") = "on" then           
  SQL = SQL & ",Calendar_Category.Subscription=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Subscription=" & CInt(False)
end if    

if request("Title_View") = "on" then           
  SQL = SQL & ",Calendar_Category.Title_View=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Title_View=" & CInt(False)
end if    

if request("Enabled") = "on" then           
  SQL = SQL & ",Calendar_Category.Enabled=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Enabled=" & CInt(False)
end if  

''Added on 28th Oct 2009
if request("Marketing_Auto") = "on" then           
  SQL = SQL & ",Calendar_Category.Marketing_Automation=" & CInt(True)
else
  SQL = SQL & ",Calendar_Category.Marketing_Automation=" & CInt(False)
end if    
  
SQL = SQL & " WHERE Calendar_Category.ID=" & CInt(request("ID"))
'response.write replace (SQL,",",",<BR>")
'response.end
conn.execute (SQL)

Call Disconnect_SiteWide
      
' Success, Go back and re-display record with updated data

BackURL = "default.asp?ID=edit_category&Site_ID=" & CInt(request("Site_ID")) & "&Category_ID=" & CInt(request("ID"))
  
response.redirect BackURL

%>
