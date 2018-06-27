<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' The purpose of this script is to perform the initial bulk loading of assets into the SiteWide
' calendar table from data scrapped from the old FNET website.  This script should not be run after
' final loading, however it should be saved since it will perform bulk loading for future acquisitions.
'
' Author: Kelly Whitlock

if 1 = 2 then ' Locked to not run

  Session.timeout = 240 ' Set to 4 Hours
  Server.ScriptTimeout = 2 * 60
  
  Call Connect_SiteWide
  
  SQL = "DELETE FROM Calendar WHERE Site_ID = 82"
  
  conn.execute SQL
  
  SQL = "INSERT INTO dbo.Calendar " &_
        "  ( Site_ID, Code,Category_ID, Sub_Category, Content_Group, Content_Group_Name, Product, Title, Description, " &_
        "    Instructions, Splash_Header, Item_Number, Item_Number_Show, Item_Number_2, PID, Revision_Code, Cost_Center, " &_
        "    Location, LDays, LDate, BDate ,VDate, XDays, PDate, UDate, PEDate, Confidential, Link, Link_PopUp_Disabled, " &_
        "    Include, Include_Size, File_Name, File_Size, File_Page_Count, Archive_Name, Archive_Size, File_Name_POD, File_Size_POD, " &_
        "    Archive_Name_POD, Archive_Size_POD, Thumbnail, Thumbnail_Size, Thumbnail_Request, Secure_Stream, Image_Locator, " &_
        "    Forum_ID, Forum_Moderated, Forum_Moderator_ID, Groups, " &_
        "    Subscription, Subscription_Early, Headline_View, Language, Country, " &_
        "    Status, Status_Comment, Clone, Locked, Submitted_By, Approved_By, Review_By, Review_By_Group, Campaign) " &_
        "SELECT Site_ID, Code, Category_ID, Sub_Category, Content_Group, Content_Group_Name, Product, Title, Description, " &_
        "    Instructions, Splash_Header, Item_Number, Item_Number_Show, Item_Number_2, PID, Revision_Code, Cost_Center, " &_
        "    Location, LDays, LDate, BDate ,VDate, XDays, PDate, UDate, PEDate, Confidential, Link, Link_PopUp_Disabled, " &_
        "    Include, Include_Size, File_Name, File_Size, File_Page_Count, Archive_Name, Archive_Size, File_Name_POD, File_Size_POD, " &_
        "    Archive_Name_POD, Archive_Size_POD, Thumbnail, Thumbnail_Size, Thumbnail_Request, Secure_Stream, Image_Locator, " &_
        "    Forum_ID, Forum_Moderated, Forum_Moderator_ID, Groups, " &_
        "    Subscription, Subscription_Early, Headline_View ,Language, Country, " &_
        "    Status, Status_Comment, Clone, Locked, Submitted_By, Approved_By, Review_By, Review_By_Group, Campaign " &_
        "FROM Calendar_FNET"
  
    conn.execute SQL
  
  Call Disconnect_SiteWide
  
  response.redirect "xCalendar_Cost_Center.asp"
  
end if
%>
