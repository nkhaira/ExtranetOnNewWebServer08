<%@ Language="VBScript" Codepage=65001%>
<!--METADATA TYPE="TypeLib" UUID="{6B16F98B-015D-417C-9753-74C0404EBC37}" -->
<%

' --------------------------------------------------------------------------------------
'
' Author: Kelly Whitlock
'
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

Session("BackURL") = ""

Dim FileUpEE_Flag, FileUpEE_Remote_Flag
FileUpEE_Flag        = true
FileUpEE_Remote_Flag = true

Dim oFileUpEEProgressWS
Dim oFileUpEEProgressClient
Dim cProgressID
Dim wProgressID
	
' Create two FileUpEEProgress instances for each stage of the transfer we want to monitor.
Set oFileUpEEProgressWS = Server.CreateObject("SoftArtisans.FileUpEEProgress")
Set oFileUpEEProgressClient = Server.CreateObject("SoftArtisans.FileUpEEProgress")

' Tell each of the progress objects which stage to monitor
oFileUpEEProgressClient.Watch = saClient
oFileUpEEProgressWS.Watch     = saWebServer

' Get a new progress ID for the client > webserver layer
cProgressID = oFileUpEEProgressClient.NextProgressID

' Get a new progress ID for the webserver > fileserver layer
wProgressID = oFileUpEEProgressWS.NextProgressID

' The client and webserver progress IDs (cProgressID and wProgressID respectively)
' are submitted to the progress indicator window and to the webserver script
' as querystring parameters.  See the JavaScript startupload().
' Note:  The progress IDs MUST be submitted in the querystring, not as a form element.

if LCase(Request("Language"))     = "xon" then
  Session("ShowTranslation")      = True
elseif LCase(Request("Language")) = "xof" then
  Session("ShowTranslation")      = False
end if

Dim Site_ID

%>
<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<!--#include virtual="/connections/connection_formdata.asp" -->
<%

Call Connect_SiteWide

%>
<!--#include virtual="/sw-administrator/CK_Admin_Credentials.asp"-->
<%

Dim BackURL
Dim HomeURL
Dim Site_Code
Dim Calendar_ID
Dim Category_ID
Dim Category_ID_Change
Dim Category_Code
Dim Content_Group
Dim Admin_Access
Dim Admin_Name
Dim Code_X_Name
Dim Login_ID          ' Portal View

Login_ID = Admin_ID   ' Assign Admin_ID to Portal View Login_ID for [Site View] rendering of asset container

Dim Path_Site         ' eFulfillment Documents are stored /[site_code]/download/[sub]/*.*
Dim Path_Site_POD     ' Common Directory for Print on Demand Files -- All of SiteWide
Dim Path_Include
Dim Path_File
Dim Path_Site_Secure
Dim Path_File_POD
Dim Path_Thumbnail

Dim Show_View
Dim Show_Detail
Dim Show_Location
Dim Show_ImageStore
Dim Show_Product_Series
Dim Show_PID
Dim Show_Sub_Category
Dim Show_Link
Dim Show_Item_Number
Dim Show_Item_Number_2
Dim Show_Link_PopUp_Disabled
Dim Show_File
Dim Show_File_POD
Dim Show_Include
Dim Show_Thumbnail
Dim Show_Subscription
Dim Show_Calendar
Dim Show_Forum
Dim Show_Content_Group
Dim Show_Date_Basic
Dim Show_Date_PRD
Dim Show_Mark_Confidential
Dim Show_Shopping_Cart
Dim Show_Country_Restrictions
Dim Show_Special_Instructions
Dim Show_Site_View
Dim Show_Reassign_Owner
Dim Show_Submission_Approve
Dim Show_Preserve_Clone

Dim PID_Enabled
Dim PID_System
Dim Preset_EEF
Dim Preset_FDL
Dim Write_Form_Show_Values
Dim Field_Editable
Dim FormName

if not isblank(request("Show_View")) and isnumeric(request("Show_View")) then
  Show_View = request("Show_View")
else
  Show_View = CInt(False)
end if    

Show_Location            = False
Show_ImageStore          = False
Show_Link                = False
Show_Link_PopUp_Disabled = False
Show_Item_Number         = False
Show_Item_Number_2       = False
Show_PID                 = False
Show_File                = False
Show_File_POD            = False
Show_Include             = False
Show_Thumbnail           = False
Show_Subscription        = False
Show_Calendar            = False
Show_Forum               = False
Show_Content_Group       = False
Show_Date_Basic          = False
Show_Date_PRD            = False
Show_Mark_Confidential   = False
Show_Shopping_Cart       = False
Show_Country_Restrictions= False
Show_Special_Instructions= False
Show_Sub_Cateogry        = False
Show_Product_Series      = False
Show_Preserve_Clone      = False

Write_Form_Show_Values   = False

Dim Icon_Type
Dim Icon_Extension

Dim Region
Region         = 0
Dim RegionValue
RegionValue    = ""
Dim RegionColorPointer
RegionColorPointer = 0
Dim RegionColor(4)
RegionColor(0) = "#0000FF"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"
RegionColor(4) = "#FFCC99"

BackURL       = "/sw-administrator/Calendar_Edit.asp" 
HomeURL       = "/sw-administrator/Default.asp"

Calendar_ID   = request("ID")
Category_ID   = request("Category_ID")
Content_Group = request("Content_Group")

Path_Include     = "Download\Content"
Path_File        = "Download"
Path_Site_Secure = "0"
Site_Path_POD    = ""
Path_File_POD    = "POD"
Path_Thumbnail   = "Download\Thumbnail"

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<!--#include virtual="/SW-Common/SW-Field_Names.asp"-->
<!--#include virtual="/SW-Common/SW-Content_Subroutines.asp"-->
<!--#include virtual="/SW-Administrator/Calendar_Show_Values.asp"-->
<%

Screen_Title   = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Content / Event Administration Screen",Alt_Language,conn)
Bar_Title      = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=MediumBoldGold>" & Translate("Content / Event Administration Screen",Login_Language,conn) & "</FONT>"
Navigation     = false
Top_Navigation = false
Content_Width  = 95  ' Percent

if IsNumeric(Calendar_ID) and Admin_Access < 2 then
  response.redirect HomeURL & "?Site_ID=" & Site_ID
end if

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Navigation.asp"-->

<IFRAME STYLE="display:none;position:absolute;width:148;height:194;z-index=100" ID="CalFrame" MARGINHEIGHT=0 MARGINWIDTH=0 NORESIZE FRAMEBORDER=0 SCROLLING=NO SRC="/SW-Common/SW-Calendar_PopUp.asp"></IFRAME>
<%

SQL = "SELECT Title FROM Calendar_Category WHERE Site_ID=" & Site_ID & " AND (Code=8000 OR Code=8001) ORDER BY Code"
Set rsCampaign = Server.CreateObject("ADODB.Recordset")
rsCampaign.Open SQL, conn, 3, 3

Dim Code_8000_Name, Code_8001_Name
Code_8000_Name = rsCampaign("Title")
rsCampaign.MoveNext
Code_8001_Name = rsCampaign("Title")
rsCampaign.close
set rsCampaign = nothing

' --------------------------------------------------------------------------------------
' Edit
' --------------------------------------------------------------------------------------

if IsNumeric(Calendar_ID) then
  %>
  <!--#include virtual="/SW-Administrator/Calendar_Edit_Update.asp"-->
  <%
end if

' --------------------------------------------------------------------------------------
' Add Record
' --------------------------------------------------------------------------------------

if Calendar_ID= "add" then
  %>
  <!--#include virtual="/SW-Administrator/Calendar_Edit_Add.asp"-->
  <%
end if

%>
<A NAME="HELP"></A>
<DIV CLASS=Medium>
<DIV ALIGN=CENTER>
<%="<B>" & Translate("Note",Login_Language,conn) & "</B>: " & Translate("Do not copy and paste text directly from a MS Office product or any application that embeds hidden characters to format the text, otherwise the hidden characters pasted into the text field and will be displayed as unintelligible gibberish.",Login_Language,conn) & " " & Translate("To circumvent this from happening, paste into Notepad, then copy and paste into the text field.",Login_Language,conn)%>
<%="<BR>(<A HREF=""#Special""><B>" & Translate("See Special Characters and Font Attribute Tags",Login_Language,conn) & "</B></A>)."%>
<P>
<B><% if lcase(request("ID")) = "add" then response.write Translate("Add",Login_Language,conn) & " " else response.write Translate("Edit",Login_Language,conn) & " " %><%=Translate("Content or Event",Login_Language,conn)%></B><BR>
<%=Translate("These are date/group specific asset or events related to general or specific Product or Product Families.",Login_Language,conn)%>
</DIV>
<BR>

  <UL>
  
  <% if isnumeric(request("ID")) then %>
  <LI><B><%=Translate("Show / Hide Site View Button",Login_Language,conn)%></B> - <%=Translate("Toggles showing a replication of how this asset or event item would appear on the site to the user.",Login_Language,conn)%><P>
  <% end if %>
  
  <LI><B><%=Translate("Content or Event ID Number",Login_Language,conn)%></B> - <%=Translate("Internal database reference number or ADD for new record.  If the record is new, after you click on the [Save / Update] button, the record will be re-displayed with a new Content / Event ID number.",Login_Language,conn)%>
  
  <% if isnumeric(request("ID")) then %>
  <LI><B><%=Translate("Locked",Login_Language,conn)%></B> - <%=Translate("A lock is applied to a Content or Event ID Number, indicated by a red [ID Number], by the Site Administrator to prevent Edit, Clone, Duplication or Delete.  If you need to have this record modified, you will need to contact your site administrator.",Login_Language,conn)%></LI>
  <% end if %>
  
  <% if isnumeric(request("ID")) then %>
  <LI><B><%=Translate("Status",Login_Language,conn)%></B> - <%=Translate("Review allows the Site Administrator to view asset or event as if it were Live, however, the user is unable to see this asset or event Item until the status is changed to Live. Asset or event Items are Archived if the Site Administrator selects this Status option or the asset or event Item expires based on date.",Login_Language,conn)%></LI>
  <% end if %>
  
  <P>
  <LI><B><%=Translate("Content Grouping",Login_Language,conn)%></B> - <%=Translate("Each asset item or event can be grouped &quot;Individually&quot; (default), or associated with a &quot;Multiple Asset Container&quot; (MAC) grouping of related assets.",Login_Language,conn) & " " & Translate("If a &quot;Multiple Asset Container&quot; is selected, another drop-down selection box will appear allowing you to select which &quot;Multiple Asset Container&quot; to associate this asset to.",Login_Language,conn) & " " & Translate("Asset items or events associated with a &quot;Multiple Asset Container&quot; may also appear as individual items under their respective categories.",Login_Language,conn)%></LI>
  <P>
  
  <LI><B><%=Translate("Category",Login_Language,conn)%></B> - <%=Translate("This is the location that the asset will appear under the Calendar or Library navigation menu.",Login_Language,conn)%></LI>
  <P>
  <LI><B><%=Translate("Product or Product Family",Login_Language,conn)%></B> - <%=Translate("This is a critical sorting/grouping field.  Please try to add new asset or event records using one of the pre-existing selections, or if you require a new Product or Product Family name, you can specify a new name by using the input box.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Title",Login_Language,conn)%></B> - <%=Translate("Topic or title of the asset or event. (Included with Subscription Service)",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Description",Login_Language,conn)%></B> - <%=Translate("Short narrative description of the asset or event. (Included with Subscription Service)",Login_Language,conn)%></LI>
  
  <% if Show_Special_Instructions = True then %>
  <LI><B><%=Translate("Special Instructions",Login_Language,conn)%></B> - <%=Translate("Short instructions of how to use, order, or other instructions related to the asset or event.",Login_Language,conn)%></LI>
  <% end if %>

  <% if Show_Item_Number = True then %>
  <P>
  <LI><B><%=Translate("Item / Reference Number",Login_Language,conn)%> 1</B> - <%=Translate("Oracle Item Number, Literature Number or other generic designator.",Login_Language,conn)%>&nbsp;&nbsp;<%=Translate("The show checkbox, if checked will display the Item Number in the asset's description.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Revision",Login_Language,conn)%> 1</B> - <%=Translate("Oracle/MfgPro Item Number, Literature Number revision designator.",Login_Language,conn)%>&nbsp;&nbsp;<%=Translate("The show checkbox, if checked will display the Item Number in the asset's description.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Show",Login_Language,conn)%></B> - <%=Translate("Displays Oracle/MfgPro Item Number, Literature Number and revision designator with description information.",Login_Language,conn)%>&nbsp;&nbsp;<%=Translate("The show checkbox, if checked will display the Item Number in the asset's description.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Item / Reference Number",Login_Language,conn)%> 2</B> - <%=Translate("Legacy or other reference designator.",Login_Language,conn)%>&nbsp;&nbsp;<%=Translate("The show checkbox, if checked will display the Item Number in the asset's description.",Login_Language,conn)%></LI>
  <% end if %>
  <P>
  <% if Show_Location = True then %>
    <LI><B><%=Translate("Location",Login_Language,conn)%></B> - <%=Translate("Building, City, State, Country information. (Included with Subscription Service)",Login_Language,conn)%></LI>
  <% end if %>
  <% if Show_Link = True then %>
    <P><LI><B><%=Translate("URL to Web Page",Login_Language,conn)%></B> - <%=Translate("If the asset or event has additional information located on another web page, supply the complete URL.",Login_Language,conn)%>&nbsp;&nbsp;
    <%=Translate("Note",Login_Language,conn)%>: <%=Translate("If the container is a multiple asset container (MMC) such as a Product Introduction or Campaign, adding a URL Link to a web page at the MMC level, assumes that all the content is off-site, therefore the container name will not appear in the MAC name dropdown.",Login_Language,conn)%></LI>
  <% end if %>
  <% if Show_Link_PopUp_Disabled = True then %>
    <LI><B><%=Translate("URL to Web Page Pop-Up Window Disable",Login_Language,conn)%></B> - <%=Translate("If disabled, then URL Link to Web Page is direct as opposed to using a separate pop-up browser window.  A Session variable, Session(&quot;BackURL&quot;) can be interrogated by the link to obtain the parent link to restore this view when the link application is done.",Login_Language,conn)%></LI>
  <% end if %>
  
  <% if Show_Location = True or Show_Link = True or Show_Link_PopUp_Disabled = True then %>
    <P>
  <% end if %>    
  
  <% if Show_File = True then %>
    <LI><B><%=Translate("Asset File - (LOW Resolution)",Login_Language,conn)%></B> - <%=Translate("Low Resolution Asset File used for web view / download / email, eFulfillment and the Subscription Service. An Oracle / MfgPro Item Number is required in field Item Reference #1 if this asset is used for eFulfillment, POD or orderable through the Shopping Cart. Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Secure Stream",Login_Language,conn)%></B> - <%=Translate("Secure Stream hides the http:// url (path) to the document from the user.  This option is used to secure document directory paths so that they cannot be viewed or linked to directly.  Use this option only for confidential or those documents that require a credential to access the document, since there is some additional pop-up overhead involved to stream the file to the user as opposed to a simple redirect to the digital file itself.",Login_Language,conn)%></LI><BR>  
  <%  if Show_File_POD = True then %>
        <LI><B><%=Translate("Asset File - (POD Resolution)",Login_Language,conn)%></B> - <%=Translate("Medium Resolution Print on Demand Asset File.  This field supplies a link to this asset used by the Everett POD process.  An Oracle / MfgPro Item Number is required in field Item Reference #1.  Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%> <B><%=Translate("Everett - Marketing Communications Use Only.",Login_Language,conn)%></B></LI>
  <%  end if %>
     <P>  
  <% end if %>
  
  
  <% if Show_Item_Number = True then %>
    <LI><B><%=Translate("Available to Electronic Email Fulfillment",Login_Language,conn) & " - (" & Translate("End-User Oracle",Login_Language,conn)%>)</B> - <%=Translate("Enables Item Number to be available for the US/Intercon Electronic Email Fulfillment (EEF) Sytem and Print-on-Demand (POD) System.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Available to Electronic Fulfillment",Login_Language,conn) & " - (" & Translate("End-User Digital Library",Login_Language,conn)%>)</B> - <%=Translate("Enables Item Number to be available for the Digital Library on ",Login_Language,conn) & Request.ServerVariables("SERVER_NAME") & "."%></LI>
    <LI><B><%=Translate("Available to Literature Order Shopping Cart",Login_Language,conn) & " - "%></B><%=Translate("Enables Item Number to be added by the user&acute;s shopping cart for US/Intercon Print on Demand (POD) system or for physical media fulfillment (DCG).",Login_Language,conn)%></LI>
    <P>
  <% end if %>  
  
  
  <% if Show_Include = True then %>
    <LI><B><%=Translate("Content File",Login_Language,conn)%></B> - <%=Translate("Copies the contents of an external HTML file that is included as additional content information for this record  Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%></LI>
    <P>
  <% end if %>
  
  <% if Show_Thumbnail = True then %>
    <LI><B><%=Translate("Thumbnail Image File",Login_Language,conn)%></B> - <%=Translate("Adds an image to this record  Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn) & " " & Translate("Thumbnails should be no wider than 80px, height proportional to width, i.e., maintain aspect ration when resizing.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Request Thumbnail",Login_Language,conn)%></B> - <%=Translate("If you do not have the ability to create your own thumbnail image file for this asset, check this checkbox to have a thumbnail image created.",Login_Language,conn)%></LI>
  <% end if %>
  <P>
  <% if Content_Group > 0 and not Show_Calendar and Show_Content_Group then %>
    <LI><B><%=Translate("Override MAC Date",Login_Language,conn)%></B> - <%=Translate("When enabled for this asset the override allows independent Pre-Announce days, Beginning Date, Ending Date and Move to Archive values even though it is attached to a MAC.",Login_Language,conn)%>&nbsp;&nbsp;<%=Translate("Typical use of the ""Override MAC Date"" is the addition of a new asset after the MAC ""GO LIVE"" date.",Login_Language,conn)%></LI>
  <% end if %>  
  
  <% if Show_Date_Basic = False then %>
    <LI><B><%=Translate("Pre-Announce",Login_Language,conn)%></B> - <%=Translate("Number of days prior to the Beginning Date to display this asset or event. Default=0 - No effect on Beginning Date.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Beginning Date",Login_Language,conn)%></B> - <%=Translate("Beginning Date of Event or ""Go Live"" date of the asset.",Login_Language,conn)%>&nbsp;(<%=Translate("Included with Subscription Service",Login_Language,conn)%>)&nbsp;<%=Translate("If you need a calendar, click on the calendar icon.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
    <LI><B><%=Translate("Ending Date",Login_Language,conn)%></B> - <%=Translate("Ending Date of Event or ""Archive"" date of the asset.",Login_Language,conn)%>&nbsp;(<%=Translate("Included with Subscription Service",Login_Language,conn)%>)&nbsp; <%=Translate("If you need a calendar, click on the calendar icon.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
    <LI><B><%=Translate("Move to Archive",Login_Language,conn)%></B> - <%=Translate("Number of days after the Ending Date to display the asset or event.  Default=0, No effect on Ending Date.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Public Release Date",Login_Language,conn)%></B> - <%=Translate("The date that this information can be released to the public.",Login_Language,conn)%>&nbsp;<%=Translate("A Public Release Date Notice will appear in the Description section of the asset item or event.",Login_Language,conn)%>&nbsp;<%=Translate("Leave this date blank if there is not a Public Release Date restriction.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
    <LI><B><%=Translate("Mark as Confidential",Login_Language,conn)%></B> - <%=Translate("The caption &quot;Confidential - Not for Public Release&quot; will appear in the Description section of the asset item or event.",Login_Language,conn)%></LI>
  <% else %>
    <LI><B><%=Translate("Beginning Date",Login_Language,conn)%></B> - <%=Translate("Beginning Date of Event or ""Go Live"" date of the asset.",Login_Language,conn)%>&nbsp;(<%=Translate("Included with Subscription Service",Login_Language,conn)%>)&nbsp;<%=Translate("If you need a calendar, click on the calendar icon.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
  <% end if %>
  <P>
  <% if Show_Subscription = True then %>
    <LI><B><%=Translate("Send Notice via Subscription Service",Login_Language,conn)%></B> - <%=Translate("Sends a customized email memo to the user containing Title, Product/Series, Date, Description and link of the record to the Channel Group enabled, whose User Profiles have Subscription Service enabled.  The information is sent on the Beginning Date or Pre-Announce Date if specified.",Login_Language,conn)%></LI>
    <P>
  <% end if %>
  <% if Show_Calendar = True then %>
    <LI><B><%=Translate("Calendar",Login_Language,conn)%></B> - <%=Translate("Shows this asset or event on the Site calendar.",Login_Language,conn)%></LI>
    <P>  
  <% end if %>  
  <LI><B><%=Translate("Groups",Login_Language,conn)%></B> - <%=Translate("Select each Channel Category or Entitlement that is allowed to view this asset or event. For a new asset or event additions, pre-selected (default) group(s) checkboxes are displayed in red.",Login_Language,conn)%></LI>
  <% if Admin_Access >= 8 then %>
    <LI><B><%=Translate("Groups - Administrator Accounts",Login_Language,conn)%></B> - <%=Translate("Select each Administrator account that is allowed to view this asset or event. These selections are only available to the Site Administrator and restrict the asset or event item to be viewed only if the user has the selected authorization.  Ensure that standard groups above are un-checked.",Login_Language,conn)%></LI>
  <% end if %>
  <P>  
  <LI><B><%=Translate("Restrict/Limit to Countries",Login_Language,conn)%></B> - <%=Translate("Select Restrict to or Limit to each Country allowed to view this asset or event or leave blank if no countries are restricted.  This is a multi-select area.  To select more than one restricted country, hold down the [CTRL] key while selecting with your mouse.",Login_Language,conn)%></LI>
  <% if Admin_Access <= 2 then %>
    <P>
    <LI><B><%=Translate("Select Group to Approve this Submission",Login_Language,conn)%></B> - <%=Translate("All new submissions to this site, require the review and approval of the group responsible for maintenance of this information.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Request Review of this Submission by Email",Login_Language,conn)%></B> - <%=Translate("All submissions will automatically appear in the approval queue of the Content / Event Administrator, however, you may want to inform the Administrator by Email of your submission for date sensitive submittals or other reasons that a review is pending.",Login_Language,conn)%></LI>
  <% end if %>
  <% if Admin_Access = 4 or Admin_Access >=8 then %>
    <P>
    <LI><B><%=Translate("Group Assigned to Approve this Submission",Login_Language,conn)%></B> - <%=Translate("As a Content / Event Administrator, you can select yourself as reviewer of this submission (default), or you can re-assign the submission to another group for review, approval and maintenance of this information.  Note: For this submission to appear in the submitter&acute;s or administrator&acute;s queue, the Status flag must be set to &quot;Review&quot;.",Login_Language,conn)%></LI>
    <LI><B><%=Translate("Request Review of this Submission by Email",Login_Language,conn)%></B> - <%=Translate("If you have selected another Content / Event Administrator, all submissions will automatically appear in the approval queue of that Administrator, however, you may want to inform the Administrator by Email of your submission for date sensitive submittals or other reasons that a review is pending.",Login_Language,conn)%></LI>
  <% end if %>
  <% if isnumeric(request("ID")) and Clone = 0 then %>
    <P>
    <LI><B><%=Translate("Clone",Login_Language,conn)%></B> - <%=Translate("Clones an existing record to new record, however preserves Parent ID Number.  You can only clone from the original English version parent record that is not a clone in itself.  You cannot clone from a subsequent cloned record.  Cloning is primarily used for relating multi-language document versions of the same information to the original master or primary English version.",Login_Language,conn)%><P> 
    <%=Translate("Certain fields such as Item Number 1, Revision, Low Resolution and POD resolution asset files are not preserved since the intent of the clone is to reload these parameters in relation to the new language version.  The original Item Number, Revision and Language of the parent is stored in Item Number 2 for reference.",Login_Language,conn)%></LI><P>
    <LI><B><%=Translate("Duplicate",Login_Language,conn)%></B> - <%=Translate("Duplicates an existing record to new record.  You can only clone from the original English version parent record that is not a clone in itself.  Duplicating is primarily used for copying similar versions of the same information such as duplicating a pre-configured record to the next asset, to preserve Product or Series, Launch Dated, Groups, Country restrictions, etc.",Login_Language,conn)%></LI>
  <% end if %>  
  
  <P>
  </UL>
  
<UL>
<LI><B><%=Translate("Special Characters",Login_Language,conn)%></B> - <%=Translate("Certain characters have special meaning in HTML documents. The following entity names are used in HTML, always prefixed by ampersand (&) and followed by a semicolon. They represent particular graphic characters, which have special meanings in places in the markup, or may not be part of the character set available to the writer.",Login_Language,conn)%></LI>
<P>
<A NAME="Special"></A>
<INDENT>
<TABLE BGCOLOR="#EEEEEE" Border=0>
  <TR>
    <TD CLASS=Medium>
      <TABLE BORDER=0 COLPADDING=4>
        <TR><TD CLASS=Medium><B><%=Translate("Glyph",Login_Language,conn)%></B></TD><TD CLASS=Medium><B><%=Translate("Name",Login_Language,conn)%></B></TD><TD CLASS=Medium><B><%=Translate("Syntax",Login_Language,conn)%></B></TD><TD CLASS=Medium><B><%=Translate("Description",Login_Language,conn)%></B></TD></TR>
        <TR><TD CLASS=Medium>&lt;</TD><TD CLASS=Medium>lt</TD><TD CLASS=Medium>&amp;lt;</TD><TD CLASS=Medium><%=Translate("Less Than",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&gt;</TD><TD CLASS=Medium>gt</TD><TD CLASS=Medium>&amp;gt;</TD><TD CLASS=Medium><%=Translate("Greater Than",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&amp;</TD><TD CLASS=Medium>amp</TD><TD CLASS=Medium>&amp;amp;</TD><TD CLASS=Medium><%=Translate("Ampersand",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&quot;</TD><TD CLASS=Medium>quot</TD><TD CLASS=Medium>&amp;quot;</TD><TD CLASS=Medium><%=Translate("Double Quote",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&quot;</TD><TD CLASS=Medium>rdquo</TD><TD CLASS=Medium>&amp;rdquo;</TD><TD CLASS=Medium><%=Translate("Right Double Quote",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&ldquo;</TD><TD CLASS=Medium>ldquo</TD><TD CLASS=Medium>&amp;ldquo;</TD><TD CLASS=Medium><%=Translate("Left Double Quote",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&acute;</TD><TD CLASS=Medium>rsquo</TD><TD CLASS=Medium>&amp;rsquo;</TD><TD CLASS=Medium><%=Translate("Right Single Quote",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&lsquo;</TD><TD CLASS=Medium>lsquo</TD><TD CLASS=Medium>&amp;lsquo;</TD><TD CLASS=Medium><%=Translate("Left Single Quote",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&reg;</TD><TD CLASS=Medium>reg</TD><TD CLASS=Medium>&amp;reg;</TD><TD CLASS=Medium><%=Translate("Registered",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&copy;</TD><TD CLASS=Medium>copy</TD><TD CLASS=Medium>&amp;copy;</TD><TD CLASS=Medium><%=Translate("Copyright",Login_Language,conn)%></TD></TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
</INDENT>
</UL>
<UL>
<LI><B><%=Translate("Font Attribute Tags",Login_Language,conn)%></B> - <%=Translate("Certain combinations of characters have special meaning in HTML documents to format the appearance of the text prefixed and suffixed by these font formatting attributes. The following font attribute names are used in HTML, always prefixed by a &gt; and suffixed by a &lt; sign.  Do not use the above conversions for these special characters.",Login_Language,conn)%></LI>
<P>
<INDENT>
<TABLE BGCOLOR="#EEEEEE" Border=0>
  <TR>
    <TD CLASS=Medium>
      <TABLE BORDER=0 COLPADDING=4>
        <TR><TD CLASS=Medium><B><%=Translate("Syntax",Login_Language,conn)%></B></TD><TD CLASS=Medium><B><%=Translate("Description",Login_Language,conn)%></B></TD></TR>
        <TR><TD CLASS=Medium>&lt;B&gt;</TD><TD CLASS=Medium><%=Translate("Bold Enabled",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&lt;/B&gt;</TD><TD CLASS=Medium><%=Translate("Bold Disabled",Login_Language,conn)%></TD></TR>  
        <TR><TD CLASS=Medium>&lt;I&gt;</TD><TD CLASS=Medium><%=Translate("Italics Enabled",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&lt;/I&gt;</TD><TD CLASS=Medium><%=Translate("Italics Disabled",Login_Language,conn)%></TD></TR>  
        <TR><TD CLASS=Medium>&lt;U&gt;</TD><TD CLASS=Medium><%=Translate("Underline Enabled",Login_Language,conn)%></TD></TR>
        <TR><TD CLASS=Medium>&lt;/U&gt;</TD><TD CLASS=Medium><%=Translate("Underline Disabled",Login_Language,conn)%></TD></TR>  
      </TABLE>
    </TD>
  </TR>
</TABLE>

</INDENT>
</UL>
</DIV>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
' --------------------------------------------------------------------------------------
' Functions and Subroutines
' --------------------------------------------------------------------------------------
%>
<!--#include virtual="/include/Pop-Up.asp"-->
<!--#include virtual="/include/core_countries_multi.inc"-->

<SCRIPT LANGUAGE="JAVASCRIPT" SRC="/SW-Common/SW-Calendar_PopUp.js"></SCRIPT>

<SCRIPT LANGUAGE="JAVASCRIPT">
<!--//

var CheckMsg = "";
var ErrorMsg = "";
var FormName = document.<%=FormName%>
var CheckFlg = false;
var CheckSts = false;
var ctr
var Menu_Button = false;
var DeleteFlag  = false;

function Grouping_Name_Check() {
  if (1==2) {
    for (var i = 0; i < FormName.Content_Group.length; i++) {
      if (FormName.Content_Group[i].selected) {
        if (FormName.Content_Group.value == '0') {
        }    
        else {
          for (var i = 0; i < FormName.Campaign.length; i++) {
            if (FormName.Campaign[i].selected) {
              if (FormName.Campaign.value == '0') {
                FormName.Campaign.style.backgroundColor = "#FFB9B9";
                FormName.Campaign.focus();                          
                alert("<%=Translate("You must select a ",Alt_Language,conn) & " " & Code_X_Name & " " & Translate(" Name before you complete the rest of the fields of this form.",Alt_Language,conn)%>");
              }    
            }
          }
        }    
      }
    }  
  }
}

function SubGroups_Check() {
  if (! CheckFlg) {
    CheckMsg =  "<%=Translate("Alert",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Please ensure that you have authority to enable this content or event item to be viewed by Group within this Region that you have just selected.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Checkboxes with red borders containing a checkmark are default presets enabled by the Site Administrator.  These presets can be individually unchecked if the content or event item is not appropriate for this group.",Alt_Language,conn)%>\r\n\n";    
    CheckMsg += "<%=Translate("Click [OK] to Continue.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("This Alert Message will only appear once.",Alt_Language,conn)%>\r\n\n";
    alert(CheckMsg);
    CheckFlg = true;
    CheckMsg = "";
  }
  return(false);
}

function SubGroups_1_Check() {
  if (FormName.SubGroups_1.checked) {
    CheckMsg = "<%=Translate("Attention",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("You have selected All Groups for the US Region.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Please ensure that you have authority to enable this content or event item to be viewed by All Groups within this Region.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Checkboxes with red borders containing a checkmark are default presets enabled by the Site Administrator.  These presets can be individually unchecked if the content or event item is not appropriate for this group.",Alt_Language,conn)%>\r\n\n";    
    CheckMsg += "<%=Translate("Click [OK] to select All Groups within this Region, or click [Cancel] to uncheck All Groups within this region.",Alt_Language,conn)%>\r\n";

    if (<%=Admin_Region%> != 1) {
      CheckSts = confirm(CheckMsg);
    }
    else {
      CheckSts = true;
    }  

    if (CheckSts == true) {
      for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
        CheckMsg = FormName.SubGroups[ctr].value.substring(0,1);
        if (! FormName.SubGroups[ctr].checked && CheckMsg == 'u') {
          FormName.SubGroups[ctr].checked = 1;
        }
      }
      FormName.SubGroups_1.checked = 0;
    }
    else {
      for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
        CheckMsg = FormName.SubGroups[ctr].value.substring(0,1);
        if (FormName.SubGroups[ctr].checked && CheckMsg == 'u') {
          FormName.SubGroups[ctr].checked = 0;
        }
      }
      FormName.SubGroups_1.checked = 0;
    }
    CheckMsg = "";
  }
  return(false);
}

function SubGroups_2_Check() {
  if (FormName.SubGroups_2.checked) {
    CheckMsg = "<%=Translate("Attention",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("You have selected All Groups for the European Region.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Please ensure that you have authority to enable this content or event item to be viewed by All Groups within this Region.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Checkboxes with red borders containing a checkmark are default presets enabled by the Site Administrator.  These presets can be individually unchecked if the content or event item is not appropriate for this group.",Alt_Language,conn)%>\r\n\n";    
    CheckMsg += "<%=Translate("Click [OK] to select All Groups within this Region, or click [Cancel] to uncheck All Groups within this region.",Alt_Language,conn)%>\r\n";

    if (<%=Admin_Region%> != 2) {
      CheckSts = confirm(CheckMsg);
    }
    else {
      CheckSts = true;
    }  

    if (CheckSts == true) {
      for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
        CheckMsg = FormName.SubGroups[ctr].value.substring(0,1);
        if (! FormName.SubGroups[ctr].checked && CheckMsg == 'e') {
          FormName.SubGroups[ctr].checked = 1;
        }
      }
      FormName.SubGroups_2.checked = 0;
    }
    else {
      for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
        CheckMsg = FormName.SubGroups[ctr].value.substring(0,1);
        if (FormName.SubGroups[ctr].checked && CheckMsg == 'e') {
          FormName.SubGroups[ctr].checked = 0;
        }
      }
      FormName.SubGroups_2.checked = 0;
    }
    CheckMsg = "";
  }
  return(false);
}

function SubGroups_3_Check() {
  if (FormName.SubGroups_3.checked) {
    CheckMsg = "<%=Translate("Attention",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("You have selected All Groups for the Intercon Region.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Please ensure that you have authority to enable this content or event item to be viewed by All Groups within this Region.",Alt_Language,conn)%>\r\n\n";
    CheckMsg += "<%=Translate("Checkboxes with red borders containing a checkmark are default presets enabled by the Site Administrator.  These presets can be individually unchecked if the content or event item is not appropriate for this group.",Alt_Language,conn)%>\r\n\n";    
    CheckMsg += "<%=Translate("Click [OK] to select All Groups within this Region, or click [Cancel] to uncheck All Groups within this region.",Alt_Language,conn)%>\r\n";

    if (<%=Admin_Region%> != 3) {
      CheckSts = confirm(CheckMsg);
    }
    else {
      CheckSts = true;
    }  

    if (CheckSts == true) {
      for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
        CheckMsg = FormName.SubGroups[ctr].value.substring(0,1);
        if (! FormName.SubGroups[ctr].checked && CheckMsg == 'i') {
          FormName.SubGroups[ctr].checked = 1;
        }
      }
      FormName.SubGroups_3.checked = 0;
    }
    else {
      for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
        CheckMsg = FormName.SubGroups[ctr].value.substring(0,1);
        if (FormName.SubGroups[ctr].checked && CheckMsg == 'i') {
          FormName.SubGroups[ctr].checked = 0;
        }
      }
      FormName.SubGroups_3.checked = 0;
    }
    CheckMsg = "";
  }
  return(false);
}

function Check_Filename(field) {
  var valid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_- :.\\/"
  var ok = "yes";
  var temp;
  var CheckMsg = "";
  for (var i=0; i<field.value.length; i++) {
    temp = "" + field.value.substring(i, i+1);
    if (valid.indexOf(temp) == "-1") ok = "no";
    }
    if (ok == "no") {
    CheckMsg  = "Invalid Character(s)in Local Directory Path and/or Filename.Ext\r\n\n";
    CheckMsg += "The file name that you just loaded contains invalid or hidden character(s).\r\n";
    CheckMsg += "Valid Characters are: [A-Z],[a-z],[0-9],[space, underscore, hyphen and period]\r\n\n";
    CheckMsg += "Invalid Characters appear in this directory path or file name below:\r\n\n";
    CheckMsg += field.value + "\r\n\n";
    CheckMsg += "To fix this problem, click [OK] to close this alert message. The illegal file name will be highlighted.\r\n\n";
    CheckMsg += "Delete the file name from the highlighted input box by pressing the [Delete] key.\r\n\n";
    CheckMsg += "Use Windows Explorer to rename the file path and/or file name to eliminate the invalid characters.\r\n\n";
    CheckMsg += "Return to this form and use [Browse] to reload the corrected file name.";
    alert(CheckMsg);
    field.focus();
    field.select();
  }
}

function Check_POD_File(field) {
  var POD_Name = field;
  if (POD_Name.value.length > 0) {
    var mystyle = "height=500,width=700,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no,resizable=yes";
    var CkItemWindow = window.open('/sw-administrator/SW-Ck_POD_File.asp?Asset_ID=' + FormName.ID.value + '&Item_Number=' + FormName.Item_Number.value + '&Site_ID=<%=Site_ID%>' + '&Language=<%=Login_Language%>','CkItemWinddow',mystyle);
    return true;
  }
  else {
    return true;
  }
}

var highlightcolor="lightyellow"
var ns6=document.getElementById&&!document.all
var previous=''
var eventobj

// Regular expression to highlight only form elements
var intended=/INPUT|TEXTAREA/

// Function to check whether element clicked is form element

function checkel(which) {
  if (which.style&&intended.test(which.tagName)) {
    if (ns6&&eventobj.nodeType==3)
      eventobj=eventobj.parentNode.parentNode
    return true;
  }
  else
  return false;
}

function highlight(e){
  eventobj=ns6? e.target : event.srcElement
  if (previous!='') {
    if (checkel(previous))
      previous.style.backgroundColor=''
      previous=eventobj
      if (checkel(eventobj))
      eventobj.style.backgroundColor=highlightcolor
  }
  else {
    if (checkel(eventobj))
      eventobj.style.backgroundColor=highlightcolor
    previous=eventobj
  }
}

function ck_item_number() {

  var mystyle        = "height=500,width=700,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no,resizable=yes";
  var CheckMsg = "";
  
  if (FormName.Item_Number.value != "" && FormName.Item_Number.value.length == 7) {
    var CkItemWindow = window.open('/sw-administrator/SW-Ck_Item_Number.asp?Asset_ID=' + FormName.Item_Number.value + '&Site_ID=<%=Site_ID%>' + '&Language=<%=Login_Language%>' + '&FormName=<%=FormName%>','CkItemWinddow',mystyle);
  }
  else {
    if (FormName.Item_Number.value != "" && !IsNumeric(FormName.Item_Number.value)) {
      CheckMsg  = "Invalid Character(s)in Item / Reference Number: 1\r\n\n";
      CheckMsg += "The Item / Reference Number: 1 value you have just entered contains invalid characters.\r\n";
      CheckMsg += "A valid Item / Reference Number: 1, consists of a 7-digit number.\r\n\n";
      alert(CheckMsg);
      
      FormName.Item_Number.style.backgroundColor = "#FFB9B9";
      FormName.Item_Number.focus();
      FormName.Item_Number.select();
    }  
  }

  if (FormName.Item_Number.value == "") {
    FormName.Revision_Code.value = "";
    FormName.Cost_Center.value = "";
    FormName.Title.value = "";
    FormName.Content_Language.options[0].selected = true;    
    FormName.Item_Number_Show.checked = false;
    FormName.SubGroups[0].checked = false;
    FormName.SubGroups[1].checked = false;
    FormName.SubGroups[2].checked = false;    
    FormName.Item_Number.focus();
  }  
}

function IsDate(datein){
	
	var indate=datein;

	if (indate.indexOf("-") != -1) {
		var sdate = indate.split("-");
	}
	else {
		var sdate = indate.split("/");
	}

	var chkDate=new Date(Date.parse(indate));
	
	var cmpDate=(chkDate.getMonth()+1) + "/" + (chkDate.getDate())+"/" + (chkDate.getYear());
	var indate2=(Math.abs(sdate[0])) + "/" + (Math.abs(sdate[1]))+"/" + (Math.abs(sdate[2]));

	if (indate2 != cmpDate) {  
    return 0;
	}
	else {
		if (cmpDate == "NaN/NaN/NaN") {
      return false;
		}
		else {
      return true;
		}	
	}
}

function IsNumeric(sText) {

  var ValidChars = "0123456789.";
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

function MAC_Date_Override() {

  FormName.LDays.disabled=true;
  FormName.BDate.disabled=true;
  FormName.EDate.disabled=true;
  FormName.PEDate.disabled=true;
  <%
  if Show_Subscription = true then
    response.write "FormName.Subscription_Early[0].disabled = true;" & vbCrLf
    response.write "FormName.Subscription_Early[1].disabled = true;" & vbCrLf
  end if
  %>
  for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
    if (FormName.SubGroups[ctr].checked == true) {
      if (FormName.SubGroups[ctr].value == 'nomac') {
        FormName.LDays.disabled=false;
        FormName.BDate.disabled=false;
        FormName.EDate.disabled=false;
        FormName.PEDate.disabled=false;
        <%
        if Show_Subscription = true then
          response.write "FormName.Subscription_Early[0].disabled = false;" & vbCrLf
          response.write "FormName.Subscription_Early[0].checked = true;" & vbCrLf          
          response.write "FormName.Subscription_Early[1].checked = false;" & vbCrLf
        end if
        %>
        alert("<%=Translate("Override MAC Date",Alt_Language,conn)%>\r\n\n<%=Translate("Caution: Overriding the MAC date could result in showing this asset to a user prior to the MAC GO LIVE date if this asset's Pre-Announce + Beginning Date preceds the MAC's Pre-Announce + Beginning Date.", Alt_Language, conn)%>\n\r\r<%=Translate("Also note:  This asset's PREVIEW / LIVE / ARCHIVE status is independent of the MAC's REVIEW / LIVE / ARCHIVE status and must be set.",Alt_Language,conn)%>");
        break
      }
    }  
  }  
}

// Added by zensar.This function calls the asp page which checks for the validation.

function fnSendXMLHTTPRequest_Post() {

	var objHTTP, strParameters,strResult,navClone,recID;
	objHTTP = new ActiveXObject('Microsoft.XMLHTTP');
	objHTTP.Open('POST',"Sw-Pcat_FNet_Validate.asp",false);
	objHTTP.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
  
	if (isNaN(FormName.ID.value)== true) {
	  navClone = "";
	}
	else {
      if(FormName.Nav_Clone) {
	    navClone=FormName.Nav_Clone.value;
      }  
	  else {
	    navClone="";
      } 
      recID=FormName.ID.value ; 
	}
	var strCountry;
	var ctrEelement = document.getElementsByName("Country_Reset");
	var strSelectedCountry;
	if(!ctrEelement==null) {
    if (FormName.Country_Reset[0].checked==true) {
      strCountry=FormName.Country_Reset[0].value;
    }
    else if (FormName.Country_Reset[1].checked==true) {
      strCountry=FormName.Country_Reset[1].value;
    }
    else {
      strCountry=FormName.Country_Reset[2].value;
    }
	  strSelectedCountry = FormName.Country.value;
	}
	else {
	    strCountry="none";
	    strSelectedCountry = "";
	}

 
  var obj;
  strParameters = "PCat=" + FormName.PCat.value + "&Nav_Clone=" + navClone + "&Category_ID=" ;
  strParameters = strParameters + FormName.Category_ID.value + "&Title=" + escape(FormName.Title.value) + "&ID=" + recID ;
  strParameters = strParameters + "&Item_Number=" + FormName.Item_Number.value + "&Content_Language=" + FormName.Content_Language.value;
  strParameters = strParameters + "&Country=" + strSelectedCountry + "&Country_Reset=" + strCountry;
  strParameters = strParameters + "&Site_ID=" + <%=Site_ID%> + "&opr=" + FormName.opr.value;
  
  objHTTP.send(strParameters);
  strResult = objHTTP.responseText;
  return strResult;
}

// Checks Required Fields before form is submitted
function CheckRequiredFields1(){
alert('dfgdgdfg');
return false;
}
function CheckRequiredFields() {
//Modified by zensar on 09-08-2006
  var Show_PID      = -1;
  var PID_System    =  0;
  var AspShow_PID   = <%=cint(Show_PID)%>;
  var AspPID_System = <%=PID_System%>;
  var Show_msg      = false;
  var cloneRecords = '';
  var strResult;

 
  if (FormName.opr.value=="D")
  {//alert('1');
        if (Show_PID == AspShow_PID)
           {//alert('2');
            if (PID_System == AspPID_System) 
                {//alert('3');
                    strResult = fnSendXMLHTTPRequest_Post();
                    if (strResult != "")
                     {
                      if (strResult.substr(0,7) == "confirm")
                       {
  		                    cloneRecords = strResult.substring(7);
	                   }
                      else 
                       {    
                            if (strResult.substr(0,5) == "Error")
                            {
                                ErrorMsg = strResult + "\r\n";
                                alert(ErrorMsg);
                                return false;
                            }
                            else
                            {
                                ErrorMsg = ErrorMsg + strResult + "\r\n";
                            }
                       }
                     }
                     if(cloneRecords != ""){
		                    var ContinueYn;
		                    ContinueYn = confirm(cloneRecords);
		                    if (ContinueYn == false) {
			                    FormName.opr.value = '';
			                    return(false);
		                    }
	                    }
                 }
           }
        return(true);
  }  
 //******************************************* 
    
  if (Menu_Button == false) {
   
    var ErrorMsg = "";
    for (var i = 0; i < FormName.Content_Group.length; i++)
    {

      if (FormName.Content_Group[i].selected) 
      {
        if (FormName.Content_Group.value == '0') 
        {
        }    
        else 
           {
                  for (var t = 0; t < FormName.Campaign.length; t++) 
                  {
                    if (FormName.Campaign[t].selected)
                     {
                      if (FormName.Campaign.value == '0') 
                      {
                        FormName.Campaign.style.backgroundColor = "#FFB9B9";            
                        ErrorMsg = ErrorMsg + "<%=Translate("Missing",Alt_Language,conn) & " " & Code_X_Name & " " & Translate("Name",Alt_Language,conn)%>\r\n";
                      }
                     }
                   }  
           }
        }
      } 
    }
    

    
    // Product or Product Series
    if (FormName.Product.value.length == 0) {
      for (var i = 0; i < FormName.Product_New.length; i++) {
        if (FormName.Product_New[i].selected) {
          if (FormName.Product_New.value == '') {
            FormName.Product_New.style.backgroundColor = "#FFB9B9";
            ErrorMsg = ErrorMsg + "<%=Translate("Missing Product or Product Family",Alt_Language,conn)%>\r\n\n";        
          }    
        }
      }
    }  
    
    // Title  
    if (FormName.Title.value.length == 0) {
      FormName.Title.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Missing Title",Alt_Language,conn)%>\r\n\n";        
    }
    
    // Language
   
    if (FormName.Content_Language[0].selected == true) {
      FormName.Content_Language.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Please select a Language for this asset",Alt_Language,conn)%>\r\n\n";        
    }
    
    // Check for Oracle Item Number and Revision Code

    if (FormName.Item_Number.value.length == 7) {
      if (!IsNumeric(FormName.Item_Number.value)) {
        FormName.Item_Number.style.backgroundColor = "#FFB9B9";    
        ErrorMsg = ErrorMsg + "<%=Translate("Invalid Item Reference Number 1.  Value must be 7 numeric characters in length. Format: ####### (Oracle Item Number)",Alt_Language,conn)%>\r\n";
      }  
      else if (FormName.Revision_Code.value.length == 0 || FormName.Revision_Code.value.length > 1 || IsNumeric(FormName.Revision_Code.value)) {
          FormName.Revision_Code.style.backgroundColor = "#FFB9B9";
          ErrorMsg = ErrorMsg + "<%=Translate("Invalid Revision Code  Value must be 1 alpha character [A-Z] in length",Alt_Language,conn)%>\r\n";
      }    
      else {
        if (FormName.File_Name.value.length == 0 && FormName.URLLink.value.length == 0) {
          if (PID_System == AspPID_System) {
            FormName.File_Name.style.backgroundColor = "#FFB9B9";        
            ErrorMsg = ErrorMsg + "<%=Translate("Missing Asset File - (LOW Resolution) for Item Reference Number 1, or URL to Web Page.",Alt_Language,conn)%>\r\n";        
          }
          else {
            if (<%=CInt(Show_File)%> == -1 && <%=CInt(Show_URL)%> == 0) {
              FormName.File_Name.style.backgroundColor = "#FFB9B9";
              FormName.File_Name.focus();
              alert("<%=Translate("Warning - Missing Asset File - (LOW Resolution) for Item Reference Number 1",Alt_Language,conn)%>\r\n");
            }
            else if (<%=CInt(Show_File)%> == 0 && <%=CInt(Show_URL)%> == -1) {
              FormName.URLLink.style.backgroundColor = "#FFB9B9";
              FormName.URLLink.focus();
              alert("<%=Translate("Warning - Missing URL to Web Page.",Alt_Language,conn)%>\r\n");          
            }
            else {
              FormName.URLLink.style.backgroundColor = "#FFB9B9";
              FormName.File_Name.style.backgroundColor = "#FFB9B9";
              FormName.URLLink.focus();              
              FormName.File_Name.focus();
              alert("<%=Translate("Warning - Missing Asset File - (LOW Resolution) for Item Reference Number 1, or URL to Web Page.",Alt_Language,conn)%>\r\n");
            }  
          }  
        }        
      }
    }              

    // Check for European Literature Number and Revision Code
  
   /* else if (FormName.Item_Number.value.length == 9) {
    
      var mypart = FormName.Item_Number.value
      var part1 = mypart.substr(0,5);
      var part2 = mypart.substr(5,1);
      var part3 = mypart.substr(6,3);
      
      if (!IsNumeric(part1) || part2 != "-" || IsNumeric(part3)) {
        FormName.Item_Number.style.backgroundColor = "#FFB9B9";
        ErrorMsg = ErrorMsg + "<%=Translate("Invalid Item Reference Number 1 - 9-character European Literature Numbers must be 5-numeric characters + '-' + 3-alpha characters. Format: #####-NNN.",Alt_Language,conn)%>\r\n";
      }  
      else {
      }
      
      if (FormName.Revision_Code.value.length == 0 || FormName.Revision_Code.value.length > 1 || IsNumeric(FormName.Revision_Code.value)) {
        if (FormName.Revision_Code.value.length == 0) {
          FormName.Revision_Code.value = "A";
        }
        else {
          FormName.Revision_Code.style.backgroundColor = "#FFB9B9";
          ErrorMsg = ErrorMsg + "<%=Translate("Invalid Revision Code - Value must be 1-alpha character [A-Z] in length. Use 'A' for the default value.",Alt_Language,conn)%>\r\n";
        }  
      }
      
      if (FormName.File_Name.value.length == 0 && FormName.URLLink.value.length == 0) {
        if (PID_System == AspPID_System) {
          FormName.File_Name.style.backgroundColor = "#FFB9B9";        
          ErrorMsg = ErrorMsg + "<%=Translate("Missing Asset File - (LOW Resolution) for Item Reference Number 1, or URL to Web Page.",Alt_Language,conn)%>\r\n";        
        }
        else {
          if (<%=CInt(Show_File)%> == -1 && <%=CInt(Show_File)%> == 0) {
            FormName.File_Name.style.backgroundColor = "#FFB9B9";
            FormName.File_Name.focus();
            alert("<%=Translate("Warning - Missing Asset File - (LOW Resolution) for Item Reference Number 1",Alt_Language,conn)%>\r\n");
          }
          else if (<%=CInt(Show_File)%> == 0 && <%=CInt(Show_URL)%> == -1) {
            FormName.URLLink.style.backgroundColor = "#FFB9B9";
            FormName.URLLink.focus();
            alert("<%=Translate("Warning - Missing URL to Web Page.",Alt_Language,conn)%>\r\n");          
          }
          else {
              FormName.URLLink.style.backgroundColor = "#FFB9B9";
              FormName.File_Name.style.backgroundColor = "#FFB9B9";
              FormName.URLLink.focus();              
              FormName.File_Name.focus();
            alert("<%=Translate("Warning - Missing Asset File - (LOW Resolution) for Item Reference Number 1, or URL to Web Page.",Alt_Language,conn)%>\r\n");
          }  
        }  
      }        
    }

    else if (FormName.Item_Number.value.length > 0 && (FormName.Item_Number.value.length != 7 && FormName.Item_Number.value.length != 9)) {
      FormName.Item_Number.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Item Reference Number 1 - Value must be 7-numeric characters in length. Format: ####### (Oracle Item Number Format), or 9-characters in length Format: #####-NNN (European Literature Number).",Alt_Language,conn)%>\r\n";
    }
*/
alert('dfgdfg');
return false;
    // Store original Enabled/Disable values and set to enabled
    /*
    if (<%=Show_Date_Basic%> == -1) {
      FormName.EDate.value = FormName.BDate.value;
    }
    
   
    var BDateDisabled  = FormName.BDate.disabled;
    var EDateDisabled  = FormName.EDate.disabled;
    var LDaysDisabled  = FormName.LDays.disabled;
    var PEDateDisabled = FormName.PEDate.disabled;
    
    FormName.BDate.disabled = false;
    FormName.EDate.disabled = false;
    FormName.LDays.disabled = false;
    FormName.PEDate.disabled = false;

    if (FormName.BDate.value.length == 0 || IsDate(FormName.BDate.value) == 0) {
      FormName.BDate.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Begin Date or Date Format.  Use: mm/dd/yyyy",Alt_Language,conn)%>\r\n";
    }
 
    if (FormName.EDate.value.length == 0 || IsDate(FormName.EDate.value) == 0) {
      FormName.EDate.style.backgroundColor = "#FFB9B9";  
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Ending Date or Date Format.  Use: mm/dd/yyyy",Alt_Language,conn)%>\r\n";
    }
    
    if (FormName.PEDate.value.length != 0) {
      if (IsDate(FormName.PEDate.value) == 0) {
        FormName.PEDate.style.backgroundColor = "#FFB9B9";
        ErrorMsg = ErrorMsg + FormName.PEDate.value + "\r\n";
        ErrorMsg = ErrorMsg + "<%=Translate("Invalid Public Release Date or Date Format.  Use: mm/dd/yyyy",Alt_Language,conn)%>\r\n";
      }  
    }
  
    var BDate = new Date(FormName.BDate.value);
    var EDate = new Date(FormName.EDate.value);
  
    if (BDate.getTime() - EDate.getTime() > 0) {
      FormName.EDate.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Date -  End Date is before Begin Date.",Alt_Language,conn)%>\r\n";
    }
        
        
    if (Show_PID == AspShow_PID) {
      if (PID_System == AspPID_System) {

        //selected product combo.

        var products;
        if (FormName.PCat_SProducts.length) {
          for (products = 0; products < FormName.PCat_SProducts.length; products++) {
            FormName.PCat_SProducts.options[products].selected = true;
          }
        }
        if (FormName.PcatRelationNone.checked == false) {
          for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
            if (FormName.SubGroups[ctr].value == "fedl") {
              if(FormName.opr.value != "D") {
                FormName.SubGroups[ctr].checked = true;
              }
              break;
            }
          }
          Show_msg = true;    
        }
      }
    }  

      
    //---------------------------------------------------------------------------------
    
    if (!IsNumeric(FormName.LDays.value)) {
      FormName.LDays.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Date -  Pre-Announce days must be a positive numeric value",Alt_Language,conn)%>\r\n";    
    }
    
    if (!IsNumeric(FormName.XDays.value)) {
      FormName.XDays.style.backgroundColor = "#FFB9B9";
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid Date -  Move to Archive days must be a positive numeric value",Alt_Language,conn)%>\r\n";    
    }
  
    if (<%=Show_Submission_Approve%> == -1) {
      if (document.getElementById("Review_By_Group")) {
        CheckSts = false;
        for (ctr=0; ctr < FormName.Review_By_Group.length; ctr++) {
          if (FormName.Review_By_Group[ctr].selected == true || FormName.Review_By_Group.value == "0") {
            CheckSts = true;
            break
          }  
        }
        if (CheckSts == false) {
          FormName.Review_By_Group.style.backgroundColor = "#FFB9B9";
          ErrorMsg = ErrorMsg + "<%=Translate("Missing -  You have not selected a group to Approve this asset.",Alt_Language,conn)%>\r\n";        
        }
      }  
    }  
    
    CheckSts = false;
    for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
      if (FormName.SubGroups[ctr].checked == true) {
        if (FormName.SubGroups[ctr].value != 'view' && FormName.SubGroups[ctr].value != 'shpcrt') {

          CheckSts = true;
          break
        }  
      }  
    }
    if (CheckSts == false) {
      ErrorMsg = ErrorMsg + "<%=Translate("Missing -  No Groups have been selected to view this asset.",Alt_Language,conn)%>\r\n";        
    }

    CheckSts = true;
    for (ctr=0; ctr < FormName.SubGroups.length; ctr++) 
    {   
        if ((FormName.SubGroups[ctr].value == 'view') || (FormName.SubGroups[ctr].value == 'fedl') || (FormName.SubGroups[ctr].value == 'shpcrt')) 
        {
            if (FormName.SubGroups[ctr].checked == true) 
            { 
              if (FormName.Item_Number.value.length != 7 && FormName.Item_Number.value.length != 9)
               {
                CheckSts = false;
                break
               }
            }    
        } 
    }

    
    if (CheckSts == false) {
      ErrorMsg = ErrorMsg + "<%=Translate("Invalid or Missing Item Reference Number 1 - This number and revision is required for all assets used for Electronic Email Fulfillment - (End-User Oracle)(EEF), Electronic Fulfillment - (End-User Digital Library)(FDL), or Print on Demand Publications (POD).",Alt_Language,conn)%> ";
      ErrorMsg = ErrorMsg + "<%=Translate("If not applicable, in the Select Groups Allowed to View this Information section, uncheck the EEF and/or FDL option(s) or supply missing Item Number 1.",Alt_Language,conn)%>\r\n\n";
      FormName.Item_Number.style.backgroundColor = "#FFB9B9";
      FormName.Revision_Code.style.backgroundColor = "#FFB9B9";
      if (Show_msg == true) {
         ErrorMsg = ErrorMsg + "<%=Translate("Invalid or Missing Item Reference Number 1 - Please select the checkbox, Do not relate to Product Catalog, if the asset is not available for End-User Digital Library)(FDL). ",Alt_Language,conn)%>\r\n\n";
      }
    }
    
    if (Show_PID == AspShow_PID) {
      if (PID_System == AspPID_System) {
        strResult = fnSendXMLHTTPRequest_Post();
        if (strResult != "") {
          if (strResult.substr(0,7) == "confirm") {
	      		cloneRecords = strResult.substring(7);
      		}
    	  	else {
                    if (strResult.substr(0,5) == "Error")
                    {
                        ErrorMsg = strResult + "\r\n";
                        alert(ErrorMsg);
                        return false;
                    }
                    else
                    {
                        ErrorMsg = ErrorMsg + strResult + "\r\n";
                    }
    			}
	      }
      }
    }
    alert('Test4');
    if (ErrorMsg.length) {
  
      // Reset Enabled/Disable back to original value
      FormName.BDate.disabled  = BDateDisabled;
      FormName.EDate.disabled  = EDateDisabled;
      FormName.LDays.disabled  = LDaysDisabled;
      FormName.PEDate.disabled = PEDateDisabled;
  
      ErrorMsg = "<%=Translate("Please enter the missing information for following REQUIRED fields (or use N/A)",Alt_Language,conn)%>:\r\n\n" + ErrorMsg;
      Menu_Button = false;
      alert (ErrorMsg);
      return (false);
    }

    else {
alert('in else Test5');
      var goStatus = true;
      
      if(cloneRecords != "") {
		    var ContinueYn;
		    ContinueYn = confirm(cloneRecords);
		    if (ContinueYn == false) {
			    FormName.opr.value = '';
			    return false;
		    }
	    }

      if ((FormName.Item_Number.value.length == 7 || FormName.Item_Number.value.length == 9) && FormName.Revision_Code.value.length == 1 && <%=Show_Item_Number%> == <%=CInt(True)%>) {
      
        goStatus = false;
        CheckMsg =  "<%=Translate("Attention",Alt_Language,conn)%>\r\n\n";
        CheckMsg += "<%=Translate("You have indicated that this asset is a Literature Item, however to make this asset available to the Electronic File Fulfilment System (EFF), the orange: Available to Electronic Fulfillment - (End-User Viewable) checkbox must be selected.",Alt_Language,conn)%>\r\n\n";
        CheckMsg += "<%=Translate("If you do not want this asset available to the EFF system no further action is needed.",Alt_Language,conn)%>\r\n\n";
        CheckMsg += "<%=Translate("Click on the [OK] button to proceed.",Alt_Language,conn)%>\r\n";
        for (ctr=0; ctr < FormName.SubGroups.length; ctr++) {
          if (FormName.SubGroups[ctr].checked == true) {
            if (FormName.SubGroups[ctr].value == "view" || FormName.SubGroups[ctr].value == "fedl") {
              goStatus = true;
            }
          }     
        }
        if (CheckMsg.length && goStatus == false) {
          alert(CheckMsg);
          goStatus = true;          
        }
      }

      <%
      if Show_Subscription = true then
        response.write "FormName.Subscription_Early[0].disabled = false;" & vbCrLf
        response.write "FormName.Subscription_Early[1].disabled = false;" & vbCrLf
      end if
      %>
     
// --------------------------------------------------------------------------------      
// For Language Versions we need to Base64 Encode strings on Submit and Decode on the Server before storing to DB
// The following provides this functionality for all strings meeting this criteria.
// --------------------------------------------------------------------------------

      if (FormName.Title != null) { 
        FormName.Title_B64.value         = encodeBase64(FormName.Title.value);
      }

      if (FormName.Description != null) {       
        FormName.Description_B64.value   = encodeBase64(FormName.Description.value);
      }
      if (FormName.Instructions != null) {      
        FormName.Instructions_B64.value  = encodeBase64(FormName.Instructions.value);
      }
      if (FormName.Splash_Header != null) {
        FormName.Splash_Header_B64.value = encodeBase64(FormName.Splash_Header.value);
      }
      if (FormName.Splash_Footer != null) {                  
        FormName.Splash_Footer_B64.value = encodeBase64(FormName.Splash_Footer.value);
      }
      if (FormName.Location != null) {       
        FormName.Location_B64.value      = encodeBase64(FormName.Location.value);
      }
        
      if (goStatus == true) {
        if (<%=CInt(FileUpEE_Flag)%> == -1) {
          startUpload();
        }
      }    
      return(goStatus);
    }
  }
              
  Menu_Button = false;
  
  if (<%=CInt(FileUpEE_Flag)%> == -1) {
    startUpload();
  }
  return (true);  */
}  

// --------------------------------------------------------------------------------
// Client Side Unicode Encoding for Multipart Form
// --------------------------------------------------------------------------------

var END_OF_INPUT = -1;

var base64Chars = new Array(
    'A','B','C','D','E','F','G','H',
    'I','J','K','L','M','N','O','P',
    'Q','R','S','T','U','V','W','X',
    'Y','Z','a','b','c','d','e','f',
    'g','h','i','j','k','l','m','n',
    'o','p','q','r','s','t','u','v',
    'w','x','y','z','0','1','2','3',
    '4','5','6','7','8','9','+','/'
);

var reverseBase64Chars = new Array();

for (var i=0; i < base64Chars.length; i++){
  reverseBase64Chars[base64Chars[i]] = i;
}

var base64Str;
var base64Count;

function setBase64Str(str){
  base64Str = str;
  base64Count = 0;
}

function readBase64(){    
  if (!base64Str) return END_OF_INPUT;
  if (base64Count >= base64Str.length) return END_OF_INPUT;
  var c = base64Str.charCodeAt(base64Count) & 0xff;
  base64Count++;
  return c;
}

function encodeBase64(str){
  str = utf8Encode(str);
  setBase64Str(str);
  var result = '';
  var inBuffer = new Array(3);
  var lineCount = 0;
  var done = false;
  while (!done && (inBuffer[0] = readBase64()) != END_OF_INPUT){
    inBuffer[1] = readBase64();
    inBuffer[2] = readBase64();
    result += (base64Chars[ inBuffer[0] >> 2 ]);
    if (inBuffer[1] != END_OF_INPUT){
      result += (base64Chars [(( inBuffer[0] << 4 ) & 0x30) | (inBuffer[1] >> 4) ]);
      if (inBuffer[2] != END_OF_INPUT){
        result += (base64Chars [((inBuffer[1] << 2) & 0x3c) | (inBuffer[2] >> 6) ]);
        result += (base64Chars [inBuffer[2] & 0x3F]);
      }
      else {
        result += (base64Chars [((inBuffer[1] << 2) & 0x3c)]);
        result += ('=');
        done = true;
      }
    }
    else {
      result += (base64Chars [(( inBuffer[0] << 4 ) & 0x30)]);
      result += ('=');
      result += ('=');
      done = true;
    }
    lineCount += 4;
    if (lineCount >= 76){
      result += ('\n');
      lineCount = 0;
    }
  }
  return result;
}

function readReverseBase64(){   
  if (!base64Str) return END_OF_INPUT;
  
  while (true){      
    if (base64Count >= base64Str.length) return END_OF_INPUT;
    var nextCharacter = base64Str.charAt(base64Count);
    base64Count++;
    if (reverseBase64Chars[nextCharacter]){
      return reverseBase64Chars[nextCharacter];
    }
    if (nextCharacter == 'A') return 0;
  }
  return END_OF_INPUT;
}

function ntos(n){
  n=n.toString(16);
  if (n.length == 1) n="0"+n;
  n="%"+n;
  return unescape(n);
}

function decodeBase64(str){
  setBase64Str(str);
  var result = "";
  var inBuffer = new Array(4);
  var done = false;
  while (!done && (inBuffer[0] = readReverseBase64()) != END_OF_INPUT
    && (inBuffer[1] = readReverseBase64()) != END_OF_INPUT){
    
    inBuffer[2] = readReverseBase64();
    inBuffer[3] = readReverseBase64();
    result += ntos((((inBuffer[0] << 2) & 0xff)| inBuffer[1] >> 4));
    if (inBuffer[2] != END_OF_INPUT){
      result +=  ntos((((inBuffer[1] << 4) & 0xff)| inBuffer[2] >> 2));
      if (inBuffer[3] != END_OF_INPUT){
        result +=  ntos((((inBuffer[2] << 6)  & 0xff) | inBuffer[3]));
      }
      else {
        done = true;
      }
    }
    else {
      done = true;
    }
  }
  return result;
}

var digitArray = new Array('0','1','2','3','4','5','6','7','8','9','a','b','c','d','e','f');

function toHex(n){
  var result = ''
  var start = true;
  for (var i=32; i>0;){
    i-=4;
    var digit = (n>>i) & 0xf;
    if (!start || digit != 0){
      start = false;
      result += digitArray[digit];
    }
  }
  return (result==''?'0':result);
}

function pad(str, len, pad){
  var result = str;
  for (var i=str.length; i<len; i++){
    result = pad + result;
  }
  return result;
}

function encodeHex(str){
  var result = "";
  for (var i=0; i<str.length; i++){
    result += pad(toHex(str.charCodeAt(i)&0xff),2,'0');
  }
  return result;
}

function decodeHex(str){
  str = str.replace(new RegExp("s/[^0-9a-zA-Z]//g"));
  var result = "";
  var nextchar = "";
  for (var i=0; i<str.length; i++){
    nextchar += str.charAt(i);
    if (nextchar.length == 2){
      result += ntos(eval('0x'+nextchar));
      nextchar = "";
    }
  }
  return result;
}

function chr(code)
{
	return String.fromCharCode(code);
}

//returns utf8 encoded charachter of a unicode value.
//code must be a number indicating the Unicode value.
//returned value is a string between 1 and 4 charachters.
function code2utf(code)
{
	if (code < 128) return chr(code);
	if (code < 2048) return chr(192+(code>>6)) + chr(128+(code&63));
	if (code < 65536) return chr(224+(code>>12)) + chr(128+((code>>6)&63)) + chr(128+(code&63));
	if (code < 2097152) return chr(240+(code>>18)) + chr(128+((code>>12)&63)) + chr(128+((code>>6)&63)) + chr(128+(code&63));
}

//it is a private function for internal use in utf8Encode function 
function _utf8Encode(str)
{	
	var utf8str = new Array();
	for (var i=0; i<str.length; i++) {
		utf8str[i] = code2utf(str.charCodeAt(i));
	}
	return utf8str.join('');
}

//Encodes a unicode string to UTF8 format.
function utf8Encode(str)
{	
	var utf8str = new Array();
	var pos,j = 0;
	var tmpStr = '';
	
	while ((pos = str.search(/[^\x00-\x7F]/)) != -1) {
		tmpStr = str.match(/([^\x00-\x7F]+[\x00-\x7F]{0,10})+/)[0];
		utf8str[j++] = str.substr(0, pos);
		utf8str[j++] = _utf8Encode(tmpStr);
		str = str.substr(pos + tmpStr.length);
	}
	
	utf8str[j++] = str;
	return utf8str.join('');
}

//it is a private function for internal use in utf8Decode function 
function _utf8Decode(utf8str)
{	
	var str = new Array();
	var code,code2,code3,code4,j = 0;
	for (var i=0; i<utf8str.length; ) {
		code = utf8str.charCodeAt(i++);
		if (code > 127) code2 = utf8str.charCodeAt(i++);
		if (code > 223) code3 = utf8str.charCodeAt(i++);
		if (code > 239) code4 = utf8str.charCodeAt(i++);
		
		if (code < 128) str[j++]= chr(code);
		else if (code < 224) str[j++] = chr(((code-192)<<6) + (code2-128));
		else if (code < 240) str[j++] = chr(((code-224)<<12) + ((code2-128)<<6) + (code3-128));
		else str[j++] = chr(((code-240)<<18) + ((code2-128)<<12) + ((code3-128)<<6) + (code4-128));
	}
	return str.join('');
}

//Decodes a UTF8 formated string
function utf8Decode(utf8str)
{
	var str = new Array();
	var pos = 0;
	var tmpStr = '';
	var j=0;
	while ((pos = utf8str.search(/[^\x00-\x7F]/)) != -1) {
		tmpStr = utf8str.match(/([^\x00-\x7F]+[\x00-\x7F]{0,10})+/)[0];
		str[j++]= utf8str.substr(0, pos) + _utf8Decode(tmpStr);
		utf8str = utf8str.substr(pos + tmpStr.length);
	}
	
	str[j++] = utf8str;
	return str.join('');
}

// PCat Module Function

function AddRemoveOptions(strFrom,strTo) {	
  var i=0;
  var objfrom = eval('document.<%=FormName%>.' + strFrom);
  var objto   = eval('document.<%=FormName%>.' + strTo);
  
  
  for(i=(objfrom.options.length-1);i>=0;i--) {	
 		if (objfrom.options[i].selected==true) {
 			var optnew;
 			optnew = document.createElement("OPTION") 
 			optnew.text=objfrom.options[i].text;
 			optnew.value=objfrom.options[i].value;
 			if(checkifexists(objfrom.options[i].value)==false)
 			{
 			    alert("Product ''" + objfrom.options[i].text + "'' is already present!");
 			    //return false;
 			}
 			else
 			{
 			    objto.options.add(optnew);
 			}    
 			//objfrom.options.remove(i);
 			//sortList(strTo);
		}
 	}
}

function RemoveOption(strFrom) {
  var i;
  var objfrom = eval('document.<%=FormName%>.' + strFrom);
  for(i=(objfrom.options.length-1);i>=0;i--) {	
    if (objfrom.options[i].selected==true) {
      objfrom.options.remove(i);
    }
  }
}

// PCat Module Function

function sortList(objfrom) {	
  var objfrom;
  var objto;
  var i;
  var pos;
  var optnew;

  strFrom=eval('document.<%=FormName%>.' + objfrom);
  
  var listarray = new Array(strFrom.options.length-1);
  
  for(i=(strFrom.options.length-1);i>=0;i--) {
 		listarray[i]=strFrom.options[i].text + '$_' + strFrom.options[i].value;
 	}

 	listarray.sort();

  for(i=(strFrom.options.length-1);i>=0;i--) {
 		strFrom.options.remove(i);
 	}

  for(i=0;i<=listarray.length-1;i++) {
		optnew = document.createElement("OPTION") ;
		pos=listarray[i].indexOf('$_');
		optnew.text=listarray[i].substring(0,pos);
		optnew.value=listarray[i].substring(pos+3,listarray[i].length-pos+3);
		strFrom.options.add(optnew);
	}
}


// FileUpEE Progress Indicator for Uploads

function startUpload() {

  if ((FormName.File_Existing.value.length == 0 && FormName.File_Name.value.length != 0) || (FormName.File_Existing_POD.value.length == 0 && FormName.File_Name_POD.value.length != 0)) {

    var FCnt = 0;
    if (FormName.Thumbnail.value.length     != 0 && FormName.Thumbnail_Existing.value.length == 0) FCnt++;
    if (FormName.File_Name.value.length     != 0 && FormName.File_Existing.value.length == 0)      FCnt++;
    if (FormName.File_Name_POD.value.length != 0 && FormName.File_Existing_POD.value.length == 0)  FCnt++;
    
    var sW = 10;
    var sH = 10;

   	var winstyle = "height=" + sH + "width=" + sW + ",status=no,scrollbars=no,toolbar=no,menubar=no,location=no";
    var Progress = window.open("/SW-FileUp_Progress/SW-File_Progress_EE.asp?wProgressID=<%=wProgressID%>&cProgressID=<%=cProgressID%>&FCnt=" + FCnt,null,winstyle);
		FormName.action = "Calendar_Admin.asp?wProgressID=<%=wProgressID%>&cProgressID=<%=cProgressID%>&FileUpEE_Flag=<%=CInt(FileUpEE_Flag)%>&FileUpEE_Remote_Flag=<%=CInt(FileUpEE_Remote_Flag)%>";
  }
}

//Added by Zensar on 23-06-2006.This functions trims the Text.
function trimAll(sString) 
{
    while (sString.substring(0,1) == ' ')
    {
        sString = sString.substring(1, sString.length);
    }
    while (sString.substring(sString.length-1, sString.length) == ' ')
    {
        sString = sString.substring(0,sString.length-1);
    }
    return sString;
}
//-->
function setFileCheckbox(sCheckbox)
{//alert('fghfghfgh')
    if(FormName.Preserve_Path != null)
    {
        if (sCheckbox=="D")
        {
            if (FormName.Delete_File.checked==true)
                FormName.Preserve_Path.checked = !(FormName.Delete_File.checked)
        }
        else
        {
            if (FormName.Preserve_Path.checked==true)
                     FormName.Delete_File.checked = !(FormName.Preserve_Path.checked)
        } 
    }   
}
function setOperation(operation)
    {
        FormName.opr.value=operation;  
        //return true;
    }
</script>


<!--#include virtual="/include/RTEditor/RTE_Editor_Launch.asp"-->

<%
Call Disconnect_SiteWide
%>