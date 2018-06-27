<%
' --------------------------------------------------------------------------------------
'
' Author: K. D. Whitlock
'
' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Establish connection to SiteWide DB
' --------------------------------------------------------------------------------------

Session("BackURL") = ""

Dim ProgressID
Set uplp = Server.CreateObject("Softartisans.FileUpProgress")
ProgressID = uplp.NextProgressID
'set uplr = nothing

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
Dim Category_Code
Dim Content_Group
Dim Admin_Access
Dim Admin_Name

Dim Path_Site         ' eFulfillment Documents are stored /[site_code]/download/[sub]/*.*
Dim Path_Site_POD     ' Common Directory for Print on Demand Files -- All of SiteWide
Dim Path_Include
Dim Path_File
Dim Path_File_POD
Dim Path_Thumbnail

Dim Show_View
Dim Show_Detail
Dim Show_Location
Dim Show_ImageStore
Dim Show_Link
Dim Show_Item_Number
Dim Show_Link_PopUp_Disabled
Dim Show_File
Dim Show_File_POD
Dim Show_Include
Dim Show_Thumbnail
Dim Show_Subscription
Dim Show_Calendar
Dim Show_Forum
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
Show_File                = False
Show_File_POD            = False
Show_Include             = False
Show_Thumbnail           = False
Show_Subscription        = False
Show_Calendar            = False
Show_Forum               = False
Write_Form_Show_Values   = False

Dim Icon_Type
Dim Icon_Extension

Dim Region
Region         = 0
Dim RegionValue
RegionValue    = ""
Dim RegionColorPointer
RegionColorPointer = 0
Dim RegionColor(3)
RegionColor(0) = "#0000FF"
RegionColor(1) = "#99FFCC"
RegionColor(2) = "#66CCFF"
RegionColor(3) = "#FFCCFF"

BackURL       = "/sw-administrator/Calendar_Edit.asp" 
HomeURL       = "/sw-administrator/Default.asp"

Calendar_ID   = request("ID")
Category_ID   = request("Category_ID")
Content_Group = request("Content_Group")

Path_Include  = "Download\Content"
Path_File     = "Download"
Site_Path_POD = ""
Path_File_POD = "POD"
Path_Thumbnail= "Download\Thumbnail"

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
<FONT CLASS=Medium>
<DIV ALIGN=CENTER>
<B><% if lcase(request("ID")) = "add" then response.write Translate("Add",Login_Language,conn) & " " else response.write Translate("Edit",Login_Language,conn) & " " %><%=Translate("Content or Event",Login_Language,conn)%></B><BR>
<%=Translate("These are date/group specific content or events related to general or specific Product or Product Families.",Login_Language,conn)%>
</DIV>
<BR>
<UL>
<% if isnumeric(request("ID")) then %>
<LI><B><%=Translate("Show / Hide Site View Button",Login_Language,conn)%></B> - <%=Translate("Toggles showing a replication of how this content / event item would appear on the site to the user.",Login_Language,conn)%><BR><BR>
<% end if %>
<LI><B><%=Translate("Content or Event ID Number",Login_Language,conn)%></B> - <%=Translate("Internal database reference number or ADD for new record.  If the record is new, after you click on the [Save / Update] button, the record will be re-displayed with a new Content / Event ID number.",Login_Language,conn)%>
<% if isnumeric(request("ID")) then %>
<LI><B><%=Translate("Locked",Login_Language,conn)%></B> - <%=Translate("A lock is applied to a Content or Event ID Number, indicated by a red [ID Number], by the Site Administrator to prevent Edit, Clone, Duplication or Delete.  If you need to have this record modified, you will need to contact your site administrator.",Login_Language,conn)%></LI>
<% end if %>
<% if isnumeric(request("ID")) then %>
<LI><B><%=Translate("Status",Login_Language,conn)%></B> - <%=Translate("Review allows the Site Administrator to view Content or Event as if it were Live, however, the user is unable to see this Content or Event Item until the status is changed to Live. Content or Event Items are Archived if the Site Administrator selects this Status option or the Content or Event Item expires based on date.",Login_Language,conn)%></LI>
<% end if %>
<BR><BR>
<LI><B><%=Translate("Category",Login_Language,conn)%></B> - <%=Translate("Pre-selected from the Main Menu.  This cannot be changed, unless you first delete this record and re-enter the data under a new Content or Event Category.",Login_Language,conn)%></LI>
<LI><B><%=Translate("Product or Product Family",Login_Language,conn)%></B> - <%=Translate("This is a critical sorting/grouping field.  Please try to add new Content or Event records using one of the pre-existing selections, or if you require a new Product or Product Family name, you can specify a new name by using the input box.",Login_Language,conn)%></LI>
<LI><B><%=Translate("Content Grouping",Login_Language,conn)%></B> - <%=Translate("Each Content Item or Event can be grouped &quot;Individually&quot; (default), or associated with a &quot;Multiple Asset Container&quot; (MAC) grouping of related assets.",Login_Language,conn) & " " & Translate("If a &quot;Multiple Asset Container&quot; is selected, another drop-down selection box will appear allowing you to select which &quot;Multiple Asset Container&quot; to associate this asset to.",Login_Language,conn) & " " & Translate("Content Items or Events associated with a &quot;Multiple Asset Container&quot; may also appear as individual items under their respective categories.",Login_Language,conn)%></LI>
<LI><B><%=Translate("Title",Login_Language,conn)%></B> - <%=Translate("Topic or title of the Content or Event. (Included with Subscription Service)",Login_Language,conn)%></LI>
<LI><B><%=Translate("Description",Login_Language,conn)%></B> - <%=Translate("Short narrative description of the Content or Event. (Included with Subscription Service)",Login_Language,conn)%></LI>
<LI><B><%=Translate("Special Instructions",Login_Language,conn)%></B> - <%=Translate("Short instructions of how to use, order, or other instructions related to the Content or Event.",Login_Language,conn)%></LI>
<% if Show_Item_Number = True then %>
<BR><BR>
<LI><B><%=Translate("Item / Reference Number",Login_Language,conn)%></B> - <%=Translate("Oracle/MfgPro Item Number, Literature Number or other reference designator.",Login_Language,conn)%>&nbsp;&nbsp;<%=Translate("The show checkbox, if checked will display the Item Number in the content's description.",Login_Language,conn)%></LI>
<% end if %>
<BR><BR>
<% if Show_Location = True then %>
  <LI><B><%=Translate("Location",Login_Language,conn)%></B> - <%=Translate("Building, City, State, Country information. (Included with Subscription Service)",Login_Language,conn)%></LI>
<% end if %>
<% if Show_Link = True then %>
  <LI><B><%=Translate("URL to Web Page",Login_Language,conn)%></B> - <%=Translate("If the Content or Event has additional information located on another web page, supply the complete URL.",Login_Language,conn)%></LI>
<% end if %>
<% if Show_Link_PopUp_Disabled = True then %>
  <LI><B><%=Translate("URL to Web Page Pop-Up Window Disable",Login_Language,conn)%></B> - <%=Translate("If disabled, then URL Link to Web Page is direct as opposed to using a separate pop-up browser window.  A Session variable, Session(&quot;BackURL&quot;) can be interogated by the link to obtain the parent link to restore this view when the link application is done.",Login_Language,conn)%></LI>
<% end if %>

<% if Show_Location = True or Show_Link = True or Show_Link_PopUp_Disabled = True then %>
  <BR><BR>
<% end if %>    

<% if Show_File = True then %>
  <LI><B><%=Translate("Asset File - (LOW Resolution)",Login_Language,conn)%></B> - <%=Translate("Low Resolution Asset File used for web view / download / email, eFulfillment and the Subscription Service. An Oracle / MfgPro Item Number is required in field Item Reference #1 if this asset is used for eFulfillment, POD or orderable through the Shopping Cart. Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%></LI>
  <% if Show_Item_Number = True then %>
    <LI><B><%=Translate("Available to Electronic Fulfillment",Login_Language,conn) & " - (" & Translate("End-User Viewable",Login_Language,conn)%>)</B> - <%=Translate("Enables Item Number to be available for the US/Intercon Electronic File Fulfillment (EFF) Sytem.",Login_Language,conn)%></LI>
  <% end if %>
  <BR><BR>  
<% end if %>

<% if Show_File_POD = True then %>
  <LI><B><%=Translate("Asset File - (POD Resolution)",Login_Language,conn)%></B> - <%=Translate("Medium Resolution Print on Demand Asset File.  This field supplies a link to this asset used by the Everett POD process.  An Oracle / MfgPro Item Number is required in field Item Reference #1.  Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%></LI>
  <% if Show_Item_Number = True then %>
    <LI><B><%=Translate("Available to Literature Order Shopping Cart",Login_Language,conn) & " - "%></B><%=Translate("Enables Item Number to be added by the user&acute;s shopping cart for US/Intercon Print on Demand (POD) system for physical media fulfillment (GAC).",Login_Language,conn)%></LI>
  <% end if %>  
  <BR><BR>
<% end if %>

<% if Show_Include = True then %>
  <LI><B><%=Translate("Content File",Login_Language,conn)%></B> - <%=Translate("Copies the contents of an external HTML file that is included as additional content information for this record  Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%></LI>
  <BR><BR>
<% end if %>

<% if Show_Thumbnail = True then %>
  <LI><B><%=Translate("Thumbnail Image File",Login_Language,conn)%></B> - <%=Translate("Adds an image to this record  Use the [Browse] button to locate the file on your local drive.  The file you selected will be uploaded to this server, once you have clicked on the [Save / Update] button below.  At a later time, if you wish to unattach this file from this record, click on the checkbox to the right of the file name.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Request Thumbnail",Login_Language,conn)%></B> - <%=Translate("If you do not have the ability to create your own thumbnail image file for this asset, check this checkbox to have a thumbnail image created.",Login_Language,conn)%></LI>
<% end if %>
<BR><BR>
<LI><B><%=Translate("Pre-Announce",Login_Language,conn)%></B> - <%=Translate("Number of days prior to the Beginning Date to display this Content or Event. Default=0 - No effect on Beginning Date.",Login_Language,conn)%></LI>
<LI><B><%=Translate("Beginning Date",Login_Language,conn)%></B> - <%=Translate("Actual Beginning Date of Event or Channel Notification Date of the content.",Login_Language,conn)%>&nbsp;(<%=Translate("Included with Subscription Service",Login_Language,conn)%>)&nbsp;<%=Translate("If you need a calendar, click on the calendar icon.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
<LI><B><%=Translate("Ending Date",Login_Language,conn)%></B> - <%=Translate("Actual Ending Date of Event or Expiration Date of the content.",Login_Language,conn)%>&nbsp;(<%=Translate("Included with Subscription Service",Login_Language,conn)%>)&nbsp; <%=Translate("If you need a calendar, click on the calendar icon.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
<LI><B><%=Translate("Move to Archive",Login_Language,conn)%></B> - <%=Translate("Number of days after the Ending Date to display the Content or Event.  Default=0, No effect on Ending Date.",Login_Language,conn)%></LI>
<LI><B><%=Translate("Public Release Date",Login_Language,conn)%></B> - <%=Translate("The date that this information can be released to the public.",Login_Language,conn)%>&nbsp;<%=Translate("A Public Release Date Notice will appear in the Description section of the Content Item or Event.",Login_Language,conn)%>&nbsp;<%=Translate("Leave this date blank if there is not a Public Release Date restriction.",Login_Language,conn)%>&nbsp;<IMG ALIGN=TOP BORDER=0 HEIGHT=21 SRC="/images/calendar/calendar_icon.gif" STYLE="POSITION: relative" WIDTH=34></LI>
<LI><B><%=Translate("Mark as Confidential",Login_Language,conn)%></B> - <%=Translate("The caption &quot;Confidential - Not for Public Release&quot; will appear in the Description section of the Content Item or Event.",Login_Language,conn)%></LI>
<BR><BR>
<% if Show_Subscription = True then %>
  <LI><B><%=Translate("Send Notice via Subscription Service",Login_Language,conn)%></B> - <%=Translate("Sends a customized email memo to the user containing Title, Product/Series, Date, Description and link of the record to the Channel Group enabled, whose User Profiles have Subscription Service enabled.  The information is sent on the Beginning Date or Pre-Announce Date if specified.",Login_Language,conn)%></LI>
  <BR><BR>
<% end if %>
<% if Show_Calendar = True then %>
  <LI><B><%=Translate("Calendar",Login_Language,conn)%></B> - <%=Translate("Shows this Content or Event on the Site calendar.",Login_Language,conn)%></LI>
  <BR><BR>  
<% end if %>  
<LI><B><%=Translate("Groups",Login_Language,conn)%></B> - <%=Translate("Select each Channel Category that is allowed to view this Content or Event. For new Content or Event additions, pre-selected (default) group(s) checkboxes are displayed in red.",Login_Language,conn)%></LI>
<% if Admin_Access >= 8 then %>
  <LI><B><%=Translate("Groups - Administrator Accounts",Login_Language,conn)%></B> - <%=Translate("Select each Administrator account that is allowed to view this Content or Event. These selections are only available to the Site Administrator and restrict the content / event item to be viewed only if the user has the selected authorization.  Ensure that standard groups above are un-checked.",Login_Language,conn)%></LI>
<% end if %>
<BR><BR>  
<LI><B><%=Translate("Restrict/Limit to Countries",Login_Language,conn)%></B> - <%=Translate("Select Restrict to or Limit to each Country allowed to view this Content or Event or leave blank if no countries are restricted.  This is a multi-select area.  To select more than one restricted country, hold down the [CTRL] key while selecting with your mouse.",Login_Language,conn)%></LI>
<% if Admin_Access <= 2 then %>
  <BR><BR>
  <LI><B><%=Translate("Select Group to Approve this Submission",Login_Language,conn)%></B> - <%=Translate("All new submissions to this site, require the review and approval of the group responsible for maintenance of this information.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Request Review of this Submission by Email",Login_Language,conn)%></B> - <%=Translate("All submissions will automatically appear in the approval queue of the Content / Event Administrator, however, you may want to inform the Administrator by Email of your submission for date sensitive submittals or other reasons that a review is pending.",Login_Language,conn)%></LI>
<% end if %>
<% if Admin_Access = 4 or Admin_Access >=8 then %>
  <BR><BR>
  <LI><B><%=Translate("Group Assigned to Approve this Submission",Login_Language,conn)%></B> - <%=Translate("As a Content / Event Administrator, you can select yourself as reviewer of this submission (default), or you can re-assign the submission to another group for review, approval and maintenance of this information.  Note: For this submission to appear in the submitter&acute;s or administrator&acute;s queue, the Status flag must be set to &quot;Review&quot;.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Request Review of this Submission by Email",Login_Language,conn)%></B> - <%=Translate("If you have selected another Content / Event Administrator, all submissions will automatically appear in the approval queue of that Administrator, however, you may want to inform the Administrator by Email of your submission for date sensitive submittals or other reasons that a review is pending.",Login_Language,conn)%></LI>
<% end if %>
<% if isnumeric(request("ID")) and Clone = 0 then %>
  <BR><BR>
  <LI><B><%=Translate("Clone",Login_Language,conn)%></B> - <%=Translate("Clones an existing record to new record, however preserves Parent ID Number.  You can only clone from the original parent record and not from a subsequent child record.  Cloning is primarily used for relating multi-language document versions of the same information to the original master or primary record.",Login_Language,conn)%></LI>
  <LI><B><%=Translate("Duplicate",Login_Language,conn)%></B> - <%=Translate("Duplicates an existing record to new record.  You can only Duplicate from the original parent record and not from a subsequent child record.  Duplicating is primarily used for copying similar versions of the same information from the original master or primary record.",Login_Language,conn)%></LI>
<% end if %>  

<BR><BR>
</UL>
<UL>
<LI><B><%=Translate("Special Characters",Login_Language,conn)%></B> - <%=Translate("Certain characters have special meaning in HTML documents. The following entity names are used in HTML, always prefixed by ampersand (&) and followed by a semicolon. They represent particular graphic characters, which have special meanings in places in the markup, or may not be part of the character set available to the writer.",Login_Language,conn)%></LI>
<BR><BR>
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
<BR><BR>
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
</FONT>

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
var FormName = document.<%=FormName%>
var CheckFlg = false;
var CheckSts = false;



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
    CheckMsg  = "Invalid Character(s)in File Name Path\r\n\n";
    CheckMsg += "The file name that you just loaded contains invalid or hidden character(s).\r\n";
    CheckMsg += "Valid Characters are: [A-Z],[a-z],[0-9],[space,underscore,hyphen and period]\r\n\n";
    CheckMsg += "To fix this problem, click [OK] to close this alert message. The illegal file name will be highlighted.\r\n\n";
    CheckMsg += "Delete the file name from the highlighted input box by pressing the [Delete] key.\r\n\n";
    CheckMsg += "Use Windows Explorer to rename the file name at the source directory path.\r\n\n";
    CheckMsg += "Return to this form and use [Browse] to reload the corrected file name.";
    alert(CheckMsg);
    field.focus();
    field.select();
  }
}

function startupload() {

//		var winstyle      = "height=220,width=400,scrollbars=no,status=no,toolbar=no,menubar=no,location=no";
//		var upload_monitor = window.open('/sw-administrator/Calendar_Edit_Upload_Monitor.asp?ProgressID=<%=ProgressID%>','upload_monitor',winstyle);
//    upload_monitor.focus();

    FormName.action = "/sw-administrator/Calendar_Admin.asp"
  	FormName.submit();
}

var highlightcolor="lightyellow"
var ns6=document.getElementById&&!document.all
var previous=''
var eventobj

//Regular expression to highlight only form elements
var intended=/INPUT|TEXTAREA|SELECT|OPTION/

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

//-->
</script>

<%
Call Disconnect_SiteWide
%>