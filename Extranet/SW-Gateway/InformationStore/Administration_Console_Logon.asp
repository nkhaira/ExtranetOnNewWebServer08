<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_Translate.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->

<%

Call Connect_Sitewide

Image_Locator = ""
if not isblank(request("Locator")) then
  if isnumeric(request("Locator")) then
    Image_Locator_Type = 0
    Image_Locator = request("Locator")
  elseif LCase(request("Locator")) = "search" then  
    Image_Locator_Type = 2
    Image_Locator = ""
  elseif LCase(request("Locator")) = "new" then  
    Image_Locator_Type = 3
    Image_Locator = ""
  elseif Mid(LCase(request("Locator")),1,5) = "edit:" then  
    Image_Locator_Type = 4
    Image_Locator = Mid(request("Locator"),6)
  else
    Image_Locator_Type = 1
    Image_Locator = request("Locator")
  end if  
else
    Image_Locator_Type = 2
    Image_Locator = ""
end if

if not isblank(request("Site_ID")) then
  Site_ID = request("Site_ID")
elseif not isblank(session("Site_ID")) then
  Site_ID = session("Site_ID")
else
  Site_ID = 0
end if      

SQL =  "SELECT UserData.* FROM UserData WHERE UserData.Site_ID=" & Site_ID & " AND UserData.NTLogin='" & Session("Logon_User") & "'"
Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.Open SQL, conn, 3, 3

if not rsUser.EOF then		

  response.write "<HTML>" & vbCrLf
  response.write "<HEAD>" & vbCrLf
  response.write "<TITLE>InformationStore Gateway</TITLE>" & vbCrLf
  response.write "</HEAD>" & vbCrLf
  response.write "<BODY BGCOLOR=""White"" onLoad='document.forms[0].submit()'>" & vbCrLf
  response.write "<FORM METHOD=""POST"" ACTION=""http://www.informationstore.net/fluke/extranet/default.asp"" ID=FORM1 NAME=FORM1>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=Account_ID           VALUE="""    & rsUser("ID") & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=Company              VALUE="""    & rsUser("Company") & """>" & vbCrLf

  if isblank(rsUser("Shipping_Address")) then
    Use_Business = True
  else
    Use_Business = False
  end if    

  select case Use_Business
    case False
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_Address     VALUE=""" & rsUser("Shipping_Address") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_City        VALUE=""" & rsUser("Shipping_City") & """>" & vbCrLf
      if isblank(rsUser("Shipping_State")) or rsUser("Shipping_State") = "ZZ" then
        response.write "<INPUT TYPE=HIDDEN NAME=Shipping_State     VALUE=""" & rsUser("Shipping_State_Other") & """>" & vbCrLf    
      else  
        response.write "<INPUT TYPE=HIDDEN NAME=Shipping_State     VALUE=""" & rsUser("Shipping_State") & """>" & vbCrLf
      end if
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_Postal_Code VALUE=""" & rsUser("Shipping_Postal_Code") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_Country     VALUE=""" & rsUser("Shipping_Country") & """>" & vbCrLf
    case True
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_Address     VALUE=""" & rsUser("Business_Address") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_City        VALUE=""" & rsUser("Business_City") & """>" & vbCrLf
      if isblank(rsUser("Business_State")) or rsUser("Business_State") = "ZZ" then
        response.write "<INPUT TYPE=HIDDEN NAME=Shipping_State     VALUE=""" & rsUser("Business_State_Other") & """>" & vbCrLf
      else  
        response.write "<INPUT TYPE=HIDDEN NAME=Shipping_State     VALUE=""" & rsUser("Business_State") & """>" & vbCrLf
      end if  
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_Postal_Code VALUE=""" & rsUser("Business_Postal_Code") & """>" & vbCrLf
      response.write "<INPUT TYPE=HIDDEN NAME=Shipping_Country     VALUE=""" & rsUser("Business_Country") & """>" & vbCrLf
  end select

  response.write "<INPUT TYPE=HIDDEN NAME=Business_Phone           VALUE=""" & rsUser("Business_Phone") & """>" & vbCrLf
  
  response.write "<INPUT TYPE=HIDDEN NAME=FirstName                VALUE=""" & rsUser("FirstName") & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=MiddleName               VALUE=""" & rsUser("MiddleName") & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=LastName                 VALUE=""" & rsUser("LastName")  & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=Language                 VALUE=""" & rsUser("Language") & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=Email                    VALUE=""" & rsUser("Email")     & """>" & vbCrLf

  response.write "<INPUT TYPE=HIDDEN NAME=NTLogin                  VALUE=""" & rsUser("NTLogin")   & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=Password                 VALUE=""" & rsUser("Password") & """>" & vbCrLf

  response.write "<INPUT TYPE=HIDDEN NAME=Site_ID                  VALUE=""" & rsUser("Site_ID") & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=Groups                   VALUE=""" & Replace(rsUser("Groups")," ","") & """>" & vbCrLf
  response.write "<INPUT TYPE=HIDDEN NAME=SubGroups                VALUE=""" & rsUser("SubGroups") & """>" & vbCrLf

  if Session("BackURL")="" or isnull(Session("BackURL")) or CInt(Request("KillBackURL")) = CInt(True) then
    BackURL = "" '"http://support.fluke.com/register/login.asp?Site_ID=" & Site_ID
  else
    BackURL = Session("BackURL")  
  end if  
  
  response.write "<INPUT TYPE=HIDDEN NAME=BackURL                  VALUE=""" & BackURL & """>" & vbCrLf

  select case Image_Locator_Type
    case 0, 1, 2  ' Individual, Collection or Search View
      response.write "<INPUT TYPE=HIDDEN NAME=Image_Locator        VALUE=""" & Image_Locator & """>" & vbCrLf
    case 3        ' New Collection
      response.write "<INPUT TYPE=HIDDEN NAME=CampaignHash         VALUE=""NEW"">" & vbCrLf
    case 4        ' Edit Collection
      response.write "<INPUT TYPE=HIDDEN NAME=CampaignHash         VALUE=""" & Image_Locator & """>" & vbCrLf
  end select    
  
  Site_Description = "Image Store"
  SQL = "SELECT Site.* FROM Site WHERE ID=" & Session("Site_ID")
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3

  if not rsSite.EOF then
    Site_Description = Translate(rsSite("Site_Description"),Alt_Language,conn)
    select case Image_Locator_Type
      case 0
        Site_Description = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>View Image</SPAN>"
      case 1
        Site_Description = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>View Image Collection</SPAN>"
      case 3
        Site_Description = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>Create New Image Collection</SPAN>"
      case 4
        Site_Description = Site_Description & "<BR><SPAN CLASS=SmallBoldGold>Add / Edit Image Collection</SPAN>"
    end select     
  end if

  rsSite.close
  set rsSite = nothing

  response.write "<INPUT TYPE=HIDDEN NAME=Site_Description         VALUE=""" & Site_Description & """>" & vbCrLf
  response.write "</FORM>" & vbCrLf
  response.write "</BODY>" & vbCrLf
  response.write "</HTML>" & vbCrLf
 
else

  response.write "Invalid User Gateway Parameters for the InformationStore."
  
end if

rsUser.close
set rsUser = nothing
Call Disconnect_Sitewide   
%>

