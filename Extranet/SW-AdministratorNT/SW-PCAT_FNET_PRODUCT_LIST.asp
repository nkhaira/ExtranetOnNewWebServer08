<HTML>
<HEAD>
<TITLE>Category / Sub Category Matrix</TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
</HEAD>
<BODY BGCOLOR="White">
<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<!--#include virtual="/sw-administratorNT/SW-PCAT_FNET_IISSERVER.asp"-->
<%

on error resume next

Site_ID = request("Site_ID")
set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
'Pass the actual url here
call Products.open("POST",striisserverpath,0,0,0)
call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
'Response.Write rs("PID")
'Response.End
''Nitin Code Changes Start
call Products.send("operation=P&assetpid=0" & "&SiteID=" & Site_ID)
''Nitin Code Changes End
strProducts=Products.responseXML.XML
'.Write strProducts
'.End
set objxml=Server.CreateObject("msxml2.domdocument")
call objxml.loadxml(strProducts)
    
if not isblank(Site_ID) then

  response.write "<SPAN CLASS=HEADING4>" & Site_Description & "</SPAN><BR>"
  response.write "<SPAN CLASS=SmallBOLD>Product Catalog - Matrix Listing</SPAN><P>"
  response.write "<SPAN CLASS=Small>The Product Catalog Matrix listing is an aid to ensure that you are<BR>adding content or event items into the correct selected products<BR>that have been pre-determined by your Site Administrator.<BR>If you need to add a new Product, see your Site Administrator.<P>"
  
  Call Table_Begin
  
  response.write "<TABLE BGCOLOR=Gray BORDER=0 CELLPADDING=4>"
  
  response.write "<TR>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Product Id</FONT></TD>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Product</FONT></TD>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Product Id</FONT></TD>"
  response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">Product</FONT></TD>"
  response.write "</TR>"
  set objcol=objxml.selectsingleNode("Info")
  if not(objcol is nothing) then
          set objcolProducts=objcol.firstChild
          set objcol=objcolProducts.firstChild
          for icol=0 to objcol.childnodes.length-1
                set objinfo=objcol.childnodes(icol)
                set objinfo1=objinfo.firstchild
                set objinfo2=objinfo.lastchild
               
                if (icol mod 2) = 0 then
                    'if icol=0 then
                    response.write "<TR>"
                    'end if
                    response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">" & objinfo1.text & "</TD><TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">" & objinfo2.text & "</TD>"
                    if icol=objcol.childnodes.length-1 then
                            response.write "<TD CLASS=SmallBold BGCOLOR=Black>" & "&nbsp;" & "</TD><TD CLASS=SmallBold BGCOLOR=Black>" & "&nbsp;" & "</TD>"
                            Response.Write "</TR>"
                    end if
                    
                    'if icol<>0 then
                    '    Response.Write "</TR>"
                    'end if
                else
                    'if icol<>1 then
                    '    response.write "<TR>"
                    'end if
                    response.write "<TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">" & objinfo1.text & "</TD><TD CLASS=SmallBold BGCOLOR=Black><FONT COLOR=""#FFCC00"">" & objinfo2.text & "</TD>"
                   
                    'if icol=1 then
                        Response.Write "</TR>"
                    'end if
                end if
          next
  else
         ShowProductRetrieveError("Unable to retrive products.")
  end if        
  response.write "</TABLE>"
  Call Table_End
  
  response.write "</BODY>"
  response.write "</HTML>"
end if

if err.number <> 0 then
    ShowProductRetrieveError(err.Description)
end if

sub ShowProductRetrieveError(strmessage)
    BackURL= "/sw-administratorNT/default.asp?Site_ID=" & Site_ID 
	response.write "<HTML>" & vbCrLf
	response.write "<HEAD>" & vbCrLf
	response.write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
	response.write "<TITLE>Error</TITLE>" & vbCrLf
	response.write "</HEAD>"
	response.write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
	response.write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """>" & vbCrLf
	response.write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """>" & vbCrLf
	response.write "<DIV ALIGN=CENTER>"
	Call Nav_Border_Begin
	response.write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
	Response.Write strmessage & "<br><br>"
	Response.Write "Unable to retrive the product records.<br><br>"
	response.write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
	response.write "</TD></TR></TABLE>" & vbCrLf
	Call Nav_Border_End
	response.write "</FORM>" & vbCrLf
	response.write "</DIV>"
	response.write "</BODY>"
	response.write "</HTML>"
	on error goto 0
	Response.End
end sub	
%> 