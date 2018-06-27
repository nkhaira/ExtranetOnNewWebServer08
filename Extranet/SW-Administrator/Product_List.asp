<HTML>
<HEAD>
<TITLE>Category / Sub Category Matrix</TITLE>
<LINK REL=STYLESHEET HREF="/SW-Common/SW-Style.css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso8859-1">
</HEAD>

<BODY BGCOLOR="White">

<!--#include virtual="/include/functions_String.asp"-->
<!--#include virtual="/include/functions_table_border.asp"-->
<%

Site_ID = request("Site_ID")
set Products=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
'Pass the actual url here
call Products.open("POST","http://dtmevtvsdv07:8080/ExtranetPcat/PcatHttpHandler.aspx",0,0,0)
call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
'Response.Write rs("PID")
'Response.End
call Products.send("operation=P&assetpid=0")
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
  end if        
  response.write "</TABLE>"
  Call Table_End
  
  response.write "</BODY>"
  response.write "</HTML>"

else

  response.redirect 

end if
  


%> 