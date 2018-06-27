<%@Language="VBScript" Codepage=65001%>
<% 
on error resume next
set obj = Server.CreateObject("SW_Read_File.SW_Binary_Read")
Response.Write(obj.Returnstring())
if err.number <> 0 then
    Response.Write err.Description
end if

%>
