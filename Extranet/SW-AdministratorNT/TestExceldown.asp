<%
file_being_created= "upload/Test.xls" 
set fso = createobject("scripting.filesystemobject") 
Set act = fso.CreateTextFile(server.mappath(file_being_created), true) 

act.WriteLine("<TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>")
act.WriteLine("<tr><th colspan=7>Asset Activity Detail Report for Asset ID: " & AssetID & "</th></tr>")
act.WriteLine("<TR>")
act.WriteLine("<TD><font size=2 face=""Arial""><b>"&  id & "</b></font></TD>")
act.WriteLine("</tr>")
act.WriteLine("</table>")
act.close 
%>