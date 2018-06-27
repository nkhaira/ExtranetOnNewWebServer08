<%

<!--#include virtual="/connections/Connection_Products.asp"-->

ConnectProducts

SQL = "SELECT prd_manual.obsolete, man_ver.*, man_type.code ManualType_Code, man_media.code ManualMedia_Code "
SQL = SQL & "from prd_manual inner join prd_manual_ver man_ver on prd_manual.manual_id = man_ver.manual_id "
SQL = SQL & "inner join prd_ManualType man_type on man_ver.ManualType_ID = man_type.ManualType_ID "
SQL = SQL & "inner join prd_ManualMediaTypes man_media on	man_ver.ManualMediaType_ID = man_media.ManualMediaType_ID	"
SQL = SQL & "where man_ver.Manual_Ver_ID = @Manual_Ver_ID"

Set rsManuals = Server.CreateObject("ADODB.Recordset")
rsManuals.Open SQL, conn, 3, 3

do while not rsManuals.EOR

response.


DisconnectProducts


%>
