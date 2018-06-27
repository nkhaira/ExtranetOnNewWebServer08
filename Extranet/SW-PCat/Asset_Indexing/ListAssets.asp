<html>
	<body>
	  <!--#include virtual="/connections/connection_SiteWide.asp"-->
	    <%

	    'whichFN=server.mappath("/inetpub/assets/tempfile.txt")

		' first, create the file out of thin air
		Set fstemp = server.CreateObject("Scripting.FileSystemObject")
		Set filetemp = fstemp.CreateTextFile("D:\Extranet\SW-PCat\Asset_Indexing\tempfile.xml", true)
		' true  = file can be over-written if it exists
		' false = file CANNOT be over-written if it exists

		filetemp.WriteLine("<?xml version='1.0' ?>")


filetemp.WriteLine("<Assets>")
'filetemp.WriteLine(" <WEBSITE>")
'filetemp.WriteLine("   <URL>15seconds.com</URL>")
'filetemp.WriteLine(" </WEBSITE>")
'filetemp.WriteLine(" <WEBSITE>")
'filetemp.WriteLine("   <URL>internet.com</URL>")
'filetemp.WriteLine(" </WEBSITE>")
'filetemp.WriteLine("</WEBSITES>")

		'filetemp.WriteLine("This is a brand new file!!!!")
		'filetemp.writeblanklines(3)
		'filetemp.WriteLine("This is the last line of the file we created!")
'		filetemp.Close

			'set objadoconn=server.CreateObject("adodb.connection")
			
			Call Connect_SiteWide
			'Put your connection string here I have used UDL.
			'objadoconn.Open "Provider=sqloledb;Data Source=flkprd18;Initial Catalog=Fluke_SiteWide;User Id=MarComWeb;Password=!?wwwProd1"

			set objadorecordset=server.CreateObject("adodb.recordset")

			sqlstr="select distinct '<a href=http://dtmevtvsdv15.danahertm.com/portweb/' + calendar.file_name +'>' + 'http://dtmevtvsdv15.danahertm.com/portweb/' + calendar.file_name + '</a>' as crawl_url,"
			sqlstr=sqlstr & "'<a href=' + 'http://dtmevtvsdv15.danahertm.com/find_it.asp?document='+calendar.item_number + '>'+'http://dtmevtvsdv15.danahertm.com/find_it.asp?document='+calendar.item_number + '</a>' as user_url,calendar.file_name,calendar.item_number "
			sqlstr=sqlstr & "from calendar inner join site on calendar.site_id=site.id where (calendar.status=1) and (calendar.subgroups like '%view%' or calendar.subgroups like '%fed1%') and "
			sqlstr=sqlstr & "(calendar.item_number is not null or calendar.item_number <> '') and (calendar.file_name is not null) and (calendar.site_id=82) and (calendar.Language='eng')"

			'objadorecordset.Open sqlstr,objadoconn,adOpenStatic
			objadorecordset.Open sqlstr,conn,adOpenStatic
			
			do while not(objadorecordset.eof or objadorecordset.BOF)
				Response.Write objadorecordset.Fields(0).Value
				filetemp.WriteLine(" <Asset")
				filetemp.WriteLine(" URI='http://dtmevtvsdv15.danahertm.com/portweb/"+objadorecordset.Fields(2).Value+"'>")
				filetemp.WriteLine(" 	<CrawlURI>")
				filetemp.WriteLine("http://dtmevtvsdv15.danahertm.com/portweb/" + objadorecordset.Fields(2).Value)
				filetemp.WriteLine(" 	</CrawlURI>")
				filetemp.WriteLine(" 	<UserURI>")
				filetemp.WriteLine("http://dtmevtvsdv15.danahertm.com/find_it.asp?document="+objadorecordset.Fields(3).Value)
				filetemp.WriteLine(" 	</UserURI>")
				filetemp.WriteLine(" </Asset>")
				Response.Write "<br>"
				'Response.Write objadorecordset.Fields(1).Value
				'Response.Write "<br>"
				objadorecordset.MoveNext
			loop
			filetemp.WriteLine("</Assets>")
			filetemp.Close
			
			Call Disconnect_SiteWide
		%>
	</body>
</html>


