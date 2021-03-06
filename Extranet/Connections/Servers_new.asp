<%
Function GetConnectionString(strDatabaseName, strServerName, strLogin)
	Dim aServer
	Dim aFiltered
	Dim strConnectionString
	Dim aServerInfo
	Dim aNodes
	Dim node
	Dim myenv
	
	strConnectionString = ""
	
	' Data structure: Domain|SQL Server IP|Database Name|User Type|UID|Password
	' User Type can be: (others can be added later)
	'	ADMIN
	'	WEBUSER
	'   REPORTING
	
		' first group is Fluke_Products - webuser
	aServerInfo = Array("DEV|EVTIBG18|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!live$",_
		_
		"*****|Fluke Products - admin|***********",_
		"DEV|EVTIBG18|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!live$",_
		_
		"**********|Fluke FormData|**************",_
		"DEV|EVTIBG18.tc.fluke.com|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!live$",_
		_
		"**********|Fluke FormData Reporting(used for running oracle extracts)|**************",_
		"DEV|EVTIBG18|FLUKE_FORMDATA|REPORTING|FORMDATA_REPORTING|form_!reporting$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_FORMDATA|REPORTING|FORMDATA_REPORTING|form_!reporting$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_FORMDATA|REPORTING|FORMDATA_REPORTING|form_!reporting$",_
		_
		"**********|Fluke Buying (connection_buying.asp)|*********",_
	 	"DEV|EVTIBG18|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Networks Secure Forms|***********",_
		"DEV|EVTIBG18|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Calibration Newsletter - webuser|*******",_
		"DEV|EVTIBG18|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		_
		"**********|Calibration Newsletter - admin|***********",_
		"DEV|EVTIBG18|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		_
		"**********|Survey|*************",_
		"DEV|EVTIBG18|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Meterman|***********",_
		"DEV|EVTIBG18|METERMAN|WEBUSER|METERMAN_USER|mmtt_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|METERMAN|WEBUSER|METERMAN_USER|mmtt_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|METERMAN|WEBUSER|METERMAN_USER|mmtt_!live$",_
		_
		"**********|Meterman admin|***********",_
		"DEV|EVTIBG18|METERMAN|ADMIN|METERMAN_ADMIN|mmtta_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|METERMAN|ADMIN|METERMAN_ADMIN|mmtta_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|METERMAN|ADMIN|METERMAN_ADMIN|mmtta_!live$",_
		_
		"**********|Virtual directory tool|**********",_
		"DEV|EVTIBG18|FLUKE_VIRTUALDIRECTORIES|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_VIRTUALDIRECTORIES|ADMIN|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_VIRTUALDIRECTORIES|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|eStore|************",_
		"DEV|EVTIBG18.tc.fluke.com|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|SiteWide|**************",_
		"DEV|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		_
		"**********|Where To Buy (connection_wtb.asp)|***********",_
		"DEV|EVTIBG18|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|W T B (connection_wtb2000.asp)|*********",_
		"DEV|EVTIBG18|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|WTB|WEBUSER|FLUKEWEBUSER|wtb2000",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
		_
		"**********|Fluke UserData|**************",_
		"DEV|EVTIBG18|FLUKE_USERDATA|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_USERDATA|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_USERDATA|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Product Catalog (use connection_sql.asp)|***********",_
		"DEV|EVTIBG18|PRODUCTCATALOG|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|PRODUCTCATALOG|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|PRODUCTCATALOG|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"*****|Product Catalog - admin|***********",_
		"DEV|EVTIBG18|PRODUCTCATALOG|ADMIN|PRODUCTS_ADMIN|products_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|PRODUCTCATALOG|ADMIN|PRODUCTS_ADMIN|products_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|PRODUCTCATALOG|ADMIN|PRODUCTS_ADMIN|products_!live$",_
		_
		"**********|BRAZIL|***********",_
		"DEV|EVTIBG18|BRAZILWEB|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|BRAZILWEB|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|BRAZILWEB|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|TM Store - called from verifysession.asp in the TMStore root|*******",_
		"DEV|EVTIBG18|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!dev$",_
		"TEST|Us-dat-sql02-t.data.ib.fluke.com|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!live$",_
		"PRD|Us-dat-sql02-p.data.ib.fluke.com|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!live$")
	
	myenv = "PRD"
	aNodes = split(strServerName,".")
	
  if instr(strServerName,"DTMEVTVSDV15") > 0 then
    myenv = "DEV"
  elseif instr(strServerName,"DTMEVTVSDV18") > 0 then
    myenv = "TEST"
  else
  	for each node in aNodes
  		select case node
  			case "DEV", "TST", "TEST", "PRD"
  				myenv = node
  				exit for
  		end select
  	next
  end if
	
	' Filter on our three pieces of information, this should leave one entry in the array
	aFiltered = Filter(aServerInfo, strDatabaseName)
	'response.write "strDatabasename: " & strdatabasename & "<BR>"
	aFiltered = Filter(aFiltered, strLogin)
	'response.write "strlogin: " & strlogin & "<BR>"
	aFiltered = Filter(aFiltered, myenv)
	'response.write "myenv: " & myenv & "<BR>"
	
	if ubound(aFiltered) > 0 then
		'response.write aFiltered(0)&"<BR>"
		'response.write aFiltered(16)&"<BR>"
		'response.write UBOUND(aFiltered)&"<BR>"
		Response.write "Failed to create the database string"
	else
		aServer = Split(aFiltered(0), "|")
		strConnectionString = "DRIVER={SQL Server};SERVER=" & aServer(1) &_
			";UID=" & aServer(4) &_
			";DATABASE=" & aServer(2) &_
			";pwd=" & aServer(5)
	end if
	
	GetConnectionString = strConnectionString
End Function


%>