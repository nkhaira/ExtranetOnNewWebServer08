<%
' ESCROW changes:
' change line 129 to: myenv = "PRD"
Function GetConnectionString(strDatabaseName, strServerName, strLogin)
	Dim aServer
	Dim aFiltered
	Dim strConnectionString
	Dim aServerInfo
	Dim aNodes
	Dim node
	Dim myenv
	Dim chosen
	Dim line
	chosen = ""
	
	strConnectionString = ""
	
	' Data structure: Domain|SQL Server IP|Database Name|User Type|UID|Password
	' User Type can be: (others can be added later)
	'	ADMIN
	'	WEBUSER
	
		' first group is Fluke_Products - webuser
	aServerInfo = Array("MILHOUSE|216.244.65.42|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!dev$",_
		"DEV|EVTIBG03|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!dev$",_
		"STG||FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!staging$",_
		"FLUKE|216.244.76.70|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!live$",_
		"PRD|FLKPRD03|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!live$",_
		_
		"*****|Fluke Products - admin|***********",_
		"MILHOUSE|216.244.65.42|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!dev$",_
		"DEV|EVTIBG03|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!dev$",_
		"STG||FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!staging$",_
		"FLUKE|216.244.76.70|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!live$",_
		"PRD|FLKPRD03|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!live$",_
		_
		"**********|Fluke FormData|**************",_
		"MILHOUSE|216.244.65.42|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!dev$",_
		"DEV|EVTIBG03|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!dev$",_
		"STG||FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!dev$",_
		"FLUKE|216.244.76.70|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!live$",_
		"PRD|FLKPRD03|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!live$",_
		_
		"**********|Fluke Buying (connection_buying.asp)|*********",_
		"MILHOUSE|216.244.65.42|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"STG||FLUKE_BUYING|WEBUSER|BUYING_WEB|",_
		"FLUKE|216.244.76.70|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|FLUKE_BUYING|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Networks Secure Forms|***********",_
		"MILHOUSE|216.244.65.42|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"STG||FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|FLUKE_PROMO|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Calibration Newsletter - webuser|*******",_
		"MILHOUSE|216.244.65.42|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		"DEV|EVTIBG03|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		"STG||FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		"FLUKE|216.244.76.70|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		"PRD|FLKPRD03|FLUKE_CALNEWSLETTER|WEBUSER|CALNEWS_WEB|!huntergatherer5",_
		_
		"**********|Calibration Newsletter - admin|***********",_
		"MILHOUSE|216.244.65.42|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		"DEV|EVTIBG03|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		"STG||FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		"FLUKE|216.244.76.70|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		"PRD|FLKPRD03|FLUKE_CALNEWSLETTER|ADMIN|CALNEWS_ADMIN|calnews5",_
		_
		"**********|Survey|*************",_
		"MILHOUSE|216.244.65.42|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"STG||FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|Meterman|***********",_
		"MILHOUSE|216.244.65.42|METERMAN|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|METERMAN|WEBUSER|METERMAN_USER|mmtt_!dev$",_
		"STG||METERMAN|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|METERMAN|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|METERMAN|WEBUSER|METERMAN_USER|mmtt_!live$",_
		_
		"**********|Meterman admin|***********",_
		"DEV|EVTIBG03|METERMAN|ADMIN|METERMAN_ADMIN|mmtta_!dev$",_
		"PRD|FLKPRD03|METERMAN|ADMIN|METERMAN_ADMIN|mmtta_!live$",_
		_
		"**********|Virtual directory tool|**********",_
		"MILHOUSE|216.244.65.42|FLUKE_VIRTUALDIRECTORIES|ADMIN|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|FLUKE_VIRTUALDIRECTORIES|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"STG||FLUKE_VIRTUALDIRECTORIES|ADMIN|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|FLUKE_VIRTUALDIRECTORIES|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|FLUKE_VIRTUALDIRECTORIES|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|eStore|************",_
		"MILHOUSE|216.244.65.42|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"STG||ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|ESTORE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|SiteWide|**************",_
		"MILHOUSE|216.244.65.42|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		"DEV|EVTIBG03|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		"STG|FLKPRD03|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		"FLUKE|216.244.76.70|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		"PRD|FLKPRD03|FLUKE_SITEWIDE|WEBUSER|SITEWIDE_WEB|tuggy_boy",_
		_
		"**********|Where To Buy (connection_wtb.asp)|***********",_
		"MILHOUSE|216.244.65.42|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"DEV|EVTIBG03|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"STG|EVTIBG03|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"PRD|FLKPRD03|FLUKE_WHERETOBUY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		_
		"**********|W T B (connection_wtb2000.asp)|*********",_
		"MILHOUSE|216.244.65.42|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
		"DEV|EVTIBG03|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
		"STG||WTB|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
		"FLUKE|216.244.76.70|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
		"PRD|FLKPRD03|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
		_
		"**********|TM Store - called from verifysession.asp in the TMStore root|*******",_
		"MILHOUSE|216.244.65.42|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!dev$",_
		"DEV|EVTIBG03|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!dev$",_
		"STG||FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!dev$",_
		"FLUKE|216.244.76.70|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!live$",_
		"PRD|FLKPRD03|FLUKE_TMSTORE|WEBUSER|TMSTORE_WEB|tmstore_!live$")
	
	myenv = "PRD"
	aNodes = split(strServerName,".")
	
	for each node in aNodes
		select case node
			case "PRD","DEV","STG","MILHOUSE"
				myenv = node
				exit for
		end select
	Next
	
	' Filter on our three pieces of information, this should leave one entry in the array
	aFiltered = Filter(aServerInfo, strDatabaseName)
	aFiltered = Filter(aFiltered, strLogin)
	
	for each line in aFiltered
		if Instr(line,myenv) = 1 then
			chosen = line
			exit for
		end if
	next
	'aFiltered = Filter(aFiltered, myenv)
	
	if len(chosen) = 0 then
		Response.write "Failed to create database connection"
	else
		aServer = Split(chosen, "|")
		strConnectionString = "DRIVER={SQL Server};SERVER=" & aServer(1) &_
			";UID=" & aServer(4) &_
			";DATABASE=" & aServer(2) &_
			";pwd=" & aServer(5)
	end if
	
	GetConnectionString = strConnectionString
End Function


%>