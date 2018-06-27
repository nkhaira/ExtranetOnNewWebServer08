<%
Class ConnectionManager
  Private m_lc  ' AS ADODB.Connection

  Private Sub class_initialize()
    set m_lc = nothing
  end sub

  Private Sub class_terminate()
    set m_lc = nothing
  end sub

'===================================================================================================
  Public Function DB_ConnectDefault(strDbName)
  	Dim strUser, strServer

  	strUser = "WEBUSER"
  	strServer = uCase(Request.ServerVariables("SERVER_NAME"))

  	set DB_ConnectDefault = DB_Connect(strDbName, strUser, strServer)
  End Function

'===================================================================================================
  Public Function DB_Connect(strDbName, strLogin, strServerName)
  	Dim strConnectionString

  	strConnectionString = GetConnectionString(strDbName, strServerName, strLogin)

  	set m_lc = Server.CreateObject("ADODB.Connection")
  	m_lc.ConnectionTimeout = 2
  	m_lc.Open strConnectionString

    set DB_Connect = m_lc
  End Function

'===================================================================================================
  Public Sub DB_disconnect(dbconn)
  	dbconn.close
  	set dbconn = nothing
  End sub

'===================================================================================================
  Private Function GetConnectionString(strDatabaseName, strServerName, strLogin)
    Dim aServer
    Dim aFiltered
    Dim strConnectionString
    Dim aServerInfo
    Dim aNodes
    Dim node
    Dim myenv

'response.write "strdatabasename: " & strdatabasename & "<BR>"
'response.write "strservername: " & strservername & "<BR>"
'response.write "strlogin: " & strlogin & "<BR>"
'response.end

    strConnectionString = ""

    ' Data structure: Environment|SQL Server/IP|Database Name|User Type|UID|Password
    ' User Type is an alias; WEBUSER & ADMIN are the only currently used ones

    aServerInfo = Array(_
      "DEV|EVTIBG03|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!dev$",_
      "STG|FLKSTG03|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!staging$",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|FLUKE_PRODUCTS|WEBUSER|PRODUCTS_WEB|products_!live$",_
      _
      "*****|Fluke Products - admin|***********",_
      "DEV|EVTIBG03|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!dev$",_
      "STG|FLKSTG03|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!staging$",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|FLUKE_PRODUCTS|ADMIN|PRODUCTS_ADMIN|products_!live$",_
      _
      "**********|Fluke FormData|**************",_
      "DEV|EVTIBG03|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!dev$",_
      "STG|FLKSTG03|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!staging$",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|FLUKE_FORMDATA|WEBUSER|FORMDATA_WEB|formdata_!live$",_
			_
			"**********|Survey|*************",_
			"DEV|EVTIBG03|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
			"STG|FLKSTG03|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
			"PRD|FLKPRD18.DATA.IB.fluke.com|FLUKE_SURVEY|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
      _
      "**********|W T B (connection_wtb2000.asp)|*********",_
      "DEV|EVTIBG03|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
      "STG|FLKSTG03|WTB|WEBUSER|FLUKEWEBUSER|wtb2000",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|WTB|WEBUSER|WTB_ADMIN|wtb2000",_
      _
      "**********|Product Catalog (use connection_sql.asp)|***********",_
      "DEV|EVTIBG18.dev.ib.fluke.com|PRODUCTCATALOG|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
      "STG|FLKTST18.data.ib.fluke.com|PRODUCTCATALOG|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|PRODUCTCATALOG|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
      _
      "*****|Product Catalog - admin|***********",_
      "DEV|EVTIBG18.dev.ib.fluke.com|PRODUCTCATALOG|ADMIN|PRODUCTS_ADMIN|products_!dev$",_
      "STG|FLKTST18.data.ib.fluke.com|PRODUCTCATALOG|ADMIN|PRODUCTS_ADMIN|products_!staging$",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|PRODUCTCATALOG|ADMIN|PRODUCTS_ADMIN|products_!live$",_
      _
      "*****|Fluke Commerce |***********",_
      "DEV|EVTIBG03|FLUKEWEB_COMMERCE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
      "STG|FLKSTG03|FLUKEWEB_COMMERCE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|FLUKEWEB_COMMERCE|WEBUSER|FLUKEWEBUSER|57OrcaSQL",_
 
      _
      "*****|Global Forms |***********",_
      "DEV|EVTIBG18|GLOBAL_FORMDATA|WEBUSER|formdata_web|formdata_!dev$",_
      "PRD|FLKPRD18.DATA.IB.fluke.com|GLOBAL_FORMDATA|WEBUSER|formdata_web|formdata_!live$"_
    )

    myenv = "DEV"
    aNodes = split(strServerName,".")

    for each node in aNodes
      select case node
        case "EVTIBG08", "DTMEVTVSDV15"
          myenv = "DEV"
          exit for
        case "DTMEVTVSDV18"
          myenv = "TEST"
          exit for

        case "PRD", "DEV"
          myenv = node
          exit for
      end select
    Next
    
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
End Class
%>