<%response.write "hello world!"

dim strConnectionString 
dim m_lc
dim rssite
dim SQL

strConnectionString = "DRIVER={SQL Server};SERVER=flkprd18.data.ib.fluke.com;UID=marcomweb;DATABASE=fluke_wheretobuy;pwd=!?wwwProd1"



set m_lc = Server.CreateObject("ADODB.Connection")
  	m_lc.ConnectionTimeout = 2
  	m_lc.Open strConnectionString



SQL = "SELECT contacttypeid FROM ContactType"
      
Set rsSite = Server.CreateObject("ADODB.Recordset")

      rsSite.Open SQL, m_lc, 3, 3
      
      if not rsSite.EOF then	
	response.write rssite("contacttypeid") & "<BR>"
	rssite.movenext
      else
	response.write "error"
      end if




%>