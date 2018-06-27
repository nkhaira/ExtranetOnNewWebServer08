<%

'''http://author.dev.fluke.com/NR/exeres/6CDA4C46-8734-4D59-929A-C2D2FC92BFFD.htm?wbc_purpose=Basic&NRMODE=Unpublished?wbc_purpose=Basic&NRMODE=Update&FRAMELESS=true&NRNODEGUID=6CDA4C46-8734-4D59-929A-C2D2FC92BFFD
'****************************************************
'	_flukeCreateRichProducts.asp
'
'	generates a new rich product pages for the specified
'	catalog and locale
'
'	NOTE: THIS USES A CUSTOM CONTENT CONNECTOR APP INCLUDE
'
'****************************************************

''''''''''''''''''''''''''' Start '''''''''''''''''''''

'on error resume next

'-----------------------------------------------------------------------------------------------------------
Dim m_lang,m_pid,CurPageType,sql
Dim conn,strConn,rspRODUCTS,m_dbconn
server.scripttimeout = 3000
'strConnPFNET= "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PCAT2;Data Source=dtmevtsvdb02"
strConnPFNET= "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PCAT2;Data Source=tcp:129.196.128.186,1433"

set conn = server.CreateObject("adodb.connection")
conn.Open strConnPFNET

'strConnPIG= "Provider=SQLOLEDB;Persist Security Info=False; Data Source = evtibg18.tc.fluke.com; Initial Catalog = PRODUCTCATALOG; User Id = marcomweb; Password=!?wwwProd1"
strConnPIG= "Provider=SQLOLEDB.1;Initial Catalog = PRODUCTCATALOG;User Id = marcomweb; Password=!?wwwProd1;Data Source=tcp:129.196.132.91,1433"
set oConn = server.CreateObject("adodb.connection")
oConn.Open strConnPIG


'strConnPIGProd= "Provider=SQLOLEDB;Persist Security Info=False; Data Source = flkprd18.data.ib.fluke.com; Initial Catalog = PRODUCTCATALOG; User Id = marcomweb; Password=!?wwwProd1"
strConnPIGProd=  "Provider=SQLOLEDB.1;Initial Catalog = PRODUCTCATALOG;User Id = marcomweb; Password=!?wwwProd1;Data Source=tcp:129.196.231.137,1433"
set oConnProd = server.CreateObject("adodb.connection")
oConnProd.Open strConnPIGProd

set m_dbconn =oConnProd 
m_dbconn.cursorlocation = 2
'set m_dbconn = oConn
'Set oConn =  objConnMan.db_ConnectDefault("PRODUCTCATALOG")

sSQL_Get = "select IG_PID,Pcat_PID from Temp_Products"
set rspRODUCTS = conn.Execute(sSQL_Get)
if err.number <> 0 then
conn.close
conn.open
set rspRODUCTS = conn.Execute(sSQL_Get)
err.clear
end if
do while not (rspRODUCTS.eof)
	set rsLanguages=  oConn.execute("select distinct iso2 from mcs_language_locale")
	do while not (rsLanguages.eof)
	m_lang=rsLanguages.fields(0).value
	m_pid=rspRODUCTS.fields("IG_PID").value
sqlinsert=""
	sqlinsert= "insert into ProductContent values(" & rspRODUCTS.fields("IG_PID").value & ",'" & m_lang & "','" & replace(RenderHTML(1),"'","''") & "',1)"
	oConn.execute sqlinsert
if err.number <> 0 then
oconn.close
oconn.open
oConn.execute sqlinsert
err.clear
end if
	'response.write sqlinsert
sqlinsert=""
	sqlinsert= "insert into ProductContent values(" & rspRODUCTS.fields("IG_PID").value & ",'" & m_lang & "','" & replace(RenderHTML(2),"'","''") & "',2)"
	oConn.execute sqlinsert
if err.number <> 0 then
oconn.close
oconn.open
oConn.execute sqlinsert
err.clear
end if
	'response.write sqlinsert
sqlinsert=""	
	sqlinsert= "insert into ProductContent values(" & rspRODUCTS.fields("IG_PID").value & ",'" & m_lang & "','" & replace(RenderHTML(4),"'","''") & "',4)"
	oConn.execute sqlinsert
if err.number <> 0 then
oconn.close
oconn.open
oConn.execute sqlinsert
err.clear
end if
	'response.write sqlinsert
	'response.write m_pid & "<BR>"
	rsLanguages.movenext
        loop
	rspRODUCTS.movenext
loop

'm_lang="en"
'm_pid=38155
'response.write RenderHTML(1)
response.end
function RenderHTML(CurPageType)
   	Dim cmd,rs, strOut
   	    
  	set cmd = CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_GetPHIDs"
  	' create all the parameters we want
		' response.write "FOO = " & m_pid	
  	cmd.Parameters.Append cmd.CreateParameter("@PID",3,1,,m_pid)
  	cmd.Parameters.Append cmd.CreateParameter("@PageType",3,1,,CurPageType)
  	set rs = cmd.execute
if err.number <> 0 then
m_dbconn.close
m_dbconn.open
set rs = cmd.execute
err.clear
end if
  	if rs.EOF then
  		RenderHTML = ""
  	else
                strOut = ""
  		do until rs.EOF
			'RESPONSE.WRITE rs("PHID")&"-"&rs("A_html")&"-"&rs("PlaceHolderType")&"-"&rs("Class")&"-"&rs("URL")&"-"&rs("TextID")&"-"&rs("AltTextID")
  			' this is the critical subroutine - can be called from various places
  			strOut = strOut & Write_PHID(rs("PHID"),rs("A_html"),rs("PlaceHolderType"), _
          		rs("Class"),rs("URL"),rs("TextID"),rs("AltTextID"))
  			rs.MoveNext
  		loop
      RenderHTML = strOut
  	end if
  	set rs = nothing
  	set cmd = nothing
end function
  
function Write_cPHID(phid,ph_html,ph_class)
  	Dim cmd,rs, strTemp, strEnd
    
    strEnd = ""
    
  	set cmd = CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_cPHID_Get"
  	' create all the parameters we want
  	cmd.Parameters.Append cmd.CreateParameter("@cPHID",3,1,,phid)
  	set rs = cmd.execute
if err.number <> 0 then
m_dbconn.close
m_dbconn.open
set rs = cmd.execute
err.clear
end if
  	if rs.EOF then
  		Write_cPHID = ""
  	else  	 
			strTemp = ph_html
			if len(ph_class&"") > 1 then
				strTemp = strTemp & "<span class=""" & ph_class & """>"
				strEnd = "</span>"
			end if
		        
  		do until rs.EOF
  			strTemp = strTemp & Write_PHID(rs("PHID"),rs("A_html"),rs("phtype"), _
					rs("Class"),rs("URL"),rs("TextID"),rs("AltTextID"))
  			rs.MoveNext
  		loop
	  	
  		Write_cPHID = strTemp & strEnd		  
  	end if
  	
  	set rs = nothing
  	set cmd = nothing
  end function

   function Write_PHID(phid,ph_html,ph_type,ph_class,ph_url,ph_Tid,ph_ATid)
  	' note that as we write each PlaceHolderType subroutine we need only
  	' pass in those attributes meaningful to that subroutine (placeholdertype)
    ' some placeholders naturally don't appear here (Embedded ones)
    ' PlaceHolderType 5 (Trailing Image) is deprecated
  	select case Cint(ph_type)
  		case 1 ' MediumHeadline
  			Write_PHID = Write_HText(ph_html,ph_class,ph_Tid)
  		case 2, 15 ' Paragraph & AnchoredPara
  			Write_PHID = Write_GText(phid,ph_type,ph_html,ph_class,ph_Tid)
  		case 3,13,14 ' All Lists
  			Write_PHID = Write_UL(phid,ph_type,ph_html,ph_class,ph_Tid)
  		case 4,10 ' Image & TranslatedImage
  			Write_PHID = Write_Img(phid,ph_html,ph_type,ph_url,ph_Tid,ph_ATid)
      case 6,11,16,20,21 ' PlainLink, TranslatedLink, BlankLink, PlacedLink, & PlacedBlankLink
        Write_PHID = Write_Link(ph_type,ph_html,ph_class,ph_url,ph_Tid,ph_ATid)
      case 9 ' SpecTable
        Write_PHID = Write_Spec(phid,ph_html,ph_class,ph_url,ph_Tid)
      case 12 ' GeneralTable
        Write_PHID = Write_Table(phid,ph_html,ph_class,ph_url,ph_Tid)
      case 19 ' CompoundPHID
        Write_PHID = Write_cPHID(phid,ph_html,ph_class)
  	end select
  end function

   function Write_HText(ph_html,ph_class,ph_Tid)
  	Dim stfnt,endfnt
    if ph_class = "" then
  		stfnt = "<BR>"
  		endfnt = ""
  	else
  		stfnt = "<BR><span class=""" & ph_class & """>"
  		endfnt = "</span>"
  	end if
  	Write_HText = ph_html & stfnt & GetPHText(ph_Tid,True) & endfnt & vbcrlf 
  end function

   function Write_GText(phid,ph_type,ph_html,ph_class,ph_Tid)
  	Dim stfnt,endfnt
    if ph_class = "" then
  		stfnt = ""
  		endfnt = ""
  	else
  		stfnt = "<span class=""" & ph_class & """>"
  		endfnt = "</span>"
  	end if
  
    if cint(ph_type) = 15 then
      stfnt = "<A name=""para" & phid & """>" & stfnt
      endfnt = endfnt & "</a>"
    end if
    
  	Write_GText = ph_html & stfnt & GetPHText(ph_Tid,True) & endfnt & vbcrlf  	   
  end function

   function Write_UL(phid,ph_type,ph_html,ph_class,ph_Tid)
  	Dim cmd,rs,strTemp, listType, strTbleStart, strTblEnd, strLiType
  	Dim iID,strLangs
    
    select case cint(ph_type)
			case 3,14 ' Unordered
  			listType = "UL"
  			strTbleStart = "<TABLE><TR><TD>"
  			strTblEnd = "</td></tr></table>"
  			strLiType = """"
  		case 13 ' Ordered
  			listType = "OL"
  			strTbleStart = "<TABLE><TR><TD>"
  			strTblEnd = "</td></tr></table>"
  			strLiType = """"
  		case 18 ' EmbeddedUnordered
  			listType = "UL"
  			strTbleStart = ""
  			strTblEnd = ""
  			strLiType = """ type=""disc"""
		end select
    
  	set cmd = CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_GetListItems"
  	' create all the parameters we want
  	cmd.Parameters.Append cmd.CreateParameter("@PHID",3,1,,phid)
  	cmd.Parameters.Append cmd.CreateParameter("@Lang",200,1,2,m_lang)
  	set rs = cmd.execute
  if err.number <> 0 then
m_dbconn.close
m_dbconn.open
set rs = cmd.execute
err.clear
end if
    ' for some strange reason I must enclose the feature list in a table cell...
	  strTemp = ph_html & strTbleStart & "<" & listType & " class=""" & ph_class & _
    strLiType & ">" & vbcrlf
    
  	do until rs.EOF
      if cint(ph_type) = 14 and len(trim(rs("Anchor")&"")) > 0 then
  		  strTemp = strTemp & "  <li><a href=""#" & rs("Anchor") & """>" & _
          InLinePHID(rs("PlaceHolderText")) & "</a></li>" & vbcrlf
      else
 	strTemp = strTemp & "  <li>" & InLinePHID(rs("PlaceHolderText")) & "</li>" & vbcrlf
      end if
        
      if m_bDoCsv then
        m_csvCmd.Parameters("@Source").value = "PHT"
        m_csvCmd.Parameters("@TextID").value = rs("TextID")
        m_csvCmd.Parameters("@TextVal").value = rs("PlaceHolderText")
        m_csvCmd.Parameters("@SpecVal").value = null
        m_csvCmd.Execute
if err.number <> 0 then
m_dbconn.close
m_dbconn.open
m_csvCmd.Execute
err.clear
end if
      end if  
  		rs.MoveNext
  	loop
  	set rs = nothing
  	set cmd = nothing
  	
  	Write_UL = strTemp & "</" & listType & ">" & strTblEnd & vbcrlf
    
  end function

   function Write_Img(phid,ph_html,ph_type,ph_url,ph_Tid,ph_ATid)
	
		Dim strTemp,cmd,rs,rsField
		' do both regular images and Translated Images here
	    
    if cint(ph_type) = 10 then
      ph_url = GetPHText(ph_Tid,False)
    end if
	    
		' Instr test for preview only, remove when we go live
		if UCase(Request.ServerVariables("SERVER_NAME")) = "WWW.DEV.FLUKE.COM" then
  		if Instr(lcase(ph_url),"/products/") > 0 then ph_url = "http://cms.dev.fluke.com" & ph_url
		end if
	    
		strTemp = ph_html & "<img border=""0"" src=""" & ph_url & """"
	    
		' is there AltText?
  	if len(ph_ATid&"") > 0 then
  		strTemp = strTemp & " alt=""" & GetPHText(ph_ATid,False) & """"
  	end if
	    
		' add ImageAttr if they exist
	    
  	set cmd = CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_ImageAttr_Get"
  	' create all the parameters we want
  	cmd.Parameters.Append cmd.CreateParameter("@PHID",3,1,,phid)
  	set rs = cmd.execute
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
set rs = cmd.execute
end if
		set cmd = nothing
	    
		if not rs.EOF then
			for each rsField in rs.Fields
				if len(rsField.Value&"") > 0 then
				strTemp = strTemp & " " & rsField.Name & "=""" & rsField.Value & """"
				end if
			next
		end if
		set rs = nothing
	    
		Write_Img = strTemp & ">" & vbcrlf
		
  end function

   function Write_Link(ph_type,ph_html,ph_class,ph_url,ph_Tid,ph_Atid)
	
    ' 6  PlainLink
    ' 11 TranslatedLink
    ' 16 BlankLink
    ' 20 PlacedLink
    ' 21 PlacedBlankLink
    Dim stfnt,endfnt,target,strHref,strLink
    
    target = ""
    strHref = ph_url
    strLink = ""
    
    if ph_class = "" then
      stfnt = ""
      endfnt = ""
    else
      stfnt = "<span class=""" & ph_class & """>"
      endfnt = "</span>"
    end if
    
    select case cint(ph_type)
      case 6 ' PlainLink, link contents are text
        strLink = GetPHText(ph_Tid,True)
      case 11 ' PlainLink, link contents are text, Href also comes from PlaceHolderText (thru AltTextID)
        strLink = GetPHText(ph_Tid,True)
        strHref = GetPHText(ph_ATid,False)
      case 16 ' BlankLink, just like PlainLink but in new window
        target = " TARGET=""_BLANK"""
        strLink = GetPHText(ph_Tid,True)
      case 20 ' PlacedLink, link contents come from TextID referencing another PHID
        strLink = GetOnePHID(ph_Tid)
      case 21 ' PlacedBlankLink
        target = " TARGET=""_BLANK"""
        strLink = GetOnePHID(ph_Tid)
    end select
    
    Write_Link = ph_html & "<A HREF=""" & strHref & """" & target & ">" & stfnt & _
      strLink & endfnt & "</a>" & vbcrlf
  end function

   function GetPHText(id,embed)
    Dim cmd,rs,strText,strLangs
    
  	set cmd = CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_GetPHText"
  	' create all the parameters we want
  	cmd.Parameters.Append cmd.CreateParameter("@TextID",3,1,,id)
  	cmd.Parameters.Append cmd.CreateParameter("@Lang",200,1,2,m_lang)
  	set rs = cmd.execute
if err.number <> 0 then
m_dbconn.close
m_dbconn.open
set rs = cmd.execute
err.clear
end if
    set cmd = nothing
    
    if rs.EOF then
      GetPHText = ""
    elseif embed then
	      GetPHText = InLinePHID(rs("PlaceHolderText"))
   
    else
      GetPHText = rs("PlaceHolderText")
    end if
    
  	'if csv functionality	
	  if m_bDoCsv then
    	m_csvCmd.Parameters("@Source").value = "PHT"
    	m_csvCmd.Parameters("@TextID").value = id
    	m_csvCmd.Parameters("@TextVal").value = rs("PlaceHolderText")
    	m_csvCmd.Parameters("@SpecVal").value = null
      m_csvCmd.Execute  
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
m_csvCmd.Execute  
end if  
	  end if		
	  
    set rs = nothing
  end function

   function Write_Spec(phid,ph_html,ph_class,ph_url,ph_Tid)
  	Dim cmd,rs,strTemp,previd,specid
  	Dim strText,strSub,strVal,iID,iSort,strLangs
    
    'strTemp = ph_html & "<A name=""spec" & phid & """><TABLE id=""tblProdSpec"" name=""tblProdSpec"" width=650 border=1 " _
	strTemp = ph_html & "<TABLE id=""tblProdSpec"" name=""tblProdSpec"" width=650 border=1 " _
  		& "bordercolor=""#C5CDD2"" cellpadding=4 cellspacing=0 >" _
  		& "<TR><TD colspan=2 class=""bodybold"" bgcolor=""#DDE6ED"" style=""padding: 3 3 3 3"">" & GetPHText(ph_Tid,True) _
  		& "</td></tr>" & vbcrlf
    
  	set cmd = Server.CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_SpecItems_Get"
    cmd.Parameters.Append cmd.CreateParameter("@PHID",3,1,,phid)
    cmd.Parameters.Append cmd.CreateParameter("@Lang",129,1,2,m_lang)
    
    set rs = cmd.Execute
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
set rs = cmd.Execute
end if
  	set cmd = nothing
    
    previd = 0
  	do until rs.EOF
      specid = cint(rs("SpecItemID"))
      
  		if cint(specid) <> cint(previd) then
  			if cint(previd) <> 0 then
    			strTemp = strTemp & "</table></td></tr>"
  			end if
  			strTemp = strTemp & "<TR><TD width=150 class=""PrdSpecItem"">" _
    			& InLinePHID(rs("PlaceHolderText")) & "</td>" _
    			& "<TD width=420 class=""PrdSpecItem"">" _
    			& "<TABLE border=0 cellpadding=0 cellspacing=0>" & vbcrlf _
    			& "<TR><TD class=""bodybold"">" & InLinePHID(rs("SpecSubject")) _
    			& "</td>" & vbcrlf _
    			& "<TD>&nbsp;" & InLinePHID(rs("SpecValue")) & "</td></TR>" _
    			& vbcrlf
          
        if m_bDoCsv then
          m_csvCmd.Parameters("@Source").value = "PHT"
      		m_csvCmd.Parameters("@TextID").value = rs("TextID")
      		m_csvCmd.Parameters("@TextVal").value = rs("PlaceHolderText")
      		m_csvCmd.Parameters("@SpecVal").value = null
      		m_csvCmd.Execute
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
m_csvCmd.Execute
end if
        end if
    		previd = specid
  		else
  			strTemp = strTemp & "<TR><TD class=""PrdSpecSubject"">" _
    			& InLinePHID(rs("SpecSubject")) & "</td>" & vbcrlf _
    			& "<TD>&nbsp;" & InLinePHID(rs("SpecValue")) & "</td></TR>" _
    			& vbcrlf
  		end if
		
  	  if m_bDoCsv then
    		m_csvCmd.Parameters("@Source").value = "SPC"
    		m_csvCmd.Parameters("@TextID").value = specid
    		m_csvCmd.Parameters("@SortOrder").value = rs("ValSort")
    		m_csvCmd.Parameters("@TextVal").value = rs("SpecSubject")
    		m_csvCmd.Parameters("@SpecVal").value = rs("SpecValue")
			  m_csvCmd.Execute
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
m_csvCmd.Execute
end if
      end if
					
  	  rs.MoveNext
  	loop
  	set rs = nothing
  	
  	'Write_Spec = strTemp & "</table></td></tr></table></a>" & vbcrlf
	Write_Spec = strTemp & "</table></td></tr></table></a>" & vbcrlf
  end function

   function Write_Table(phid,ph_html,ph_class,ph_url,ph_Tid)
  	Dim cmd,rs,rs2,strTemp,prevrow,strPHID
    
    Write_Table = ""
    
    ' Get Table Attributes & start the table. pdb_GTable_Get
    ' Loop on the Table contents, for each:
    '   write close row (if necessary)
    '   write open row (if necessary)
    '   write open TD
    '   get the phid & call Write_PHID
    '   close the TD
    
  	set cmd = Server.CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_GTable_Get"
       
    cmd.Parameters.Append cmd.CreateParameter("@PHID",3,1,,phid)    
    set rs = cmd.Execute
 if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
 set rs = cmd.Execute
end if
    if rs.EOF then exit function
    		
		' setup the getphid command
    cmd.CommandText = "pdb_Get1PHID"
		cmd.Parameters.Append cmd.CreateParameter("@Lang",129,1,2,m_lang)
			
		strTemp = ph_html & "<A name=""gtable" & phid & """><TABLE"
		    
		if len(rs("Border")&"") > 0 then strTemp = strTemp & " border=""" & rs("Border") & """"
		if len(rs("CellPadding")&"") > 0 then strTemp = strTemp & " CellPadding=""" _
		& rs("CellPadding") & """"
		if len(rs("Cellspacing")&"") > 0 then strTemp = strTemp & " Cellspacing=""" _
		& rs("Cellspacing") & """"
		if len(rs("Width")&"") > 0 then strTemp = strTemp & " Width=""" & rs("Width") & """"
		if len(rs("Align")&"") > 0 then strTemp = strTemp & " Align=""" & rs("Align") & """"
		if len(rs("BorderColor")&"") > 0 then strTemp = strTemp & " BorderColor=""" _
		& rs("BorderColor") & """"
		if len(rs("BGColor")&"") > 0 then strTemp = strTemp & " BGColor=""" & rs("BGColor") & """"
		  
		strTemp = strTemp & ">" & vbcrlf    
    set rs = rs.NextRecordSet
    
    ' write the first row before we get into the loop
    strTemp = strTemp & "<TR>" & vbcrlf
    prevrow = rs("TableRow")
    
  	do until rs.EOF
WaitFor 1,0
      ' we have these values
      ' rs("TableRow")
      ' rs("Rowspan")
      ' rs("TableCell") -- part of the sort
      ' rs("Colspan")
      ' rs("Align")
      ' rs("Valign")
      ' rs("Width")
      ' rs("CellPHID") -- this is where we get the contents
      
      if prevrow <> rs("TableRow") then
				strTemp = strTemp & "</TR>" & vbcrlf & "<TR>" & vbcrlf
				prevrow = rs("TableRow")
			end if
		      
			strTemp = strTemp & "<TD"
			if len(rs("Rowspan")&"") > 0 then strTemp = strTemp & " Rowspan=""" & rs("Rowspan") & """"
			if len(rs("Colspan")&"") > 0 then strTemp = strTemp & " Colspan=""" & rs("Colspan") & """"
			if len(rs("Align")&"") > 0 then strTemp = strTemp & " Align=""" & rs("Align") & """"
			if len(rs("Valign")&"") > 0 then strTemp = strTemp & " Valign=""" & rs("Valign") & """"
			if len(rs("Width")&"") > 0 then strTemp = strTemp & " Width=""" & rs("Width") & """"
			if len(rs("Class")&"") > 0 then strTemp = strTemp & " Class=""" & rs("Class") & """"
			strTemp = strTemp & ">" & vbcrlf
		        'Parag
			' Write_PHID(phid,ph_html,ph_type,ph_class,ph_url,ph_Tid,ph_ATid)
                        if isnull(rs("CellPHID"))= false then
				cmd.Parameters("@PHID").value = clng(rs("CellPHID"))
			else
				cmd.Parameters("@PHID").value = 0
			end if
on error resume next
			set rs2 = cmd.Execute
response.write err.description
'2147467259 
if err.number <> 0  then
response.write m_dbconn.state
if m_dbconn.state =1 then m_dbconn.close
m_dbconn.open
set rs2 = cmd.Execute
end if
on error goto 0
			strPHID = Write_PHID(rs("CellPHID"),"",rs2("PlaceHolderType"),rs2("Class"),rs2("URL") _
  						,rs2("TextID"),rs2("AltTextID"))
			if len(strPHID&"") > 0 then strTemp = strTemp & strPHID
			set rs2 = nothing
		      
			strTemp = strTemp & "</td>" & vbcrlf					
  		rs.MoveNext
  	loop
  	
  	set rs = nothing  	
		Write_Table = strTemp & "</tr></table></a>" & vbcrlf  		
  end function  

private function InLinePHID(strInit)
    ' regex objects & collections
    Dim regEx,reg2Ex,aMatch,aTail
    ' ADO stuff
    Dim cmd,rs,cmd1
    ' arrays
    Dim aList,aCloseMap,aCloseTags
    ' strings
    Dim strTar
    ' ints
    Dim ctr,tctr,iID,phType,iSort,iPos
    ' booleans
    Dim bCloseAtEnd,bFoundClose
    
    aCloseMap = Array()
    aCloseTags = Array()
    bCloseAtEnd = False
    
    InLinePHID = strInit
    
    if len(strInit&"") < 1 then
      exit function
    end if
    
    Set regEx = New RegExp   ' Create a regular expression.
    regEx.Pattern = "<PHID\s+id\s*=\s*(\d+\.?\d*)\s*>"   ' Set pattern
    regEx.IgnoreCase = True   ' Set case insensitivity.
    regEx.Global = True   ' Set global applicability.
    
    if not regEx.Test(strInit) then
      set regEx = nothing
      exit function
    end if
    
    Set reg2Ex = New RegExp   ' Create a regular expression.
    reg2Ex.Pattern = "</PHID>"   ' Set pattern.
    reg2Ex.IgnoreCase = True   ' Set case insensitivity.
    reg2Ex.Global = True   ' Set global applicability.
    
  	set cmd = CreateObject("ADODB.Command")
  	cmd.ActiveConnection = m_dbconn
  	cmd.CommandType = 4
  	cmd.CommandText = "pdb_Get1PHID"
  	' create all the parameters we want
  	cmd.Parameters.Append cmd.CreateParameter("@PHID",3,1)
    
    set aMatch = regEx.Execute(strInit)
    set aTail = reg2Ex.Execute(strInit)
    
    if aMatch.Count > 1 and aTail.Count <> aMatch.Count then
      InLinePHID = "<span style=""color: red"">mismatched phids in: </span>" & strInit
      exit function
    elseif aTail.Count = 0 then
      ' note that bCloseAtEnd will only be true if aMatch.Count = 1 and aTail.Count = 0
      bCloseAtEnd = true
    end if
    
    ReDim aCloseMap(aMatch.Count-1)
    ReDim aCloseTags(aMatch.Count-1)
    
    ' build the map of opening to closing tags such that the value of each indice (of aCloseMap)
    ' indicates which opening tag that position will close
    ' i.e. aCloseMap(0) = 1 indicates that the first </phid> gets the closing tag associated to
    ' the second <PHID ...>
    if aMatch.Count > 1 then
      ' this is the general case (even number of tags), the logic is to start from the
      ' end of the string and match the opening tag to the first unused closing tag
      ' that appears after the opening tag
      
      for ctr = (aMatch.Count-1) to 0 Step -1
        iPos = aMatch(ctr).FirstIndex
        bFoundClose = False
        
        for tctr = 0 to (aTail.Count-1)
          if aTail(tctr).FirstIndex > iPos then
            if not (PcatInArray(aCloseMap,tctr)) then
              aCloseMap(ctr) = tctr
              bFoundClose = True
              exit for
            end if
          end if
        next
        
        if not bFoundClose then
          InLinePHID = "<span style=""color: red"">mismatched phids in: </span>" & strInit
          exit function
        end if
      next
    else
      ' since there was only one PHID don't spend time figuring out where it is
      aCloseMap(0) = 0
    end if
    
    ' now walk through the collection of opening <PHID>s, perform the substitution
    ' and record the appropriate closing tag in aCloseTags
    for ctr = 0 to ubound(aCloseMap)
      
      ' get the ID value out of the open PHID tag
      iId = regEx.Replace(aMatch(ctr).Value,"$1")
      
      ' we'll be doing this to just the first match
      regEx.Global = False
      
      ' this is for AnchoredLists, the phid ID values are floats
      ' with the integer portion the PHID ID and the decimal portion the Sort position
      if Instr(iID,".") > 0 then
        aList = split(iID,".")
        iID = aList(0)
        iSort = aList(1)
      end if
      
      ' using stored procedure "pdb_Get1PHID", get the db values for the PHID's ID
      cmd.Parameters("@PHID").Value = cint(iID)
    	set rs = cmd.execute
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
set rs = cmd.execute
end if
      if rs.EOF then
        ' it wasn't in the database
        strInit = regEx.Replace(strInit,"")
        aCloseTags(ctr) = ""
      else
        ' this is analagous to Write_PHID's case statement, but only the embedded PH types
        ' the HTML is usually easy so we don't usually call more functions
        select case cint(rs("PlaceHolderType"))
          case 7, 17 ' EmbeddedLink
            strTar = ""
            if cint(rs("PlaceHolderType")) = 17 then strTar = " TARGET=""_BLANK"""
            if IsNull(rs("class")) then
              strInit = regEx.Replace(strInit,"<A HREF=""" & rs("url") & """" & strTar & ">")
            else
              strInit = regEx.Replace(strInit,"<A HREF=""" & rs("url") & """ class=""" & _
                rs("class") & """" & strTar & ">")
            end if
            aCloseTags(ctr) = "</a>"
          case 8 ' EmbeddedClass
            strInit = regEx.Replace(strInit,"<span class=""" & rs("class") & """>")
            aCloseTags(ctr) = "</span>"
          case 14 ' AnchoredList
            set rs = nothing
            set cmd1 = CreateObject("ADODB.Command")
          	cmd1.ActiveConnection = m_dbconn
          	cmd1.CommandType = 4
          	cmd1.CommandText = "pdb_AnchorList1_Get"
          	' create all the parameters we want
          	cmd1.Parameters.Append cmd.CreateParameter("@PHID",3,1,,cint(iID))
          	cmd1.Parameters.Append cmd.CreateParameter("@Sort",3,1,,cint(iSort))
            set rs = cmd1.execute
if err.number = -2147217908 then
m_dbconn.close
m_dbconn.open
set rs = cmd1.execute
end if
            if rs.EOF then
              strInit = regEx.Replace(strInit,"")
              aCloseTags(ctr) = ""
            else
              strInit = regEx.Replace(strInit,"<A name=""" & rs("Anchor") & """>")
              aCloseTags(ctr) = "</a>"
            end if
            set cmd1 = nothing
          case 18 ' Embedded List
            'Write_UL(phid,ph_type,ph_html,ph_class,ph_Tid)
            strInit = regEx.Replace(strInit,Write_UL(iID,18,"",rs("Class"),""))
            aCloseTags(ctr) = ""
        end select
      end if
      set rs = nothing
      
    next
    
    ' now substitute the ending tags, this is easy because the tags are stored
    ' in aCloseTags and the positions in aCloseMap, one might think that we
    ' need to sort aCloseMap with respect to the aTail collection, but we don't...
    if bCloseAtEnd then
      strInit = strInit & aCloseTags(0)
    else
      reg2Ex.Global = false
      for ctr = 0 to ubound(aCloseTags)
        strInit = reg2Ex.Replace(strInit,aCloseTags(aCloseMap(ctr)))
      next
    end if
    
    ' clean up
    set aMatch = nothing
    set regEx = nothing
    set reg2Ex = nothing
    set cmd = nothing
    InLinePHID = strInit
  end function

private function PcatInArray(aTemp,str)
    dim val
    'Response.write "testing " & str
    PcatInArray = false
    for each val in aTemp
      'Response.write " against " & val & "<BR>"
      'Response.write (cstr(val) = cstr(str))
      'Response.write "<BR>"
      if cstr(val) = cstr(str) then
        PcatInArray = True
        exit function
      end if
    next
  end function


Response.Buffer = true 
 
Function WaitFor(SecDelay,ShowMsg) 
    timeStart = Timer() 
    timeEnd = timeStart + SecDelay 
 
    Msg = "Timer started at " & timeStart & "<br>" 
    Msg = Msg & "Script will continue in " 
 
    i = SecDelay 
    Do While timeStart < timeEnd 
        If i = Int(timeEnd) - Int(timeStart) Then 
        Msg = Msg & i 
        If i <> 0 Then Msg = Msg & ", " 
        If ShowMsg = 1 Then Response.Write Msg 
%> 
 
<%         Response.Flush() %> 
 
<% 
        Msg = "" 
        i = i - 1 
        End if 
        timeStart = Timer() 
    Loop 
    Msg = "...<br>Slept for " & SecDelay & " seconds (" & _ 
        Timer() & ")" 
    If ShowMsg = 1 Then Response.Write Msg 
End Function 

'-----------------------------------------------------------------------------------------------------------


response.End

Dim objTemplate
'-----------------------------------------------------------------------------------------------------------
if Request.QueryString("Old")= "true" then
	Set objTemplate = Autosession.Searches.GetByPath("/Templates/Web2_0/Fluke Core Templates/Generic/NoNav_onePlaceholder")
	objTemplate.SourceFile="/NR/exeres/Templates/Authors/Generic/NoNav_onePlaceholder_staging.asp"
else
	Set objTemplate = Autosession.Searches.GetByPath("/Templates/Web2_0/Fluke Core Templates/Generic/NoNav_onePlaceholder")
	objTemplate.SourceFile="/NR/ExeRes/Templates/Web2_0/Fluke_CoreTemplates/Generic/NoNav_onePlaceholder_staging.asp"
end if

Call Autosession.CommitAll
'-----------------------------------------------------------------------------------------------------------
Response.Write "Done"
Response.End

server.scripttimeout = 3000

'response.buffer=false
Dim rsGetPosting, objConnFWC, oConnFWC, pSql
Dim pGetPosting, pGetPostingNew
Dim dispTitleOld,dispDescrOld, dispKeywordsOld, prodID
Dim dispTitleNew,dispDescrNew, dispKeywordsNew

Dim upd
upd=Request.QueryString("update")


dim strCatalogName
dim strCatalogLanguage
Dim strTable
Dim strUrlOldPost,strUrlNewPost

strTable="<table border=1 cellpadding=0 cellspacing=0><tr><td>"
strTable=strTable & "<b><font face=Arial size=2>Product ID</font></b> </td>"
strTable=strTable & "<td><b><font face=Arial size=2>Old Posting Title</b> </td>"
strTable=strTable & "<td><b><font face=Arial size=2>Old Posting Keywords</b> </td>"
strTable=strTable & "<td><b><font face=Arial size=2>Old Posting Description</b> </td>"
strTable=strTable & "<td><b><font face=Arial size=2>New Posting Title</b> </td>"
strTable=strTable & "<td><b><font face=Arial size=2>New Posting Description</b> </td>"
strTable=strTable & "<td><b><font face=Arial size=2>New Posting Keywords</b> </td></tr>"
strCatalogName = "FlukeUnitedStates"
strCatalogLanguage = "en-us"

Set pProductUtilities = New NRCSProductUtilities
call InitializeContentConnector(strCatalogName, strCatalogLanguage)

Set objConnFWC = new ConnectionManager
Set oConnFWC =  objConnFWC.db_ConnectDefault("FLUKEWEB_COMMERCE")


pSql="select distinct productid,richcontentpostingid,newrichcontentpostingid from "&_
"[flukeproducts_en-us_catalog] a ,flukeproducts_catalogproducts b" &_
" where [#Catalog_Lang_Oid] = b.oid" &_
" and richcontentpostingid is not null AND richcontentpostingid <> ''"

'response.write psql
Set rsGetPosting = Server.CreateObject("ADODB.RecordSet")
rsGetPosting.open pSql, oConnFWC, 3, 3
If not rsGetPosting.EOF then

Do Until rsGetPosting.EOF


strTable=strTable & "<tr>"

'''' Get Product ID ''''

prodID=""
prodID=rsGetPosting("productid")

''''' End Product ID '''''

''''' Display Product ID '''''

	
	strTable=strTable & "<td><font face=Arial size=2>" & prodID & "</font></td>"

'''''' End Display Product ID ''''''''


'''' Get Old Posting Title, Description & Keywords  '''''''
set pGetPosting=nothing
if (isnull(rsGetPosting("richcontentpostingid")) = false and rsGetPosting("richcontentpostingid") <> "") then

set pGetPosting = g_pNRCSFactory.GetPostingByGUID(rsGetPosting("richcontentpostingid"))

if not pGetPosting is nothing then

	


	dispTitleOld=""
	dispDescrOld=""
	dispKeywordsOld =""

	dispTitleOld=pGetPosting.Posting.Placeholders("METADATA_TITLE").HTML
	dispDescrOld=pGetPosting.Posting.Placeholders("METADATA_DESCRIPTION").HTML
	dispKeywordsOld=pGetPosting.Posting.Placeholders("METADATA_KEYWORDS").HTML

	


'''''''' Display Old Posting Title, Description & Keywords  ''''''''
	
	strTable=strTable & "<td><font face=Arial size=2>" & dispTitleOld & "</td>"	
	strTable=strTable & "<td><font face=Arial size=2>" & dispDescrOld & "</td>"
	strTable=strTable & "<td><font face=Arial size=2>" & dispKeywordsOld & "</td>"
	

''''''' End Display Old Posting ID ''''	



	''''' Get New Posting '''''

	set pGetPostingNew=nothing
	if (isnull(rsGetPosting("newrichcontentpostingid"))=false and rsGetPosting("newrichcontentpostingid") <> "") then
	set pGetPostingNew= g_pNRCSFactory.GetPostingByGUID(rsGetPosting("newrichcontentpostingid"))

	if not pGetPostingNew is nothing then
		
''' Check for the update querystring variable. If the value is true, then proceed with the change '''

		If (upd = "true") then 
			
			if trim(pGetPostingNew.Posting.Placeholders("METADATA_TITLE").HTML) = "" then
				pGetPostingNew.Posting.Placeholders("METADATA_TITLE").HTML=dispTitleOld
			end if

			if trim(pGetPostingNew.Posting.Placeholders("METADATA_DESCRIPTION").HTML) = "" then
				pGetPostingNew.Posting.Placeholders("METADATA_DESCRIPTION").HTML=dispDescrOld
			end if

			if trim(pGetPostingNew.Posting.Placeholders("METADATA_KEYWORDS").HTML) = "" then
				pGetPostingNew.Posting.Placeholders("METADATA_KEYWORDS").HTML=dispKeywordsOld
			end if

			pGetPostingNew.Posting.approve
			Call Autosession.CommitAll

		end if

		''' End Check ''''''''''''''''''

		dispTitleNew=""
		dispTitleNew=pGetPostingNew.Posting.Placeholders("METADATA_TITLE").HTML
		dispDescrNew=""
		dispDescrNew=pGetPostingNew.Posting.Placeholders("METADATA_DESCRIPTION").HTML
		dispKeywordsNew=""
		dispKeywordsNew=pGetPostingNew.Posting.Placeholders("METADATA_KEYWORDS").HTML


'''''''' Display New Posting Title, Description & Keywords  ''''''''
	
	strTable=strTable & "<td><font face=Arial size=2>" & dispTitleNew & "</td>"
	strTable=strTable & "<td><font face=Arial size=2>" & dispDescrNew & "</td>"
	strTable=strTable & "<td><font face=Arial size=2>" & dispKeywordsNew & "</td>"


''''''' End Display New Posting ID ''''	
	else
		
	end if
else
	strTable=strTable & "<td>&nbsp;</td>"
	strTable=strTable & "<td>&nbsp;</td>"
	strTable=strTable & "<td>&nbsp;</td>"
strTable=strTable & "<td>&nbsp;</td>"

	end if
	''' End New Posting Name ''''
else
strTable=strTable & "<td>&nbsp;</td>"
strTable=strTable & "<td>&nbsp;</td>"
strTable=strTable & "<td>&nbsp;</td>"
strTable=strTable & "<td>&nbsp;</td>"
strTable=strTable & "<td>&nbsp;</td>"
strTable=strTable & "<td>&nbsp;</td>"

end if
else
end if
'''''' End Old Posting Name  ''''''

strTable=strTable & "</tr>"

rsGetPosting.movenext
loop
end if

strTable=strTable & "</table>"
response.write strTable
Set objConnFWC=Nothing
Set oConnFWC=Nothing


Response.end
''''''''''''''''''''''''''' End '''''''''''''''''''''



dim pSimplePosting
dim pProduct

dim strMessage
Dim pProductUtilities
dim sProductID
dim rsCategories
dim rsProductProps
dim pCategory
dim strRichProductTemplate
dim strRichProductChannel	
dim aList

Server.ScriptTimeout = 3000

Set pProductUtilities = New NRCSProductUtilities

'Get post data
'strCatalogName = Request("hidCatalogName")
'strCatalogLanguage = request("hidLanguage")

strCatalogName = "FlukeUnitedStates"
strCatalogLanguage = "en-us"
'aList = split(sProductId,", ") 
Dim strData

Dim rsProductPIDs, objConnMan, oConn, sSQL

sSQL = "SELECT  DISTINCT P.PID As PID,ProductBrandId  FROM dbo.mcs_Products AS P INNER JOIN dbo.mcs_MultiLingual AS M ON P.PID = M.PID" &_  
" WHERE (P.ProductBrandID IN (1, 4)) AND (P.PID IN " &_
"(" &_
"38514,"&_
"36706,"&_
"36895,"&_
"38363,"&_
"36141,"&_
"37793,"&_
"34238,"&_
"35911,"&_
"35969,"&_
"34935,"&_
"35481,"&_
"36851,"&_
"36016,"&_
"36463,"&_
"38414,"&_
"37036,"&_
"37115,"&_
"37783,"&_
"32067,"&_
"38553,"&_
"31598,"&_
"35744,"&_
"38101,"&_
"36791,"&_
"36864,"&_
"36899,"&_
"6574,"&_
"35975,"&_
"35635,"&_
"36479,"&_
"36764,"&_
"37116,"&_
"3034,"&_
"36519,"&_
"36770,"&_
"36594,"&_
"30579,"&_
"36073,"&_
"36504,"&_
"8129,"&_
"1707,"&_
"34827,"&_
"34936,"&_
"36584,"&_
"36673,"&_
"37803,"&_
"36521,"&_
"36776,"&_
"36938,"&_
"36771,"&_
"36520,"&_
"36644,"&_
"38246,"&_
"126,"&_
"35665,"&_
"7629,"&_
"38508,"&_
"34826,"&_
"36740,"&_
"36859,"&_
"34880,"&_
"36252,"&_
"36751,"&_
"37430,"&_
"36018,"&_
"36074,"&_
"36083,"&_
"35995,"&_
"36678,"&_
"36157,"&_
"36769,"&_
"36548,"&_
"36554,"&_
"34273,"&_
"35723,"&_
"3627,"&_
"9065,"&_
"3395,"&_
"35480,"&_
"36028,"&_
"36616,"&_
"38510,"&_
"303,"&_
"36227,"&_
"36477,"&_
"5179,"&_
"3435,"&_
"34881,"&_
"36097,"&_
"23,"&_
"29983,"&_
"34723,"&_
"36263,"&_
"25524,"&_
"31084,"&_
"34552,"&_
"36215,"&_
"36643,"&_
"36087,"&_
"36426,"&_
"36574,"&_
"38223,"&_
"36092,"&_
"36093,"&_
"36612,"&_
"36225,"&_
"35686,"&_
"36286,"&_
"38264,"&_
"36590,"&_
"38220,"&_
"34645,"&_
"35998,"&_
"36144,"&_
"3617,"&_
"37767,"&_
"34472,"&_
"35667,"&_
"35970,"&_
"35287,"&_
"36855,"&_
"38292,"&_
"19351,"&_
"36515,"&_
"37914,"&_
"8542,"&_
"29419,"&_
"34931,"&_
"35719,"&_
"36558,"&_
"36767,"&_
"3635,"&_
"36638,"&_
"24896,"&_
"37795,"&_
"34933,"&_
"36790,"&_
"36868,"&_
"35993,"&_
"28764,"&_
"315,"&_
"35971,"&_
"1982,"&_
"35656,"&_
"36166,"&_
"37035,"&_
"31015,"&_
"36015,"&_
"36077,"&_
"34889,"&_
"36244,"&_
"38156,"&_
"20400,"&_
"2744,"&_
"32582,"&_
"30595,"&_
"34673,"&_
"36082,"&_
"3039,"&_
"36160,"&_
"38212,"&_
"36725,"&_
"36786,"&_
"36154,"&_
"36551,"&_
"34929,"&_
"36220,"&_
"36254,"&_
"38257,"&_
"38206,"&_
"5178,"&_
"35708,"&_
"36246,"&_
"7998,"&_
"36779,"&_
"36844,"&_
"35478,"&_
"35716,"&_
"36509,"&_
"36834,"&_
"36956,"&_
"35962,"&_
"36210,"&_
"36274,"&_
"30494,"&_
"34301,"&_
"35475,"&_
"7786,"&_
"21,"&_
"36072,"&_
"36744,"&_
"17773,"&_
"36591,"&_
"38213,"&_
"34473,"&_
"36414,"&_
"36754,"&_
"34930,"&_
"35866,"&_
"36893,"&_
"35746,"&_
"36672,"&_
"36854,"&_
"34330,"&_
"35741,"&_
"36480,"&_
"35489,"&_
"36211,"&_
"8368,"&_
"36774,"&_
"36284,"&_
"36550,"&_
"36642,"&_
"38199,"&_
"38556,"&_
"3410,"&_
"36042,"&_
"36739,"&_
"35445,"&_
"36204,"&_
"36860,"&_
"34355,"&_
"35484,"&_
"36276,"&_
"36288,"&_
"38511,"&_
"27281,"&_
"3611,"&_
"36736,"&_
"37151,"&_
"35989,"&_
"36094,"&_
"36731,"&_
"2338,"&_
"36522,"&_
"36919,"&_
"36027,"&_
"36084,"&_
"36159,"&_
"36147,"&_
"36656,"&_
"37725,"&_
"22,"&_
"36017,"&_
"34651,"&_
"36468,"&_
"36861,"&_
"36858,"&_
"23515,"&_
"24390,"&_
"36741,"&_
"34948,"&_
"36541,"&_
"36777,"&_
"36549,"&_
"36784,"&_
"3396,"&_
"34568,"&_
"35391,"&_
"23623,"&_
"34927,"&_
"30602,"&_
"3394,"&_
"37792,"&_
"34884,"&_
"36614,"&_
"9853,"&_
"34888,"&_
"36021,"&_
"5574,"&_
"36010,"&_
"36170,"&_
"37895,"&_
"34236,"&_
"37827,"&_
"35296,"&_
"36987,"&_
"37513,"&_
"17969,"&_
"36540,"&_
"17500,"&_
"34272,"&_
"36571,"&_
"36462,"&_
"36559,"&_
"36712,"&_
"37965,"&_
"2279,"&_
"34298,"&_
"35615,"&_
"8221,"&_
"36709,"&_
"8059,"&_
"20144,"&_
"36737,"&_
"38208,"&_
"1704,"&_
"36676,"&_
"36248,"&_
"36789,"&_
"36951,"&_
"36675,"&_
"34934,"&_
"36155,"&_
"7875,"&_
"25136,"&_
"36977,"&_
"36213,"&_
"36677,"&_
"1666,"&_
"32023,"&_
"36158,"&_
"35994,"&_
"36163,"&_
"36772,"&_
"38207,"&_
"21473,"&_
"36547,"&_
"36555,"&_
"38237,"&_
"38507,"&_
"38211,"&_
"32833,"&_
"33309,"&_
"36266,"&_
"36019,"&_
"38256,"&_
"36957,"&_
"34239,"&_
"36867,"&_
"7144,"&_
"36011,"&_
"36029,"&_
"37062,"&_
"38183,"&_
"32946,"&_
"36223,"&_
"36556,"&_
"31768,"&_
"6516,"&_
"9852,"&_
"34277,"&_
"35490,"&_
"36832,"&_
"36156,"&_
"36467,"&_
"36765,"&_
"36557,"&_
"36649,"&_
"36937,"&_
"7845,"&_
"38364,"&_
"38513,"&_
"34882,"&_
"35681,"&_
"36164,"&_
"3574,"&_
"36639,"&_
"37794,"&_
"35482,"&_
"36148,"&_
"3030,"&_
"36224,"&_
"36865,"&_
"33061,"&_
"36546,"&_
"38304,"&_
"36281,"&_
"38153,"&_
"36218,"&_
"36866,"&_
"36898,"&_
"37737,"&_
"37819,"&_
"36422,"&_
"36512,"&_
"36593,"&_
"36575,"&_
"36214,"&_
"36475,"&_
"38184,"&_
"34237,"&_
"35972,"&_
"36761,"&_
"38290,"&_
"36729,"&_
"36766,"&_
"38222,"&_
"36792,"&_
"17187,"&_
"35045,"&_
"35745,"&_
"36592,"&_
"35633,"&_
"36024,"&_
"30956,"&_
"9854,"&_
"36894,"&_
"6575,"&_
"36078,"&_
"34299,"&_
"36090,"&_
"38512,"&_
"29011,"&_
"37034,"&_
"35476,"&_
"36768,"&_
"38178,"&_
"36249,"&_
"36569,"&_
"38049,"&_
"35583,"&_
"35628,"&_
"36535,"&_
"38209,"&_
"35743,"&_
"36081,"&_
"6556,"&_
"34828,"&_
"36843,"&_
"1708,"&_
"30405,"&_
"36637,"&_
"15270,"&_
"36099,"&_
"36560,"&_
"36000,"&_
"36091,"&_
"36954,"&_
"24317,"&_
"34937,"&_
"35616,"&_
"36145,"&_
"36473,"&_
"34300,"&_
"36076,"&_
"36080,"&_
"24565,"&_
"36167,"&_
"35595,"&_
"34353,"&_
"36647,"&_
"38019,"&_
"38140,"&_
"27383,"&_
"29214,"&_
"36038,"&_
"19035,"&_
"38282,"&_
"8130,"&_
"38517,"&_
"36586,"&_
"38265,"&_
"38291,"&_
"35706,"&_
"36478,"&_
"36833,"&_
"38187,"&_
"36699,"&_
"36469,"&_
"36763,"&_
"38210,"&_
"37790,"&_
"36095,"&_
"36671,"&_
"36711,"&_
"35742,"&_
"36086,"&_
"37955,"&_
"36098,"&_
"37736,"&_
"35485,"&_
"37495,"&_
"33326,"&_
"3393,"&_
"34354,"&_
"36753,"&_
"38103,"&_
"38503,"&_
"2278,"&_
"38205,"&_
"7521,"&_
"36161,"&_
"37796,"&_
"38464,"&_
"38189,"&_
"34474,"&_
"34807,"&_
"36006,"&_
"37758,"&_
"32084,"&_
"36009,"&_
"36415,"&_
"37786,"&_
"36788,"&_
"36847,"&_
"37770,"&_
"36245,"&_
"36418,"&_
"36472,"&_
"36773,"&_
"36955,"&_
"7544,"&_
"28688,"&_
"35270,"&_
"36572,"&_
"29300,"&_
"36096,"&_
"6569,"&_
"36001,"&_
"20,"&_
"36089,"&_
"36209,"&_
"36041,"&_
"36640,"&_
"36735,"&_
"34885,"&_
"26774,"&_
"26864,"&_
"29980,"&_
"37724,"&_
"36168,"&_
"26695,"&_
"34551,"&_
"34928,"&_
"36617,"&_
"34853,"&_
"36040,"&_
"35707,"&_
"16273,"&_
"35414,"&_
"36783,"&_
"34932,"&_
"36143,"&_
"36762,"&_
"34886,"&_
"36079,"&_
"36008,"&_
"36845,"&_
"37881,"&_
"35479,"&_
"36088,"&_
"36745,"&_
"36862,"&_
"36013,"&_
"36169,"&_
"36836,"&_
"35291,"&_
"36255,"&_
"36778,"&_
"38515,"&_
"1710,"&_
"35704,"&_
"36775,"&_
"35465,"&_
"36474,"&_
"27075,"&_
"36423,"&_
"36982,"&_
"36207,"&_
"34806,"&_
"36022,"&_
"36205,"&_
"36287,"&_
"36583,"&_
"36953,"&_
"36413,"&_
"36897,"&_
"36075,"&_
"36952,"&_
"38144,"&_
"25829,"&_
"34925,"&_
"36755,"&_
"24480,"&_
"37377,"&_
"36039,"&_
"35447,"&_
"36681,"&_
"35566,"&_
"5177,"&_
"15138,"&_
"36153,"&_
"36615,"&_
"36251,"&_
"20221,"&_
"38260,"&_
"38554,"&_
"36585,"&_
"23932,"&_
"35477,"&_
"36573,"&_
"35674,"&_
"36257,"&_
"36891,"&_
"38524,"&_
"36890,"&_
"38155,"&_
"38448,"&_
"19200,"&_
"35526,"&_
"36226,"&_
"36856,"&_
"23382,"&_
"36425,"&_
"36619,"&_
"36165,"&_
"36857,"&_
"26533,"&_
"36142,"&_
"36680,"&_
"36542,"&_
"21454,"&_
"29927,"&_
"36012,"&_
"36543,"&_
"38201,"&_
"34348,"&_
"36710,"&_
"30588,"&_
"5308,"&_
"36785,"&_
"36863,"&_
"7699,"&_
"30760,"&_
"35657,"&_
"36005,"&_
"2574,"&_
"36085,"&_
"7227"&_
")" &_
")"  &_
"AND (M.mcsLocale = 'en-us')"

'"AND (M.StartDate <= GETDATE()) AND (M.EndDate >= GETDATE())  ORDER BY P.PID "

'sSQL = "SELECT distinct PID FROM dbo.mcs_Products WHERE (ProductBrandID IN (1, 4)) AND (ProductType IN ('M')) ORDER BY PID"
Set objConnMan = new ConnectionManager
Set oConn =  objConnMan.db_ConnectDefault("PRODUCTCATALOG")


Set rsProductPIDs = Server.CreateObject("ADODB.RecordSet")
rsProductPIDs.open sSQL, oConn, 3, 3

'Response.Write rsProductPIDs.Fields("PID").Value
If not rsProductPIDs.EOF then

Do Until rsProductPIDs.EOF

On Error Resume Next
aList = split(CStr(rsProductPIDs.Fields("PID").Value) & "(FlukeProducts)",", ") 

call InitializeContentConnector(strCatalogName, strCatalogLanguage)
			Dim rsM
if not g_pCSCatalog is nothing then

	' loop thru list of products sent in
	for each sProductID in aList
		dim pProductPosting, rsCurrentProps
		
		set pCategory = nothing		
		set pProduct = g_pCSCatalog.GetProduct(sProductID)
		'if not(pProduct is nothing) then

'			set rsM = pProduct.GetProductProperties
'			rsM.fields("newrichcontentpostingid").value = null
'			call pProduct.SetProductProperties(rsM,True)
'set rsM= nothing
'		end if
'		Call Autosession.CommitAll
		set pProductPosting = g_pNRCSFactory.GetPostingByProduct(pProduct)
if rsProductPIDs("ProductBrandId").value = 1 then
strRichProductChannel = "/Channels/nusen/Products/" 'trim(rsCatProps("NewRichProductChannel"))     
strData="<P align=left><A href=""http://register.fluke.com/globalforms/global.asp?formnum=466"" target=_blank><IMG class=prdButton_top src=""/NR/rdonlyres/A22B3F2C-BC2B-4C3F-976E-2E2DF6B9B2F7/0/prd_detail_contact_btn_295px_x_20px.gif"" border=0></A><BR>"&_
"<A href=""http://register.fluke.com/globalforms/global.asp?formnum=468"" target=_blank><IMG class=prdButton_bottom src=""/NR/rdonlyres/C7BD077F-5DAD-46C5-B879-E9DA8438AF44/0/prd_detail_demo_btn_295px_x_20px.gif"" border=0></A> </P>"
elseif rsProductPIDs("ProductBrandId").value = 4 then
strRichProductChannel = "/Channels/nbusen/Products/"
strData="<P align=left><A href=""http://www.fluke.com/globalforms/fbc_lead3.asp?formnum=193"" target=_blank><IMG class=prdButton_bottom src=""/NR/rdonlyres/C7BD077F-5DAD-46C5-B879-E9DA8438AF44/0/prd_detail_demo_btn_290px_x_20px.gif"" border=0></A></P>"&_
"<P align=left><A href=""http://www.fluke.com/globalforms/fbc_lead3.asp?formnum=193"" target=_blank><IMG src=""/NR/rdonlyres/9835277E-E2F8-41FD-8A15-D79EDAA341FA/0/prd_detail_quote_btn_290px_x_20px.gif"" border=0></A></P>"
end if
		if (pProductPosting.IsRichProductPosting) then
			Response.Write pProductPosting.Posting.Name & " is already a rich posting<br>"
			'pProductPosting.Posting.Placeholders("CONTENT_AREA_1").HTML = ""
			if rsProductPIDs("ProductBrandId").value = 4 then
				
pProductPosting.Posting.Placeholders("CONTENT_AREA_2").HTML = strData
				pProductPosting.Posting.approve
				Call Autosession.CommitAll
			end if
		else
				dim sCategoryName
				sCategoryName = pProduct.PrimaryParent
				set pCategory = g_pCSCatalog.GetCategory(sCategoryName)

				' must have category to get template and channel targets
				if not (pCategory is nothing) then
					dim rsCatProps
					dim pRichProductChannel
					dim pRichProductTemplate
					
					set pRichProductChannel = nothing
					set pRichProductTemplate = nothing
					
					set rsCatProps = pCategory.GetCategoryProperties
					strRichProductTemplate = "/Templates/web2_0/Authors/Products/NewRichFlukeProduct" 
					'trim(rsCatProps("NewRichProductTemplate"))
					
					dim node
					if (len(strRichProductChannel) > 0 ) then
						' get richproductchannel
						set pRichProductChannel = Autosession.Searches.GetByPath(strRichProductChannel)
					end if			
					if (len(strRichProductTemplate) > 0 ) then
						' get richproducttemplate
						set pRichProductTemplate = Autosession.Searches.GetByPath(strRichProductTemplate)
					end if	
					'on error resume next	
					'response.Write pRichProductChannel.CanCreatePostings 
					' check for actual template and channel objects
					if  ((pRichProductTemplate is nothing) OR (pRichProductChannel is nothing)) then
						Response.Write "No Rich Product Channel or Rich ProductTemplate for product " & sProductID & "<br>"
					else
						dim pPosting
						dim pRichPosting

						set pRichPosting = nothing
						set pPosting = nothing
						' create posting, passing in template
						set pPosting = pRichProductChannel.CreatePosting(pRichProductTemplate)
if rsProductPIDs("ProductBrandId").value = 1 then
strData="<P align=left><A href=""http://register.fluke.com/globalforms/global.asp?formnum=466"" target=_blank><IMG class=prdButton_top src=""/NR/rdonlyres/A22B3F2C-BC2B-4C3F-976E-2E2DF6B9B2F7/0/prd_detail_contact_btn_295px_x_20px.gif"" border=0></A><BR>"&_
"<A href=""http://register.fluke.com/globalforms/global.asp?formnum=468"" target=_blank><IMG class=prdButton_bottom src=""/NR/rdonlyres/C7BD077F-5DAD-46C5-B879-E9DA8438AF44/0/prd_detail_demo_btn_295px_x_20px.gif"" border=0></A> </P>"
elseif rsProductPIDs("ProductBrandId").value = 4 then
strData= "<P align=left><A href=""http://register.fluke.com/globalforms/global.asp?formnum=465"" target=_blank><IMG class=prdButton_top src=""/NR/rdonlyres/A22B3F2C-BC2B-4C3F-976E-2E2DF6B9B2F7/0/prd_detail_contact_btn_295px_x_20px.gif"" border=0></A>"&_
"<BR><A href=""http://register.fluke.com/globalforms/global.asp?formnum=467"" target=_blank><IMG class=prdButton_bottom src=""/NR/rdonlyres/C7BD077F-5DAD-46C5-B879-E9DA8438AF44/0/prd_detail_demo_btn_295px_x_20px.gif"" border=0></A> </P>"
end if
if trim(pProductPosting.Posting.Placeholders("CONTENT_AREA_2").HTML) = "" then
	pPosting.Placeholders("CONTENT_AREA_2").HTML = strData
end if
						'response.write err.description & "<br>"
pPosting.approve
								'pRichPosting.Posting.commit
								Call Autosession.CommitAll
						if not pPosting is nothing then
							' set props
							pPosting.Description = ""
							pPosting.StartDate = Date
				
							set rsProductProps = pProduct.GetProductProperties

							if not rsProductProps.eof then
								dim sName
								sName = rsProductProps("Name")
								'Response.Write sProductID & "<br>"
								pPosting.Name = trimSpecials(sName)
							else
								pPosting.Name = trimSpecials(sProductID)
							end if
							' response.End					
							' get the NRCSPosting object (cast)
							

							set pPosting = g_pNRCSFactory.GetPostingByGUID(pPosting.GUID)
							'response.Write "fdgdfgdfg11111111111"
							'response.write err.description & "<br>"

							Set pRichPosting = pProductUtilities.CreateRichProductMappingForPosting(pPosting, pProduct)
						        'response.write err.description & "<br>"
							response.Write pRichPosting.Posting.template.name
							
							if not (pRichPosting is nothing ) then
								pRichPosting.Posting.approve
								'pRichPosting.Posting.commit
								Call Autosession.CommitAll
								'Pause(2)
								Response.Write "New Posting for PID: " & sProductID & "<a href=""" & pRichPosting.Posting.URLModePublished & """ target=""_new"">" & pRichPosting.Posting.Name & "</a><br>"
							else
								' don't leave orphans
								pPosting.Delete
								pPosting.Commit
								Response.Write "Failed to create new rich product posting for PID: " & sProductID
								Response.Write "<br>"
							end if
						else
							Response.Write "Unable to create posting in channel " & pRichProductChannel.Path & "<br>"
						end if 'not pRichPosting is nothing
					end if '((pRichProductTemplate is nothing) OR (pRichProductChannel is nothing))
			end if 'not (pCategory is nothing)
		end if 'simple
		' output result to screen
		Response.Flush
	next 'for each sProductID in aList
end if 'not g_pCSCatalog is nothing
'if err.number <> 0 then
'	Response.Write err.Description & "<BR>"
	'err.Clear
'end if 

rsProductPIDs.MoveNext
Loop
End if
 
rsProductPIDs.Close
Set rsProductPIDs = Nothing
Set oConn = Nothing
Set oConn = Nothing
Set objConnMan = Nothing

' trim special characters from name
function trimSpecials(byval strIn)
	dim regEx, sPattern
	sPattern = "[~`!@#$%&*/\\]"
	set regEx = new RegExp
	regEx.Pattern  =sPattern
	regEx.Global = true
	trimSpecials = regEx.Replace(strIn, " ")
end function

Sub Pause(intSeconds)
	Dim startTime
	startTime = Time()
	Do Until DateDiff("s", startTime, Time(), 0, 0) > intSeconds
	Loop
End Sub

%>