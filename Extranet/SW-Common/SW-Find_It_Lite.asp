<%
' --------------------------------------------------------------------------------------
' Author:     Kelly Whitlock
' Date:       2/5/2005
' Name:       Asset File Find It Lite Version
' Purpose:    Checks to see if asset is available.  Returns: FIND_IT.asp Path and Title in formated HTML string or NULL)
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

response.write Find_It_Link(request("Document"), request("Title"), request("CMS_Site"), "" ,"<BR>")


Function Find_It_Link(Item_Number, Title, CMS_Site, BHTML, EHTML)

  ' Item_Number - Single or | separated string of Oracle 7-digit Item Number(s)
  ' Title - Single Title or | separated string of Titles (Override for Portal Title)
  ' CMS_Site - page where function is called (reference and tracking purposes)
  ' BHTML - Beginning HTML string (optional)
  ' EHTML - Ending HTML string    (optional or <BR>)

  Dim SQL, aServer, aNodes, node, Find_It_Server, Find_It_Links, x, Item_Numbers_Max, Titles_Max
  
  if instr(1,Item_Number,"|") > 0 then
  
    Item_Numbers = split(Item_Number,"|")
    Item_Numbers_Max = UBound(Item_Numbers)

    if instr(1,Title,"|") > 0 then
      Titles = split(Title,"|")
      Titles_Max = UBound(Titles)
    else
      reDim Titles(Item_Numbers_Max)
      for x = 0 to Item_Numbers_Max
        Titles(x)  = ""
      next
      Titles_Max = Item_Numbers_Max
    end if
    
  else
  
    reDim Item_Numbers(0)
    Item_Numbers(0)  = Item_Number
    Item_Numbers_Max = 0
    
    reDim Titles(0)
    Titles(0)  = Title
    Titles_Max = 0

  end if
  
  if Item_Numbers_Max <> Titles_Max then

    Find_It_Link = "Unmatched Item_Number(s) (" & Item_Numbers_Max & ") and Title(s) (" & Titles_Max & ")"
  
  else

    ' --------------------------------------------------------------------------------------  
    Call Connect_SiteWide
    ' --------------------------------------------------------------------------------------

    aServer = request.ServerVariables("SERVER_NAME")
    aNodes  = split(aServer,".")
	
    Find_IT_Server = "Support.Fluke.com"
  
	  for each node in aNodes
		  select case node
			  case "DEV","TST"
				  Find_IT_Server = "Support.DEV.Fluke.com"
        case "DTMEVTVSDV15"
				  Find_IT_Server = "dtmevtvsdv15.danahertm.com"        
  				exit for
	  	end select
    next
  
    for x = 0 to Item_Numbers_Max

      SQL = "SELECT Title " &_
            "FROM   dbo.Calendar " &_
            "WHERE (Item_Number = '" & Trim(Item_Numbers(x)) & "') " &_
            "       AND (File_Name IS NOT NULL) " &_
            "       AND (SubGroups LIKE '%view%' OR SubGroups LIKE '%fedl%') " &_
            "       AND (Status = 1) " &_
            "ORDER  BY Revision_Code DESC, UDate DESC"
          
      'response.write SQL & "<P>"
      
      Set rsAsset = Server.CreateObject("ADODB.Recordset")
      rsAsset.Open SQL, conn, 3, 3
     
      if not rsAsset.EOF then
    
        if len(Titles(x)) = 0 then
          Find_It_Links = Find_It_Links & BHTML & "<A HREF=""http://" & Find_IT_Server & "/Find_It.asp?Document=" & Item_Numbers(x) & "&SRC=FDL&Server=" & aServer & "&CMS_Site=" & CMS_Site & """>" & rsAsset("Title") & "</A>" & EHTML
        else
          Find_It_Links = Find_It_Links & BHTML & "<A HREF=""http://" & Find_IT_Server & "/Find_It.asp?Document=" & Item_Numbers(x) & "&SRC=FDL&Server=" & aServer & "&CMS_Site=" & CMS_Site & """>" & Titles(x) & "</A>" & EHTML
        end if
      
      end if
    
      rsAsset.Close
      set rsAsset = nothing
      
    next
    
    Call Disconnect_SiteWide
    
    Find_It_Link = Find_It_Links
    
  end if
  
end function
%>