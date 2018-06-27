<%

'strFTPFilePath = FTPFile(Server.MapPath("/" & oFileUpEE.form("Path_Site") & "/" & rsFile("File_Name")), rsFile("File_Name")) ''FTP the file

function CurrentTier()
    DIM strTier      
    strTier = "PRD"   ''Default
    if instr(UCase(Request.ServerVariables("SERVER_NAME")),"DTMEVTVSDV15") > 0 OR instr(UCase(Request.ServerVariables("SERVER_NAME")),"DEV") > 0  then
        strTier = "DEV"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"TEST") > 0 then
        strTier = "TEST"
    elseif instr(UCase(Request.ServerVariables("SERVER_NAME")),"PRD") > 0 then
        strTier = "PRD"
    end if 
    CurrentTier = strTier
end function


function FTPFile()
 
    Dim sFTPLoc, sFTPLocPrefix 
    sFTPLocPrefix = ""
    set obj = server.CreateObject("FTPclient.FTP")     
       if CurrentTier() = "DEV" then
        
       if(CInt(Site_ID) = 3)then           
            FTPFile = "content.fluke.com/dev/fluke"
             
            'RI#1705 -gpd
       elseif(CInt(Site_ID) = 29)then  
            '24 May 2017
            FTPFile = "content.fluke.com/Datapaq_Dev"
              
        'Biomed
        elseif(CInt(Site_ID) = 46)then           
            FTPFile = "content.fluke.com/Biomed_Dev"
            
        elseif(CInt(Site_ID) = 1 OR CInt(Site_ID) = 4 OR CInt(Site_ID) = 5)then           
            FTPFile = "content.fluke.com/FCal_Dev"
                      
        else
           FTPFile = ""
                                
        end if
        
    elseif CurrentTier() = "TEST"  then
                
        if(CInt(Site_ID) = 3)then         
             FTPFile = "content.fluke.com/Test/fluke"
            
            'RI#1705 -gpd
       elseif(CInt(Site_ID) = 29)then           
            FTPFile = "content.fluke.com/Test/datapaq"
            
       'Biomed
       elseif(CInt(Site_ID) = 46)then         
            FTPFile = "content.fluke.com/Biomed_Test"
            
       elseif(CInt(Site_ID) = 1 OR CInt(Site_ID) = 4 OR CInt(Site_ID) = 5)then           
            FTPFile = "content.fluke.com/FCal_Test"
            
        else
            FTPFile = ""
            
        end if
    elseif CurrentTier() = "PRD"  then
        if(CInt(Site_ID) = 3)then
             FTPFile = "download.fluke.com/pricelist"
            
            'RI#1705 -gpd
       elseif(CInt(Site_ID) = 29)then           
            FTPFile = "download.fluke.com/datapaq"
            
            'Biomed
       elseif(CInt(Site_ID) = 46)then         
            FTPFile = "download.fluke.com/Biomedical"
            
       elseif(CInt(Site_ID) = 1 OR CInt(Site_ID) = 4 OR CInt(Site_ID) = 5)then           
            FTPFile = "download.fluke.com/FCal"
           

        else
            FTPFile = ""
            
        end if
    else
	    'Default to PRODUCTION
	    if(CInt(Site_ID) = 3)then
	        FTPFile = "download.fluke.com/pricelist"
            
            'RI#1705 -gpd
       elseif(CInt(Site_ID) = 29)then           
            FTPFile = "download.fluke.com/datapaq"
            
            'Biomed
       elseif(CInt(Site_ID) = 46)then         
            FTPFile = "download.fluke.com/Biomedical"
         
       elseif(CInt(Site_ID) = 1 OR CInt(Site_ID) = 4 OR CInt(Site_ID) = 5)then           
            FTPFile = "download.fluke.com/FCal"   

         else
            FTPFile = ""
	        
         end if   
	end if 
	''End 
		
    'sFTPLoc = sFTPLocPrefix '& sFTPLocation
        
      
        
    'FTPFile = sFTPLocation
end function


%>