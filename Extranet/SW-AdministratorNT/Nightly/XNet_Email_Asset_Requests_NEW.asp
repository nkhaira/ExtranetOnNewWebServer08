<%                          ' Rem out for VBS version (also see end of script)
Option Explicit

Dim DevEnv
DevEnv = true               ' Set to True for Debuging with browser on EVTIBG18
'Dim Server                 ' Rem Out for DevEnv = False
Server.scripttimeout = 360  ' Rem out for VBS version

' --------------------------------------------------------------------------------------
' Title:  Externet Missing Asset Report Email
'
' Author: Kelly Whitlock
' Date:   01/16/2004
'
' --------------------------------------------------------------------------------------

'On Error Resume Next

' --------------------------------------------------------------------------------------

' Globals
Dim dbConnSiteWide

' Objects
Dim Mailer, rsSite, rsCostCenter, rsThumbs, rsEmail, rsPOD

' Numeric Counters

Dim e, x, y, Asset_Owner_Email_Max

' Strings
Dim Site_Root, SQLSite, SQLThumbs, SQLCostCenter, SQLPod, SQLEmail
Dim strBody
Dim Site_ID, Cost_Center, Asset_Owner, Asset_Owner_Email, Owner_Admin_List, Owner_Email_List, Name
Dim Edit_Content, View_Content
Dim ErrorMsg

' Arrays
Dim Status(3)
Status(0) = "<FONT STYLE=""font-size:8.5pt;font-weight:Normal;color:Black;background:#FFFF00;text-decoration:none;font-family:Arial,Verdana;"">&nbsp;In Review&nbsp;</FONT>"
Status(1) = "<FONT STYLE=""font-size:8.5pt;font-weight:Normal;color:Black;background:#00FF33;text-decoration:none;font-family:Arial,Verdana;"">&nbsp;Live&nbsp;</FONT>"
Status(2) = "<FONT STYLE=""font-size:8.5pt;font-weight:Normal;color:Black;background:#FF99CC;text-decoration:none;font-family:Arial,Verdana;"">&nbsp;Archived&nbsp;</FONT>"

Dim Lit_Name
'Lit_Name   = Split("PRT,PDF,POD,CD,DSK,DSP,VTN,VTP,TMB",",")
Lit_Name   = Split("PRT,PDF,POD,CD,TMB",",")

Dim Lit_Field
'Lit_Field  = Split("Lit_Print,Lit_PDF,Lit_POD,Lit_CD,Lit_Disks,Lit_Display,Lit_Video_NTSC,Lit_Video_PAL,Thumbnail",",")
Lit_Field  = Split("Lit_Print,Lit_PDF,Lit_POD,Lit_CD,Thumbnail",",")

Dim Lit_Names_Max
Lit_Names_Max = UBound(Lit_Field)

ReDim Lit_Type(2,Lit_Names_Max)

for x = 0 to Lit_Names_Max
  for y = 0 to 2
    Lit_Type(y,x) = "<FONT STYLE=""Background:"
    select case y
      case 0
        Lit_Type(y,x) = Lit_Type(y,x) & "#FFFFFF"
      case 1
        Lit_Type(y,x) = Lit_Type(y,x) & "#00CC00"
      case 2
        Lit_Type(y,x) = Lit_Type(y,x) & "#FFFF00"
    end select
    Lit_Type(y,x) = Lit_Type(y,x) & ";"">&nbsp;" & Lit_Name(x) & "&nbsp;</FONT>"
  next
next

Lit_Names_Max = Lit_Names_Max - 1       ' Do not include last array value (Thumbnail) since this is not a Boolean Value but we will use the Lit_Type(y,x) value.

Dim Asset_Status_Max
Asset_Status_Max = 5
Dim Asset_Status(5)       ' Set to match Asset_Status_Max

' Numbers
Dim xItem_Number, xFile_Name, xThumbnail, xFile_Name_POD, xRevision, xView
xItem_Number   = 0
xFile_Name     = 1
xThumbnail     = 2
xFile_Name_POD = 3
xRevision      = 4
xView          = 5

Dim Lit_Counter, Old_L_ID, Old_P_ID

' True/false
Dim Asset_Missing

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------

Call Connect_SiteWideDatabase

Call Main

if err.number <> 0 then
  ErrorMsg = ErrorMsg = "A fatal error has occured in Nightly\XNet_Email_Missing_Assets.vbs<BR>" & vbCrLf &_
                        "Error Number: " & err.number & "<BR>" & vbCrLf &_
                        "Error Description: " & err.description & "<BR>" & vbcrlf
	err.clear
end if

if not isblank(ErrorMsg) then  
  Call Send_Error_Email
end if

Call Disconnect_SiteWideDatabase

' --------------------------------------------------------------------------------------
' Subroutines and Functions
' --------------------------------------------------------------------------------------

sub Main

  ' If PRINT or POD, then PDF is required so Update DB Before beginning analysis for missing items

  SQLThumbs = "UPDATE dbo.Literature_Items_US SET PDF=-1 WHERE ([PRINT]=-1) OR (POD=-1) OR (CD=-1)"
  dbConnSiteWide.execute SQLThumbs

  ' Now get records for analysis

SQLThumbs = "SELECT   dbo.Lit_Cost_Center.Site_ID AS Site_ID, dbo.Literature_Items_US.ITEM AS L_Item, dbo.Literature_Items_US.REVISION AS L_Revision,  " &_
            "         dbo.Literature_Items_US.[ACTION] AS Action, dbo.Literature_Items_US.STATUS AS L_Status, dbo.Literature_Items_US.END_USER AS L_End_User,  " &_
            "         dbo.Literature_Items_US.STATUS_NAME AS L_Status_Name, dbo.Literature_Items_US.Literature_Type AS L_Category, dbo.Literature_Items_US.INVENTORY_RULE AS L_Inventory_Rule,  " &_
            "         dbo.Literature_Items_US.INVENTORY_RULE_CODE AS L_Inventory_Rule_Code, dbo.Literature_Items_US.[LANGUAGE] AS L_Language, " &_ 
            "         dbo.Calendar.[Language] AS P_Language, dbo.Literature_Items_US.COST_CENTER AS Cost_Center,  " &_
            "         dbo.Literature_Items_US.DESCRIPTION AS Description, dbo.Literature_Items_US.MARCOM_ITEM_DESC AS Description_Marcom,  " &_
            "         dbo.Literature_Items_US.EFULFILLMENT AS Description_EDF, dbo.Literature_Items_US.[PRINT] AS Lit_Print, dbo.Literature_Items_US.PDF AS Lit_PDF, " &_
            "         dbo.Literature_Items_US.WEB AS Lit_Web, dbo.Literature_Items_US.CD AS Lit_CD, dbo.Literature_Items_US.DISPLAY AS Lit_Display,  " &_
            "         dbo.Literature_Items_US.DISKS AS Lit_Disks, dbo.Literature_Items_US.VIDEO_NTSC AS Lit_Video_NTSC,  " &_
            "         dbo.Literature_Items_US.VIDEO_PAL AS Lit_Video_PAL, dbo.Literature_Items_US.POD AS Lit_POD,  " &_
            "         dbo.Literature_Items_US.MARCOM_MANAGER AS L_Owner, dbo.Calendar.ID AS P_ID, dbo.Calendar.Item_Number AS P_Item,  " &_
            "         dbo.Calendar.Site_ID AS P_Site_ID, dbo.Calendar.Revision_Code AS P_Revision, dbo.Calendar.File_Name AS File_Name,  " &_
            "         dbo.Calendar.File_Name_POD AS File_Name_POD, dbo.Calendar.Thumbnail AS Thumbnail, dbo.Calendar.Status AS P_Status,  " &_
            "         dbo.Calendar.Campaign AS P_Campaign, dbo.Calendar.Submitted_By AS P_Owner_ID,  " &_
            "         dbo.UserData.LastName + ', ' + dbo.UserData.FirstName AS P_Owner, dbo.UserData.Email AS P_Email, dbo.Calendar.SubGroups AS Subgroups " &_
            "FROM     dbo.Lit_Cost_Center RIGHT OUTER JOIN " &_
            "         dbo.Literature_Items_US ON dbo.Lit_Cost_Center.Cost_Center = dbo.Literature_Items_US.COST_CENTER LEFT OUTER JOIN " &_
            "         dbo.UserData INNER JOIN " &_
            "         dbo.Calendar ON dbo.UserData.ID = dbo.Calendar.Submitted_By ON dbo.Literature_Items_US.ITEM = dbo.Calendar.Item_Number " &_
            "WHERE   (dbo.Literature_Items_US.STATUS = 'Active') AND (dbo.Literature_Items_US.[ACTION] = 'Complete') AND  " &_
            "        (dbo.Literature_Items_US.STATUS_NAME = 'Final Loaded') AND (dbo.Lit_Cost_Center.Site_ID <> 19) AND (dbo.Lit_Cost_Center.Site_ID <> 80) AND  " &_
            "        (dbo.Literature_Items_US.COST_CENTER <> 'No Value') AND (dbo.Literature_Items_US.INVENTORY_RULE_CODE <> 'Destroy') OR " &_
            "        (dbo.Literature_Items_US.STATUS = 'Active') AND (dbo.Literature_Items_US.[ACTION] = 'Complete') AND  " &_
            "        (dbo.Literature_Items_US.STATUS_NAME = 'Reprint') AND (dbo.Lit_Cost_Center.Site_ID <> 19) AND (dbo.Lit_Cost_Center.Site_ID <> 80) AND  " &_
            "        (dbo.Literature_Items_US.COST_CENTER <> 'No Value') AND (dbo.Literature_Items_US.INVENTORY_RULE_CODE <> 'Destroy') " &_
            "ORDER BY dbo.Literature_Items_US.COST_CENTER, dbo.Literature_Items_US.[ACTION], dbo.Literature_Items_US.ITEM "
            
'response.write sqlthumbs
'response.end            

  if DevEnv then
    Set rsThumbs = Server.CreateObject("ADODB.Recordset")
  else
    Set rsThumbs = CreateObject("ADODB.Recordset")
  end if      
  rsThumbs.Open SQLThumbs, dbConnSiteWide, 3, 3

  strBody = ""
  
  ' --------------------------------------------------------------------------------------

  if not rsThumbs.EOF then
  
    Cost_Center = -1
    Lit_Counter = 0
    Old_P_ID    = null
    Old_L_ID    = null
    
    do while not rsThumbs.EOF
    
      if isnull(rsThumbs("P_ID")) or rsThumbs("P_ID") <> Old_P_ID then
      
        if Old_L_ID <> rsThumbs("L_Item") then
        
          'if rsThumbs("Lit_PDF")        = CInt(true) or _
          '   rsThumbs("Lit_Print")      = CInt(true) or _
          '   rsThumbs("Lit_POD")        = CInt(true) or _
          '   rsThumbs("Lit_CD")         = CInt(true) or _
          '   rsThumbs("Lit_Disks")      = CInt(true) or _         
          '   rsThumbs("Lit_Display")    = CInt(true) or _
          '   rsThumbs("Lit_Video_NTSC") = CInt(true) or _
          '   rsThumbs("Lit_Video_PAL")  = CInt(true) then
             
          '   Lit_Counter = Lit_Counter + 1
  
          'end if
           
          if rsThumbs("Lit_PDF")        = CInt(true) or _
             rsThumbs("Lit_Print")      = CInt(true) or _
             rsThumbs("Lit_POD")        = CInt(true) then
             
             Lit_Counter = Lit_Counter + 1
  
           end if


        end if
         
      end if

      Old_L_ID = rsThumbs("L_Item")   
      Old_P_ID = rsThumbs("P_ID")   
       
      rsThumbs.MoveNext
      
    loop
    
'response.write Lit_Counter
'response.end

    rsThumbs.MoveFirst  

    Old_P_ID    = null
    Old_L_ID    = null
      
    do while not rsThumbs.EOF
    
      if (isnull(rsThumbs("P_ID")) and rsThumbs("L_Item") <> Old_L_ID) or rsThumbs("P_ID") <> Old_P_ID and rsThumbs("L_Item") <> Old_L_ID then    
    
        if LCase(rsThumbs("L_Language"))     = "eng" _
           and CInt(rsThumbs("L_End_User"))  = CInt(True) _
           and CInt(rsThumbs("Lit_Display")) = CInt(false) _
           and CInt(rsThumbs("Lit_CD")) = CInt(false) _
           and CInt(rsThumbs("Lit_Disks")) = CInt(false) _
           and CInt(rsThumbs("Lit_Video_NTSC")) = CInt(false) _
           and CInt(rsThumbs("Lit_Video_PAL")) = CInt(false) _                      
           or LCase(rsThumbs("L_Language"))  = LCase(rsThumbs("P_Language")) _
           and CInt(rsThumbs("L_End_User"))  = CInt(True) _
           and CInt(rsThumbs("Lit_Display")) = CInt(false) _
           and CInt(rsThumbs("Lit_CD")) = CInt(false) _
           and CInt(rsThumbs("Lit_Disks")) = CInt(false) _
           and CInt(rsThumbs("Lit_Video_NTSC")) = CInt(false) _
           and CInt(rsThumbs("Lit_Video_PAL")) = CInt(false) then                      
               
          for x = 0 to Asset_Status_Max
            Asset_Status(x) = false
          next  
    
          ' Determine Status of Extranet Asset (if any one true, PP asset container is required)
          
          if rsThumbs("Lit_PDF")        = CInt(true) or _
             rsThumbs("Lit_Print")      = CInt(true) or _
             rsThumbs("Lit_POD")        = CInt(true) or _
             rsThumbs("Lit_CD")         = CInt(true) then
             
             'rsThumbs("Lit_Disks")      = CInt(true) or _
             'rsThumbs("Lit_Display")    = CInt(true) or _
             'rsThumbs("Lit_Video_NTSC") = CInt(true) or _
             'rsThumbs("Lit_Video_PAL")  = CInt(true) then
    
            if CInt(Cost_Center) <> CInt(rsThumbs("Cost_Center")) then
          
              if Cost_Center <> -1 then
                strBody = strBody & vbCrLf
              end if
              
              ' This is where the last cost center email is sent out.
              
              if not isblank(strBody) then
              
                Call Get_Mail_Body
                Call Send_Mail
                
                strBody          = ""
                Asset_Owner      = ""
                Owner_Admin_List = ""
                Owner_Email_List = ""
              
              end if
              
              ' New cost center report starts here.
            
              Cost_Center = rsThumbs("Cost_Center")          
              
              strBody = strBody & "<FONT Style=""color:#FFCC00;background:#000000;font-weight:bold;font-size:12pt;font-family:Arial,Verdana;"">" &_
                                  "&nbsp;&nbsp;Cost Center: " & Cost_Center & "&nbsp;&nbsp;</FONT><BR>"
  
              SQLCostCenter = "SELECT dbo.Site.Site_Description AS Description " &_
                              "FROM   dbo.Lit_Cost_Center LEFT OUTER JOIN " &_
                              "       dbo.Site ON dbo.Lit_Cost_Center.Site_ID = dbo.Site.ID " &_
                              "WHERE (dbo.Lit_Cost_Center.Cost_Center=" & Cost_Center & ") " &_
                              "ORDER BY dbo.Lit_Cost_Center.Site_ID"
    
              if DevEnv then
                Set rsCostCenter = Server.CreateObject("ADODB.Recordset")
              else
                Set rsCostCenter = CreateObject("ADODB.Recordset")
              end if      
              rsCostCenter.Open SQLCostCenter, dbConnSiteWide, 3, 3
              
              strBody = strBody & "<FONT Style=""font-weight:Bold; font-size:10pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">"
              do while not rsCostCenter.EOF
                strBody = strBody & rsCostCenter("Description") & "<BR>"
                rsCostCenter.MoveNext
              loop
              strBody = strBody & "</FONT><BR>" & vbCrLf 
                          
              rsCostCenter.close
              set rsCostCenter  = nothing
              set SQLCostCenter = nothing
              
            end if  
            
            if isblank(rsThumbs("P_Item")) then                                         ' If P_Item is NULL then asset does not exsist on the Extranet
              Asset_Status(xItem_Number) = True
            else
              Asset_Status(xItem_Number) = False  
            end if
                
            ' False if Asset Container Exists on the Extranet, So Do the Rest of the Checks
                
            if CInt(Asset_Status(xItem_Number)) = CInt(False) then
    
              if CInt(rsThumbs("P_Status")) <= 1 then       ' In-Review or Live Only
      
                ' Check for PDF / POD File(s) and Thumbnail
                    
                if rsThumbs("Lit_PDF")   = CInt(True) or _
                   rsThumbs("Lit_Print") = CInt(True) or _
                   rsThumbs("Lit_CD")    = CInt(True) or _
                   rsThumbs("Lit_POD")   = CInt(True) then  ' These assets require PDF and Thumbnail
                   
                  if not isblank(rsThumbs("File_Name")) then
                    Asset_Status(xFile_Name) = False
                  else
                    Asset_Status(xFile_Name) = True
                  end if
                  
                  if not isblank(rsThumbs("Thumbnail")) then
                    Asset_Status(xThumbnail) = False
                  else
                    Asset_Status(xThumbnail) = True
                  end if
                  
                  ' Check for POD (Special Case)
                  
                  if rsThumbs("Lit_POD") = CInt(True) then  ' Search DB for POD File (Must Be Live)
                  
                    SQLPod = "SELECT dbo.Calendar.Subgroups, dbo.Calendar.File_Name_POD FROM dbo.Calendar WHERE dbo.Calendar.Item_Number='" & rsThumbs("L_Item") & "' AND dbo.Calendar.Status=1 AND dbo.Calendar.File_Name_POD IS NOT NULL"
                    if DevEnv then
                      Set rsPOD = Server.CreateObject("ADODB.Recordset")
                    else
                      Set rsPOD = CreateObject("ADODB.Recordset")
                    end if      

                    rsPOD.Open SQLPod, dbConnSiteWide, 3, 3

                    if not rsPOD.EOF then
                      Asset_Status(xFile_Name_POD) = False
                      if instr(1,LCase(rsPOD("Subgroups")),"view") > 0 or instr(1,LCase(rsPOD("Subgroups")),"fedl") > 0 then
                        Asset_Status(xView) = False
                      else
                        Asset_Status(xView) = True
                      end if  
                    else
                      Asset_Status(xFile_Name_POD) = True
                    end if
                    
                    rsPod.close
                    set RsPod  = nothing
                    set SQLPod = nothing
                  end if  
                    
                end if    

                ' Check for Thumbnails

                if rsThumbs("Lit_CD")         = CInt(true) or _
                   rsThumbs("Lit_Disks")      = CInt(true) or _            
                   rsThumbs("Lit_Display")    = CInt(true) or _
                   rsThumbs("Lit_Video_NTSC") = CInt(true) or _
                   rsThumbs("Lit_Video_PAL")  = CInt(true) then  ' These Lit Items only require a Thumbnail as a minimum.
                   
                  if not isblank(rsThumbs("Thumbnail")) then
                    Asset_Status(xThumbnail) = False
                  else
                    Asset_Status(xThumbnail) = True
                  end if
                end if   

                ' Compare Revision Codes
                
                if not isblank(rsThumbs("P_Revision")) then
                  if Replace(UCase(trim(rsThumbs("L_Revision"))),"REV ","") = Replace(UCase(trim(rsThumbs("P_Revision"))),"REV ","") then
                    Asset_Status(xRevision) = False
                  else
                    Asset_Status(xRevision) = True
                  end if
                else
                  Asset_Status(xRevision) = True
                end if
                
              end if
            end if
            
            ' Check for Any Missing Assets
            
            Asset_Missing = False
            for x = 1 to Asset_Status_Max                 ' Bypass Revision (0)
              if CInt(Asset_Status(x)) = CInt(True) then
                Asset_Missing = True
                exit for
              end if
            next
            
            'xItem_Number   = 0
            'xFile_Name     = 1
            'xThumbnail     = 2
            'xFile_Name_POD = 3
            'xRevision      = 4
            'xView          = 5
    
    '        for x = 1 to Asset_Status_Max                 ' Bypass Revision (0)
    '          response.write x & " " & Asset_Status(x)  & "<BR>"
    '        next
    
    
            ' Missing Assets then Add Detail to Report
            
            if CInt(Asset_Missing) = CInt(True) then    
    
              strBody = strBody & "<FONT Style=""font-weight:Normal; font-size:8.5pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">" &vbCrLf &_
                                  "Item Number: <B>" & rsThumbs("L_Item") & "</B><BR>" & vbCrLf &_
                                  "Revision: Oracle <B>" & UCase(rsThumbs("L_Revision")) & "</B>" & vbCrLf
  
              if CInt(Asset_Status(xRevision)) = CInt(True) then
                strBody = strBody & " <FONT COLOR=""RED""><B>Mismatch</B></FONT> Portal Asset Revision: <B>" & UCase(rsThumbs("P_Revision")) & "</B>" & vbCrLf
              end if
              
              strBody = strBody & "<BR>" &_
              "Literature Type: <B>" & rsThumbs("L_Category") & "</B><BR>" & vbCrLf &_                                     
              "Description: " & rsThumbs("Description") & "<BR>" & vbCrLf &_
              "Description EDF: " & rsThumbs("Description_EDF") & "<BR>" & vbCrLf &_
              "Missing: <FONT COLOR=""RED""><B>"
              
              if CInt(Asset_Status(xFile_Name)) = CInt(True) then
                strBody = strBody & "Low-Resolution PDF File, "
              end if  

              if CInt(Asset_Status(xFile_Name_POD)) = CInt(True) then
                strBody = strBody & "High-Resolution POD File, "
              end if  

              if CInt(Asset_Status(xThumbnail)) = CInt(True) then
                strBody = strBody & "Thumbnail, "
              end if
              
              if Mid(strBody,Len(strBody)-1,2) = ", " then
                strBody = Mid(strBody,1,Len(strBody)-2)
              end if  
              
              strBody = strBody & "</B></FONT><BR>" & vbCrLf
              
              ' Low Resolution PDF or High Resolution POD must be enabled for View to work with EDF and Find_It
              
              if (CInt(Asset_Status(xFile_Name)) = CInt(False)  or _
                  CInt(Asset_Status(xFile_Name_POD)) = CInt(False)) and _
                  CInt(Asset_Status(xView)) = CInt(True) then
                
                  strBody = strBody & "End User Viewable: <FONT COLOR=""Red"">Must be checked to be available to the "
                  if CInt(Asset_Status(xFile_Name)) = CInt(False) then strBody = strBody & "EDF "
                  if CInt(Asset_Status(xFile_Name_POD)) = CInt(False) then strBody = strBody & " and POD "
                  strBody = strBody & " Fulfillment System.</FONT><BR>"
              end if
              
              if instr(1,Asset_Owner,LCase(rsThumbs("P_Owner"))) = 0 then
                if not isblank(Asset_Owner) then
                  Asset_Owner = Asset_Owner & "|"
                end if
                Asset_Owner = Asset_Owner & LCase(rsThumbs("P_Owner"))
              end if
              
              Edit_Content = "http://Support.Fluke.com/SW-Administrator/Calendar_Edit.asp?ID=" & rsThumbs("P_ID") & "&Site_ID=" & rsThumbs("P_Site_ID")
              View_Content = "http://Support.Fluke.com/Find_It.asp?Document=" & rsThumbs("P_Item")
              
              strBody = strBody & "Owner: "    & rsThumbs("P_Owner") & "<BR>" & vbCrLf &_
                                  "Asset ID: " & rsThumbs("P_ID") & "&nbsp;&nbsp;&nbsp;" & vbCrLf
                                  
              if DevEnv then                                                    
              strBody = strBody & "<A HREF=""javascript:void(0);"" onclick=""window.open('" & View_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""Click to View this Asset"">"
              else
              strBody = strBody & "<A HREF=""" & View_Content & """ TITLE=""Click to View this Asset"">"
              end if
              strBody = strBody & "View</A>&nbsp;&nbsp;&nbsp;" & vbCrLf
              
              if DevEnv then                                                    
              strBody = strBody & "<A HREF=""javascript:void(0);"" onclick=""window.open('" & Edit_Content & "','MyPop1','fullscreen=no,toolbar=yes,status=no,menubar=no,scrollbars=yes,resizable=yes,directories=no,location=no,width=760,height=560,left=250,top=20'); MyPop1.focus(); return false;"" TITLE=""Click to View this Asset"">"
              else
              strBody = strBody & "<A HREF=""" & Edit_Content & """ TITLE=""Click to View this Asset"">"
              end if
              
              strBody = strBody & "Edit Asset</A><BR>" & vbCrLf &_
                                  "Status: " & Status(rsThumbs("P_Status")) & vbCrLf &_
                                  "<BR>" & vbCrLf

              strBody = strBody & "Asset Container Digital File Requirements: "
                        
              for x = 0 to Lit_Names_Max
                Asset_Missing = 0        
                select case LCase(Lit_Field(x))
                  case "lit_print"
                    Asset_Missing = ABS(CInt(rsThumbs("Lit_Print"))) + ABS(CInt(Asset_Status(xFile_Name)))
                  case "lit_pdf"
                    if CInt(rsThumbs("Lit_Print")) = CInt(True) or _
                       CInt(rsThumbs("Lit_POD")) = CInt(True)   or _
                       CInt(rsThumbs("Lit_PDF")) = CInt(True)   then
                        Asset_Missing = 1 + ABS(CInt(Asset_Status(xFile_Name)))
                    end if    
                  case "lit_pod"
                    Asset_Missing = ABS(CInt(rsThumbs("Lit_POD"))) + ABS(CInt(Asset_Status(xFile_Name_POD)))
                  case else
                    Asset_Missing = ABS(CInt(rsThumbs(Lit_Field(x))))
                end select
                                                                
                strBody = strBody & Lit_Type(Asset_Missing,x)
              next
              
              ' check thumbnail by different method then above
              
              select case CInt(Asset_Status(xThumbnail))
                case CInt(False)
                  strBody = strBody & Lit_Type(1,Lit_Names_Max + 1)
                case CInt(True)
                  strBody = strBody & Lit_Type(2,Lit_Names_Max + 1)
                case else
                  strBody = strBody & Lit_Type(0,Lit_Names_Max + 1)
              end select
              
              strBody = strBody & "<BR>" & vbCrLf
              
              strBody = strBody & "Oracle Status: "
              
              strBody = strBody & rsThumbs("Action")
              if not isblank(rsThumbs("L_Inventory_Rule")) then
                strBody = strBody & " | " & rsThumbs("L_Inventory_Rule")
              end if
                                                
              strBody = strBody & "<P>" & vbCrLf        
    
            ' Asset does not exsist on Externet
            
            elseif CInt(Asset_Status(xItem_Number)) = CInt(True) then  
            
              strBody = strBody & "<FONT Style=""font-weight:Normal; font-size:8.5pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">" &vbCrLf &_
                                  "Item Number: <B>" & rsThumbs("L_Item") & "</B><BR>" & vbCrLf &_
                                  "Revision: Oracle <B>" & rsThumbs("L_Revision") & "</B><BR>" & vbCrLf &_
                                  "Literature Type: <B>" & rsThumbs("L_Category") & "</B><BR>" & vbCrLf &_
                                  "Description: " & rsThumbs("Description") & "<BR>" & vbCrLf &_
                                  "Description EDF: " & rsThumbs("Description_EDF") & "<BR>" & vbCrLf &_
                                  "Missing: <FONT COLOR=""RED""><B>No Asset Container was found for this Item Number.</B></FONT><BR>" &_
                                  "Owner: "
  
                if isblank(rsThumbs("L_Owner")) then
                  strBody = strBody & "Unknown" & "<BR>" & vbCrLf
                else
                  if instr(1,Asset_Owner,LCase(rsThumbs("L_Owner"))) = 0 then
                    if not isblank(Asset_Owner) then
                      Asset_Owner = Asset_Owner & "|"
                    end if
                    Asset_Owner = Asset_Owner & LCase(rsThumbs("L_Owner"))
                  end if
                  strBody = strBody & rsThumbs("L_Owner") & "<BR>" & vbCrLf
                end if  
    
              strBody = strBody & "Asset Container Digital File Requirements: "
              
                for x = 0 to Lit_Names_Max
                  select case CInt(rsThumbs(Lit_Field(x)))
                    case CInt(True)
                      strBody = strBody & Lit_Type(2,x)                                
                    case else
                      strBody = strBody & Lit_Type(0,x)
                  end select  
                next
                                    
              strBody = strBody & Lit_Type(2,Lit_Names_Max + 1)
              
              strBody = strBody & "<BR>" & vbCrLf
              
              strBody = strBody & "Oracle Status: "
              
              strBody = strBody & rsThumbs("Action")
              if not isblank(rsThumbs("L_Inventory_Rule")) then
                strBody = strBody & " | " & rsThumbs("L_Inventory_Rule")
              end if
  
              strBody = strBody & "<P>" & vbCrLf                                        
  
            end if
            
          end if
          
        end if
        
      end if      
                              
      Old_P_ID = rsThumbs("P_ID")
      Old_L_ID = rsThumbs("L_Item")      
      
      rsThumbs.MoveNext         

    loop              
        
    rsThumbs.close
    set rsThumbs = nothing
    
    if not isblank(strBody) then
      Call Get_Mail_Body
      Call Send_Mail
    end if
    
  end if                  

end sub

' --------------------------------------------------------------------------------------

sub Get_Mail_Body()

  strBody = "<HTML>" & vbCrLf & "<BODY>" & vbCrLf &_
            "<FONT Style=""font-weight:Normal; font-size:10pt;font-family: Arial,Verdana;color:#000000; font-style: Normal;"">" &_
            "You are receiving this notification because you are listed as the Site Administrator of a Partner Portal Site on Support.Fluke.com, a Content Administrator with an incomplete Asset Container, or the owner of the Literature Item as specified in the Oracle Literature Database.<P>" &_
            "The following is an automated weekly report that lists literature items with item numbers that have missing assets or do not have an asset container on the Partner Portal Extranet site.<P>" & vbCrLf &_
            "<UL><LI>To edit the asset container, click on the <U>Edit Asset</U>.  The first time you click on the <U>Edit Asset</U> link. You may have to log into your account, then click on the <U>Edit Asset</U> link again.</LI><P>" & vbCrLf &_
            "<LI>If there is not an asset container noted (e.g., no <U>Edit Asset</U> link), you should create an Asset Container for the Literature Item on the respective Partner Portal site.  These Asset Containers are used to supply the various files and images to the Fluke.com, the Digital Library, the Electronic Document Fulfillment Center (EDF), the Print-On-Demand Print Center (POD), the Partner Portal Literature Order System (LOS), the Partner Portal Email Subscription Service, and for users viewing literature items on the Partner Portal Sites at Support.Fluke.com</LI>" &_
            "</UL></FONT><P>" &_
            "<TABLE BORDER=0>" &_
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">XXX</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Not Required</FONT></TD></TR>" &_
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;background:#FFFF00;"">XXX</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Asset Required / Not Found</FONT></TD></TR>" &_                
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;background:#339900;"">XXX</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Asset Required / Found</FONT></TD></TR>" &_
            "</TABLE><P>" &_
            "<TABLE BORDER=0>" &_
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">PRT</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Print Version - Fulfillment through DCG (Requires PDF Version for Download)</FONT></TD></TR>" &_
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">PDF</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Low  - Resolution PDF File or other Dital Asset File - Fulfillment through Support.Fluke.com</FONT></TD></TR>" &_                
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">POD</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">High - Resolution PDF File - Fulfillment through Fluke Park Print Services (Requires LR PDF Version for POD System)</FONT></TD></TR>" &_                
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">CD </FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">CD   - Fulfillment through DCG (Only ""CD Lable"" Low - Resolution PDF and Thumbnail Required)</FONT></TD></TR>" &_
            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">THB</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Thumbnail Image File (JPG / 72dpi / 80px Wide / Xpx Height)</FONT></TD></TR>" &_                
            "</TABLE><P>" &_
           
            strBody &_
            "</BODY>" & vbCrLf & "</HTML>" & vbCrLf
            
'            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">DSK</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Floppy Disk - Fulfillment through DCG (Only Thumbnail Required)</FONT></TD></TR>" &_
'            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">DSP</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Display - Fulfillment through DCG (Only Thumbnail Required)</FONT></TD></TR>" &_                
'            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">VTN</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Video NTSC Format - Fulfillment through DCG (Only Thumbnail Required)</FONT></TD></TR>" &_                
'            "<TR><TD><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">VTN</FONT></TD><TD><FONT STYLE=""font-size:8.5pt;font-family: Arial;"">Video PAL  Format - Fulfillment through DCG (Only Thumbnail Required)</FONT></TD></TR>" &_                            

end sub

' --------------------------------------------------------------------------------------

sub Send_Mail
            
  Call Connect_Mailer

    'Mailer.FromName    = "Digital Library @ Support.Fluke.com"
    'Mailer.FromAddress = "Webmaster@fluke.com"

    msg.From = """Digital Library @ Support.Fluke.com""" & "Webmaster@fluke.com"
    ' Notify Site Administrator(s)
    
    SQLEmail = "SELECT DISTINCT " &_
               "       dbo.Lit_Cost_Center.Site_ID AS Site_ID, dbo.Site.MailToName AS Owner, dbo.Site.MailTo AS Owner_Email, dbo.Site.MailCCName AS CC_Owner, " &_
               "       dbo.Site.MailCC AS CC_Owner_Email " &_
               "FROM   dbo.Lit_Cost_Center " &_
               "LEFT OUTER JOIN " &_
                       "dbo.Site ON dbo.Lit_Cost_Center.Site_ID = dbo.Site.ID " & _
               "WHERE dbo.Lit_Cost_Center.Cost_Center='" & Cost_Center & "'"

    if DevEnv then
      Set rsEmail = Server.CreateObject("ADODB.Recordset")
    else
      Set rsEmail = CreateObject("ADODB.Recordset")
    end if              

    rsEmail.Open SQLEmail, dbConnSiteWide, 3, 3

    On error resume next
    Site_ID = rsEmail("Site_ID")
    if err.number <> 0 then
      Site_ID = 90
    end if
    on error goto 0
      
    if not rsEmail.EOF then
      if not DevEnv then    
        'Mailer.AddRecipient rsEmail("Owner"), rsEmail("Owner_Email")
        if not isblank(rsEmail("Owner_CC_Email")) then
          'Mailer.AddRecipient rsEmail("Owner_CC"), rsEmail("Owner_CC_Email")        
        end if  
      else
        'Owner_Admin_List = Owner_Admin_List & rsEmail("Owner") & " " & rsEmail("Owner_Email") & "<BR>"
      end if    
    end if
      
    rsEmail.close
    set rsEmail  = nothing
    set SQLEmail = nothing
    
    ' Notify Literature Item Owners

    if not isblank(Asset_Owner) then
    
      Asset_Owner_Email     = Split(Asset_Owner,"|")
      Asset_Owner_Email_Max = UBound(Asset_Owner_Email)

      for e = 0 to Asset_Owner_Email_Max
      
        if instr(1,Asset_Owner_Email(e),",") > 0 then

          Name = Split(Asset_Owner_Email(e),",")
          Name(1) = Mid(Trim(Name(1)),1,2) ' First 2 Initial
            
          SQLEmail = "SELECT  DISTINCT LastName, FirstName + ' ' + LastName AS Owner, Email AS Owner_Email " &_
                     "FROM    dbo.UserData " &_
                     "WHERE   (LastName = N'"  & Trim(Name(0)) & "') AND " &_
                     "        (FirstName LIKE N'" & Trim(Name(1)) & "%') AND " &_
                     "        (Site_ID=" & Site_ID & ") AND " &_
                     "        (SubGroups LIKE '%administrator%' OR SubGroups LIKE '%content%') ORDER By LastName"

          if DevEnv then
            Set rsEmail = Server.CreateObject("ADODB.Recordset")
          else
            Set rsEmail = CreateObject("ADODB.Recordset")
          end if              
    
          rsEmail.Open SQLEmail, dbConnSiteWide, 3, 3

          if not rsEmail.EOF then
            if not DevEnv then    
              'Mailer.AddRecipient rsEmail("Owner"), rsEmail("Owner_Email")
            else
              'Owner_Email_List = Owner_Email_List & rsEmail("Owner") & " " & rsEmail("Owner_Email") & "<BR>" & vbCrLf
            end if  
          end if
          
          rsEmail.close
          set rsEmail  = nothing
          set SQLEmail = nothing
  
        end if
        
      next
      
    end if
    
    if DevEnv then
      strBody = "<B><FONT STYLE=""font-size:8.5pt;font-family:Arial;"">" & Owner_Admin_List & Owner_Email_List & "</FONT></B><P>" & strBody
    end if  
    
      
    ' Domain Administrator
    
    'Mailer.AddBCC "Kelly Whitlock","Kelly.Whitlock@Fluke.com"
    msg.Bcc = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"

    'Mailer.Subject = "Cost Center: " & Cost_Center & " " & "Digital Library Missing Asset Report - " & Date()
    'Mailer.BodyText =  strBody
    
    msg.Subject = "Cost Center: " & Cost_Center & " " & "Digital Library Missing Asset Report - " & Date()
    msg.TextBody = strBody

    'if Mailer.SendMail then
    'else
    '  ErrorMsg = ErrorMsg & "Send Email Failure<BR><BR>" & "Error Description: " & Mailer.Response & ". "
    'end if

    msg.Configuration = conf
    On Error Resume Next
    msg.Send
    If Err.Number = 0 then
      'Success
    Else
      ErrorMsg = ErrorMsg & "Send Email Failure<BR><BR>" & "Error Description: " & Err.Description & ". "
    End If
    
    if DevEnv then
      response.write ErrorMsg & "<P>"     
      response.write strBody
    end if
      
  Call Disconnect_Mailer
    
end sub

' --------------------------------------------------------------------------------------

sub Connect_SiteWideDatabase()

	Dim strConnectionString_SiteWide
	
	set dbConnSiteWide = CreateObject("ADODB.Connection")
	
	if DevEnv then
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=EVTIBG18.DEV.IB.FLUKE.COM; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	else
		strConnectionString_SiteWide = "Driver={SQL Server}; SERVER=FLKPRD18.DATA.IB.FLUKE.COM; " &_
			"UID=sitewide_email;DATABASE=fluke_SiteWide;pwd=f6sdW"
	end if
	
	dbConnSiteWide.ConnectionTimeOut = 120
	dbConnSiteWide.CommandTimeout = 120
	dbConnSiteWide.Open strConnectionString_SiteWide

end sub

' --------------------------------------------------------------------------------------

sub Disconnect_SiteWideDatabase()

	if IsObject(dbConnSiteWide) then
		dbConnSiteWide.Close
		set dbConnSiteWide = nothing
	end if

end sub

' --------------------------------------------------------------------------------------

sub Connect_Mailer

  'Set Mailer = CreateObject("SMTPsvg.Mailer")
  'adding new email method
  %>
  <!--#include virtual="/connections/connection_email_new.asp"-->
  <%

  if DevEnv then
    'Mailer.RemoteHost = "mailhost.tc.fluke.com"
  else
    'Mailer.RemoteHost = "mail.evt.danahertm.com:25"
  end if

  'Mailer.ReturnReceipt = false
  'Mailer.ConfirmRead = false
  'Mailer.WordWrap = True
  'Mailer.WordWrapLen = 85
  'Mailer.QMessage = True
  'Mailer.ClearAttachments
  'Mailer.ContentType = "text/html; charset=us-ascii"
 
end sub

' --------------------------------------------------------------------------------------

sub Disconnect_Mailer

  'set Mailer = Nothing

end sub

' --------------------------------------------------------------------------------------

sub Send_Error_Email

  if not isblank(ErrorMsg) then
  
    Call Connect_Mailer

      'Mailer.FromName    = "EVTIBG01 admin"
      'Mailer.FromAddress = "webmaster@fluke.com"

      msg.From = """EVTIBG01 admin""" & "webmaster@fluke.com"
  
      'Mailer.AddRecipient "Kelly Whitlock","Kelly.Whitlock@fluke.com"
      msg.To = """Santosh Tembhare""" & "santosh.tembhare@fluke.com"
  
    	'Mailer.Subject     = "Error: Nightly - Extranet Missing Asset Report - " & Date()
    	'Mailer.BodyText    = ErrorMsg

      msg.Subject = "Error: Nightly - Extranet Missing Asset Report - " & Date()
      msg.TextBody = ErrorMsg

      'Mailer.SendMail
      msg.Configuration = conf
      On Error Resume Next
      msg.Send
      If Err.Number = 0 then
        'Success
      Else
        'Fail
      End If
    
    Call Disconnect_Mailer    
  
  end if
  
end sub

' --------------------------------------------------------------------------------------

function IsBlank(MyString)

  if isnull(MyString) then
    IsBlank = True
  elseif not isnull(MyString) then
    if Len(Trim(MyString)) = 0 then
      IsBlank = True
    elseif UCase(Trim(MyString)) = "NO VALUE" then
      IsBlank = True
    else
      IsBlank = False
    end if
  else
    IsBlank = False
  end if
  
end function

' --------------------------------------------------------------------------------------
'%>