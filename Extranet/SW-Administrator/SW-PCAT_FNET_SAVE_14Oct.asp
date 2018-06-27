<!--#include virtual="/sw-administrator/SW-PCAT_FNET_IISSERVER.asp"-->
<!--#include virtual="/include/functions_string.asp"-->

<%
Dim SaveRecord, sSql, iPID, sRelations
SaveRecord=true
''on error resume next
sRelations = "0"

    optiontext = false
    iPID = 0
    
    if (Site_ID = 3 Or Site_ID = 46) then
        if IsNumeric(CStr(oFileUpEE.form("ID"))) then
            SQL = "SELECT Item_Number, Count(*) AS Counts FROM dbo.Calendar WHERE (Item_Number = '" & CStr(oFileUpEE.form("Item_Number")) & "') GROUP BY Item_Number"
            Set rsAllItems = Server.CreateObject("ADODB.Recordset")
            rsAllItems.Open SQL, conn, 3, 3

            if not rsAllItems.EOF then
	            if(isnumeric(rsAllItems("Counts")) and CInt(rsAllItems("Counts")) > 1) then
		            SQL = "SELECT [ID], Item_Number, Revision_Code, PID FROM dbo.Calendar WHERE (Item_Number = '" & CStr(oFileUpEE.form("Item_Number")) & "') AND PID > 0"
		            Set rsDuplicateItems = Server.CreateObject("ADODB.Recordset")
		            rsDuplicateItems.Open SQL, conn, 3, 3
		            ''Response.Write SQL
		            ''Response.End
		            if not rsDuplicateItems.EOF then
		              iPID = CLng(rsDuplicateItems.Fields("PID"))
		              ''Response.End
		              if (CStr(rsDuplicateItems("ID")) = CStr(oFileUpEE.form("ID"))) then
                            optiontext = false
                            ''Response.Write optiontext
                            ''Response.End
                      else
                            optiontext = true
                            ''Response.Write optiontext
                            ''Response.End
                            sRelations = "1"
                            if(Asc(rsDuplicateItems.Fields("Revision_Code")) <= Asc(oFileUpEE.form("Revision_Code"))) then
                                 optiontext = false
			                     sSql = "UPDATE Calendar Set PID = 0 WHERE [ID] = " & rsDuplicateItems.Fields("ID")
			                     'response.Write sSql
			                     conn.Execute sSql

			                     sSql = "UPDATE Calendar Set PID = " & CStr(iPID) & " WHERE [ID] = " & CStr(oFileUpEE.form("ID"))
			                     'response.Write sSql
			                     'response.End
			                     conn.Execute sSql
			                     
			                     sSql = "UPDATE Calendar Set Clone = " & oFileUpEE.form("ID") & " WHERE Clone = " & rsDuplicateItems("ID")
			                     'response.Write sSql
			                     'response.End
			                     conn.Execute sSql
			                end if
			            end if
		            end if
		            rsDuplicateItems.close
		            set rsDuplicateItems = nothing		
	            end if
            end if

            rsAllItems.close
            set rsAllItems = nothing
        end if
    end if
'RESPONSE.END
' Clone
if CLng(oFileUpEE.form("Clone")) = 0 then
			strclone = false
else
			if CLng(oFileUpEE.form("Clone")) = CLng(oFileUpEE.form("ID")) then
				strclone = false
			else
				strclone = true
			end if
end if 

' Language
set rsLanguage = conn.execute("select iso2 from language where code='" & oFileUpEE.form("Content_Language") & "'")

if not(rsLanguage.eof) then
		strLanguage = rsLanguage.fields(0).value
end if

rsLanguage.close
set rsLanguage = nothing

strOldLanguage = ""

if trim(oFileUpEE.form("oldLanguage")) <> "" then
					set rsLanguage = conn.execute("select iso2 from language where code='" & oFileUpEE.form("oldLanguage") & "'")
					if not(rsLanguage.eof) then
							strOldLanguage=rsLanguage.fields(0).value
					end if
					rsLanguage.close
					set rsLanguage = nothing
					if strclone = true then
								set rsLanguage = conn.execute("select id from calendar where PID=" & oFileUpEE.form("PCat") & " and language= '" & oFileUpEE.form("oldLanguage") & "'")
								if not(rsLanguage.eof) then
									strOldLanguage = strLanguage
								end if
								rsLanguage.close
								set rsLanguage = nothing
						end if	    
end if

if(site_id = 3 Or Site_id = 46) then
    if(oFileUpEE.form("PCat") > 0) then
      iPID = oFileUpEE.form("PCat")
    end if
else
    iPID = oFileUpEE.form("PCat")
    optiontext=false
end if	

'RESPONSE.Write IPID				
'RESPONSE.End
if (oFileUpEE.form("PCat") >= 0 and optiontext=false) then
			on error resume next
			'Save
		lngpid=0
        
        
		'Whether to check for the product duplication also.
		if isblank(oFileUpEE.form("Nav_Clone")) and isblank(oFileUpEE.form("Nav_Duplicate")) then 
					Dim Oprn  
					Dim calendarId
		
					if not isblank(oFileUpEE.form("Nav_Update")) then
													if (not isblank(oFileUpEE.form("PCat")) Or iPID > 0) then
																if (clng(oFileUpEE.form("PCat")) > 0 Or iPID > 0) then
																							Oprn = "U"
																							if IsItChecked(oFileUpEE,"Delete_File") = "on" then
																									PcatSaveFilePath = " "
																									Bytes = 0
																							else
																													if not isblank(oFileUpEE.form("File_Existing")) then
																																			PcatSaveFilePath=oFileUpEE.form("File_Existing")
																																			' Remove Path, only filename will be stored in the Productengine database.
																																			if InstrRev(PcatSaveFilePath, "\") > 1 then
																																				PcatSaveFilePath = Mid(PcatSaveFilePath, InstrRev(PcatSaveFilePath, "\") + 1)
																																			end if
																																			if InstrRev(PcatSaveFilePath, "/") > 1 then
																																				PcatSaveFilePath = Mid(PcatSaveFilePath, InstrRev(PcatSaveFilePath, "/") + 1)
																																			end if
																																			Bytes = oFileUpEE.form("File_Size")
																													end if	
																								end if
																								'original
																								calendarId=oFileUpEE.form("ID")
																else
																								Oprn = "A"
																								if isnumeric(oFileUpEE.form("ID")) then
																													calendarId=oFileUpEE.form("ID") 
																													if not isblank(oFileUpEE.form("File_Existing")) then
																																		PcatSaveFilePath=oFileUpEE.form("File_Existing")
																																		' Remove Path, only filename will be stored in the Productengine database.
																																		if InstrRev(PcatSaveFilePath, "\") > 1 then
																																			PcatSaveFilePath = Mid(PcatSaveFilePath, InstrRev(PcatSaveFilePath, "\") + 1)
																																		end if
																																		if InstrRev(PcatSaveFilePath, "/") > 1 then
																																			PcatSaveFilePath = Mid(PcatSaveFilePath, InstrRev(PcatSaveFilePath, "/") + 1)
																																		end if
																																		Bytes=oFileUpEE.form("File_Size")
																													end if	
																													'original
																									else
																													calendarId = -1
																								end if    
																end if    
											            else
															            Oprn = "A"
															            calendarId = -1
											            end if    
						elseif not isblank(oFileUpEE.form("Nav_Delete")) then
											Oprn = "D"
						end if

					if ((trim(PcatSaveFilePath) <> "" or trim(oFileUpEE.form("URLLink")) <> "" or trim(oFileUpEE.form("File_Existing")) <> "") and trim(Oprn) <> "D")  then
									Country = ""
									with response
												for each oFormItem in oFileUpEE.Form
															if oFormItem.Name ="PCat_SProducts" then
																			if IsObject(oFormItem.Value) then
    																			for each oSubItem in oFormItem.Value
       																		ProductList = ProductList & oSubItem.Value & ","
																							next
																			else
  																					ProductList = ProductList & oFormItem.Value & "," 
																			end if
															end if
															
															if(Site_ID <> 82) then
																if oFormItem.Name ="PCat_SWebCats" then
																				if IsObject(oFormItem.Value) then
    																				for each oSubItem in oFormItem.Value
       																					strIndustry = strIndustry & oSubItem.Value & ","
																					next
																				else
  																						strIndustry = strIndustry & oFormItem.Value & ","
																				end if
																end if
															end if															
															
															if oFormItem.Name ="Country" then
																	if IsObject(oFormItem.Value) then
																						for each oSubItem in oFormItem.Value
  																					Country = Country & oSubItem.Value & "," 
																						next
																	else
  																				Country = Country & oFormItem.Value & "," 
																	end if
															end if
												next
									end with

									SubGroups=""
									'Code added for gold silo aggregation(gold options) -- by zensar(11/17/08)
									    SubGroups = oFileUpEE.Form("EntSubGroups") & ", "	              
									if IsObject(oFileUpEE.form("SubGroups"))=true then
											for each oSubItem in oFileUpEE.form("SubGroups")
													SubGroups = Subgroups & oSubItem.Value & ", "
											next
									else
											SubGroups= Subgroups & oFileUpEE.form("SubGroups")  
									end if    
					    		        	  
									set Products = server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")

                                    'Pass the actual url here

									Call Products.open("POST",striisserverpath,0,0,0)
									Call Products.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
					    		              
									if(Site_ID = 82) then
									  if isobject(oFileUpEE.form("IndustryCode"))=true then
												for each oSubItem in oFileUpEE.form("IndustryCode")
															  strIndustry = strIndustry & oSubItem.Value & ", "
												next 
									  else
												    strIndustry = oFileUpEE.form("IndustryCode")
									  end if
									end if
										    
									'Virtual Demo Link
									'This is really dumb that we have to hard wire the Link functionality of the AMS to just handling Virtual Demos
									if not isblank(oFileUpEE.form("URLLink")) then
											' Virtual Demo Path - Remove prefix but preserve /subdirectory/path
											' Added on 8th June 2010, for Fnet Item 805, to do not upper case if link have http
											if (Site_ID = 82) then
											    if  (Instr(1,LCase(oFileUpEE.form("URLLink")),"http://") = 0 and Instr(1,LCase(oFileUpEE.form("URLLink")),"https://") = 0) then
											        PcatSaveFilePath = replace(ucase(ReplaceQuote(LCase(oFileUpEE.form("URLLink")))),ucase("/virtual_demo/fnet/"),"")
											    else
											        PcatSaveFilePath = oFileUpEE.form("URLLink")
											    end if
											else
											    PcatSaveFilePath = replace(ucase(ReplaceQuote(LCase(oFileUpEE.form("URLLink")))),ucase("/virtual_demo/fnet/"),"")
											end if
											' End
									end if
					    
									' Country
									if Not isblank(oFileUpEE.form("Country")) and instr(1,oFileUpEE.form("Country_Reset"),"none") = 0 then
											if instr(1,oFileUpEE.form("Country_Reset"),"0") > 0 then
													Country = "0," & killquote(Country)             ' Exclude these countries
											else      
													Country = "1," & killquote(Country)             ' Include these countries
											end if  
									else
											Country = "none"                                  ' No Restrictions 
									end if
					    		              
									' Asset Status
									'RI-934 
								'response.write "Status is " & ofileUpEE.form("Status") 
								'Response.write "<br>AssetStatus is " & oFileUpEE.form("assetStatus") 
								if not isblank(ofileUpEE.form("Status")) then
									if ofileUpEE.form("Status") = 1 then
												if IsItChecked(oFileUpEE,"Delete_File") = "on" then
														AssetStatus = false
												else
														AssetStatus = true
												end if		
									else
										AssetStatus = false
									end if
								else
												'RI-934 AssetStatus = false 
									if oFileUpEE.form("assetStatus") = 1 then
												AssetStatus = true
									else
												AssetStatus = false
									end if

									
								end if
									'response.write "<br>Value is " & AssetStatus

									' Duplicate Title
									if DuplicateTitle = true then
												'AssetTitle = mid("[" & Record_ID & "] " & ReplaceQuote(DuplicateTitleText),1,128)
												AssetTitle = utf8Decode(decodeBase64(oFileUpEE.form("Title_B64")))
									else
												AssetTitle = utf8Decode(decodeBase64(oFileUpEE.form("Title_B64")))
									end if
					    
									if IsItChecked(oFileUpEE,"Preserve_Path") = "on" and isblank(error_msg) then                     
															PreservePath = true
									else
															PreservePath = false
									end if
									'Response.Write iPID & Oprn
									'Response.End 
									'Start. RI 1038 by shailendra 16Aug2010 because some HTML tags creating problem
									if (Site_ID = 3 Or Site_ID = 46) then
									'varDescription =server.URLEncode(Server.HTMLEncode(utf8Decode(decodeBase64(oFileUpEE.form("Description_B64")))))
									'below if conditions add for RI 1896 :Unable to update assets on Portal:Error while adding\updating the records Pcatalog  
									    if(strLanguage="es" or strLanguage="pt") then
									    
									  varDescription =utf8Decode(decodeBase64(oFileUpEE.form("Description_B64")))
									   else
									       varDescription =server.URLEncode(Server.HTMLEncode(utf8Decode(decodeBase64(oFileUpEE.form("Description_B64")))))
									    end if     
									      
									else
									    varDescription = utf8Decode(decodeBase64(oFileUpEE.form("Description_B64")))
									end if 
									'End
									'Modified by zensar on 12-09-2006 for avoiding dup item number		              
									strparameters =  "operation="       & Oprn &_
																										"&isclone="        & strclone &_
																										"&assetpid="       & iPID &_
																										"&title="          & AssetTitle &_
																										"&description="    & varDescription &_
																										"&filename="       & PcatSaveFilePath &_
																										"&filesize="       & Bytes &_
																										"&begindate="      & oFileUpEE.form("BDate") &_
																										"&products="       & ProductList &_
																										"&language="       & strLanguage &_
																										"&Category_Type="  & oFileUpEE.form("Category_ID") &_
																										"&oraclenumber="   & newItemNumber &_
																										"&access="         & SubGroups &_
																										"&industry="       & strIndustry & _
																										"&IncludeExclude=" & Country &_
																										"&status="         & AssetStatus &_
																										"&calendarId="     & calendarId &_
																										"&oldLanguage="    & strOldLanguage &_
																										"&oldItemNumber="  & oFileUpEE.form("oldItemNumber") &_
																										"&AssetId="        & Record_ID &_ 
																										"&Preserve="       & PreservePath & "&SiteID=" & Site_ID & "&Relations=" & sRelations
																										
																										
							        
									'>>>>>>>>>>>>>>         
									if err.number <> 0 then
											Showerror(err.Description)
											response.end
									end if    
									Call Products.send(strparameters)
									strProducts = Products.responseXML.XML
									set objxml = Server.CreateObject("msxml2.domdocument")
									Call objxml.loadxml(strProducts)

									if not(objxml is nothing ) then
													set objcol = objxml.selectsingleNode("ProductId")
													if IsNumeric(objcol.text) then
																	lngpid = objcol.text
													else
																	Showerror(objcol.text)
													end if      
									end if
									'Response.Write lngpid
									'Response.End
					end if

					if err.number <> 0 then
								Showerror(err.Description)
					end if
					on error goto 0
end if	
        
if isblank(oFileUpEE.form("Nav_Duplicate")) then        
				if not isblank(oFileUpEE.form("PCat"))  then
						if clng(oFileUpEE.form("PCat")) <> -1 and clng(oFileUpEE.form("PCat")) <> 0 then
								lngpid = oFileUpEE.form("PCat")
						end if
				end if
else
				lngpid = 0   
end if   
		  
SQL = "update Calendar set PID=" & lngpid & " WHERE Calendar.ID = " & Record_ID & " and site_id=" & cint(Site_ID)
conn.execute (SQL)

SQL = "update Calendar set PID=" & lngpid & ",Category_id=" & oFileUpEE.form("Category_ID") & " WHERE Calendar.clone = " & Record_ID & " and site_id=" & cint(Site_ID)
conn.execute (SQL)

'---------------------------------------------------------------------------------------  
' Sub Routines and Functions
'---------------------------------------------------------------------------------------  

sub Showerror(errordesc)

  BackURL = "/sw-administrator/default.asp?Site_ID=" & Site_ID & "&ID=edit_record&Category_ID=" & oFileUpEE.form("Category_ID")

  with response
  .write "<HTML>" & vbCrLf
  .write "<HEAD>" & vbCrLf
  .write "<LINK REL=STYLESHEET HREF=""/SW-Common/SW-Style.css"">" & vbCrLf
  .write "<TITLE>Error</TITLE>" & vbCrLf
  .write "</HEAD>" & vbCrLf
  .write "<BODY BGCOLOR=""White"" LINK =""#000000"" VLINK=""#000000"" ALINK=""#000000"">" & vbCrLf
  .write "<FORM METHOD=""POST"" NAME=""foo"" ACTION=""" & BackUrl & """>" & vbCrLf
  .write "<INPUT TYPE=""HIDDEN"" VALUE=""" & BackURL & """>" & vbCrLf
  .write "<DIV ALIGN=CENTER>" & vbCrLf

  Call Nav_Border_Begin

  .write "<TABLE CELLPADDING=10><TR><TD CLASS=NORMALBOLD BGCOLOR=WHITE ALIGN=CENTER>" & vbCrLf
  '.write "If your browser does not automatically return to the edit screen<BR>within 5 seconds, click on the [Continue] link below.<P>"
  .write "Error while adding\updating the records in Pcatalog<br><br>"
  .write errordesc & "<br><br>"
  '.write "Please delete this record and insert it again." & "<br><br>"
  .write "<SPAN CLASS=NavLeftHighlight1>&nbsp;&nbsp;<A HREF=""" & BackURL & """>Continue</A>&nbsp;&nbsp;</SPAN>"
  .write "</TD>" & vbCrLf
  .write "</TR>" & vbCrLf
  .write "</TABLE>" & vbCrLf

  Call Nav_Border_End

  .write "</FORM>" & vbCrLf
  .write "</DIV>" & vbCrLf
  .write "</BODY>" & vbCrLf
  .write "</HTML>" & vbCrLf
  
  end with

  on error goto 0
  response.end
  end sub

else
				if oFileUpEE.form("pidDelete") > 0 then  
										set objAssetDelete=server.CreateObject("Msxml2.SERVERXMLHTTP.6.0")
										'Pass the actual url here
										Call objAssetDelete.open("POST",striisserverpath,0,0,0)
										Call objAssetDelete.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
								    
										if strclone=false then
															strparameters = "operation=D" &_
																														 "&isclone="   & strclone &_
																													 	"&assetpid="  & oFileUpEE.form("pidDelete") &_
                											    "&language="  & strLanguage &_
																														 "&DeleteAll=true&setRelationship=true&itemNumber=" &_
																													 	oFileUpEE.form("oldItemNumber") & "&SiteID=" & Site_ID
										else
														strparameters = "operation=D" &_
																														"&isclone="   & strclone &_
																														"&assetpid="  & oFileUpEE.form("pidDelete") &_
																											   "&language="  & strLanguage &_
																														"&DeleteAll=false&setRelationship=false&itemNumber=" &_
																														oFileUpEE.form("oldItemNumber") & "&SiteID=" & Site_ID
										end if	

										Call objAssetDelete.send(strparameters)
										set objAssetDelete=nothing
				end if
				  
				SQl = "update Calendar set PID=-1 where Calendar.clone = " & Record_ID & " and site_id=" & cint(Site_ID)
				conn.execute (SQL)
				  
				SQL = "update Calendar set PID=-1  WHERE Calendar.ID = " & Record_ID & " and site_id=" & cint(Site_ID)
				conn.execute (SQL)
end if  
%>