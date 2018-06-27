<%
sub Get_Show_Values
  SQL = "SELECT * FROM Calendar_Category WHERE Site_ID=" & CInt(Site_ID) & " and ID=" & CInt(Category_ID)
  Set rsCategory = Server.CreateObject("ADODB.Recordset")
  rsCategory.Open SQL, conn, 3, 3
    
  if not rsCategory.EOF then

    Show_Location             = CInt(rsCategory("Location"))
    
    Show_ImageStore           = CInt(rsCategory("ImageStore"))
    Show_Link                 = CInt(rsCategory("Link"))
    Show_Link_PopUp_Disabled  = CInt(rsCategory("Link_PopUp_Disabled"))
    Show_Item_Number          = CInt(rsCategory("Item_Number"))
    Show_Item_Number_2        = CInt(rsCategory("Item_Number_2"))
    Show_File                 = CInt(rsCategory("File_Name"))
    Show_File_POD             = CInt(rsCategory("File_Name_POD"))
    Show_Include              = CInt(rsCategory("Include"))
    Show_Thumbnail            = CInt(rsCategory("Thumbnail"))
    Show_Subscription         = CInt(rsCategory("Subscription"))
    Show_Calendar             = CInt(rsCategory("Calendar_View"))
    Show_Forum                = CInt(rsCategory("Forum"))
    Show_Content_Group        = CInt(rsCategory("Content_Group"))
    Show_Date_Basic           = CInt(rsCategory("Date_Basic"))
    Show_Date_PRD             = CInt(rsCategory("Date_PRD"))
    Show_Mark_Confidential    = CInt(rsCategory("Mark_Confidential"))
    Show_Shopping_Cart        = CInt(rsCategory("Shopping_Cart"))
    Show_Country_Restrictions = CInt(rsCategory("Country_Restrictions"))
    Show_Site_View            = CInt(rsCategory("Site_View"))
    Show_Reassign_Owner       = CInt(rsCategory("Reassign_Owner"))
    Show_Submission_Approve   = CInt(rsCategory("Submission_Approve"))
    Show_Sub_Category         = CInt(rsCategory("Sub_Category"))
    Show_Product_Series       = CInt(rsCategory("Product_Series"))
    Show_Special_Instructions = CInt(rsCategory("Instructions"))
    Show_PID                  = CInt(rsCategory("PID"))
    Show_Preserve_Clone       = CInt(rsCategory("Preserve_Path_Clone"))
    Preset_EEF                = CInt(rsCategory("Preset_EEF"))
    Preset_FDL                = CInt(rsCategory("Preset_FDL"))
    Category_Code             = rsCategory("Code")
    Preset_Sub_Category       = rsCategory("Preset_Sub_Category")
    
   ''Added on 28th Oct 2009
if rsCategory("Marketing_Automation") <> null or rsCategory("Marketing_Automation") <> "" then
    Show_Marketing_Automation= CInt(rsCategory("Marketing_Automation"))
    if Show_Marketing_Automation = True then
        response.write "<INPUT TYPE=""HIDDEN"" NAME=""Show_Marketing_Automation"" VALUE=""" & Show_Marketing_Automation & """>" 
    end if
end if
    
    if Write_Form_Show_Values = True then
       
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Show_Forum"" VALUE=""" & Show_Forum & """>"                    
      response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" VALUE=""" & rsCategory("ID") & """ >"
      response.write "<B>" & Translate(rsCategory("Title"),Login_Language,conn) & "</B>"
      response.write "<INPUT TYPE=""Hidden"" NAME=""Code"" VALUE=""" & rsCategory("Code") & """ >"
      response.write "<INPUT TYPE=""Hidden"" NAME=""Path_Category"" VALUE=""" & rsCategory("Category") & """ >"
    end if      
  else
    if Write_Form_Show_Values = True then  
      response.write "<INPUT TYPE=""Hidden"" NAME=""Category_ID"" VALUE="""">"
      response.write "<INPUT TYPE=""Hidden"" NAME=""Code"" VALUE="""">"
    end if  
  end if
  
  rsCategory.close
  set rsCategory = nothing
  
  SQL = "SELECT PID_Enabled, PID_System FROM Site WHERE ID=" & CInt(Site_ID)
  Set rsSite = Server.CreateObject("ADODB.Recordset")
  rsSite.Open SQL, conn, 3, 3  

  if not rsSite.EOF then
    PID_Enabled = rsSite("PID_Enabled")
    PID_System  = rsSite("PID_System")

    if CInt(PID_Enabled) = CInt(False) then
      Show_PID = False
    end if
  end if
  rsSite.close
  set rsSite = nothing
  
end sub
%>
