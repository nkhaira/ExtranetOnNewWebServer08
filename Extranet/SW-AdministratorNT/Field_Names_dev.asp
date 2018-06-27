<!--#include virtual="/include/functions_date_formatting.asp"-->
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

SC5 = 50
Email_Delimited = 50

    if SC5 = Email_Delimited then
    
      SQLField_Names = "SELECT * from Field_Names WHERE Field_Names.Table_Name='UserData' AND Field_Names.Enabled=" & CInt(True) & " ORDER BY Field_Names.ID"
      Set rsField_Names = Server.CreateObject("ADODB.Recordset")
      rsField_Names.Open SQLField_Names, conn, 3, 3

      response.write "<FORM ""Field_Names"">"
      response.write "<TABLE WIDTH=""50%"">"
      
      do while not rsField_Names.EOF
      
        response.write "<TR>"
        response.write "<TD WIDTH=""5%"" CLASS=SMALL>"
        response.write "<INPUT TYPE=""CHECKBOX"" VALUE=""" & rsField_Names("ID") & """"
        if CInt(rsField_Names("Default_Select")) = CInt(True) then
          response.write " CHECKED"
        end if
        response.write " CLASS=Small>"
        response.write "</TD>"
        response.write "<TD CLASS=SMALL>"       
        response.write rsField_Names("Description")
        response.write "</TD>"
        response.write "</TR>" & vbCrLf
        
        rsField_Names.MoveNext
                        
      loop

      response.write "<TR>"
      response.write "<TD WIDTH=""5%"" CLASS=SMALL>"
      response.write "&nbsp;"
      response.write "</TD>"      
      response.write "<TD CLASS=SMALL>"       
      response.write "<INPUT TYPE=""Submit"" NAME=""Email"" VALUE=""" & Translate("Submit",Login_Language,conn) & """>"
      response.write "</TD>"
      response.write "</TR>" & vbCrLf
     
      response.write "</TABLE>"
      response.write "</FORM>"

      rsField_Names.close
      set rsField_Names = nothing

    else
    end if
    
    Call Disconnect_SiteWide          
%>    

