<%

      response.write "<FONT CLASS=Medium><IMG SRC=""/images/lock.gif"" BORDER=0 WIDTH=37 ALIGN=""ABSMIDDLE"">&nbsp;&nbsp;&nbsp;" & Translate("This is a secure site connection to protect your personal information.",Login_Language,conn) & "</FONT><BR><BR>"
      response.write "<FORM NAME=""Contact_Us"" ACTION=""" & HomeURL & """ METHOD=""POST"">"
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""Site_ID"" VALUE=""" & Site_ID & """>"
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""NS"" VALUE=""" & Top_Navigation & """>"
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""CID"" VALUE=""" & CID & """>"      
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""SCID"" VALUE=""" & SCID & """>"
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""PCID"" VALUE=""" & PCID & """>"
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""CIN"" VALUE=""" & "1" & """>"                        
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""LANG"" VALUE=""" & Login_Language & """>"      
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""FCM_Name"" VALUE=""" & FCM_Name & """>"
      response.write "<INPUT TYPE=""HIDDEN"" NAME=""FCM_EMail"" VALUE=""" & FCM_EMail & """>"
                              
      response.write Translate("Select one of the &quot;Contact Us&quot; recipients below",Login_Language,conn) & ":<BR><BR>"
      
      Call Nav_Border_Begin
      
      
      response.write "<TABLE WIDTH=""100%"" CLASS=tablebackground><TR>"
      if not (isblank(Login_Fcm_ID) and CLng(Login_Fcm_ID) > 0) and not isblank(Fcm_EMail) then     
        response.write "<TD CLASS=Medium WIDTH=20>"
        response.write "<INPUT TYPE=""RADIO"" NAME=""MessageTo"" VALUE=3 CHECKED></TD>"
        response.write "<TD CLASS=Medium>&nbsp;&nbsp;" & Translate("Support Group",Login_Language,conn) & "</TD>"
        response.write "<TD CLASS=Medium> - " & Translate("Questions regarding the content of this site",Login_Language,conn) & ".</TD></TR>"
      end if
      if not (isblank(Login_Fcm_ID) and CLng(Login_Fcm_ID) > 0) and not isblank(Fcm_EMail) then     
        response.write "<TD CLASS=Medium WIDTH=20>"
        response.write "<INPUT TYPE=""RADIO"" NAME=""MessageTo"" VALUE=1 ></TD>"
        response.write "<TD CLASS=Medium>&nbsp;&nbsp;" & Translate("Administrator",Login_Language,conn) & "</TD>"
        response.write "<TD CLASS=Medium> - " & Translate("Questions regarding the operation of this site",Login_Language,conn) & ".</TD></TR>"
      else
        response.write "<TD CLASS=Medium WIDTH=20>"
        response.write "<INPUT TYPE=""RADIO"" NAME=""MessageTo"" VALUE=1 CHECKED></TD>"
        response.write "<TD CLASS=Medium>&nbsp;&nbsp;" & Translate("Administrator",Login_Language,conn) & "</TD>"
        response.write "<TD CLASS=Medium> - " & Translate("Questions regarding the operation of this site",Login_Language,conn) & ".</TD></TR>"
      end if          
      response.write "<TD CLASS=Medium WIDTH=20>"
      response.write "<INPUT TYPE=""RADIO"" NAME=""MessageTo"" VALUE=2></TD><TD CLASS=Medium>&nbsp;&nbsp;" & Translate("Webmaster",Login_Language,conn) & "</TD><TD CLASS=Medium> - " & Translate("Questions regarding the operation of this server or site features",Login_Language,conn) & ".</TD></TR>"
      response.write "</TABLE>"
      
      'Response.Write("Site ID" & Site_ID )
      'Response.Write("Login_Fcm_ID" & Login_Fcm_ID )
	'  Response.End()
      
      
      Call Nav_Border_End
      
      response.write "<BR>"
            
      response.write Translate("Enter your question(s) in the text area below",Login_Language,conn) & ". "
      response.write Translate("There is no need to include your name, EMail address or phone number, since this information will automatically be sent with your &quot;Contact Us - Request&quot;",Login_Language,conn) & ".<BR><BR>"
      if UCASE(Login_Language) <> "ENG" then
        response.write Translate("If possible, please compose your message in English.",Login_Language, conn) & "<BR><BR>"
      end if
      
      Call Nav_Border_Begin
      
      response.write "<TABLE WIDTH=""100%"" CLASS=tablebackground>"
      response.write "<TR>"
      response.write "<TD>"     
      response.write "<TEXTAREA NAME=""UserMessage"" COLS=80 ROWS=10  MAXLENGTH=""2000"" CLASS=Medium></TEXTAREA>"
      response.write "</TD>"
      response.write "</TR>"
      response.write "<TR>"
      response.write "<TD CLASS=tablebackground ALIGN=CENTER>"
      response.write "<INPUT TYPE=""SUBMIT"" VALUE="" " & Translate("Send",Login_Language,conn) & " "" CLASS=NAVLEFTHighlight1 onmouseover=""this.className='NavLeftButtonHover'"" onmouseout=""this.className='Navlefthighlight1'"">"
      response.write "</TD>"
      response.write "</TR>"
      response.write "</TABLE>"
      
      Call Nav_Border_End
      
      response.write "</FORM>"
%>