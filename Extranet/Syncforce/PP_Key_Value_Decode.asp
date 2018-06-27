<%
  with response
  .write "<HTML><BODY><TABLE Border=1 WIDTH=""100%"">"
  .write "<TR><TD><FONT FACE=""Arial"" SIZE=2>Key</FONT></TD><TD WIDTH=""40%""><FONT FACE=""Arial"" SIZE=2>Value</FONT></TD><TD><FONT FACE=""Arial"" SIZE=2>SQL</FONT></TD></TR>"

  for each item in request.form

    .write "<TR>"
    .write "<TD><FONT FACE=ARIAL SIZE=2>" & item & "</FONT></TD>"
    
    select case UCase(item)                                            ' Action
      case "ACTION", "UPDATE", "ADD", "DELETE", "VERIFY"
        .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=GREEN><B>"
        .write request.form(item)
        .write "</B></FONT></TD>"
        .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=GREEN><B>"
        .write "ACTION"
        .write "</B></TD>"
        
      case else                                                         ' All Other Form Keys
      
        select case UCase(item)                                         ' Validate Record Data Field Names
    
          ' Numeric Fields
          case "ID", "ACCOUNT_ID", "SITE_ID", "FCM", "FCM_ID", "TYPE_CODE", "GENDER", "REGION", _
               "ACCOUNT_REGION", "EMAIL_METHOD", "CONNECTION_SPEED", "SUBSCRIPTION", _
               "CHANGEID", "REG_ADMIN", "NEWFLAG", "CM_ID"
    
            if isnumeric(request.form(item)) then
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write "," & item & "=" & Replace(request.form(item),"'","''")       ' Build SQL String
              .write "</FONT></TD>"
            elseif LCase(request.form(item)) = "on" then
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write "," & item & "=-1"                                            ' Build SQL String
              .write "</FONT></TD>"
            elseif LCase(request.form(item)) = "off" then
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write "," & item & "=0"                                             ' Build SQL String
              .write "</FONT></TD>"          
            elseif request.form(item) = "" or isnull(request.form(item)) then      ' Reset Field to NULL
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "No Data reset to NULL"
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "," & item & "=NULL"                                          ' Build SQL String
              .write "</TD>"                    
            elseif UCase(request.form(item)) = "NULL" then                         ' Build SQL String
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "," & item & "=NULL"                                          ' Build SQL String
              .write "</TD>"                    
            else  ' Invalid Data Type OR No Value
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=RED>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=RED>"
              .write "Key is Valid - Invalid Data"
              .write "</FONT></TD>"
            end if
          
          ' Varchar Fields
          case "NTLOGIN", "PASSWORD", "PASSWORD_CHANGE", "GROUPS", "SUBGROUPS", "GROUPS_AUX", _
               "PREFIX", "FIRSTNAME", "MIDDLENAME", "LASTNAME", "SUFFIX", "INITIALS", _
               "COMPANY", "COMPANY_WEBSITE", "JOB_TITLE", _
               "BUSINESS_ADDRESS", "BUSINESS_ADDRESS_2", _
               "BUSINESS_MAILSTOP", "BUSINESS_CITY", "BUSINESS_STATE", "BUSINESS_STATE_OTHER", _
               "BUSINESS_POSTAL_CODE", "BUSINESS_COUNTRY", _         
               "POSTAL_ADDRESS", "POSTAL_CITY", "POSTAL_STATE", "POSTAL_STATE_OTHER", _
               "POSTAL_POSTAL_CODE", "POSTAL_COUNTRY", _        
               "SHIPPING_ADDRESS", "SHIPPING_ADDRESS_2", _
               "SHIPPING_MAILSTOP", "SHIPPING_CITY", "SHIPPING_STATE", "SHIPPING_STATE_OTHER", _
               "SHIPPING_POSTAL_CODE", "SHIPPING_COUNTRY", _
               "BUSINESS_PHONE", "BUSINESS_PHONE_EXTENSION", "BUSINESS_PHONE_2", "BUSINESS_PHONE_2_EXTENSION", _
               "BUSINESS_FAX", "MOBILE_PHONE", "PAGER", "EMAIL", "EMAIL_2", "LANGUAGE", "FLUKE_ID", "FLUKE_ID_REP", "COMMENT", _
               "AUX_0", "AUX_1", "AUX_2", "AUX_3", "AUX_4", "AUX_5", "AUX_6", "AUX_7", "AUX_8", "AUX_9", _
               "BUSINESS_SYSTEM", "CORE_ID", "ESTORE_ID"
               
            if request.form(item) = "" or isnull(request.form(item)) then          ' Reset Field to NULL
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "No Data reset to NULL"
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "," & item & "=NULL"                                          ' Build SQL String
              .write "</FONT></TD>"        
            elseif UCase(request.form(item)) = "NULL" then                         ' Build SQL String
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "," & item & "=NULL"                                          ' Build SQL String
              .write "</TD>"                    
            elseif request.form(item) <> "" and not isnull(request.form(item)) then
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write "," & item & "='" & Replace(request.form(item),"'","''") & "'"' Build SQL String
              .write "</FONT></TD>"        
            end if
          
          case "EXPIRATIONDATE", "CHANGEDATE", "REG_REQUEST_DATE", "REG_APPROVAL_DATE", "LOGON"
    
            if request.form(item) = "" or isnull(request.form(item)) then          ' Reset Field to NULL
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "No Data reset to NULL"
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "," & item & "=NULL"                                          ' Build SQL String
              .write "</FONT></TD>"        
            elseif request.form(item) = "NULL" then
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=BLUE>"
              .write "," & item & "=NULL"                                          ' Build SQL String
              .write "</FONT></TD>"        
            elseif isdate(request.form(item)) then
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2>"
              .write "," & item & "='" & request.form(item) & "'"                  ' Build SQL String
              .write "</FONT></TD>"        
            else  ' Invalid Data Type OR No Value
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=RED>"
              .write request.form(item)
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=RED>"
              .write "Key is Valid - Data Value Invalid"
              .write "</FONT></TD>"
            end if
          
          case else                                                         ' Internal PP Fields IGNORE 
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=ORANGE>"
              if request.form(item) <> "" and not isnull(request.form(item)) then
                .write request.form(item)
              else
                .write "No Data"
              end if               
              .write "</FONT></TD>"
              .write "<TD><FONT FACE=ARIAL SIZE=2 COLOR=ORANGE>"
              .write "Ignore this Key, Data is Internal to PP"
              .write "</FONT></TD>"
       
        end select
    
    end select    

    .write "</TR>"

  next

  .write "</TABLE></BODY></HTML>"
  
  end with
%>