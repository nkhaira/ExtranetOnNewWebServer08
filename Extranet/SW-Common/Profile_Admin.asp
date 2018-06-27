<%
' --------------------------------------------------------------------------------------
' Author: D. Whitlock
' Date:   02/01/2000
' Update Account Profile - Limited Data Version
' --------------------------------------------------------------------------------------

Dim SQL
Dim BackURL

BackURL = request("BackURL")


' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

Call Connect_SiteWide

SQL = "UPDATE UserData SET "   
  
  ' Prefix
  if not isblank(request("Prefix")) then 
    SQL = SQL & "Prefix="& "'" & replacequote(request("Prefix")) & "'"
  else
    SQL = SQL & "Prefix=NULL"
  end if

  ' First Name
  if not isblank(request("FirstName")) then 
    SQL = SQL & ",FirstName="& "'" & replacequote(request("FirstName")) & "'"
  else
    SQL = SQL & ",FirstName=NULL"    
  end if

  ' Middle Name
  if not isblank(request("MiddleName")) then 
    SQL = SQL & ",MiddleName="& "'" & replacequote(request("MiddleName")) & "'"         
  else
    SQL = SQL & ",MiddleName=NULL"    
  end if

  ' Last Name
  if not isblank(request("LastName")) then 
    SQL = SQL & ",LastName="& "'" & replacequote(request("LastName")) & "'"         
  else
    SQL = SQL & ",LastName=NULL"   
  end if

  ' Suffix
  if not isblank(request("Suffix")) then 
    SQL = SQL & ",Suffix="& "'" & replacequote(request("Suffix")) & "'"
  else
    SQL = SQL & ",Suffix=NULL"             
  end if
        
  ' Company
  if not isblank(request("Company")) then 
    SQL = SQL & ",Company="& "'" & replacequote(request("Company")) & "'"
  else
    SQL = SQL & ",Company=NULL"             
  end if

  ' Job Title
  if not isblank(request("Job_Title")) then 
    SQL = SQL & ",Job_Title="& "'" & replacequote(request("Job_Title")) & "'"         
  else
    SQL = SQL & ",Job_Title=NULL"    
  end if
  
  ' Business Mail Stop
  if not isblank(request("Business_MailStop")) then 
    SQL = SQL & ",Business_MailStop="& "'" & replacequote(request("Business_MailStop")) & "'"
  else
    SQL = SQL & ",Business_MailStop=NULL"
    
  end if
  
  ' Business Address
  if not isblank(request("Business_Address")) then 
    SQL = SQL & ",Business_Address="& "'" & replacequote(request("Business_Address")) & "'"         
  else
    SQL = SQL & ",Business_Address=NULL"
  end if
  
  ' Business Address 2
  if not isblank(request("Business_Address_2")) then 
    SQL = SQL & ",Business_Address_2="& "'" & replacequote(request("Business_Address_2")) & "'"
  else
    SQL = SQL & ",Business_Address_2=NULL"    
  end if

  ' Business City
  if not isblank(request("Business_City")) then 
    SQL = SQL & ",Business_City="& "'" & replacequote(request("Business_City")) & "'"         
  else
    SQL = SQL & ",Business_City=NULL"    
  end if

  ' Business State
  if not isblank(request("Business_State")) or not isblank(request("Business_State_Other")) then 
    if not isblank(request("Business_State")) then
      SQL = SQL & ",Business_State="& "'" & replacequote(request("Business_State")) & "'"
    else
      SQL = SQL & ",Business_State=NULL"               
    end if              
    if not isblank(request("Business_State_Other")) then
      SQL = SQL & ",Business_State_Other="& "'" & replacequote(request("Business_State_Other")) & "'"         
    else
      SQL = SQL & ",Business_State_Other=NULL"      
    end if              
  end if

  ' Business Postal Code
  if not isblank(request("Business_Postal_Code")) then 
    SQL = SQL & ",Business_Postal_Code="& "'" & replacequote(request("Business_Postal_Code")) & "'"         
  else
    SQL = SQL & ",Business_Postal_Code=NULL"
  end if

  ' Business_Country
  if not isblank(request("Business_Country")) then 
    SQL = SQL & ",Business_Country="& "'" & replacequote(request("Business_Country")) & "'"         
  else
    SQL = SQL & ",Business_Country=NULL"
  end if

  ' Shipping Mail Stop
  if not isblank(request("Shipping_MailStop")) then 
    SQL = SQL & ",Shipping_MailStop="& "'" & replacequote(request("Shipping_MailStop")) & "'"
  else
    SQL = SQL & ",Shipping_MailStop=NULL"
  end if

  ' Shipping Address
  if not isblank(request("Shipping_Address")) then 
    SQL = SQL & ",Shipping_Address="& "'" & replacequote(request("Shipping_Address")) & "'"
  else
    SQL = SQL & ",Shipping_Address=NULL"
  end if
  
  ' Shipping Address 2
  if not isblank(request("Shipping_Address_2")) then 
    SQL = SQL & ",Shipping_Address_2="& "'" & replacequote(request("Shipping_Address_2")) & "'"
  else
    SQL = SQL & ",Shipping_Address_2=NULL"
  end if

  ' Shipping City
  if not isblank(request("Shipping_City")) then 
    SQL = SQL & ",Shipping_City="& "'" & replacequote(request("Shipping_City")) & "'"
  else
    SQL = SQL & ",Shipping_City=NULL"    
  end if

  ' Shipping State  
  if not isblank(request("Shipping_State")) > 0 or not isblank(request("Shipping_State_Other")) then 
    if not isblank(request("Shipping_State")) then
      SQL = SQL & ",Shipping_State="& "'" & replacequote(request("Shipping_State")) & "'"
    else
      SQL = SQL & ",Shipping_State=NULL"               
    end if              
    if not isblank(request("Shipping_State_Other")) then
      SQL = SQL & ",Shipping_State_Other="& "'" & replacequote(request("Shipping_State_Other")) & "'"         
    else
      SQL = SQL & ",Shipping_State_Other=NULL"      
    end if              
  end if

  ' Shipping Postal Code
  if not isblank(request("Shipping_Postal_Code")) then 
    SQL = SQL & ",Shipping_Postal_Code="& "'" & replacequote(request("Shipping_Postal_Code")) & "'"         
  else
    SQL = SQL & ",Shipping_Postal_Code=NULL"
  end if

  ' Shipping_Country
  if not isblank(request("Shipping_Country")) then 
    SQL = SQL & ",Shipping_Country="& "'" & replacequote(request("Shipping_Country")) & "'"         
  else
    SQL = SQL & ",Shipping_Country=NULL"
  end if

  ' Business Phone
  if not isblank(request("Business_Phone")) then 
    SQL = SQL & ",Business_Phone="& "'" & FormatPhone(replacequote(request("Business_Phone"))) & "'"         
  else
    SQL = SQL & ",Business_Phone=NULL"
  end if
  
  ' Business Phone Extension
  if not isblank(request("Business_Phone_Extension")) or not isblank(request("Business_Phone")) then 
    SQL = SQL & ",Business_Phone_Extension="& "'" & replacequote(request("Business_Phone_Extension")) & "'"         
  else
    SQL = SQL & ",Business_Phone_Extension=NULL"
  end if

  ' Business_Phone_2
  if not isblank(request("Business_Phone_2")) then 
    SQL = SQL & ",Business_Phone_2="& "'" & FormatPhone(replacequote(request("Business_Phone_2"))) & "'"
  else
    SQL = SQL & ",Business_Phone_2=NULL"    
  end if

  ' Business Phone Extension
  if not isblank(request("Business_Phone_2_Extension")) and not isblank(request("Business_Phone_2")) then 
    SQL = SQL & ",Business_Phone_2_Extension="& "'" & replacequote(request("Business_Phone_2_Extension")) & "'"
  else
    SQL = SQL & ",Business_Phone_2_Extension=NULL"             
  end if  
  
  ' Business_Fax
  if not isblank(request("Business_Fax")) then 
    SQL = SQL & ",Business_Fax="& "'" & FormatPhone(replacequote(request("Business_Fax"))) & "'"         
  else
    SQL = SQL & ",Business_Fax=NULL"    
  end if
  
  ' Mobile_Phone
  if not isblank(request("Mobile_Phone")) then 
    SQL = SQL & ",Mobile_Phone="& "'" & FormatPhone(replacequote(request("Mobile_Phone"))) & "'"         
  else
    SQL = SQL & ",Mobile_Phone=NULL"    
  end if
  
  ' Pager
  if not isblank(request("Pager")) then 
    SQL = SQL & ",Pager="& "'" & FormatPhone(replacequote(request("Pager"))) & "'"         
  else
    SQL = SQL & ",Pager=NULL"    
  end if

  ' Email
  if not isblank(request("Email")) then 
    SQL = SQL & ",Email="& "'" & replacequote(request("Email")) & "'"         
  else
    SQL = SQL & ",Email=NULL"    
  end if
 
  ' Email
  if not isblank(request("Email_Method")) then 
    SQL = SQL & ",Email_Method="& "" & replacequote(request("Email_Method")) & ""
  else
    SQL = SQL & ",Email_Method="& "0"    
  end if

  ' Email_2
  if not isblank(request("Email_2")) then 
    SQL = SQL & ",Email_2="& "'" & replacequote(request("Email_2")) & "'"         
  else
    SQL = SQL & ",Email_2=NULL"    
  end if
  
  ' Connection_Speed
  if not isblank(request("Connection_Speed")) then 
    SQL = SQL & ",Connection_Speed="& "" & replacequote(request("Connection_Speed")) & ""
  else
    SQL = SQL & ",Connection_Speed=0"
  end if

 
  ' Language
  if not isblank(request("Language")) then 
    SQL = SQL & ",Language="& "'" & replacequote(request("Language")) & "'"
  else
    SQL = SQL & ",Language="& "'eng'"    
  end if

  ' Subscription 
  if request("Subscription") = "on" then           
    SQL = SQL & ",Subscription=-1"
  else
    SQL = SQL & ",Subscription=0"
  end if        

  ' Change ID   
  if not isblank(request("Account_ID")) and isnumeric(request("Account_ID")) then               
    SQL = SQL & ",ChangeID="& killquote(request("Account_ID"))
  end if

  ' Change Date
  if not isblank(request("ChangeDate")) then
    SQL = SQL & ",ChangeDate="& "'" & request("ChangeDate") & "'"         
  end if
      
  ' Auxiliary Fields
  for aux = 0 to 9
    if not isblank(request("Aux_" & Trim(aux))) then 
      SQL = SQL & ",Aux_" & Trim(aux) & "="& "'" & replacequote(request("Aux_" & Trim(aux))) & "'"        
    end if
  next
  
  ' Build Ending SQL Statement  (Use Session NTLogin as opposed to Request("NTLogin")
  ' to prevent user from modifying Script to update another account
  
  SQL = SQL & " WHERE UserData.NTLogin='" & Session("LOGON_USER") & "'"
  response.write SQL
  response.flush
response.end  
  
  conn.Execute (SQL) 
  
  Call Disconnect_SiteWide

response.write SQL
response.end  
  
'  BackURL = replace(BackURL,"CINN=0","CINN=1")
  
'  response.redirect BackURL

%>

