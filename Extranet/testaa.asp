<%

Dim Account_Info_ID, Shipping_Address_ID, Address_Text
Dim FirstName, MiddleName, LastName, Company, Job_Title
Dim Business_Address, Business_Address_2, Business_City, Business_City_Other, Business_State
Dim Business_Postal_Code, Business_Country, Business_Email, Business_Fax, Business_Phone
Dim Shipping_Address, Shipping_Address_2, Shipping_City, Shipping_City_Other, Shipping_State
Dim Shipping_Postal_Code, Shipping_Country, Shipping_Comment
Dim	strLocalHostName, strLocalHostIP, strProtocol, strMethod, strRemoteHostName, strRemoteHostIP, iRemoteHostPort, strRemoteHostTargetFile
Dim strKeyValueDelimiter, strPairDelimiter, cProtocol, cMethod, strLocalHostReferrerFile
Dim strResponse, strPost_QueryString
Dim Form_Name, Action_Count
Dim Ship_To_Alternates, Save_Cart, Order_Number, Order_Number_Results, Order_Status_Info, Order_Status_Info_Alt
Dim Screen_Width, Screen_Height
Dim MailSubject, MailMessage, DotLine
Dim ErrorFlag

	strConnectionString = "DRIVER={SQL Server};SERVER=flkprd18.data.ib.fluke.com" &_
	";UID=marcomweb" &_
	";DATABASE=fluke_sitewide" &_
	";pwd=!?wwwProd1"
	
	Dim conn
',FirstName, MiddleName, LastName
	set conn=server.CreateObject("adodb.connection")
	conn.Open strConnectionString
	
	if conn.State =1 then
		'Response.Write "Opened"
	end if
	
'Account_Info_ID = 0
'Call Get_Account_Information


      
	'Site_id =24

 Order_Date           = Date()
 Order_Time           = Time()
 PostErrMessage       = ""
 Order_Number_Results = ""
	
	strPost_QueryString = "Action=Order_Post"
        
' Get Order By Profile Info
Account_Info_ID = 0
logon_name ="cludwig"
site_id =3
	Call Get_Account_Information        
        
    strPost_QueryString = strPost_QueryString &_
    "&SORDNH=" & Server.URLEncode("") &_
    "&SRDSTS=" & Server.URLEncode("04") &_
    
    "&OUSERN=" &_
    "&OPSSWD=" &_
    
    "&OCUNAM=" & Server.URLEncode(FormatFullName(FirstName, MiddleName, LastName)) &_
    "&OCMPNM=" & Server.URLEncode(Company) &_
    "&OTITL="  & Server.URLEncode(Job_Title) &_
    
    "&OADRS1=" & Server.URLEncode(Business_Address) &_
    "&OADRS2=" & Server.URLEncode(Business_Address_2) &_
    "&OCITY="  & Server.URLEncode(Business_City) &_
    "&OSTATE=" & Server.URLEncode(Business_State) &_
    "&OZIP="   & Server.URLEncode(Business_Postal_Code) &_
    "&OCNTRY=" & Server.URLEncode(Business_Country) &_
    
    "&OPHONE=" & Server.URLEncode(FormatPhone(Business_Phone)) &_
    "&OFAX="   & Server.URLEncode(FormatPhone(Business_Fax)) &_
    "&OEMAIL=" & Server.URLEncode(Business_Email)
            
	' Get Ship To Profile Info

	Account_Info_ID = 0
 'Ship_To_Addresses(stc)
	select case Account_Info_ID
		case 0
				Call Get_Account_Information
		case else                           ' Drop Ship To
				Call Get_Account_Information
	end select

    ' DCG ORDER SYSTEM
    strPost_QueryString = strPost_QueryString      &_
    "&SCUSID=" & Server.URLEncode("FLUKEPP:" & Common_ID & ":" & Account_Info_ID) &_
    "&SCUNAM=" & Server.URLEncode(FormatFullName(FirstName, MiddleName, LastName)) &_
    "&STITL="  & Server.URLEncode(Job_Title) &_
    "&SCMPNM=" & Server.URLEncode(Company) &_
    "&SADRS1=" & Server.URLEncode(Shipping_Address) &_
    "&SADRS2=" & Server.URLEncode(Shipping_Address_2) &_
    "&SCITY="  & Server.URLEncode(Shipping_City) &_
    "&SSTATE=" & Server.URLEncode(Shipping_State) &_
    "&SZIP="   & Server.URLEncode(Shipping_Postal_Code) &_
    "&SCNTRY=" & Server.URLEncode(Shipping_Country) &_
    "&SPHONE=" & Server.URLEncode(FormatPhone(Business_Phone)) &_
    "&SFAX="   & Server.URLEncode(FormatPhone(Business_Fax)) &_
    "&SEMAIL=" & Server.URLEncode(Business_Email) &_
    
    "&SRDDT="  & Server.URLEncode(Order_Date) &_
    "&SRDTM="  & Server.URLEncode(Order_Time) &_
    "&SDTAP="  & Server.URLEncode(Order_Date) &_
    "&SDTSP="  & Server.URLEncode("") &_
    
    "&SSHPV="  & Server.URLEncode("UPS") &_
    "&SSHPIN=" & Server.URLEncode(Shipping_Comment) &_
    "&SPOST="  & Server.URLEncode("") &_
    "&SMATL="  & Server.URLEncode("") &_
    "&SHAND="  & Server.URLEncode("") &_
    "&SDTIV="  & Server.URLEncode("") &_
    "&SPONUM=" & Server.URLEncode("") &_
    "&SCCNUM=" & Server.URLEncode("") &_
    "&SDTDU="  & Server.URLEncode("") &_
    "&SCOMM="  & Server.URLEncode("") &_
    "&SPPKG="  & Server.URLEncode("") &_
    "&SPCAR="  & Server.URLEncode("") &_
    "&SPSRV="  & Server.URLEncode("") &_
    "&SSHPI2=" & Server.URLEncode("")
    
    response.write strPost_QueryString
    response.end

  function FormatFullName(FirstName, MiddleName, LastName)

		Dim TempStr
		TempStr = ""
		  
		if not isblank(FirstName) then
			TempStr = TempStr & FirstName
		end if
		  
		if not isblank(MiddleName) then
			if len(Trim(MiddleName)) > 1 and instr(1,MiddleName,".") = 0 then
			TempStr = TempStr & " " & MiddleName
			end if
		end if
		  
		if not isblank(LastName) then
			TempStr = TempStr & " " & LastName
		end if
		  
		FormatFullName = TempStr
	  
	end function
	
	sub Get_Account_Information
	 Logon_Name = "cludwig"
  Site_ID = 3
  Account_Info_ID = 0
		select case Account_Info_ID
			case 0
			SQL =  "SELECT * FROM UserData WHERE NTLogin='" & Logon_Name & "' AND Site_ID=" & Site_ID & " AND NewFlag=0"
			case else
			SQL =  "SELECT * FROM Shopping_Cart_Ship_To WHERE ID=" & Account_Info_ID
		end select    

		Set rsLogin = Server.CreateObject("ADODB.Recordset")
		rsLogin.Open SQL, conn, 3, 3
		  
		if not rsLogin.EOF then
		  
			'if isblank(rsLogin("Shipping_Address")) or isblank(rsLogin("Shipping_City")) or isblank(rsLogin("Shipping_Postal_Code")) or isblank(rsLogin("Shipping_Country")) then
			'ErrMessage = ErrMessage &_
			'"<LI>" & Translate("We are unable to process your order, missing shipping information.",Login_Language,conn) &_
			'"<LI>" & Replace(Replace(Translate("Click on the [ Edit Account ] button to provide this information.",Login_Language,conn),"[","<SPAN CLASS=NavLeftHighlight1>&nbsp;"),"]","</SPAN>&nbsp;")
			'else
			' User Information
			if not isblank(rsLogin("FirstName"))  then FirstName = rsLogin("FirstName") else FirstName = ""
			if not isblank(rsLogin("MiddleName")) then MiddleName = rsLogin("MiddleName") else MiddleName = ""
			if not isblank(rsLogin("LastName"))   then LastName = rsLogin("LastName") else LastName = ""
			if not isblank(rsLogin("Company"))    then Company = rsLogin("Company") else Company = ""
			if not isblank(rsLogin("Job_Title"))  then Job_Title = rsLogin("Job_Title") else Job_Title = ""
			PPAccount_ID = rsLogin("ID")
		      
			Business_Phone = rsLogin("Business_Phone")
			Business_Email = rsLogin("Email")
		            
			if Account_Info_ID = 0 then
				' Business Address Information
				Business_Address = rsLogin("Business_Address")
				if not isblank(rsLogin("Business_Address_2")) then
				Business_Address_2 = rsLogin("Business_Address_2")
				else
				Business_Address_2 = ""  
				end if
				Business_City = rsLogin("Business_City")
				if rsLogin("Business_State") <> "ZZ" then
				Business_State = rsLogin("Business_State")
				else
				Business_State = ""
				end if    
				if not isblank(rsLogin("Business_State_Other")) then
				Business_State_Other = rsLogin("Business_State_Other")
				else
				Business_State_Other = ""  
				end if  
				Business_Postal_Code = rsLogin("Business_Postal_Code")
				Business_Country = rsLogin("Business_Country")
			end if
		      
			' Shipping Address Information
			Shipping_Address = rsLogin("Shipping_Address")
			if not isblank(rsLogin("Shipping_Address_2")) then
				   Shipping_Address_2 = rsLogin("Shipping_Address_2")
			else
			    Shipping_Address_2 = ""  
			end if
			Shipping_City = rsLogin("Shipping_City")
			if rsLogin("Shipping_State") <> "ZZ" then
				    Shipping_State = rsLogin("Shipping_State")
			else
				Shipping_State = ""
			end if    
			if not isblank(rsLogin("Shipping_State_Other")) then
				Shipping_State_Other = rsLogin("Shipping_State_Other")
			else
				Shipping_State_Other = ""  
			end if  
			Shipping_Postal_Code = rsLogin("Shipping_Postal_Code")
			Shipping_Country = rsLogin("Shipping_Country")
		      
			if Account_Info_ID <> 0 then
				    if not isblank(rsLogin("Comment")) then
				        Shipping_Comment = rsLogin("Comment")
				        else
				        Shipping_Comment = ""
				        end if  
			    else
				        Shipping_Comment = ""
			    end if    
			 end if
		'end if
		  
		rsLogin.close
		set rsLogin = Nothing

		end sub
		
		' --------------------------------------------------------------------------------------
' NOTE - this function is copied into subscription email tool!

function IsBlank(MyString)

  Dim tmpType

  IsBlank = false

  select case VarType(MyString)

    case 0, 1                       ' Empty & Null
      IsBlank = true
    case 8                          ' String
      if Len(MyString) = 0 then
        IsBlank = true
      elseif UCase(MyString) = "NULL" or UCase(MyString) = "NO VALUE" or Trim(MyString & " ") = "" then
        IsBlank = true
      end if
    case 9                          ' Object  
      tmpType = TypeName(MyString)
      if (tmpType = "Nothing") or (tmpType = "Empty") then
        IsBlank = True
      end if
    case 8192, 8204, 8209            ' Array
      ' Does it have at least one element?
      if UBound(MyString) = 0 then
        IsBlank = True
      end if
  end select
  
end function

function FormatPhone(MyString)

  Dim tempstr, tempchr, tempout
  
  tempstr = MyString
  tempchr = 0
  tempout = ""
  
  if not isblank(Trim(tempstr)) then
  
    for x = 1 to len(tempstr)
    
      tempchr = UCase(mid(tempstr,x,1))
      tempasc = asc(tempchr)

      select case tempasc
        case 48,49,50,51,52,53,54,55,56,57
          tempout = tempout & tempchr        
        case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90
          select case tempchr
            case "A","B","C"
              tempout = tempout & "2"
            case "D","E","F"
              tempout = tempout & "3"
            case "G","H","I"
              tempout = tempout & "4"
            case "J","K","L"
              tempout = tempout & "5"
            case "M","N","O"
              tempout = tempout & "6"
            case "P","Q","R","S"
              tempout = tempout & "7"
            case "T","U","V"
              tempout = tempout & "8"
            case "W","X","Y","Z"
              tempout = tempout & "9"
           end select
      end select

    next
    
    if len(tempout) = 7 then  
      tempout = Mid(tempout,1,3) & "." & Mid(tempout,4,4)
    elseif len(tempout) = 10 then
      tempout = Mid(tempout,1,3) & "." & Mid(tempout,4,3) & "." & Mid(tempout,7,4)
    elseif len(tempout) = 11 then
      tempout = Mid(tempout,1,1) & "." & Mid(tempout,2,3) & "." & Mid(tempout,5,3) & "." & Mid(tempout,8,4)    
    else
      tempout = tempstr  
    end if
    
    FormatPhone = tempout

  else
    FormatPhone = ""
  end if    
  
end function
%>
