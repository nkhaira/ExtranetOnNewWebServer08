<%
  ' Declaration of Arrays for Content / Item Display
  
  Dim Field_Max
  Field_Max = 37            ' If adding more Field_X then change Field_Max value and Pointer
  Dim Field_Name(37)
  Dim Field_Data(37)
  Dim Field_Flag(37)
  
  Field_Name(0)  = "ID"
  Field_Name(1)  = "Status"
  Field_Name(2)  = "SubGroups"
  Field_Name(3)  = "Sub_Category"
  Field_Name(4)  = "Product"
  Field_Name(5)  = "Title"
  Field_Name(6)  = "Description"
  Field_Name(7)  = "Thumbnail"
  Field_Name(8)  = "Location"
  Field_Name(9)  = "Link"
  Field_Name(10) = "Link_PopUp_Disabled"
  Field_Name(11) = "Include"
  Field_Name(12) = "File_Name"
  Field_Name(13) = "File_Size"
  Field_Name(14) = "Language"
  Field_Name(15) = "LDays"                ' Days prior to Begin Date to show
  Field_Name(16) = "LDate"                ' Date prior to Begin Date to show
  Field_Name(17) = "BDate"                ' Beginning Date
  Field_Name(18) = "EDate"                ' End Date or Same as Begin Date
  Field_Name(19) = "XDays"                ' Days after End Date to Expire or if 0 do not expire
  Field_Name(20) = "XDate"                ' Expiration Date or if XDays = 0 do not expire
  Field_Name(21) = "PDate"                '
  Field_Name(22) = "UDate"                ' Last Update
  Field_Name(23) = "Country"
  Field_Name(24) = "Clone"                ' Parent ID Number
  Field_Name(25) = "Archive_Name"
  Field_Name(26) = "Archive_Size"
  Field_Name(27) = "PEDate"               ' Public Release (Embargo) Date
  Field_Name(28) = "Confidential"
  Field_Name(29) = "Content_Group"        ' Individual / Campaign / Product Introduction
  Field_Name(30) = "Content_Group_Name"   ' Name
  Field_Name(31) = "Instructions"         ' Instructions for Asset, Appended to Description.  Shows on Asset but not in What's New or Email
  Field_Name(32) = "Item_Number"          ' Item / Reference Number
  Field_Name(33) = "Item_Number_Show"     ' Show Item / Reference Number
  Field_Name(34) = "Campaign"             ' Product Introduction / Campaign Container ID
  Field_Name(35) = "Code"                 ' Code Category ID
  Field_Name(36) = "Category_ID"          ' Code Category ID
  Field_Name(37) = "Revision_Code"  
  
  xID            = 0                      ' Pointers to above fields - Must Match
  xStatus        = 1
  xSubGroups     = 2
  xSub_Category  = 3
  xProduct       = 4
  xTitle         = 5
  xDescription   = 6
  xThumbnail     = 7
  xLocation      = 8
  xLink          = 9
  xLink_PopUp_Disabled = 10               ' Note: True / False Values are treated as strings in this array
  xInclude       = 11
  xFile_Name     = 12
  xFile_Name_Size= 13
  xLanguage      = 14
  xLDays         = 15
  xLDate         = 16
  xBDate         = 17
  xEDate         = 18
  xXDays         = 19
  xXDate         = 20
  xPDate         = 21
  xUDate         = 22
  xCountry       = 23
  xClone         = 24
  xArchive_Name  = 25
  xArchive_Size  = 26
  xPEDate        = 27
  xConfidential  = 28
  xContent_Group = 29
  xContent_Group_Name = 30
  xInstructions  = 31
  xItem_Number   = 32
  xItem_Number_Show   = 33
  xCampaign      = 34
  xCode          = 35
  xCategory_ID   = 36
  xRevision_Code = 37  
  
  Show_Thumbnail = True
  
%>