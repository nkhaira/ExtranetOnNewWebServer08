

strComputer = "."
strNameSpace = "root\MicrosoftIISv2"
strClass = "IIsWebVirtualDir"
 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MicrosoftIISv2")
 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM IIsWebVirtualDir",,48) 
For Each objItem in colItems 
    Wscript.Echo  objItem.Name
Next


