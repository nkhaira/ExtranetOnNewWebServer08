<%
on error resume next
'Set WShell = Server.CreateObject("WSCRIPT.Shell")
'WShell.Run("d:\extranet\sw-administrator\test.bat")
'Set WShell = Nothing
'Response.Write err.Description
'Set oComponent = Server.CreateObject("Test.WSC")
'response.Write oComponent.Test("vxcvxc")


strComputer = "."
strNameSpace = "root\MicrosoftIISv2"
strClass = "IIsWebVirtualDir"
 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MicrosoftIISv2")
 
Set colItems = objWMIService.ExecQuery("SELECT * FROM IIsWebVirtualDirSetting",,48) 
For Each objItem in colItems 
    Response.Write  objItem.Name & " - " &  objItem.Path & "<BR>"
Next

if err.number <> 0 then
    Response.Write  err.Description
end if
%>