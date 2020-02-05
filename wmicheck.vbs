
strComputer = "PCNAME"
strNamespace = "root\cimv2"
strUser = "DOMAIN\admin"
strPassword = "Pa55w0rd"

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objwbemLocator.ConnectServer _
    (strComputer, strNamespace, strUser, strPassword)
objWMIService.Security_.authenticationLevel = 6

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_OperatingSystem")
For Each objItem in ColItems
    Wscript.Echo strComputer & ": " & objItem.Caption
Next
