
strComputer = "DYW042D001"
strNamespace = "root\cimv2"
strUser = "RDW000i004\VMWare"
strPassword = "P2Vteam11668"

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objwbemLocator.ConnectServer _
    (strComputer, strNamespace, strUser, strPassword)
objWMIService.Security_.authenticationLevel = 6

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

 

Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

 

For Each objitem in colitems

 

strcomputername = objitem.Name

'wscript.echo "Computername: " & objitem.Name

 

Next

 

Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

 

For Each objitem in colitems

strOSname = objitem.Caption

strVersion = objitem.Version

strRAM = ROUND((objitem.TotalVisibleMemorySize /1024) / 1024,1)

 

Next

 

Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DriveType = '3'")

 

For Each objItem in colItems

Wscript.Echo "Caption: " & objItem.Caption

Wscript.Echo "Description: " & objItem.Description

Wscript.Echo "DeviceID: " & objItem.DeviceID

Wscript.Echo "Size: " & objItem.Size

Wscript.Echo "FreeSpace: " & objItem.FreeSpace



IF NOT ISNULL(objItem.Size) Then
decvol = ROUND((((objItem.Size - objItem.FreeSpace) / 1024) / 1024) /1024,1)

'wscript.echo objitem.caption
'wscript.echo objItem.Size
'wscript.echo objItem.FreeSpace
'wscript.echo "VOLUME USED ON DISK = " & decvol
Counter = Counter + decvol
End if



 

Next

wscript.echo counter
 
wscript.echo "Computername,OSName,OSVersion,RAM_VISABLE(GB),Data_Volume(GB)"
wscript.echo strcomputername & "," & strOSname & "," & strVersion & "," & strRAM & "," & Counter

