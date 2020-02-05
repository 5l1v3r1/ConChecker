ON ERROR RESUME NEXT

Const ForReading = 1
Const ForAppending = 8
Const ForWriting = 2


strNamespace = "root\cimv2"
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("servers.txt", ForReading)


Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile("output.csv", True)


objfile.writeline "Computername,OSName,OSVersion,RAM_VISABLE(GB),Data_Volume(GB)"
objFile.Close


Do Until objTextFile.AtEndOfStream
    strLine = objTextFile.ReadLine
    'Wscript.Echo strLine
arrayline = split(strLine,",")

'Call PingTest(arrayline(0))

Call ConnectWMI(arrayline(0),arrayline(1),arrayline(2))

'Call ConnectSSH(arrayline(0),arrayline(1),arrayline(2))


Loop

objTextFile.Close

Function ConnectWMI(strtarget,strusername,strpassword)

wscript.echo "WMI Connection Attempt to " & strtarget
ON ERROR RESUME NEXT

wmiuser = strtarget & "/" & strusername


strComputer = strtarget
strNamespace = "root\cimv2"
strUser = strtarget & "\" & strusername
strPassword = strpassword

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objwbemLocator.ConnectServer _
    (strComputer, strNamespace, strUser, strPassword)
objWMIService.Security_.authenticationLevel = 6

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_OperatingSystem")
For Each objItem in ColItems
    Wscript.Echo strComputer & ": " & objItem.Caption
Next

wscript.echo "WMI Connection Return Code = " & Err.Number
wscript.echo "WMI Connection Return Description = " & Err.Description

ON ERROR GOTO 0
 

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

'Wscript.Echo "Caption: " & objItem.Caption

'Wscript.Echo "Description: " & objItem.Description

'Wscript.Echo "DeviceID: " & objItem.DeviceID

'Wscript.Echo "Size: " & objItem.Size

'Wscript.Echo "FreeSpace: " & objItem.FreeSpace



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

Set objFile = objFSO.OpenTextFile("output.csv", ForAppending)

strosname = Replace(strOSname,","," ")



objFile.WriteLine(strcomputername & "," & strOSname & "," & strVersion & "," & strRAM & "," & Counter)


objFile.Close
 
wscript.echo "Computername,OSName,OSVersion,RAM_VISABLE(GB),Data_Volume(GB)"



wscript.echo strcomputername & "," & strOSname & "," & strVersion & "," & strRAM & "," & Counter





End Function



Function PingTest(strtarget)
'#####
'ping a remote machine
'return the error code 0 = success 1 = fail

pingtest = WshShell.Run("ping.exe "& strtarget,1,1)

wscript.echo "Ping Return code = " & Err.Number

End Function


Function ConnectSSH(strtarget,strusername,strpassword)

strcommand = "Plink.exe " & strusername & "@" & strtarget & " -auto_store_key_in_cache -pw " & strpassword & " df -h"

wscript.echo "Attempting to SSH to server: " & strtarget

test = WshShell.Run(strcommand,1,1)
wscript.echo "Test Connection Return Code = " & test

Set oExec = WshShell.Exec(strcommand)

'We wait for the end of process
Do While oExec.Status = 0
     WScript.Sleep 100
Loop

wscript.echo oExec.StdOut.ReadAll()

End Function

