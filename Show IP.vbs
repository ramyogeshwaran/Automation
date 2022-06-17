Set wshShell = WScript.CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
'strComputerName = wshShell.ExpandEnvironmentStrings( "%Logonserver%" )
set objNetwork = CreateObject("WScript.Network")
strComputerName = objNetwork.Computername
Set wmiobj = GetObject("winmgmts://localhost/root/cimv2:Win32_BIOS")
For Each ver In wmiobj.Instances_


Set IPConfigSet = GetObject("winmgmts://.").ExecQuery("select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    for each IPConfig In IPConfigSet
            If Not IsNull(IPConfig.IPAddress) Then
                   'For i=LBound(IPConfig.IPAddress) To UBound(IPConfig.IPAddress)
info = ("IPAddress: " & IPConfig.IPAddress(i) & vbCrLf & "Computer Name: " & strComputerName & vbCrLf & "DELL Service Tag: " & ver.SerialNumber)
	msgbox info ,vbInformation,"Blue-Net"
                    'WScript.Echo ("IPAddress: " & IPConfig.IPAddress(i) & vbCrLf & "Computer Name: " & strComputerName), "Title"
                   'Next
            End If
    Next
	next
End
