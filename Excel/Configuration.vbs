Option Explicit

Dim i, j, result

' ------------------------ Configuration:
Dim nPlatforms, nDevices
Dim clooConfiguration, device

Set clooConfiguration = CreateObject("ClooWrapperVBA.Configuration")

WScript.Echo "---------------Configuration test"

nPlatforms = clooConfiguration.Platforms
WScript.Echo "nPlatforms = " + CStr(nPlatforms)

If nPlatforms = 0 then
	WScript.Echo "Something went wrong: No available platforms found!"
	Wscript.Quit
End If

For i = 1 To nPlatforms
	result = clooConfiguration.SetPlatform(i - 1)
	If result Then
		WScript.Echo "Platform " + CStr(i - 1) + ":"
		WScript.Echo "	PlatformName = " + clooConfiguration.Platform.PlatformName
		WScript.Echo "	PlatformVendor = " + clooConfiguration.Platform.PlatformVendor
		WScript.Echo "	Number of devices = " + CStr(clooConfiguration.Platform.Devices)
		nDevices = clooConfiguration.Platform.Devices
		
		For j = 1 To nDevices
			result = clooConfiguration.Platform.SetDevice(j - 1)
			WScript.Echo "	Device " + CStr(j - 1) + ":" 
			WScript.Echo "		DeviceType = " + clooConfiguration.Platform.Device.DeviceType
			WScript.Echo "		DeviceName = " + clooConfiguration.Platform.Device.DeviceName
			WScript.Echo "		DeviceVendor = " + clooConfiguration.Platform.Device.DeviceVendor
			WScript.Echo "		MaxComputeUnits = " + CStr(clooConfiguration.Platform.Device.MaxComputeUnits)
			WScript.Echo "		DeviceAvailable = " + GetBooleanAsString(clooConfiguration.Platform.Device.DeviceAvailable)
			WScript.Echo "		CompilerAvailable = " + GetBooleanAsString(clooConfiguration.Platform.Device.CompilerAvailable)
		Next
	End If
Next

Function GetBooleanAsString(x)
	If x Then
		GetBooleanAsString = "true"
	Else
		GetBooleanAsString = "false"
	End If
End Function