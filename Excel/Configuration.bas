Attribute VB_Name = "Configuration"
Option Explicit

Sub GetConfiguration()
    Dim clooConfiguration As New ClooWrapperVBA.Configuration
    Dim nPlatforms As Long
    Dim nDevices As Long
    Dim i%, j%, k%, currentColumn%, nColumns%, nRows%
    Dim tmpStrings() As String
    Dim result As Boolean
    Dim wsConfiguration As Worksheet
    
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    
    ' Cleanup.
    nColumns = wsConfiguration.Cells(1, Columns.Count).End(xlToLeft).Column
    For i = 1 To nColumns
        nRows = wsConfiguration.Cells(Rows.Count, i).End(xlUp).Row
        wsConfiguration.Range(wsConfiguration.Cells(1, i), wsConfiguration.Cells(nRows, i)).ClearContents
    Next i
    
    nPlatforms = clooConfiguration.platforms
    
    For i = 1 To nPlatforms
        result = clooConfiguration.SetPlatform(i - 1)
        If result Then
            currentColumn = currentColumn + 1
            wsConfiguration.Cells(1, currentColumn) = "Platform"
            wsConfiguration.Cells(2, currentColumn) = "Name"
            wsConfiguration.Cells(3, currentColumn) = "Vendor"
            wsConfiguration.Cells(4, currentColumn) = "Profile"
            wsConfiguration.Cells(5, currentColumn) = "Version"
            wsConfiguration.Cells(6, currentColumn) = "Extensions"
            currentColumn = currentColumn + 1
            wsConfiguration.Cells(1, currentColumn) = i - 1
            wsConfiguration.Cells(2, currentColumn) = clooConfiguration.Platform.PlatformName
            wsConfiguration.Cells(3, currentColumn) = clooConfiguration.Platform.PlatformVendor
            wsConfiguration.Cells(4, currentColumn) = clooConfiguration.Platform.PlatformProfile
            wsConfiguration.Cells(5, currentColumn) = clooConfiguration.Platform.PlatformVersion
            tmpStrings = clooConfiguration.Platform.PlatformExtensions
            For j = 0 To UBound(tmpStrings)
                wsConfiguration.Cells(6 + j, currentColumn) = tmpStrings(j)
            Next j
            
            nDevices = clooConfiguration.Platform.Devices
            For j = 1 To nDevices
                result = clooConfiguration.Platform.SetDevice(j - 1)
                If result Then
                    currentColumn = currentColumn + 1
                    wsConfiguration.Cells(1, currentColumn) = "Platform"
                    wsConfiguration.Cells(2, currentColumn) = "Device"
                    wsConfiguration.Cells(3, currentColumn) = "Type"
                    wsConfiguration.Cells(4, currentColumn) = "Name"
                    wsConfiguration.Cells(5, currentColumn) = "Vendor"
                    wsConfiguration.Cells(6, currentColumn) = "MaxComputeUnits"
                    wsConfiguration.Cells(7, currentColumn) = "AddressBits"
                    wsConfiguration.Cells(8, currentColumn) = "DeviceAvailable"
                    wsConfiguration.Cells(9, currentColumn) = "CompilerAvailable"
                    wsConfiguration.Cells(10, currentColumn) = "CommandQueueFlags"
                    wsConfiguration.Cells(11, currentColumn) = "DeviceVersion"
                    wsConfiguration.Cells(12, currentColumn) = "DriverVersion"
                    wsConfiguration.Cells(13, currentColumn) = "EndianLittle"
                    wsConfiguration.Cells(14, currentColumn) = "ErrorCorrectionSupport"
                    wsConfiguration.Cells(15, currentColumn) = "SingleCapabilites"
                    wsConfiguration.Cells(16, currentColumn) = "ExecutionCapabilities"
                    wsConfiguration.Cells(17, currentColumn) = "DeviceExtensions"
                    wsConfiguration.Cells(18, currentColumn) = "GlobalMemoryCacheLineSize, bytes"
                    wsConfiguration.Cells(19, currentColumn) = "GlobalMemoryCacheSize, bytes"
                    wsConfiguration.Cells(20, currentColumn) = "GlobalMemoryCacheType"
                    wsConfiguration.Cells(21, currentColumn) = "GlobalMemorySize, bytes"
                    wsConfiguration.Cells(22, currentColumn) = "HostUnifiedMemory"
                    wsConfiguration.Cells(23, currentColumn) = "ImageSupport"
                    wsConfiguration.Cells(24, currentColumn) = "Image2DMaxHeight"
                    wsConfiguration.Cells(25, currentColumn) = "Image2DMaxWidth"
                    wsConfiguration.Cells(26, currentColumn) = "Image3DMaxDepth"
                    wsConfiguration.Cells(27, currentColumn) = "Image3DMaxHeight"
                    wsConfiguration.Cells(28, currentColumn) = "Image3DMaxWidth"
                    wsConfiguration.Cells(29, currentColumn) = "LocalMemorySize, bytes"
                    wsConfiguration.Cells(30, currentColumn) = "LocalMemoryType"
                    wsConfiguration.Cells(31, currentColumn) = "MaxClockFrequency, MHz"
                    wsConfiguration.Cells(32, currentColumn) = "MaxConstantArguments"
                    wsConfiguration.Cells(33, currentColumn) = "MaxConstantBufferSize, bytes"
                    wsConfiguration.Cells(34, currentColumn) = "MaxMemoryAllocationSize, bytes"
                    wsConfiguration.Cells(35, currentColumn) = "MaxParameterSize, bytes"
                    wsConfiguration.Cells(36, currentColumn) = "MaxReadImageArguments"
                    wsConfiguration.Cells(37, currentColumn) = "MaxSamplers"
                    wsConfiguration.Cells(38, currentColumn) = "MaxWorkGroupSize"
                    wsConfiguration.Cells(39, currentColumn) = "MaxWorkItemDimensions"
                    
'                    For k = 0 To UBound(cloo.Platform.device.MaxWorkItemSizes)
'                        wsConfiguration.Cells(40 + k, currentColumn) = "MaxWorkItem[" & CStr(k) & "] Size"
'                    Next k
                    
                    wsConfiguration.Cells(41 + k, currentColumn) = "MaxWriteImageArguments"
                    wsConfiguration.Cells(42 + k, currentColumn) = "MemoryBaseAddressAlignment, bits"
                    wsConfiguration.Cells(43 + k, currentColumn) = "MinDataTypeAlignmentSize, bytes"
                    wsConfiguration.Cells(44 + k, currentColumn) = "NativeVectorWidthChar"
                    wsConfiguration.Cells(45 + k, currentColumn) = "NativeVectorWidthDouble"
                    wsConfiguration.Cells(46 + k, currentColumn) = "NativeVectorWidthFloat"
                    wsConfiguration.Cells(47 + k, currentColumn) = "NativeVectorWidthHalf"
                    wsConfiguration.Cells(48 + k, currentColumn) = "NativeVectorWidthInt"
                    wsConfiguration.Cells(49 + k, currentColumn) = "NativeVectorWidthLong"
                    wsConfiguration.Cells(50 + k, currentColumn) = "NativeVectorWidthShort"
                    wsConfiguration.Cells(51 + k, currentColumn) = "OpenCLCVersionString"
                    wsConfiguration.Cells(52 + k, currentColumn) = "PreferredVectorWidthChar"
                    wsConfiguration.Cells(53 + k, currentColumn) = "PreferredVectorWidthDouble"
                    wsConfiguration.Cells(54 + k, currentColumn) = "PreferredVectorWidthFloat"
                    wsConfiguration.Cells(55 + k, currentColumn) = "PreferredVectorWidthHalf"
                    wsConfiguration.Cells(56 + k, currentColumn) = "PreferredVectorWidthInt"
                    wsConfiguration.Cells(57 + k, currentColumn) = "PreferredVectorWidthLong"
                    wsConfiguration.Cells(58 + k, currentColumn) = "PreferredVectorWidthShort"
                    wsConfiguration.Cells(59 + k, currentColumn) = "Profile"
                    wsConfiguration.Cells(60 + k, currentColumn) = "ProfilingTimerResolution, ns"
                    wsConfiguration.Cells(61 + k, currentColumn) = "VendorId"
                    currentColumn = currentColumn + 1
                    wsConfiguration.Cells(1, currentColumn) = i - 1
                    wsConfiguration.Cells(2, currentColumn) = j - 1
                    wsConfiguration.Cells(3, currentColumn) = clooConfiguration.Platform.device.deviceType
                    wsConfiguration.Cells(4, currentColumn) = clooConfiguration.Platform.device.DeviceName
                    wsConfiguration.Cells(5, currentColumn) = clooConfiguration.Platform.device.DeviceVendor
                    wsConfiguration.Cells(6, currentColumn) = clooConfiguration.Platform.device.MaxComputeUnits
                    wsConfiguration.Cells(7, currentColumn) = clooConfiguration.Platform.device.AddressBits
                    wsConfiguration.Cells(8, currentColumn) = clooConfiguration.Platform.device.DeviceAvailable
                    wsConfiguration.Cells(9, currentColumn) = clooConfiguration.Platform.device.CompilerAvailable
                    wsConfiguration.Cells(10, currentColumn) = clooConfiguration.Platform.device.CommandQueueFlags
                    wsConfiguration.Cells(11, currentColumn) = clooConfiguration.Platform.device.DeviceVersion
                    wsConfiguration.Cells(12, currentColumn) = clooConfiguration.Platform.device.DriverVersion
                    wsConfiguration.Cells(13, currentColumn) = clooConfiguration.Platform.device.EndianLittle
                    wsConfiguration.Cells(14, currentColumn) = clooConfiguration.Platform.device.ErrorCorrectionSupport
                    wsConfiguration.Cells(15, currentColumn) = clooConfiguration.Platform.device.SingleCapabilites
                    wsConfiguration.Cells(16, currentColumn) = clooConfiguration.Platform.device.ExecutionCapabilities
                    wsConfiguration.Cells(17, currentColumn) = clooConfiguration.Platform.device.DeviceExtensions
                    wsConfiguration.Cells(18, currentColumn) = clooConfiguration.Platform.device.GlobalMemoryCacheLineSize
                    wsConfiguration.Cells(19, currentColumn) = clooConfiguration.Platform.device.GlobalMemoryCacheSize
                    wsConfiguration.Cells(20, currentColumn) = clooConfiguration.Platform.device.GlobalMemoryCacheType
                    wsConfiguration.Cells(21, currentColumn) = clooConfiguration.Platform.device.GlobalMemorySize
                    wsConfiguration.Cells(22, currentColumn) = clooConfiguration.Platform.device.HostUnifiedMemory
                    wsConfiguration.Cells(23, currentColumn) = clooConfiguration.Platform.device.ImageSupport
                    wsConfiguration.Cells(24, currentColumn) = clooConfiguration.Platform.device.Image2DMaxHeight
                    wsConfiguration.Cells(25, currentColumn) = clooConfiguration.Platform.device.Image2DMaxWidth
                    wsConfiguration.Cells(26, currentColumn) = clooConfiguration.Platform.device.Image3DMaxDepth
                    wsConfiguration.Cells(27, currentColumn) = clooConfiguration.Platform.device.Image3DMaxHeight
                    wsConfiguration.Cells(28, currentColumn) = clooConfiguration.Platform.device.Image3DMaxWidth
                    wsConfiguration.Cells(29, currentColumn) = clooConfiguration.Platform.device.LocalMemorySize
                    wsConfiguration.Cells(30, currentColumn) = clooConfiguration.Platform.device.LocalMemoryType
                    wsConfiguration.Cells(31, currentColumn) = clooConfiguration.Platform.device.MaxClockFrequency
                    wsConfiguration.Cells(32, currentColumn) = clooConfiguration.Platform.device.MaxConstantArguments
                    wsConfiguration.Cells(33, currentColumn) = clooConfiguration.Platform.device.MaxConstantBufferSize
                    wsConfiguration.Cells(34, currentColumn) = clooConfiguration.Platform.device.MaxMemoryAllocationSize
                    wsConfiguration.Cells(35, currentColumn) = clooConfiguration.Platform.device.MaxParameterSize
                    wsConfiguration.Cells(36, currentColumn) = clooConfiguration.Platform.device.MaxReadImageArguments
                    wsConfiguration.Cells(37, currentColumn) = clooConfiguration.Platform.device.MaxSamplers
                    wsConfiguration.Cells(38, currentColumn) = clooConfiguration.Platform.device.MaxWorkGroupSize
                    wsConfiguration.Cells(39, currentColumn) = clooConfiguration.Platform.device.MaxWorkItemDimensions
                    
'                    For k = 0 To UBound(cloo.Platform.device.MaxWorkItemSizes)
'                        wsConfiguration.Cells(40 + k, currentColumn) = cloo.Platform.device.MaxWorkItemSizes(k)
'                    Next k
                    
                    wsConfiguration.Cells(41 + k, currentColumn) = clooConfiguration.Platform.device.MaxWriteImageArguments
                    wsConfiguration.Cells(42 + k, currentColumn) = clooConfiguration.Platform.device.MemoryBaseAddressAlignment
                    wsConfiguration.Cells(43 + k, currentColumn) = clooConfiguration.Platform.device.MinDataTypeAlignmentSize
                    wsConfiguration.Cells(44 + k, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthChar
                    wsConfiguration.Cells(45 + k, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthDouble
                    wsConfiguration.Cells(46 + k, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthFloat
                    wsConfiguration.Cells(47 + k, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthHalf
                    wsConfiguration.Cells(48 + k + k, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthInt
                    wsConfiguration.Cells(49, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthLong
                    wsConfiguration.Cells(50 + k, currentColumn) = clooConfiguration.Platform.device.NativeVectorWidthShort
                    wsConfiguration.Cells(51 + k, currentColumn) = clooConfiguration.Platform.device.OpenCLCVersionString
                    wsConfiguration.Cells(52 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthChar
                    wsConfiguration.Cells(53 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthDouble
                    wsConfiguration.Cells(54 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthFloat
                    wsConfiguration.Cells(55 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthHalf
                    wsConfiguration.Cells(56 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthInt
                    wsConfiguration.Cells(57 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthLong
                    wsConfiguration.Cells(58 + k, currentColumn) = clooConfiguration.Platform.device.PreferredVectorWidthShort
                    wsConfiguration.Cells(59 + k, currentColumn) = clooConfiguration.Platform.device.Profile
                    wsConfiguration.Cells(60 + k, currentColumn) = clooConfiguration.Platform.device.ProfilingTimerResolution
                    wsConfiguration.Cells(61 + k, currentColumn) = clooConfiguration.Platform.device.VendorId
                End If
            Next j
        End If
    Next i
End Sub
