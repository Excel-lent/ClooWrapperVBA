Attribute VB_Name = "HelloWorld"
Option Explicit

Sub HelloWorld()
    Dim wsHelloWorld As Worksheet
    Dim nRows%, currentRow&, nPlatforms&, nDevices&, i&, j&, result As Boolean
    Dim deviceType$, platformName$, platformVendor$, platformVersion$, deviceVendor$, deviceVersion$, driverVersion$, openCLCVersionString$
    Dim maxComputeUnits&, globalMemorySize#, maxClockFrequency#, maxMemoryAllocationSize#, deviceName$, sources$, cpuCounter&, gpuCounter&
    Dim buildLogs$, platformId&, DeviceId&, errorString$
    Dim deviceAvailable As Boolean, compilerAvailable As Boolean
    Dim m1!(1, 1), m2!(1, 1), vecM1!(), vecM2!(), vecQ&(0), vecResp!(3), globalWorkOffset&(), globalWorkSize&(1), localWorkSize&()
    Dim p&, q&, r&, resp!()
    
    Dim clooConfiguration As New ClooWrapperVBA.Configuration
    Dim progDevice As ClooWrapperVBA.ProgramDevice
    
    Set wsHelloWorld = ThisWorkbook.Worksheets("Hello World!")
    
    ' Cleanup.
    nRows = wsHelloWorld.Cells(Rows.Count, 4).End(xlUp).Row
    If nRows >= 2 Then
        wsHelloWorld.Range(wsHelloWorld.Cells(2, 4), wsHelloWorld.Cells(nRows, 4)).ClearContents
    End If
    
    ' Read configuration.
    nPlatforms = clooConfiguration.platforms
    
    currentRow = 2
    For i = 1 To nPlatforms
        result = clooConfiguration.SetPlatform(i - 1)
        If result Then
            platformName = clooConfiguration.Platform.platformName
            platformVendor = clooConfiguration.Platform.platformVendor
            platformVersion = clooConfiguration.Platform.platformVersion
            
            wsHelloWorld.Cells(currentRow, 1) = "Platform": wsHelloWorld.Cells(currentRow, 2) = i - 1: currentRow = currentRow + 1
            wsHelloWorld.Cells(currentRow, 2) = "Name": wsHelloWorld.Cells(currentRow, 2) = platformName: currentRow = currentRow + 1
            wsHelloWorld.Cells(currentRow, 2) = "Vendor": wsHelloWorld.Cells(currentRow, 3) = platformVendor: currentRow = currentRow + 1
            wsHelloWorld.Cells(currentRow, 2) = "Version": wsHelloWorld.Cells(currentRow, 3) = platformVersion: currentRow = currentRow + 1
            
            nDevices = clooConfiguration.Platform.Devices
            For j = 1 To nDevices
                result = clooConfiguration.Platform.SetDevice(j - 1)
                
                If result Then
                    deviceType = clooConfiguration.Platform.device.deviceType
                    deviceName = clooConfiguration.Platform.device.deviceName
                    deviceVendor = clooConfiguration.Platform.device.deviceVendor
                    maxComputeUnits = clooConfiguration.Platform.device.maxComputeUnits
                    deviceAvailable = clooConfiguration.Platform.device.deviceAvailable
                    compilerAvailable = clooConfiguration.Platform.device.compilerAvailable
                    deviceVersion = clooConfiguration.Platform.device.deviceVersion
                    driverVersion = clooConfiguration.Platform.device.driverVersion
                    globalMemorySize = clooConfiguration.Platform.device.globalMemorySize
                    maxClockFrequency = clooConfiguration.Platform.device.maxClockFrequency
                    maxMemoryAllocationSize = clooConfiguration.Platform.device.maxMemoryAllocationSize
                    openCLCVersionString = clooConfiguration.Platform.device.openCLCVersionString
                    
                    wsHelloWorld.Cells(currentRow, 2) = "Device": wsHelloWorld.Cells(currentRow, 3) = j - 1: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "Type": wsHelloWorld.Cells(currentRow, 4) = deviceType: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "Name": wsHelloWorld.Cells(currentRow, 4) = deviceName: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "Vendor": wsHelloWorld.Cells(currentRow, 4) = deviceVendor: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "MaxComputeUnits": wsHelloWorld.Cells(currentRow, 4) = maxComputeUnits: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "DeviceAvailable": wsHelloWorld.Cells(currentRow, 4) = deviceAvailable: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "CompilerAvailable": wsHelloWorld.Cells(currentRow, 4) = compilerAvailable: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "DeviceVersion": wsHelloWorld.Cells(currentRow, 4) = deviceVersion: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "DriverVersion": wsHelloWorld.Cells(currentRow, 4) = driverVersion: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "GlobalMemorySize, bytes": wsHelloWorld.Cells(currentRow, 4) = globalMemorySize: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "MaxClockFrequency, MHz": wsHelloWorld.Cells(currentRow, 4) = maxClockFrequency: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "MaxMemoryAllocationSize, bytes": wsHelloWorld.Cells(currentRow, 4) = maxMemoryAllocationSize: currentRow = currentRow + 1
                    wsHelloWorld.Cells(currentRow, 3) = "OpenCLCVersionString": wsHelloWorld.Cells(currentRow, 4) = openCLCVersionString: currentRow = currentRow + 1
                End If
            Next j
        End If
    Next i
    
    ' Multiplication of two matrices.
    ' Read the OpenCL sources.
    Open Application.ActiveWorkbook.Path & "\cl\FloatMatrixMultiplication.cl" For Binary As #1
    sources = Space$(LOF(1))
    Get #1, , sources
    Close #1
    
    ' Find the first found device that can compile the code.
    platformId = 0
    Do While platformId <= clooConfiguration.platforms - 1
        result = clooConfiguration.SetPlatform(platformId)
        cpuCounter = 0
        gpuCounter = 0
        For DeviceId = 0 To clooConfiguration.Platform.Devices - 1
            result = clooConfiguration.Platform.SetDevice(DeviceId)
            
            If clooConfiguration.Platform.device.compilerAvailable Then
                If clooConfiguration.Platform.device.deviceType = "CPU" Then
                    Set progDevice = New ClooWrapperVBA.ProgramDevice
                    result = progDevice.Build(sources, "", platformId, DeviceId, cpuCounter, buildLogs)
                    If result Then
                        Exit Do
                    Else
                        cpuCounter = cpuCounter + 1
                    End If
                End If
                If clooConfiguration.Platform.device.deviceType = "GPU" Then
                    Set progDevice = New ClooWrapperVBA.ProgramDevice
                    result = progDevice.Build(sources, "", platformId, DeviceId, gpuCounter, buildLogs)
                    gpuCounter = gpuCounter + 1
                    If result Then
                        Exit Do
                    Else
                        gpuCounter = gpuCounter + 1
                    End If
                End If
            End If
        Next DeviceId
        platformId = platformId + 1
    Loop
    
    errorString = progDevice.errorString
    result = progDevice.CreateKernel("FloatMatrixMult")
    
    ' Initialization of arrays:
    p = 2: q = 2: r = 2
    For i = 0 To p - 1
        For j = 0 To q - 1
            m1(i, j) = wsHelloWorld.Cells(i + 1, j + 7)
        Next j
    Next i
    vecM1 = MatrixToVectorSingle(m1, p, q)
    For i = 0 To q - 1
        For j = 0 To r - 1
            m2(i, j) = wsHelloWorld.Cells(i + 3, j + 7)
        Next j
    Next i
    vecM2 = MatrixToVectorSingle(m2, q, r)
    vecQ(0) = q
    
    result = progDevice.SetMemoryArgument_Single(0, vecResp)
    result = progDevice.SetMemoryArgument_Single(1, vecM1)
    result = progDevice.SetMemoryArgument_Single(2, vecM2)
    result = progDevice.SetMemoryArgument_Long(3, vecQ)
    
    globalWorkSize(0) = p
    globalWorkSize(1) = r
    
    result = progDevice.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
    
    result = progDevice.GetMemoryArgument_Single(0, vecResp)
    
    resp = VectorToMatrixSingle(vecResp, p, r)
    
    For i = 0 To p - 1
        For j = 0 To r - 1
            wsHelloWorld.Cells(i + 5, j + 7) = resp(i, j)
        Next j
    Next i
    
    result = progDevice.ReleaseMemObject(3)
    result = progDevice.ReleaseMemObject(2)
    result = progDevice.ReleaseMemObject(1)
    result = progDevice.ReleaseMemObject(0)
    result = progDevice.ReleaseKernel
    result = progDevice.ReleaseProgram
End Sub
