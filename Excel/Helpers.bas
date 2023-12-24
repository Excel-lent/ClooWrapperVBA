Attribute VB_Name = "Helpers"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
#End If

Function GetFirstDeviceOfType(progDevices As Collection, deviceTypeToSearch$) As Variant
    Dim i%
    
    Set GetFirstDeviceOfType = Nothing
    For i = 1 To progDevices.Count: Do
        If progDevices.Item(i).deviceType = deviceTypeToSearch Then
            Set GetFirstDeviceOfType = progDevices.Item(i).ProgramDevice
            Exit Do
        End If
    Loop While False: Next i
End Function

Function CreateDeviceCollection(sources$)
    Dim progDevices As New Collection
    Dim clooConfiguration As New ClooWrapperVBA.Configuration
    Dim i&, j&, cpuCounter&, gpuCounter&
    Dim result As Boolean
    Dim progDevice As ProgrammingDevice
    Dim buildLogs$
    
    ' Adding of all CPU and GPU devices to collection.
    For i = 1 To clooConfiguration.platforms
        result = clooConfiguration.SetPlatform(i - 1)
        For j = 1 To clooConfiguration.Platform.Devices
            result = clooConfiguration.Platform.SetDevice(j - 1)
            
            Set progDevice = New ProgrammingDevice
            
            If clooConfiguration.Platform.device.deviceType = "CPU" Then
                result = progDevice.ProgramDevice.Build(sources, "", i - 1, j - 1, cpuCounter, buildLogs)
                progDevice.DeviceId = cpuCounter
                progDevice.deviceType = "CPU"
                If result = True Then cpuCounter = cpuCounter + 1
            ElseIf clooConfiguration.Platform.device.deviceType = "GPU" Then
                result = progDevice.ProgramDevice.Build(sources, "", i - 1, j - 1, gpuCounter, buildLogs)
                progDevice.DeviceId = gpuCounter
                progDevice.deviceType = "GPU"
                If result = True Then gpuCounter = gpuCounter + 1
            Else
                result = False
            End If
            
            If result Then
                Call progDevices.Add(progDevice)
            End If
        Next j
    Next i
    
    If cpuCounter + gpuCounter = 0 Then
        Set CreateDeviceCollection = Nothing
    Else
        Set CreateDeviceCollection = progDevices
    End If
End Function

Function MatrixToVectorSingle(m() As Single, maxi As Long, maxj As Long) As Single()
    Dim v() As Single
    Dim i&, j&
    
    ReDim v(maxi * maxj - 1)
    
    For i = 0 To maxi - 1
        For j = 0 To maxj - 1
            v(i + maxi * j) = m(i, j)
        Next j
    Next i
    
    MatrixToVectorSingle = v
End Function

Function VectorToMatrixSingle(v() As Single, maxi As Long, maxj As Long) As Single()
    Dim i&, j&
    Dim m() As Single
    
    ReDim m(maxi - 1, maxj - 1)
    
    For i = 0 To maxi - 1
        For j = 0 To maxj - 1
            m(i, j) = v(i + maxi * j)
        Next j
    Next i
    
    VectorToMatrixSingle = m
End Function

Function MatrixToVectorDouble(m() As Double, maxi As Long, maxj As Long) As Double()
    Dim v() As Double
    Dim i&, j&
    
    ReDim v(maxi * maxj - 1)
    
    For i = 0 To maxi - 1
        For j = 0 To maxj - 1
            v(i + maxi * j) = m(i, j)
        Next j
    Next i
    
    MatrixToVectorDouble = v
End Function

Function VectorToMatrixDouble(v() As Double, maxi As Long, maxj As Long) As Double()
    Dim i&, j&
    Dim m() As Double
    
    ReDim m(maxi - 1, maxj - 1)
    
    For i = 0 To maxi - 1
        For j = 0 To maxj - 1
            m(i, j) = v(i + maxi * j)
        Next j
    Next i
    
    VectorToMatrixDouble = m
End Function
