Attribute VB_Name = "Performance"
Option Explicit

Private Const ARRAY_SIZE = 1000

Sub VBA_PerformanceTest()
    Dim wsPerformanceTest As Worksheet
    Dim m1!(), m2!(), vecM1!(), vecM2!(), vecResp!(), resultVba!(), vecQ&(0)
    Dim x1!(0), x2!(0), res!(0)
    Dim finalResults!()
    Dim i&, j&, k&, p&, q&, r&
    Dim buildLogs$, sources$, result As Boolean
    Dim cTime As New CTimer
    Dim globalWorkSize&(1), localWorkSize&(), globalWorkOffset&()
    Dim calcCorrect As Boolean
    Dim programDevice_Performance As ClooWrapperVBA.ProgramDevice
    Dim progDevices As Collection
    
    Set wsPerformanceTest = ThisWorkbook.Worksheets("Performance")
    wsPerformanceTest.Range("B2:C4").ClearContents
    
    p = ARRAY_SIZE: q = ARRAY_SIZE: r = ARRAY_SIZE
    
    ReDim resultVba(p - 1, r - 1)
    
    ' Dimensions of matrices:
    ReDim m1(p - 1, q - 1)
    ReDim m2(q - 1, r - 1)
    ReDim vecResp(p * r - 1)
    
    Randomize
    For i = 0 To p - 1
        For j = 0 To q - 1
            m1(i, j) = CInt((Rnd() - 0.5) * 100#)
        Next j
    Next i
    
    For i = 0 To q - 1
        For j = 0 To r - 1
            m2(i, j) = CInt((Rnd() - 0.5) * 100#)
        Next j
    Next i
    vecM1 = MatrixToVectorSingle(m1, p, q)
    vecM2 = MatrixToVectorSingle(m2, q, r)
    
    ' VBA matrix multiplication:
    cTime.StartCounter
    For i = 0 To p - 1
        For j = 0 To r - 1
            For k = 0 To q - 1
                resultVba(i, j) = resultVba(i, j) + m1(i, k) * m2(k, j)
            Next k
        Next j
    Next i
    wsPerformanceTest.Cells(2, 2) = cTime.TimeElapsed
    
    Open Application.ActiveWorkbook.Path & "\cl\FloatMatrixMultiplication.cl" For Binary As #1
    sources = Space$(LOF(1))
    Get #1, , sources
    Close #1
    
    ' Adding of all CPU and GPU devices to collection.
    Set progDevices = CreateDeviceCollection(sources)
    
    If progDevices Is Nothing Then
        MsgBox ("No devices found! Something is wrong!")
        Exit Sub
    End If
    
    If Not (GetFirstDeviceOfType(progDevices, "CPU") Is Nothing) Then
        ' CPU calculations.
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "CPU")
    
        result = programDevice_Performance.CreateKernel("FloatMatrixMult")
    
        globalWorkSize(0) = p
        globalWorkSize(1) = r
        vecQ(0) = q
    
        result = programDevice_Performance.SetMemoryArgument_Single(0, vecResp)
        result = programDevice_Performance.SetMemoryArgument_Single(1, vecM1)
        result = programDevice_Performance.SetMemoryArgument_Single(2, vecM2)
        result = programDevice_Performance.SetMemoryArgument_Long(3, vecQ)
    
        ' Start once to update cashes.
        result = programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
    
        ' Start real measurements.
        cTime.StartCounter
        result = programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        wsPerformanceTest.Cells(3, 2) = cTime.TimeElapsed
    
        result = programDevice_Performance.GetMemoryArgument_Single(0, vecResp)
        finalResults = VectorToMatrixSingle(vecResp, p, r)
    
        ' Comparison to VBA result.
        calcCorrect = True
        For i = 0 To p - 1
            For j = 0 To r - 1
                If Abs(finalResults(i, j) - resultVba(i, j)) > 1E-20 Then
                    calcCorrect = False
                End If
            Next j
        Next i
        wsPerformanceTest.Cells(3, 3) = calcCorrect
        
        result = programDevice_Performance.ReleaseMemObject(3)
        result = programDevice_Performance.ReleaseMemObject(2)
        result = programDevice_Performance.ReleaseMemObject(1)
        result = programDevice_Performance.ReleaseMemObject(0)
        result = programDevice_Performance.ReleaseKernel
        result = programDevice_Performance.ReleaseProgram
    Else
        wsPerformanceTest.Cells(3, 2) = CVErr(2042)
        wsPerformanceTest.Cells(3, 3) = CVErr(2042)
    End If
    
    ' GPU calculations.
    If Not (GetFirstDeviceOfType(progDevices, "GPU") Is Nothing) Then
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "GPU")
        
        result = programDevice_Performance.CreateKernel("FloatMatrixMult")
        
        globalWorkSize(0) = p
        globalWorkSize(1) = r
        vecQ(0) = q
        
        ReDim vecResp(p * r - 1)
        result = programDevice_Performance.SetMemoryArgument_Single(0, vecResp)
        result = programDevice_Performance.SetMemoryArgument_Single(1, vecM1)
        result = programDevice_Performance.SetMemoryArgument_Single(2, vecM2)
        result = programDevice_Performance.SetMemoryArgument_Long(3, vecQ)
        
        ' Start once to update cashes.
        Call programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        
        ' Start real measurements.
        cTime.StartCounter
        result = programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        wsPerformanceTest.Cells(4, 2) = cTime.TimeElapsed
        
        result = programDevice_Performance.GetMemoryArgument_Single(0, vecResp)
        finalResults = VectorToMatrixSingle(vecResp, p, r)
        
        ' Comparison to VBA result.
        calcCorrect = True
        For i = 0 To p - 1
            For j = 0 To r - 1
                If Abs(finalResults(i, j) - resultVba(i, j)) > 1E-20 Then
                    calcCorrect = False
                End If
            Next j
        Next i
        wsPerformanceTest.Cells(4, 3) = calcCorrect
        result = programDevice_Performance.ReleaseMemObject(3)
        result = programDevice_Performance.ReleaseMemObject(2)
        result = programDevice_Performance.ReleaseMemObject(1)
        result = programDevice_Performance.ReleaseMemObject(0)
        result = programDevice_Performance.ReleaseKernel
        result = programDevice_Performance.ReleaseProgram
    Else
        wsPerformanceTest.Cells(4, 2) = CVErr(2042)
        wsPerformanceTest.Cells(4, 3) = CVErr(2042)
    End If
End Sub

Sub GpuCpu_FloatDouble_PerformanceTest()
    GpuCpu_Float_PerformanceTest
    GpuCpu_Double_PerformanceTest
End Sub

Sub GpuCpu_Float_PerformanceTest()
    Dim wsPerformanceTest As Worksheet
    Dim upper&, singles!(), aSingle!, i&
    Dim sources$, result As Boolean
    Dim progDevices As Collection
    Dim programDevice_Performance As ClooWrapperVBA.ProgramDevice
    
    Set wsPerformanceTest = ThisWorkbook.Worksheets("Performance")
    wsPerformanceTest.Range("E3:E4").ClearContents
    
    upper = 10000000
    ReDim singles(upper)
    
    For i = 0 To upper - 1
        singles(i) = i
    Next i
    aSingle = 2!
    
    Open Application.ActiveWorkbook.Path & "\cl\FloatPerformance.cl" For Binary As #1
    sources = Space$(LOF(1))
    Get #1, , sources
    Close #1
    
    ' Adding of all CPU and GPU devices to collection.
    Set progDevices = CreateDeviceCollection(sources)
    
    If progDevices Is Nothing Then
        MsgBox ("No devices found! Something is wrong!")
        Exit Sub
    End If
    
    If GetFirstDeviceOfType(progDevices, "GPU") Is Nothing Then
        wsPerformanceTest.Cells(4, 5) = CVErr(2042)
    Else
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "GPU")
        wsPerformanceTest.Cells(4, 5) = PerformanceTest_Single(upper, singles, aSingle, programDevice_Performance)
        
        result = programDevice_Performance.ReleaseMemObject(1)
        result = programDevice_Performance.ReleaseMemObject(0)
        result = programDevice_Performance.ReleaseKernel
        result = programDevice_Performance.ReleaseProgram
    End If
    
    If GetFirstDeviceOfType(progDevices, "CPU") Is Nothing Then
        wsPerformanceTest.Cells(3, 5) = CVErr(2042)
    Else
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "CPU")
        wsPerformanceTest.Cells(3, 5) = PerformanceTest_Single(upper, singles, aSingle, programDevice_Performance)
        
        result = programDevice_Performance.ReleaseMemObject(1)
        result = programDevice_Performance.ReleaseMemObject(0)
        result = programDevice_Performance.ReleaseKernel
        result = programDevice_Performance.ReleaseProgram
    End If
End Sub

Sub GpuCpu_Double_PerformanceTest()
    Dim wsPerformanceTest As Worksheet
    Dim upper&, doubles#(), aDouble#, i&
    Dim sources$, result As Boolean
    Dim progDevices As Collection
    Dim programDevice_Performance As ClooWrapperVBA.ProgramDevice
    
    Set wsPerformanceTest = ThisWorkbook.Worksheets("Performance")
    wsPerformanceTest.Range("F3:F4").ClearContents
    
    upper = 10000000
    ReDim doubles(upper)
    
    For i = 0 To upper - 1
        doubles(i) = i
    Next i
    aDouble = 2#
    
    Open Application.ActiveWorkbook.Path & "\cl\DoublePerformance.cl" For Binary As #1
    sources = Space$(LOF(1))
    Get #1, , sources
    Close #1
    
    ' Adding of all CPU and GPU devices to collection.
    Set progDevices = CreateDeviceCollection(sources)
    
    If progDevices Is Nothing Then
        MsgBox ("No devices found! Something is wrong!")
        Exit Sub
    End If
    
    If GetFirstDeviceOfType(progDevices, "GPU") Is Nothing Then
        wsPerformanceTest.Cells(4, 6) = CVErr(2042)
    Else
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "GPU")
        wsPerformanceTest.Cells(4, 6) = PerformanceTest_Double(upper, doubles, aDouble, programDevice_Performance)
        
        result = programDevice_Performance.ReleaseMemObject(1)
        result = programDevice_Performance.ReleaseMemObject(0)
        result = programDevice_Performance.ReleaseKernel
        result = programDevice_Performance.ReleaseProgram
    End If
    
    If GetFirstDeviceOfType(progDevices, "CPU") Is Nothing Then
        wsPerformanceTest.Cells(3, 6) = CVErr(2042)
    Else
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "CPU")
        wsPerformanceTest.Cells(3, 6) = PerformanceTest_Double(upper, doubles, aDouble, programDevice_Performance)
        
        result = programDevice_Performance.ReleaseMemObject(1)
        result = programDevice_Performance.ReleaseMemObject(0)
        result = programDevice_Performance.ReleaseKernel
        result = programDevice_Performance.ReleaseProgram
    End If
End Sub

' Single precision performance at CPU / GPU.
Function PerformanceTest_Single(upper&, singles!(), aSingle!, programDevice_Performance As ClooWrapperVBA.ProgramDevice)
    Dim buildLogs$, result As Boolean
    
    result = programDevice_Performance.CreateKernel("FloatPerformance")
    
    result = programDevice_Performance.SetMemoryArgument_Single(0, singles)
    result = programDevice_Performance.SetValueArgument_Single(1, aSingle)
    
    PerformanceTest_Single = PerformanceTestExecution(upper, programDevice_Performance)
End Function

' Double precision performance at CPU / GPU.
Function PerformanceTest_Double(upper&, doubles#(), aDouble#, programDevice_Performance As ClooWrapperVBA.ProgramDevice)
    Dim buildLogs$, result As Boolean
    
    result = programDevice_Performance.CreateKernel("DoublePerformance")
    
    result = programDevice_Performance.SetMemoryArgument_Double(0, doubles)
    result = programDevice_Performance.SetValueArgument_Double(1, aDouble)
    
    PerformanceTest_Double = PerformanceTestExecution(upper, programDevice_Performance)
End Function

Function PerformanceTestExecution(upper&, programDevice_Performance)
    Dim globalWorkSize&(0), localWorkSize&(), globalWorkOffset&()
    Dim elTime#
    Dim cTime As New CTimer
    
    cTime.StartCounter
    globalWorkSize(0) = 10
    Do While cTime.TimeElapsed < 25
        cTime.StartCounter

        Call programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)

        If globalWorkSize(0) = upper Then
            Exit Do
        End If
        
        globalWorkSize(0) = globalWorkSize(0) * 2
        
        If globalWorkSize(0) > upper Then
            elTime = cTime.TimeElapsed / 1000#
            globalWorkSize(0) = upper
        End If
    Loop
    
    elTime = cTime.TimeElapsed / 1000#
    
    PerformanceTestExecution = (4096# * globalWorkSize(0) / elTime / 1000000000#)
End Function

Sub Test_OneAfterAnother()
    Dim m1!(), m2!(), vecM1!(), vecM2!(), vecResp!(), resultVba!(), vecQ&(0)
    Dim x1!(0), x2!(0), res!(0)
    Dim finalResults!()
    Dim i&, j&, k&, p&, q&, r&
    Dim buildLogs$, sources$, result As Boolean
    Dim cTime As New CTimer
    Dim globalWorkSize&(1), localWorkSize&(), globalWorkOffset&()
    Dim calcCorrect As Boolean
    Dim programDevice_Performance As ClooWrapperVBA.ProgramDevice
    Dim progDevices As Collection
    
    p = 2: q = 2: r = 2
    
    ReDim resultVba(p - 1, r - 1)
    
    ' Dimensions of matrices:
    ReDim m1(p - 1, q - 1)
    ReDim m2(q - 1, r - 1)
    ReDim vecResp(p * r - 1)
    
    m1(0, 0) = 1: m1(0, 1) = 2: m1(1, 0) = 3: m1(1, 1) = 4
    m2(0, 0) = 2: m2(0, 1) = 3: m2(1, 0) = 4: m2(1, 1) = 5
    
    vecM1 = MatrixToVector(m1, p, q)
    vecM2 = MatrixToVector(m2, q, r)
    
    ' VBA matrix multiplication:
    For i = 0 To p - 1
        For j = 0 To r - 1
            For k = 0 To q - 1
                resultVba(i, j) = resultVba(i, j) + m1(i, k) * m2(k, j)
            Next k
        Next j
    Next i
    
    Open Application.ActiveWorkbook.Path & "\cl\FloatMatrixMultiplication.cl" For Binary As #1
    sources = Space$(LOF(1))
    Get #1, , sources
    Close #1
    
    ' Adding of all CPU and GPU devices to collection.
    Set progDevices = CreateDeviceCollection(sources)
    
    If progDevices Is Nothing Then
        MsgBox ("No devices found! Something is wrong!")
        Exit Sub
    End If
    
    If Not (GetFirstDeviceOfType(progDevices, "CPU") Is Nothing) Then
        ' CPU calculations.
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "CPU")
    
        result = programDevice_Performance.CreateKernel("FloatMatrixMult")
    
        globalWorkSize(0) = p
        globalWorkSize(1) = r
        vecQ(0) = q
    
        result = programDevice_Performance.SetMemoryArgument_Single(0, vecResp)
        result = programDevice_Performance.SetMemoryArgument_Single(1, vecM1)
        result = programDevice_Performance.SetMemoryArgument_Single(2, vecM2)
        result = programDevice_Performance.SetMemoryArgument_Long(3, vecQ)
        
        result = programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        
        result = programDevice_Performance.GetMemoryArgument_Single(0, vecResp)
        finalResults = VectorToMatrix(vecResp, p, r)
        
        ' Comparison to VBA result.
        calcCorrect = True
        For i = 0 To p - 1
            For j = 0 To r - 1
                If Abs(finalResults(i, j) - resultVba(i, j)) > 1E-20 Then
                    calcCorrect = False
                End If
            Next j
        Next i
        
        result = programDevice_Performance.SetMemoryArgument_Single(1, vecM2)
        result = programDevice_Performance.SetMemoryArgument_Single(2, vecM1)
        
        result = programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        result = programDevice_Performance.GetMemoryArgument_Single(0, vecResp)
        finalResults = VectorToMatrix(vecResp, p, r)
        
        ' VBA matrix multiplication:
        ReDim resultVba(p - 1, r - 1)
        For i = 0 To p - 1
            For j = 0 To r - 1
                For k = 0 To q - 1
                    resultVba(i, j) = resultVba(i, j) + m2(i, k) * m1(k, j)
                Next k
            Next j
        Next i
        
        ' Comparison to VBA result.
        calcCorrect = True
        For i = 0 To p - 1
            For j = 0 To r - 1
                If Abs(finalResults(i, j) - resultVba(i, j)) > 1E-20 Then
                    calcCorrect = False
                End If
            Next j
        Next i
    End If
End Sub
