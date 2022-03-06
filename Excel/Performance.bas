Attribute VB_Name = "Performance"
Option Explicit

Private Const ARRAY_SIZE = 1200

Sub VBA_PerformanceTest()
    Dim wsPerformanceTest As Worksheet
    Dim m1#(), m2#(), vecM1#(), vecM2#(), vecResp#(), resultVba#(), vecQ&(0)
    Dim x1#(0), x2#(0), res#(0)
    Dim finalResults#()
    Dim i&, j&, k&, p&, q&, r&
    Dim buildLogs$, sources$
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
            m1(i, j) = (Rnd() - 0.5) * 10#
        Next j
    Next i
    
    For i = 0 To q - 1
        For j = 0 To r - 1
            m2(i, j) = (Rnd() - 0.5) * 10#
        Next j
    Next i
    vecM1 = MatrixToVector(m1, p, q)
    vecM2 = MatrixToVector(m2, q, r)
    
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
    
    Open Application.ActiveWorkbook.Path & "\cl\MatrixMultiplication.cl" For Binary As #1
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
    
        Call programDevice_Performance.CreateKernel("DoubleMatrixMult")
    
        globalWorkSize(0) = p
        globalWorkSize(1) = r
        vecQ(0) = q
    
        Call programDevice_Performance.SetMemoryArgument_Double(0, vecResp)
        Call programDevice_Performance.SetMemoryArgument_Double(1, vecM1)
        Call programDevice_Performance.SetMemoryArgument_Double(2, vecM2)
        Call programDevice_Performance.SetMemoryArgument_Long(3, vecQ)
    
        ' Start once to update cashes-
        Call programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
    
        ' Start real measurements.
        cTime.StartCounter
        Call programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        wsPerformanceTest.Cells(3, 2) = cTime.TimeElapsed
    
        Call programDevice_Performance.GetMemoryArgument_Double(0, vecResp)
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
        wsPerformanceTest.Cells(3, 3) = calcCorrect
    Else
        wsPerformanceTest.Cells(3, 2) = CVErr(2042)
        wsPerformanceTest.Cells(3, 3) = CVErr(2042)
    End If
    
    ' GPU calculations.
    If Not (GetFirstDeviceOfType(progDevices, "GPU") Is Nothing) Then
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "GPU")
        
        Call programDevice_Performance.CreateKernel("DoubleMatrixMult")
        
        globalWorkSize(0) = p
        globalWorkSize(1) = r
        vecQ(0) = q
        
        ReDim vecResp(p * r - 1)
        Call programDevice_Performance.SetMemoryArgument_Double(0, vecResp)
        Call programDevice_Performance.SetMemoryArgument_Double(1, vecM1)
        Call programDevice_Performance.SetMemoryArgument_Double(2, vecM2)
        Call programDevice_Performance.SetMemoryArgument_Long(3, vecQ)
        
        ' Start once to update cashes-
        Call programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        
        ' Start real measurements.
        cTime.StartCounter
        Call programDevice_Performance.ExecuteSync(globalWorkOffset, globalWorkSize, localWorkSize)
        wsPerformanceTest.Cells(4, 2) = cTime.TimeElapsed
        
        Call programDevice_Performance.GetMemoryArgument_Double(0, vecResp)
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
        wsPerformanceTest.Cells(4, 3) = calcCorrect
    Else
        wsPerformanceTest.Cells(4, 2) = CVErr(2042)
        wsPerformanceTest.Cells(4, 3) = CVErr(2042)
    End If
End Sub

Sub GpuCpu_SingleDouble_PerformanceTest()
    Dim wsPerformanceTest As Worksheet
    Dim upper&, singles!(), doubles#(), aSingle!, aDouble#, i&
    Dim sources$
    Dim progDevices As Collection
    Dim programDevice_Performance As ClooWrapperVBA.ProgramDevice
    
    Set wsPerformanceTest = ThisWorkbook.Worksheets("Performance")
    wsPerformanceTest.Range("E3:F4").ClearContents
    
    upper = 10000000
    ReDim singles(upper)
    ReDim doubles(upper)
    
    For i = 0 To upper - 1
        singles(i) = i
        doubles(i) = i
    Next i
    aSingle = 2!
    aDouble = 2#
    
    Open Application.ActiveWorkbook.Path & "\cl\Performance.cl" For Binary As #1
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
        wsPerformanceTest.Cells(4, 6) = CVErr(2042)
    Else
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "GPU")
        wsPerformanceTest.Cells(4, 5) = GPU_PerformanceTest_Single(upper, singles, aSingle, programDevice_Performance)
        wsPerformanceTest.Cells(4, 6) = GPU_PerformanceTest_Double(upper, doubles, aDouble, programDevice_Performance)
    End If
    
    If GetFirstDeviceOfType(progDevices, "CPU") Is Nothing Then
        wsPerformanceTest.Cells(3, 5) = CVErr(2042)
        wsPerformanceTest.Cells(3, 6) = CVErr(2042)
    Else
        Set programDevice_Performance = GetFirstDeviceOfType(progDevices, "CPU")
        wsPerformanceTest.Cells(3, 5) = CPU_PerformanceTest_Single(upper, singles, aSingle, programDevice_Performance)
        wsPerformanceTest.Cells(3, 6) = CPU_PerformanceTest_Double(upper, doubles, aDouble, programDevice_Performance)
    End If
End Sub

' Single precision performance at GPU.
Function GPU_PerformanceTest_Single(upper&, singles!(), aSingle!, programDevice_Performance As ClooWrapperVBA.ProgramDevice)
    Dim buildLogs$
    
    Call programDevice_Performance.CreateKernel("SinglePerformance")
    
    Call programDevice_Performance.SetMemoryArgument_Single(0, singles)
    Call programDevice_Performance.SetValueArgument_Single(1, aSingle)
    
    GPU_PerformanceTest_Single = PerformanceTestExecution(upper, programDevice_Performance)
End Function

' Single precision performance at CPU.
Function CPU_PerformanceTest_Single(upper&, singles!(), aSingle!, programDevice_Performance As ClooWrapperVBA.ProgramDevice)
    Dim buildLogs$
    
    Call programDevice_Performance.CreateKernel("SinglePerformance")
    
    Call programDevice_Performance.SetMemoryArgument_Single(0, singles)
    Call programDevice_Performance.SetValueArgument_Single(1, aSingle)
    
    CPU_PerformanceTest_Single = PerformanceTestExecution(upper, programDevice_Performance)
End Function

' Double precision performance at GPU.
Function GPU_PerformanceTest_Double(upper&, doubles#(), aDouble#, programDevice_Performance As ClooWrapperVBA.ProgramDevice)
    Dim buildLogs$
    
    Call programDevice_Performance.CreateKernel("DoublePerformance")
    
    Call programDevice_Performance.SetMemoryArgument_Double(0, doubles)
    Call programDevice_Performance.SetValueArgument_Double(1, aDouble)
    
    GPU_PerformanceTest_Double = PerformanceTestExecution(upper, programDevice_Performance)
End Function

' Double precision performance at CPU.
Function CPU_PerformanceTest_Double(upper&, doubles#(), aDouble#, programDevice_Performance As ClooWrapperVBA.ProgramDevice)
    Dim buildLogs$
    
    Call programDevice_Performance.CreateKernel("DoublePerformance")
    
    Call programDevice_Performance.SetMemoryArgument_Double(0, doubles)
    Call programDevice_Performance.SetValueArgument_Double(1, aDouble)
    
    CPU_PerformanceTest_Double = PerformanceTestExecution(upper, programDevice_Performance)
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
