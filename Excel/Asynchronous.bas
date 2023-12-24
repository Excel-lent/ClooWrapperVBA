Attribute VB_Name = "Asynchronous"
Option Explicit

Const MAX_TASKS = 20
Private Const ARRAY_SIZE = 2000

' Thread priority as integer (0 - "Lowest", 1 - "BelowNormal", 2 - "Normal", 3 - "AboveNormal", 4 - "Highest").
Const THREAD_PRIORITY = 0

Dim progDevices As Collection
Dim currentTaskId&()

Dim vecResp!()
Dim wsAsynchronous As Worksheet
Dim globalWorkSize&(1), localWorkSize&(), globalWorkOffset&()
Dim logLine%
Dim currentProgress%, prevProgress%, startedTasks%, finishedTasks%

Dim result As Boolean
Dim allTasks_Completed As Boolean

Sub MainLoop()
    Dim i%
    
    ' Infinite loop until all tasks are not completed.
    While Not allTasks_Completed
        For i = 1 To progDevices.Count
            If progDevices.Item(i).ProgramDevice.ExecutionCompleted Then
                result = progDevices.Item(i).ProgramDevice.GetMemoryArgument_Single(0, vecResp)    ' Extract the results and do something with received data here.
                
                wsAsynchronous.Cells(logLine, 1) = "Task " & currentTaskId(i) & ", " & progDevices.Item(i).ProgramDevice.deviceType & _
                    progDevices.Item(i).DeviceId & ": completed"
                logLine = logLine + 1
                
                finishedTasks = finishedTasks + 1
                
                ' Start new task
                If startedTasks < MAX_TASKS Then
                    ReDim vecResp(UBound(vecResp))  ' Erase output vector.
                    result = progDevices.Item(i).ProgramDevice.SetMemoryArgument_Single(0, vecResp)
                    
                    ' If you want to use callbacks, than use function below
                    ' "CPU_Task_Completed" is a function that will obtain the callback.
                    ' Call progDevices.Item(i).ProgramDevice.ExecuteAsync(globalWorkOffset, globalWorkSize, localWorkSize, THREAD_PRIORITY, AddressOf Asynchronous.CPU_Task_Completed)
                    
                    result = progDevices.Item(i).ProgramDevice.ExecuteBackground(globalWorkOffset, globalWorkSize, localWorkSize, THREAD_PRIORITY)
                    startedTasks = startedTasks + 1
                    currentTaskId(i) = startedTasks
                Else
                    ' If the maximal number of tasks is reached, then set "ExecutionCompleted" to false to avoid additional outputs.
                    progDevices.Item(i).ProgramDevice.ExecutionCompleted = False
                End If
                
                If startedTasks = finishedTasks Then
                    allTasks_Completed = True
                End If
            End If
        Next i
        
        ' Progress-bar.
        wsAsynchronous.Cells(2, prevProgress).Interior.Color = RGB(255, 255, 255)
        currentProgress = currentProgress + 1
        If currentProgress = 50 Then
            currentProgress = 1
        End If
        prevProgress = currentProgress
        wsAsynchronous.Cells(2, currentProgress).Interior.Color = RGB(0, 255, 0)
        
        DoEvents
        Sleep (100)
    Wend
    
    wsAsynchronous.Cells(2, currentProgress).Interior.Color = RGB(255, 255, 255)
    
    For i = 1 To progDevices.Count
        result = progDevices.Item(i).ProgramDevice.ReleaseMemObject(3)
        result = progDevices.Item(i).ProgramDevice.ReleaseMemObject(2)
        result = progDevices.Item(i).ProgramDevice.ReleaseMemObject(1)
        result = progDevices.Item(i).ProgramDevice.ReleaseMemObject(0)
        result = progDevices.Item(i).ProgramDevice.ReleaseKernel
        result = progDevices.Item(i).ProgramDevice.ReleaseProgram
    Next i
End Sub

Sub RunAsynchronous()
    Dim vecM1!(), vecM2!()
    Dim vecQ&(1)
    Dim i&, j&, p&, q&, r&, nRows&
    Dim buildLogs$, sources$
    
    Set wsAsynchronous = ThisWorkbook.Worksheets("Asynchronous")
    
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
    
    logLine = 7
    nRows = wsAsynchronous.Cells(Rows.Count, 1).End(xlUp).Row
    wsAsynchronous.Range(wsAsynchronous.Cells(logLine, 1), wsAsynchronous.Cells(nRows, 1)).ClearContents
    
    p = ARRAY_SIZE: q = ARRAY_SIZE: r = ARRAY_SIZE
    
    ' Dimensions of matrices:
    ReDim vecM1(p * q - 1)
    ReDim vecM2(q * r - 1)
    ReDim vecResp(p * r - 1)
    
    Randomize
    For i = 0 To p - 1
        For j = 0 To q - 1
            vecM1(i + p * j) = (Rnd() - 0.5) * 10#
        Next j
    Next i
    
    For i = 0 To q - 1
        For j = 0 To r - 1
            vecM2(i + q * j) = (Rnd() - 0.5) * 10#
        Next j
    Next i
    
    globalWorkSize(0) = p
    globalWorkSize(1) = r
    vecQ(0) = q
    
    ReDim currentTaskId(progDevices.Count)
    For i = 1 To progDevices.Count
        result = progDevices.Item(i).ProgramDevice.CreateKernel("FloatMatrixMult")
        result = progDevices.Item(i).ProgramDevice.SetMemoryArgument_Single(0, vecResp)
        result = progDevices.Item(i).ProgramDevice.SetMemoryArgument_Single(1, vecM1)
        result = progDevices.Item(i).ProgramDevice.SetMemoryArgument_Single(2, vecM2)
        result = progDevices.Item(i).ProgramDevice.SetMemoryArgument_Long(3, vecQ)
    Next i
    
    startedTasks = 0
    ' Start execution on all found devices almost simultaneously.
    For i = 1 To progDevices.Count
        result = progDevices.Item(i).ProgramDevice.ExecuteBackground(globalWorkOffset, globalWorkSize, localWorkSize, THREAD_PRIORITY)
        
        ' If you want to use callbacks, than use function below
        ' "CPU_Task_Completed" is a function that will obtain the callback.
        ' Call progDevices.Item(i).ProgramDevice.ExecuteAsync(globalWorkOffset, globalWorkSize, localWorkSize, THREAD_PRIORITY, AddressOf Asynchronous.CPU_Task_Completed)
        
        startedTasks = startedTasks + 1
        currentTaskId(i) = startedTasks
    Next i
    
    prevProgress = 1
    currentProgress = 2
    allTasks_Completed = False
    finishedTasks = 0
    
    Call MainLoop
End Sub

'Sub CPU_Task_Completed()
'
'End Sub
