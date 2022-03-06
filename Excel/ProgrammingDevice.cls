VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgrammingDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ProgramDevice As ClooWrapperVBA.ProgramDevice
Public deviceType As String
Public DeviceId As Integer

Private Sub Class_Initialize()
    Set ProgramDevice = New ClooWrapperVBA.ProgramDevice
End Sub
