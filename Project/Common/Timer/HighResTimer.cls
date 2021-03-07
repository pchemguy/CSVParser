VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighResTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Execution timer via API QueryPerformanceCounter"
'@ModuleDescription "Execution timer via API QueryPerformanceCounter"
'@PredeclaredId
'@Folder "Common.Timer"
Option Explicit

Implements IExecutionTimer

Private Type THighResTimer
    CountsPerSecond As Currency
    StartTime As Currency
    EndTime As Currency
End Type
Private this As THighResTimer


Private Sub Class_Initialize()
    QueryPerformanceFrequency this.CountsPerSecond
End Sub

Public Sub Start()
    QueryPerformanceCounter this.StartTime
End Sub

Public Property Get TimeElapsed() As Double
    QueryPerformanceCounter this.EndTime
    TimeElapsed = (this.EndTime - this.StartTime) / this.CountsPerSecond
    Debug.Print "Time elapsed " & TimeElapsed & " s."
End Property

Private Sub IExecutionTimer_Start()
    Start
End Sub

Private Function IExecutionTimer_TimeElapsed() As Double
    IExecutionTimer_TimeElapsed = TimeElapsed
End Function
