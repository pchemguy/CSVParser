VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Common.Timer")
'@PredeclaredId
Option Explicit

Implements IExecutionTimer

Private Type TBasicTimer
    StartTime As Currency
    EndTime As Currency
End Type
Private this As TBasicTimer


Public Sub Start()
    this.StartTime = Timer
End Sub

Public Property Get TimeElapsed() As Double
    this.EndTime = Timer
    TimeElapsed = CDbl(this.EndTime - this.StartTime)
    Debug.Print "Time elapsed " & TimeElapsed & " s."
End Property

Private Sub IExecutionTimer_Start()
    Start
End Sub

Private Function IExecutionTimer_TimeElapsed() As Double
    IExecutionTimer_TimeElapsed = TimeElapsed
End Function

