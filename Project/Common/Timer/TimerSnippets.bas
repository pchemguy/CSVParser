Attribute VB_Name = "TimerSnippets"
'@Folder("Common.Timer")
Option Explicit


Private Sub BasicTimerTest()
    Dim TimeMan As IExecutionTimer
    Set TimeMan = New BasicTimer
    
    TimeMan.Start
    Sleep 1
    TimeMan.TimeElapsed
End Sub


Private Sub HighResTimerTest()
    Dim TimeMan As IExecutionTimer
    Set TimeMan = New HighResTimer
    
    TimeMan.Start
    Sleep 1
    TimeMan.TimeElapsed
End Sub

