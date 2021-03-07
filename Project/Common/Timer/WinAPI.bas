Attribute VB_Name = "WinAPI"
'@Folder("Common.Timer")
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Public Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If


