Attribute VB_Name = "Demo"
Option Explicit

Sub DemoMain()
    Dim t As Date: t = NowMs
    Debug.Print t
    '
    Debug.Print CreateTimer(ThisWorkbook, "OnTime1", t + TimeSerial(0, 0, 2), 200, True), "(TimerID #1)"
    Debug.Print CreateTimer(ThisWorkbook, "OnTime2", t, 0), "(TimerID #2)"
    Debug.Print RemoveTimer(CreateTimer(ThisWorkbook, "OnTime4", t + TimeSerial(0, 0, 2))) & " - was removed"
    Debug.Print CreateTimer(ThisWorkbook, "OnTime3", t + TimeSerial(0, 0, 5)), "(TimerID #3)"
    '
    Debug.Print Now
    Debug.Print "---"
End Sub

Public Sub OnTime1(ByVal timerID As String, Optional ByVal argData As Variant)
    Static i As Long
    Debug.Print "OnTime1", Now, Round(CDbl(Timer), 3), 1
    i = i + 1
    If i = 20 Then
        RemoveTimer timerID
        i = 0
    End If
End Sub
Public Function OnTime2(ByVal timerID As String, Optional ByVal argData As Variant)
    Debug.Print "OnTime2", Now, Round(CDbl(Timer), 3)
End Function
Public Sub OnTime3(ByVal timerID As String)
    Debug.Print "OnTime3", Now, Round(CDbl(Timer), 3)
    Debug.Print CreateTimer(ThisWorkbook, "OnTime4", Now + TimeSerial(0, 0, 2)), "(TimerID)"
End Sub
Public Sub OnTime4(ByVal timerID As String)
    Debug.Print "OnTime4", Now, Round(CDbl(Timer), 3)
End Sub
