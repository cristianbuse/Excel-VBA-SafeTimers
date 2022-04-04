Attribute VB_Name = "Demo"
Option Explicit

Private Const FIXED_ID As LongPtr = 5
Private m_dynamicID As LongPtr

Public Sub DemoMain()
    SetTimer ThisWorkbook.Windows(1).hWnd, FIXED_ID, 200, AddressOf TimerProc
    m_dynamicID = SetTimer(0, 0, 100, AddressOf TimerProc)
End Sub

Private Sub TimerProc(ByVal hWnd As LongPtr _
                    , ByVal wMsg As Long _
                    , ByVal nIDEvent As LongPtr _
                    , ByVal wTime As Long)
    Select Case nIDEvent
    Case FIXED_ID
        FixedIDTimer hWnd, nIDEvent
    Case m_dynamicID
        DynamicIDTimer hWnd, nIDEvent
    Case Else
        RemoveAllTimers
    End Select
End Sub

Private Sub FixedIDTimer(ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr)
    Static c As Long
    c = c + 1
    If c = 20 Then KillTimer hWnd, nIDEvent
    Debug.Print Round(CDbl(Timer), 3), hWnd, nIDEvent, "Fixed"
End Sub

Private Sub DynamicIDTimer(ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr)
    Static c As Long
    c = c + 1
    If c = 20 Then KillTimer hWnd, nIDEvent
    Debug.Print Round(CDbl(Timer), 3), hWnd, nIDEvent, "Dynamic"
End Sub
