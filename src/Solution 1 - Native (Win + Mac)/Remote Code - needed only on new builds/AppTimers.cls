VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bookTimers As Collection

Private Sub Class_Initialize()
    Set m_bookTimers = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_bookTimers = Nothing
End Sub

Public Sub Add(ByVal bTimers As BookTimers)
    On Error Resume Next
    m_bookTimers.Remove bTimers.ID
    On Error GoTo 0
    m_bookTimers.Add bTimers, bTimers.ID
End Sub

Public Sub CheckRefs()
    Const localRefs As Long = 2 'm_bookTimers.Item + bt
    Dim bt As BookTimers
    '
    For Each bt In m_bookTimers
        If bt.RefsCount = localRefs Then
            m_bookTimers.Remove bt.ID
        End If
    Next bt
End Sub

Public Function Count() As Long
    Count = m_bookTimers.Count
End Function

Public Function PopIfNeeded() As Boolean
    Dim bt As BookTimers
    Dim minBT As BookTimers
    '
    For Each bt In m_bookTimers
        If bt.Count > 0 Then
            If minBT Is Nothing Then
                Set minBT = bt
            ElseIf bt.EarliestTime < minBT.EarliestTime Then
                Set minBT = bt
            End If
        End If
    Next bt
    If Not minBT Is Nothing Then
        PopIfNeeded = minBT.PopIfNeeded()
    End If
End Function
