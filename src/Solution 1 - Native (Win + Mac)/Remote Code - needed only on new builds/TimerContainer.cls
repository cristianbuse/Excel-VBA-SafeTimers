VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_callback As Object
Private m_id As String
Private m_delayMs As Long
Private m_earliestTime As Date
Private m_originalTime As Date
Private m_daysDelay As Double

Public Sub Init(ByRef tCallback As Object _
              , ByRef timerID As String _
              , ByRef earliestCallTime As Date _
              , ByRef delayMs As Long)
    Const msPerDay As Long = 24& * 60& * 60& * 1000&
    Dim mNow As Date: mNow = NowMSec()
    '
    Set m_callback = tCallback
    m_id = timerID
    If delayMs > 0 Then m_delayMs = delayMs
    m_daysDelay = m_delayMs / msPerDay
    '
    If earliestCallTime > mNow Then
        m_earliestTime = earliestCallTime
        m_originalTime = earliestCallTime
    Else
        m_earliestTime = mNow + m_daysDelay
        m_originalTime = mNow
    End If
End Sub
Public Property Get TimerCallback() As Object
    Set TimerCallback = m_callback
End Property
Public Property Get ID() As String
    ID = m_id
End Property
Public Property Get EarliestTime() As Date
    EarliestTime = m_earliestTime
End Property
Public Property Get Delay() As Long
    Delay = m_delayMs
End Property
Public Function Self() As TimerContainer
    Set Self = Me
End Function

Public Sub UpdateTime()
    Dim skipCount As Long
    '
    skipCount = Int((NowMSec() - m_originalTime) / m_daysDelay)
    m_earliestTime = m_originalTime + (skipCount + 1) * m_daysDelay
End Sub
