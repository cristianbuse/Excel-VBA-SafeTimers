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

Private m_hWnd As LongPtr
Private m_nIDEvent As LongPtr
Private m_id As String
Private m_delayMs As Long
Private m_earliestTime As Date
Private m_originalTime As Date

Public Sub Init(ByRef hWnd As LongPtr _
              , ByRef nIDEvent As LongPtr _
              , ByRef sID As String _
              , ByRef delayMs As Long _
              , ByRef callTime As Date)
    m_hWnd = hWnd
    m_nIDEvent = nIDEvent
    m_id = sID
    m_delayMs = delayMs
    m_earliestTime = callTime
    m_originalTime = m_earliestTime
End Sub
Public Function Self() As TimerContainer
    Set Self = Me
End Function
Public Property Get hWnd() As LongPtr
    hWnd = m_hWnd
End Property
Public Property Get EventID() As LongPtr
    EventID = m_nIDEvent
End Property
Public Property Get ID() As String
    ID = m_id
End Property
Public Property Get Delay() As Long
    Delay = m_delayMs
End Property
Public Property Get EarliestTime() As Date
    EarliestTime = m_earliestTime
End Property

Public Sub UpdateTime()
    Const msPerDay As Long = 24& * 60& * 60& * 1000&
    Dim daysDelay As Double
    Dim skipCount As Long
    '
    daysDelay = m_delayMs / msPerDay
    skipCount = Int((NowMSec - m_originalTime) / daysDelay)
    m_earliestTime = m_originalTime + (skipCount + 1) * daysDelay
End Sub
