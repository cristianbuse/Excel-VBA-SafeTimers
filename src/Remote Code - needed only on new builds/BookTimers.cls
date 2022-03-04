VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

#If Win64 Then
    Private Const PTR_SIZE As Long = 8
#Else
    Private Const PTR_SIZE As Long = 4
#End If

Private m_canPost As Boolean
Private m_id As String
Private m_refCount As Variant
Private m_timers As Collection
Private m_tProc As LongPtr

Public Sub Init(ByVal bookID As String, ByVal tProc As LongPtr)
    m_id = bookID
    m_tProc = tProc
End Sub

Private Sub Class_Initialize()
    Set m_timers = New Collection
    SetRefCount
    m_canPost = True
End Sub

Private Sub Class_Terminate()
    Set m_timers = Nothing
    On Error Resume Next
    DeleteSetting "SafeTimers", m_id
    On Error GoTo 0
End Sub

Private Sub SetRefCount()
    Const VT_BYREF As Long = &H4000
    Dim iUnk As IUnknown: Set iUnk = Me
    m_refCount = ObjPtr(iUnk) + PTR_SIZE
    CopyMemory m_refCount, vbLong + VT_BYREF, 2
End Sub

Public Property Get RefsCount() As Long
    RefsCount = GetLongByRef(m_refCount)
End Property
Private Function GetLongByRef(ByRef v As Variant) As Long
    GetLongByRef = v
End Function

Public Function Count() As Long
    Count = m_timers.Count
End Function

Public Property Get ID() As String
    ID = m_id
End Property

Public Function Self() As BookTimers
    Set Self = Me
End Function

Public Property Get CanPost() As Boolean
    If Not m_canPost Then
        m_canPost = (GetSetting("SafeTimers", m_id, "CanPost") = "True")
        If m_canPost Then
            Dim lostID As String
            lostID = GetSetting("SafeTimers", m_id, "LostID")
            If LenB(lostID) > 0 Then
                DeleteTimer lostID
                DeleteSetting "SafeTimers", m_id, "LostID"
            End If
        End If
    End If
    CanPost = m_canPost
End Property

Public Property Get EarliestTime() As Date
    EarliestTime = m_timers(1).EarliestTime
End Property

Public Function AddTimer(ByVal hWnd As LongPtr _
                       , ByVal nIDEvent As LongPtr _
                       , ByVal sID As String _
                       , ByVal delayMs As Long) As Boolean
    DeleteTimer sID
    With New TimerContainer
        Const msPerDay As Long = 24& * 60& * 60& * 1000&
        Dim nextRun As Date: nextRun = NowMSec() + delayMs / msPerDay
        '
        .Init hWnd, nIDEvent, sID, delayMs, nextRun
        InsertTimer .Self
    End With
    AddTimer = True
End Function

Public Function DeleteTimer(ByVal sID As String) As Boolean
    On Error Resume Next
    m_timers.Remove sID
    DeleteTimer = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function DeleteAllTimers() As Boolean
    Set m_timers = New Collection
    DeleteAllTimers = True
End Function

Private Sub InsertTimer(ByRef container As TimerContainer)
    Dim tc As TimerContainer
    Dim i As Long: i = 1
    '
    For Each tc In m_timers
        If tc.EarliestTime > container.EarliestTime Then Exit For
        i = i + 1
    Next tc
    If m_timers.Count = 0 Or i > m_timers.Count Then
        m_timers.Add Item:=container, Key:=container.ID
    Else
        m_timers.Add Item:=container, Key:=container.ID, Before:=i
    End If
End Sub

Public Function PopIfNeeded() As Boolean
    Const WM_TIMER As Long = &H113
    Dim tc As TimerContainer: Set tc = m_timers(1)
    '
    If tc.EarliestTime > NowMSec() Then Exit Function
    If Not IsIDEReady() Then Exit Function
    '
    m_timers.Remove 1
    If PostMessage(tc.hWnd, WM_TIMER, tc.EventID, m_tProc) = 0& Then
        If IsWindow(tc.hWnd) = 0& Then Exit Function
    End If
    m_canPost = False
    PopIfNeeded = True
    '
    tc.UpdateTime
    InsertTimer tc
End Function

Public Sub KillBookTimer()
    Const WM_TIMER As Long = &H113
    Dim rHWnd As LongPtr: rHWnd = GetReadyHWnd()
    Dim tID As LongPtr: tID = VBA.Int(m_id)
    '
    Do While IsWindow(rHWnd)
        If IsIDEReady() Then
            If PostMessage(rHWnd, WM_TIMER, tID, m_tProc) <> 0& Then Exit Do
        End If
        Sleep 1
    Loop
End Sub
