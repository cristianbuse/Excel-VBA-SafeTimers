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

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If
#If Win64 Then
    Private Const PTR_SIZE As Long = 8
#Else
    Private Const PTR_SIZE As Long = 4
#End If

Private m_id As String
Private m_refCount As Variant
Private m_timers As Collection

Public Sub Init(ByVal bookID As String)
    m_id = bookID
End Sub

Private Sub Class_Initialize()
    Set m_timers = New Collection
    SetRefCount
End Sub

Private Sub Class_Terminate()
    Set m_timers = Nothing
End Sub

Private Sub SetRefCount()
    Const VT_BYREF As Long = &H4000
    Dim iUnk As IUnknown: Set iUnk = Me
    m_refCount = ObjPtr(iUnk) + PTR_SIZE
    CopyMemory m_refCount, vbLong + VT_BYREF, 2
End Sub

Public Property Get RefsCount() As Long
    RefsCount = GetLongByRef(m_refCount) - 1 '-1 for Me
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

Public Property Get EarliestTime() As Date
    EarliestTime = m_timers(1).EarliestTime
End Property

Public Function AddTimer(ByVal tCallback As Object _
                       , ByVal timerID As String _
                       , ByVal earliestCallTime As Date _
                       , ByVal delayMs As Long) As String
    If tCallback Is Nothing Then Exit Function
    If LenB(timerID) = 0 Then Exit Function
    '
    With New TimerContainer
        .Init tCallback, timerID, earliestCallTime, delayMs
        InsertTimer .Self
    End With
    AddTimer = timerID
End Function

Public Function DeleteTimer(ByVal timerID As String) As Boolean
    On Error Resume Next
    m_timers.Remove timerID
    On Error GoTo 0
    DeleteTimer = True
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
    Const errMissingArgument As Long = 449
    Const errNotAvailable As Long = 1004
    Const errObjDisconnected As Long = -2147417848
    Const errRunFailed As Long = 50290
    Const errTypeMismatch As Long = 13
    Const errWrongArguments As Long = 450
    '
    Dim tc As TimerContainer: Set tc = m_timers(1)
    Dim remoteErrCode As Long
    '
    If tc.EarliestTime > NowMSec() Then Exit Function
    PopIfNeeded = True
    '
    On Error Resume Next
    remoteErrCode = tc.TimerCallback.TimerProc() 'Possible re-entry point!
    If Err.Number = errObjDisconnected Then Exit Function
    Err.Clear
    m_timers.Remove tc.ID
    If Err.Number <> 0 Then 'Timer was removed via re-entry
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    '
    If remoteErrCode = errMissingArgument Then Exit Function
    If remoteErrCode = errTypeMismatch Then Exit Function
    If remoteErrCode = errWrongArguments Then Exit Function
    If remoteErrCode = errNotAvailable Then
        Const notAvailableLimit As Long = 10
        tc.NotAvailableCount = tc.NotAvailableCount + 1
        If tc.NotAvailableCount > notAvailableLimit Then Exit Function
    Else
        tc.NotAvailableCount = 0
    End If
    '
    If tc.Delay > 0 Then
        tc.UpdateTime
        InsertTimer tc
    ElseIf remoteErrCode = errRunFailed Or remoteErrCode = errNotAvailable Then
        InsertTimer tc '0 delay timers are guaranteed to be called once!
    End If
End Function
