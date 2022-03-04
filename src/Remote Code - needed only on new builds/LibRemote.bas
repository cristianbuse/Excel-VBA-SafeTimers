Attribute VB_Name = "LibRemote"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

#If VBA7 Then
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private m_appTimers As AppTimers
Private m_readyHWnd As LongPtr

Public Function GetReadyHWnd() As LongPtr
    GetReadyHWnd = m_readyHWnd
End Function

Public Function GetBookTimers(ByVal readyHWnd As LongPtr _
                            , ByVal bookID As String _
                            , ByVal tProc As LongPtr) As BookTimers
    If m_readyHWnd = 0 Then
        m_readyHWnd = readyHWnd
        Set m_appTimers = New AppTimers
        Application.OnTime Now(), "MainLoop"
    End If
    With New BookTimers
        .Init bookID, tProc
        m_appTimers.Add .Self
        Set GetBookTimers = .Self
    End With
End Function

Public Sub MainLoop()
    Do While IsWindow(m_readyHWnd)
        m_appTimers.CheckRefs
        If m_appTimers.Count > 0 And m_appTimers.CanPost Then
            If Not m_appTimers.PopIfNeeded Then Sleep 1
        Else
            Sleep 1
        End If
        DoEvents
    Loop
    Set m_appTimers = Nothing
    Application.Quit
End Sub

Public Function IsIDEReady() As Boolean
    Const readyLabelCurANSI As String = "1758492059378.1308" '<Ready>
    Static readyLabel As Currency
    Const WM_GETTEXT As Long = &HD
    Dim buff As Currency
    '
    If readyLabel = 0 Then readyLabel = CCur(readyLabelCurANSI)
    If SendMessage(m_readyHWnd, WM_GETTEXT, 8, VarPtr(buff)) = 0 Then Exit Function
    IsIDEReady = (buff = readyLabel)
End Function

Public Function NowMSec() As Date
    Const secondsPerDay As Long = 24& * 60& * 60&
    NowMSec = Date + Round(Timer, 3) / secondsPerDay
End Function
