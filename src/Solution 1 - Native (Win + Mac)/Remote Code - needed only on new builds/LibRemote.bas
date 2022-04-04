Attribute VB_Name = "LibRemote"
Option Explicit

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Sub USleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
    #Else
        Private Declare Sub USleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
    #End If
#Else 'Windows
    #If VBA7 Then
        Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #Else
        Public Declare  Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    #End If
#End If

Private m_appTimers As AppTimers
Private m_cb As CommandBar 'Self-close safety

#If Mac Then
Public Sub Sleep(ByVal dwMilliseconds As Long)
    USleep dwMilliseconds * 1000&
End Sub
#End If

Public Function GetBookTimers(ByVal bookID As String, ByVal app As Object) As BookTimers
    If m_appTimers Is Nothing Then
        Set m_appTimers = New AppTimers
        #If Mac Then
            TimerResolution
        #End If
        Set m_cb = app.CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
        m_cb.Controls.Add msoControlButton
        Application.OnTime Now(), "MainLoop"
    End If
    With New BookTimers
        .Init bookID
        m_appTimers.Add .Self
        Set GetBookTimers = .Self
    End With
End Function

Public Sub MainLoop()
    Do While IsConnected() Or m_appTimers.Count > 0
        m_appTimers.CheckRefs
        If m_appTimers.Count > 0 Then
            If Not m_appTimers.PopIfNeeded Then SleepIfNoEntry
        Else
            SleepIfNoEntry
        End If
        DoEvents
    Loop
    Application.Quit
End Sub

Private Function IsConnected() As Boolean
    Static lastCheck As Date
    Dim mNow As Date: mNow = NowMSec()
    '
    If lastCheck < mNow Then
        Dim bCount As Long
        On Error Resume Next
        bCount = m_cb.Controls.Count
        On Error GoTo 0
        If bCount = 0 Then Exit Function
        lastCheck = mNow + TimeSerial(0, 0, 1)
    End If
    IsConnected = True
End Function

Private Sub SleepIfNoEntry()
    Dim isEntryRequested As Boolean
    Static isStarted As Boolean
    Static endTime As Date
    '
    On Error Resume Next
    isEntryRequested = (GetSetting("RemoteTimers", "Flags", "EntryNeeded") = "1")
    On Error GoTo 0
    If isStarted Then
        If isEntryRequested Then
            If endTime < NowMSec() Then isStarted = False
        Else
            isStarted = False
        End If
    Else
        If isEntryRequested Then
            endTime = NowMSec() + TimeSerial(0, 0, 1) / 10 '100ms timeout
            isStarted = True
        Else
            Sleep 1
        End If
    End If
End Sub

Public Function NowMSec() As Date
#If Mac Then
    Const evalResolution As Double = 0.01
    Const evalFunc As String = "=Now()"
    Static useEval As Boolean
    Static isSet As Boolean
    '
    If Not isSet Then
        useEval = TimerResolution > evalResolution
        isSet = True
    End If
    If useEval Then
        NowMSec = Evaluate(evalFunc)
        Exit Function
    End If
#End If
    Const secondsPerDay As Long = 24& * 60& * 60&
    NowMSec = Date + Round(Timer, 3) / secondsPerDay
End Function

#If Mac Then
Private Function TimerResolution() As Double
    Const secondsPerDay As Long = 24& * 60& * 60&
    Static r As Double
    If r = 0 Then
        Dim t As Double: t = Timer
        Do
            r = Round(Timer - t, 3)
            If r < 0# Then r = r + secondsPerDay 'Passed midnight
        Loop Until r > 0#
    End If
    TimerResolution = r
End Function
#End If
