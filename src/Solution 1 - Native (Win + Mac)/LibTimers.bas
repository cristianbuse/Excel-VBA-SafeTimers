Attribute VB_Name = "LibTimers"
'''=============================================================================
''' Excel-VBA-SafeTimers
''' ----------------------------------------------------
''' https://github.com/cristianbuse/Excel-VBA-SafeTimers
''' ----------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2022 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit
Option Private Module

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Sub USleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
    #Else
        Private Declare Sub USleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #Else
        Private Declare  Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    #End If
#End If

Private m_remoteTimers As Object
Private Const BOOK_NAME As String = "RemoteTimers_V1.xlam"
Private Const MODULE_NAME As String = "LibTimers"

#If Mac Then
Private Sub Sleep(ByVal dwMilliseconds As Long)
    USleep dwMilliseconds * 1000&
End Sub
#End If

'*******************************************************************************
'The main loop of the remote process will stop 'sleeping' when requested
'This speeds up the entry considerably
'*******************************************************************************
Private Function RequestEntry()
    SaveSetting "RemoteTimers", "Flags", "EntryNeeded", "1"
End Function
Private Function CancelEntry()
    SaveSetting "RemoteTimers", "Flags", "EntryNeeded", "0"
End Function

'*******************************************************************************
'Creates a timer callback and passes it to the remote application
'Returns:
'   - a unique timer ID
'Parameters:
'   - callbackBook: clearly defines the book containing the macro
'   - callbackName: the name of the macro to be called by the timer
'   - [earliestCallTime]: defer to a future date & time if needed
'   - [delayMs]: if value is greater than 0 then macro will be called repeatedly
'                in intervals approximately equal to the provided delay and will
'                only stop if the timer is removed or state is lost
'                if value is 0 (or less) then the callback will only be called
'                once but the call is guaranteed to happen
'   - [argData]: any data to be used as the second argument in the callback
'Notes:
'   - the callback must expect the 'timerID' as the first parameter and
'     it must have a second parameter only if the 'argData' argument was set.
'     The callback definition could look like this:
'     Function CB(ByVal timerID As String, Optional ByVal argData As Variant)
'*******************************************************************************
Public Function CreateTimer(ByVal callbackBook As Workbook _
                          , ByVal callbackName As String _
                          , Optional ByVal earliestCallTime As Date _
                          , Optional ByVal delayMs As Long = 0 _
                          , Optional ByVal argData As Variant) As String
    Const fullMethodName As String = MODULE_NAME & ".CreateTimer"
    '
    If callbackBook Is Nothing Then
        Err.Raise 91, fullMethodName, "Book not set"
    ElseIf LenB(callbackName) = 0 Then
        Err.Raise 5, fullMethodName, "Invalid macro name"
    End If
    With New TimerCallback
        .Init "'" & callbackBook.Name & "'!" & callbackName, argData
        On Error Resume Next
        Do
            RequestEntry
            CreateTimer = m_remoteTimers.AddTimer(.Self, .ID, earliestCallTime, delayMs)
            If LenB(CreateTimer) = 0 Then
                If Not InitTimers(reCreateBook:=False) Then Exit Do
                Sleep 1
            End If
        Loop Until LenB(CreateTimer) > 0
        CancelEntry
        On Error GoTo 0
    End With
End Function

'*******************************************************************************
'Removes a timer - returns False only if timers are disconnected
'*******************************************************************************
Public Function RemoveTimer(ByVal timerID As String) As Boolean
    On Error Resume Next
    Do
        RequestEntry
        RemoveTimer = m_remoteTimers.DeleteTimer(timerID)
        If Not RemoveTimer Then
            If Not InitTimers(reCreateBook:=False) Then Exit Do
            Sleep 1
        End If
    Loop Until RemoveTimer
    CancelEntry
    On Error GoTo 0
End Function

'*******************************************************************************
'Removes all existing timers - returns False only if timers are disconnected
'*******************************************************************************
Public Function RemoveAllTimers() As Boolean
    On Error Resume Next
    Do
        RequestEntry
        RemoveAllTimers = m_remoteTimers.DeleteAllTimers()
        If Not RemoveAllTimers Then
            If Not InitTimers(reCreateBook:=False) Then Exit Do
            Sleep 1
        End If
    Loop Until RemoveAllTimers
    CancelEntry
    On Error GoTo 0
End Function

'*******************************************************************************
'Returns 'True' only if the object is set and still connected
'*******************************************************************************
Private Function IsConnected(ByVal obj As Object) As Boolean
    If Not obj Is Nothing Then
        IsConnected = TypeName(obj) <> "Object"
    End If
End Function

'*******************************************************************************
'Initializes the remote application instance and its resources e.g. code book
'Works regardless if VB Object Model access is on or off
'*******************************************************************************
Public Function InitTimers(Optional ByVal reCreateBook As Boolean = False) As Boolean
    If IsConnected(m_remoteTimers) Then
        InitTimers = True
        Exit Function
    End If
    '
    Dim app As Application
    Dim bookExists As Boolean: bookExists = IsFile(GetBookPath())
    '
    If reCreateBook And bookExists Then
        On Error Resume Next
        Kill GetBookPath()
        bookExists = (Err.Number <> 0)
        On Error GoTo 0
        If bookExists Then Exit Function
    End If
    If bookExists Then
        Set app = GetRemoteApp()
    Else
        Set app = CreateBookInRemoteApp()
    End If
    '
    On Error Resume Next 'In case another book is running timers in the same app
    Do
        RequestEntry
        Set m_remoteTimers = app.Run("GetBookTimers", CStr(ObjPtr(ThisWorkbook)), Application)
        If m_remoteTimers Is Nothing Then
            If Not IsConnected(m_remoteTimers) Then Exit Function
            Sleep 1
        End If
    Loop Until Not m_remoteTimers Is Nothing
    CancelEntry
    On Error GoTo 0
    InitTimers = True
End Function
Private Function IsFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    IsFile = ((GetAttr(filePath) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function
Private Function GetBookPath() As String
    Dim folderPath As String
#If Mac Then
    folderPath = Application.DefaultFilePath
    If LenB(folderPath) = 0 Then folderPath = ThisWorkbook.Path
#Else
    folderPath = Environ$("temp")
#End If
    GetBookPath = folderPath & Application.PathSeparator & BOOK_NAME
End Function

'*******************************************************************************
'Opens the needed addin book in a new remote app and returns the app
'*******************************************************************************
Private Function GetRemoteApp() As Application
    Const maxSuffix As Long = 50
    Const ERR_SUBSCRIPT_OUT_OF_RANGE As Long = 9
    Dim app As Object  'Faster than early binding i.e. As Application
    Dim book As Object 'Faster than early binding i.e. As Workbook
    Dim i As Long
    Dim wndName As String
    '
    On Error Resume Next
    For i = 1 To maxSuffix
        Set book = Nothing
        Set book = GetObject("Book" & i)
        If Not book Is Nothing Then
            RequestEntry
            wndName = vbNullString
            Do While LenB(wndName) = 0
                wndName = book.Windows(1).Caption
                If LenB(wndName) = 0 Then Sleep 1
            Loop
            If wndName = CStr(ObjPtr(Application)) Then
                Set app = book.Application
                Exit For
            Else
                Stop
            End If
            CancelEntry
        End If
    Next i
    On Error GoTo 0
    If app Is Nothing Then
        Set app = NewApp()
        app.Workbooks.Open GetBookPath(), False, False
        app.Workbooks.Add.Windows(1).Caption = CStr(ObjPtr(Application))
    End If
    Set GetRemoteApp = app
End Function

'*******************************************************************************
'Creates a new app instance and sets certain properties
'*******************************************************************************
Private Function NewApp() As Object 'Faster than early binding i.e. As Application
    Set NewApp = New Application
    With NewApp
        .Visible = False
        .PrintCommunication = False
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
End Function

'*******************************************************************************
'Creates a book, adds the code, saves it and returns the Application instance
'*******************************************************************************
Private Function CreateBookInRemoteApp() As Application
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2
    Const publicNotCreatable As Long = 2
    '
    Dim app As Application
    Dim book As Workbook
    Dim isVBOMOn As Boolean: isVBOMOn = IsVBOMEnabled()
    '
    If Not isVBOMOn Then
        If Not EnableOfficeVBOM(True) Then Exit Function
    End If
    '
    On Error GoTo SafeExit
    Set app = NewApp()
    Set book = app.Workbooks.Add
    '
    With book.VBProject.VBComponents.Add(vbext_ct_ClassModule).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString TimerContainerCode()
        .Parent.Name = "TimerContainer"
    End With
    With book.VBProject.VBComponents.Add(vbext_ct_ClassModule).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString BookTimersCode()
        .Parent.Name = "BookTimers"
        .Parent.Properties("Instancing") = publicNotCreatable
    End With
    With book.VBProject.VBComponents.Add(vbext_ct_ClassModule).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString AppTimersCode()
        .Parent.Name = "AppTimers"
    End With
    With book.VBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString LibRemoteCode()
        .Parent.Name = "LibRemote"
    End With
    book.SaveAs GetBookPath(), XlFileFormat.xlOpenXMLAddIn
    '
    If Not isVBOMOn Or book.FullName <> GetBookPath() Then
        book.Close False
        app.Quit
        Set app = Nothing
        If Not isVBOMOn Then EnableOfficeVBOM False
        Set app = NewApp()
        app.Workbooks.Open GetBookPath(), False, False
    End If
    app.Workbooks.Add.Windows(1).Caption = CStr(ObjPtr(Application))
    Set CreateBookInRemoteApp = app
SafeExit:
End Function

'*******************************************************************************
'Checks if VBProject is accessible programmatically. Setting is app level
'*******************************************************************************
Private Function IsVBOMEnabled() As Boolean
    On Error Resume Next
    IsVBOMEnabled = Not Application.VBE.ActiveVBProject Is Nothing
    On Error GoTo 0
End Function

'*******************************************************************************
'Apps like AutoCAD have the Object Model access on by default. This method is
'   desgined for Microsoft Office VBA-capable applications
'*******************************************************************************
Private Function EnableOfficeVBOM(ByVal newValue As Boolean) As Boolean
    Dim i As Long: i = IIf(newValue, 1, 0)
#If Mac Then
    Shell "defaults write com.microsoft.Excel AccessVBOM -int " & i
#Else
    Dim rKey As String
    rKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version _
         & "\" & Replace(Application.Name, "Microsoft ", vbNullString) _
         & "\Security\AccessVBOM"
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite rKey, i, "REG_DWORD"
    EnableOfficeVBOM = (Err.Number = 0)
    On Error GoTo 0
#End If
End Function

'*******************************************************************************
'An enhanced 'Now' - returns the date and time including milliseconds
'On Mac resolution depends on Excel version and can be 10ms or less (e.g. 4ms)
'On Win resolution is around 4ms
'*******************************************************************************
Public Function NowMs() As Date
    Const secondsPerDay As Long = 24& * 60& * 60&
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
        NowMs = Evaluate(evalFunc)
        Exit Function
    End If
#End If
    NowMs = Date + Round(Timer, 3) / secondsPerDay
End Function
#If Mac Then
Private Function TimerResolution() As Double
    Const secondsPerDay As Long = 24& * 60& * 60&
    Static r As Double
    If r = 0 Then
        Dim t As Double: t = Timer
        Do
            r = Round(Timer - t, 3)
            If r < 0# Then r = r + secondsPerDay
        Loop Until r > 0#
    End If
    TimerResolution = r
End Function
#End If

'*******************************************************************************
'Code running 'on the other side'
'*******************************************************************************
Private Function LibRemoteCode() As String
Dim s As String
Const n As String = vbNewLine
s = s & "Option Explicit" & n
s = s & "" & n
s = s & "#If Mac Then" & n
s = s & "    #If VBA7 Then" & n
s = s & "        Private Declare PtrSafe Sub USleep Lib ""/usr/lib/libc.dylib"" Alias ""usleep"" (ByVal dwMicroseconds As Long)" & n
s = s & "    #Else" & n
s = s & "        Private Declare Sub USleep Lib ""/usr/lib/libc.dylib"" Alias ""usleep"" (ByVal dwMicroseconds As Long)" & n
s = s & "    #End If" & n
s = s & "#Else 'Windows" & n
s = s & "    #If VBA7 Then" & n
s = s & "        Public Declare PtrSafe Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)" & n
s = s & "    #Else" & n
s = s & "        Public Declare  Sub Sleep Lib ""kernel32"" Alias ""Sleep"" (ByVal dwMilliseconds As Long)" & n
s = s & "    #End If" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "Private m_appTimers As AppTimers" & n
s = s & "Private m_cb As CommandBar 'Self-close safety" & n
s = s & "" & n
s = s & "#If Mac Then" & n
s = s & "Public Sub Sleep(ByVal dwMilliseconds As Long)" & n
s = s & "    USleep dwMilliseconds * 1000&" & n
s = s & "End Sub" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "Public Function GetBookTimers(ByVal bookID As String, ByVal app As Object) As BookTimers" & n
s = s & "    If m_appTimers Is Nothing Then" & n
s = s & "        Set m_appTimers = New AppTimers" & n
s = s & "        #If Mac Then" & n
s = s & "            TimerResolution" & n
s = s & "        #End If" & n
s = s & "        Set m_cb = app.CommandBars.Add(Position:=msoBarPopup, Temporary:=True)" & n
s = s & "        m_cb.Controls.Add msoControlButton" & n
s = s & "        Application.OnTime Now(), ""MainLoop""" & n
s = s & "    End If" & n
s = s & "    With New BookTimers" & n
s = s & "        .Init bookID" & n
s = s & "        m_appTimers.Add .Self" & n
s = s & "        Set GetBookTimers = .Self" & n
s = s & "    End With" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Sub MainLoop()" & n
s = s & "    Do While IsConnected() Or m_appTimers.Count > 0" & n
s = s & "        m_appTimers.CheckRefs" & n
s = s & "        If m_appTimers.Count > 0 Then" & n
s = s & "            If Not m_appTimers.PopIfNeeded Then SleepIfNoEntry" & n
s = s & "        Else" & n
s = s & "            SleepIfNoEntry" & n
s = s & "        End If" & n
s = s & "        DoEvents" & n
s = s & "    Loop" & n
s = s & "    Application.Quit" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Function IsConnected() As Boolean" & n
s = s & "    Static lastCheck As Date" & n
s = s & "    Dim mNow As Date: mNow = NowMSec()" & n
s = s & "    '" & n
s = s & "    If lastCheck < mNow Then" & n
s = s & "        Dim bCount As Long" & n
s = s & "        On Error Resume Next" & n
s = s & "        bCount = m_cb.Controls.Count" & n
s = s & "        On Error GoTo 0" & n
s = s & "        If bCount = 0 Then Exit Function" & n
s = s & "        lastCheck = mNow + TimeSerial(0, 0, 1)" & n
s = s & "    End If" & n
s = s & "    IsConnected = True" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Private Sub SleepIfNoEntry()" & n
s = s & "    Dim isEntryRequested As Boolean" & n
s = s & "    Static isStarted As Boolean" & n
s = s & "    Static endTime As Date" & n
s = s & "    '" & n
s = s & "    On Error Resume Next" & n
s = s & "    isEntryRequested = (GetSetting(""RemoteTimers"", ""Flags"", ""EntryNeeded"") = ""1"")" & n
s = s & "    On Error GoTo 0" & n
s = s & "    If isStarted Then" & n
s = s & "        If isEntryRequested Then" & n
s = s & "            If endTime < NowMSec() Then isStarted = False" & n
s = s & "        Else" & n
s = s & "            isStarted = False" & n
s = s & "        End If" & n
s = s & "    Else" & n
s = s & "        If isEntryRequested Then" & n
s = s & "            endTime = NowMSec() + TimeSerial(0, 0, 1) / 10 '100ms timeout" & n
s = s & "            isStarted = True" & n
s = s & "        Else" & n
s = s & "            Sleep 1" & n
s = s & "        End If" & n
s = s & "    End If" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Function NowMSec() As Date" & n
s = s & "#If Mac Then" & n
s = s & "    Const evalResolution As Double = 0.01" & n
s = s & "    Const evalFunc As String = ""=Now()""" & n
s = s & "    Static useEval As Boolean" & n
s = s & "    Static isSet As Boolean" & n
s = s & "    '" & n
s = s & "    If Not isSet Then" & n
s = s & "        useEval = TimerResolution > evalResolution" & n
s = s & "        isSet = True" & n
s = s & "    End If" & n
s = s & "    If useEval Then" & n
s = s & "        NowMSec = Evaluate(evalFunc)" & n
s = s & "        Exit Function" & n
s = s & "    End If" & n
s = s & "#End If" & n
s = s & "    Const secondsPerDay As Long = 24& * 60& * 60&" & n
s = s & "    NowMSec = Date + Round(Timer, 3) / secondsPerDay" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "#If Mac Then" & n
s = s & "Private Function TimerResolution() As Double" & n
s = s & "    Const secondsPerDay As Long = 24& * 60& * 60&" & n
s = s & "    Static r As Double" & n
s = s & "    If r = 0 Then" & n
s = s & "        Dim t As Double: t = Timer" & n
s = s & "        Do" & n
s = s & "            r = Round(Timer - t, 3)" & n
s = s & "            If r < 0# Then r = r + secondsPerDay 'Passed midnight" & n
s = s & "        Loop Until r > 0#" & n
s = s & "    End If" & n
s = s & "    TimerResolution = r" & n
s = s & "End Function" & n
s = s & "#End If"
LibRemoteCode = s
End Function
Private Function TimerContainerCode() As String
Dim s As String
Const n As String = vbNewLine
s = s & "Option Explicit" & n
s = s & "" & n
s = s & "Private m_callback As Object" & n
s = s & "Private m_id As String" & n
s = s & "Private m_delayMs As Long" & n
s = s & "Private m_earliestTime As Date" & n
s = s & "Private m_originalTime As Date" & n
s = s & "Private m_daysDelay As Double" & n
s = s & "" & n
s = s & "Public Sub Init(ByRef tCallback As Object _" & n
s = s & "              , ByRef timerID As String _" & n
s = s & "              , ByRef earliestCallTime As Date _" & n
s = s & "              , ByRef delayMs As Long)" & n
s = s & "    Const msPerDay As Long = 24& * 60& * 60& * 1000&" & n
s = s & "    Dim mNow As Date: mNow = NowMSec()" & n
s = s & "    '" & n
s = s & "    Set m_callback = tCallback" & n
s = s & "    m_id = timerID" & n
s = s & "    If delayMs > 0 Then m_delayMs = delayMs" & n
s = s & "    m_daysDelay = m_delayMs / msPerDay" & n
s = s & "    '" & n
s = s & "    If earliestCallTime > mNow Then" & n
s = s & "        m_earliestTime = earliestCallTime" & n
s = s & "        m_originalTime = earliestCallTime" & n
s = s & "    Else" & n
s = s & "        m_earliestTime = mNow + m_daysDelay" & n
s = s & "        m_originalTime = mNow" & n
s = s & "    End If" & n
s = s & "End Sub" & n
s = s & "Public Property Get TimerCallback() As Object" & n
s = s & "    Set TimerCallback = m_callback" & n
s = s & "End Property" & n
s = s & "Public Property Get ID() As String" & n
s = s & "    ID = m_id" & n
s = s & "End Property" & n
s = s & "Public Property Get EarliestTime() As Date" & n
s = s & "    EarliestTime = m_earliestTime" & n
s = s & "End Property" & n
s = s & "Public Property Get Delay() As Long" & n
s = s & "    Delay = m_delayMs" & n
s = s & "End Property" & n
s = s & "Public Function Self() As TimerContainer" & n
s = s & "    Set Self = Me" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Sub UpdateTime()" & n
s = s & "    Dim skipCount As Long" & n
s = s & "    '" & n
s = s & "    skipCount = Int((NowMSec() - m_originalTime) / m_daysDelay)" & n
s = s & "    m_earliestTime = m_originalTime + (skipCount + 1) * m_daysDelay" & n
s = s & "End Sub"
TimerContainerCode = s
End Function
Private Function AppTimersCode() As String
Dim s As String
Const n As String = vbNewLine
s = s & "Option Explicit" & n
s = s & "" & n
s = s & "Private m_bookTimers As Collection" & n
s = s & "" & n
s = s & "Private Sub Class_Initialize()" & n
s = s & "    Set m_bookTimers = New Collection" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Sub Class_Terminate()" & n
s = s & "    Set m_bookTimers = Nothing" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Sub Add(ByVal bTimers As BookTimers)" & n
s = s & "    On Error Resume Next" & n
s = s & "    m_bookTimers.Remove bTimers.ID" & n
s = s & "    On Error GoTo 0" & n
s = s & "    m_bookTimers.Add bTimers, bTimers.ID" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Sub CheckRefs()" & n
s = s & "    Const localRefs As Long = 2 'm_bookTimers.Item + bt" & n
s = s & "    Dim bt As BookTimers" & n
s = s & "    '" & n
s = s & "    For Each bt In m_bookTimers" & n
s = s & "        If bt.RefsCount = localRefs Then" & n
s = s & "            m_bookTimers.Remove bt.ID" & n
s = s & "        End If" & n
s = s & "    Next bt" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Function Count() As Long" & n
s = s & "    Count = m_bookTimers.Count" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function PopIfNeeded() As Boolean" & n
s = s & "    Dim bt As BookTimers" & n
s = s & "    Dim minBT As BookTimers" & n
s = s & "    '" & n
s = s & "    For Each bt In m_bookTimers" & n
s = s & "        If bt.Count > 0 Then" & n
s = s & "            If minBT Is Nothing Then" & n
s = s & "                Set minBT = bt" & n
s = s & "            ElseIf bt.EarliestTime < minBT.EarliestTime Then" & n
s = s & "                Set minBT = bt" & n
s = s & "            End If" & n
s = s & "        End If" & n
s = s & "    Next bt" & n
s = s & "    If Not minBT Is Nothing Then" & n
s = s & "        PopIfNeeded = minBT.PopIfNeeded()" & n
s = s & "    End If" & n
s = s & "End Function"
AppTimersCode = s
End Function
Private Function BookTimersCode() As String
Dim s As String
Const n As String = vbNewLine
s = s & "Option Explicit" & n
s = s & "" & n
s = s & "#If Mac Then" & n
s = s & "    #If VBA7 Then" & n
s = s & "        Private Declare PtrSafe Function CopyMemory Lib ""/usr/lib/libc.dylib"" Alias ""memmove"" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr" & n
s = s & "    #Else" & n
s = s & "        Private Declare Function CopyMemory Lib ""/usr/lib/libc.dylib"" Alias ""memmove"" (Destination As Any, Source As Any, ByVal Length As Long) As Long" & n
s = s & "    #End If" & n
s = s & "#Else 'Windows" & n
s = s & "    #If VBA7 Then" & n
s = s & "        Private Declare PtrSafe Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As LongPtr)" & n
s = s & "    #Else" & n
s = s & "        Private Declare Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As Long)" & n
s = s & "    #End If" & n
s = s & "#End If" & n
s = s & "#If Win64 Then" & n
s = s & "    Private Const PTR_SIZE As Long = 8" & n
s = s & "#Else" & n
s = s & "    Private Const PTR_SIZE As Long = 4" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "Private m_id As String" & n
s = s & "Private m_refCount As Variant" & n
s = s & "Private m_timers As Collection" & n
s = s & "" & n
s = s & "Public Sub Init(ByVal bookID As String)" & n
s = s & "    m_id = bookID" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Sub Class_Initialize()" & n
s = s & "    Set m_timers = New Collection" & n
s = s & "    SetRefCount" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Sub Class_Terminate()" & n
s = s & "    Set m_timers = Nothing" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Sub SetRefCount()" & n
s = s & "    Const VT_BYREF As Long = &H4000" & n
s = s & "    Dim iUnk As IUnknown: Set iUnk = Me" & n
s = s & "    m_refCount = ObjPtr(iUnk) + PTR_SIZE" & n
s = s & "    CopyMemory m_refCount, vbLong + VT_BYREF, 2" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Property Get RefsCount() As Long" & n
s = s & "    RefsCount = GetLongByRef(m_refCount) - 1 '-1 for Me" & n
s = s & "End Property" & n
s = s & "Private Function GetLongByRef(ByRef v As Variant) As Long" & n
s = s & "    GetLongByRef = v" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function Count() As Long" & n
s = s & "    Count = m_timers.Count" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Property Get ID() As String" & n
s = s & "    ID = m_id" & n
s = s & "End Property" & n
s = s & "" & n
s = s & "Public Function Self() As BookTimers" & n
s = s & "    Set Self = Me" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Property Get EarliestTime() As Date" & n
s = s & "    EarliestTime = m_timers(1).EarliestTime" & n
s = s & "End Property" & n
s = s & "" & n
s = s & "Public Function AddTimer(ByVal tCallback As Object _" & n
s = s & "                       , ByVal timerID As String _" & n
s = s & "                       , ByVal earliestCallTime As Date _" & n
s = s & "                       , ByVal delayMs As Long) As String" & n
s = s & "    If tCallback Is Nothing Then Exit Function" & n
s = s & "    If LenB(timerID) = 0 Then Exit Function" & n
s = s & "    '" & n
s = s & "    With New TimerContainer" & n
s = s & "        .Init tCallback, timerID, earliestCallTime, delayMs" & n
s = s & "        InsertTimer .Self" & n
s = s & "    End With" & n
s = s & "    AddTimer = timerID" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function DeleteTimer(ByVal timerID As String) As Boolean" & n
s = s & "    On Error Resume Next" & n
s = s & "    m_timers.Remove timerID" & n
s = s & "    On Error GoTo 0" & n
s = s & "    DeleteTimer = True" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function DeleteAllTimers() As Boolean" & n
s = s & "    Set m_timers = New Collection" & n
s = s & "    DeleteAllTimers = True" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Private Sub InsertTimer(ByRef container As TimerContainer)" & n
s = s & "    Dim tc As TimerContainer" & n
s = s & "    Dim i As Long: i = 1" & n
s = s & "    '" & n
s = s & "    For Each tc In m_timers" & n
s = s & "        If tc.EarliestTime > container.EarliestTime Then Exit For" & n
s = s & "        i = i + 1" & n
s = s & "    Next tc" & n
s = s & "    If m_timers.Count = 0 Or i > m_timers.Count Then" & n
s = s & "        m_timers.Add Item:=container, Key:=container.ID" & n
s = s & "    Else" & n
s = s & "        m_timers.Add Item:=container, Key:=container.ID, Before:=i" & n
s = s & "    End If" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Function PopIfNeeded() As Boolean" & n
s = s & "    Const errMissingArgument As Long = 449" & n
s = s & "    Const errNotAvailable As Long = 1004" & n
s = s & "    Const errObjDisconnected As Long = -2147417848" & n
s = s & "    Const errRunFailed As Long = 50290" & n
s = s & "    Const errTypeMismatch As Long = 13" & n
s = s & "    '" & n
s = s & "    Dim tc As TimerContainer: Set tc = m_timers(1)" & n
s = s & "    Dim remoteErrCode As Long" & n
s = s & "    '" & n
s = s & "    If tc.EarliestTime > NowMSec() Then Exit Function" & n
s = s & "    PopIfNeeded = True" & n
s = s & "    '" & n
s = s & "    On Error Resume Next" & n
s = s & "    remoteErrCode = tc.TimerCallback.TimerProc() 'Possible re-entry point!" & n
s = s & "    If Err.Number = errObjDisconnected Then Exit Function" & n
s = s & "    Err.Clear" & n
s = s & "    m_timers.Remove tc.ID" & n
s = s & "    If Err.Number <> 0 Then 'Timer was removed via re-entry" & n
s = s & "        Err.Clear" & n
s = s & "        Exit Function" & n
s = s & "    End If" & n
s = s & "    On Error GoTo 0" & n
s = s & "    '" & n
s = s & "    If remoteErrCode = errMissingArgument Then Exit Function" & n
s = s & "    If remoteErrCode = errNotAvailable Then Exit Function" & n
s = s & "    If remoteErrCode = errTypeMismatch Then Exit Function" & n
s = s & "    '" & n
s = s & "    If tc.Delay > 0 Then" & n
s = s & "        tc.UpdateTime" & n
s = s & "        InsertTimer tc" & n
s = s & "    ElseIf remoteErrCode = errRunFailed Then" & n
s = s & "        InsertTimer tc '0 delay timers are guaranteed to be called once!" & n
s = s & "    End If" & n
s = s & "End Function"
BookTimersCode = s
End Function
