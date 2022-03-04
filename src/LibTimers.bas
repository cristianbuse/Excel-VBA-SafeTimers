Attribute VB_Name = "LibTimers"
'''=============================================================================
''' VBA MemoryTools
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

#If VBA7 = 0 Then       'LongPtr trick discovered by @Greedo (https://github.com/Greedquest)
    Public Enum LongPtr
        [_]
    End Enum            'Kindly given here:
#End If                 'https://github.com/cristianbuse/VBA-MemoryTools/issues/3

Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As LongPtr, ByVal dwId As Long, riid As GUID, ppvObject As Object) As Long
    Private Declare PtrSafe Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function KillTimerAPI Lib "user32" Alias "KillTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function SetTimerAPI Lib "user32" Alias "SetTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
#Else
    Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As Long, ByVal dwId As Long, riid As GUID, ppvObject As Object) As Long
    Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function KillTimerAPI Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function SetTimerAPI Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
#End If

Const USER_TIMER_MAXIMUM As Long = &H7FFFFFFF 'Around 25 days
Private Const BOOK_NAME As String = "RemoteTimersAPI_V1.xlam"

Private m_localTimers As Collection
Private m_remoteTimers As Object
Private m_VBIDEHWnd As LongPtr
Private m_bookID As String

'*******************************************************************************
'An enhanced 'Now' - returns the date and time including milliseconds
'*******************************************************************************
Public Function NowMs() As Date
    Const secondsPerDay As Long = 24& * 60& * 60&
    NowMs = Date + Round(Timer, 3) / secondsPerDay
End Function

'*******************************************************************************
'Safe wrapper around Win API
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer
'Parameters:
'   - hWnd: a handle to the window to be associated with the timer
'   - nIDEvent: a nonzero timer identifier
'   - uElapse: the time-out value, in milliseconds
'   - lpTimerFunc: a pointer to the function to be notified
'*******************************************************************************
Public Function SetTimer(ByVal hWnd As LongPtr _
                       , ByVal nIDEvent As LongPtr _
                       , ByVal uElapse As Long _
                       , ByVal lpTimerFunc As LongPtr) As LongPtr
    Const minDelay As Long = 1
    Dim result As LongPtr
    '
    If lpTimerFunc = 0 Then Exit Function
    If uElapse < minDelay Then uElapse = minDelay
    '
    If Not InitTimers(False) Then Exit Function
    result = SetTimerAPI(hWnd, nIDEvent, USER_TIMER_MAXIMUM, lpTimerFunc)
    If result = 0 Then Exit Function
    KillTimerAPI hWnd, nIDEvent 'No longer needed
    '
    If hWnd = 0 Then
        hWnd = Application.hWnd 'Save the implicit hWnd
        nIDEvent = result
    End If
    '
    Dim sID As String: sID = GetTimerID(hWnd, nIDEvent)
    Dim remoteResult As Boolean
    '
    On Error Resume Next
    remoteResult = m_remoteTimers.AddTimer(hWnd, nIDEvent, sID, uElapse)
    On Error GoTo 0
    If Not remoteResult Then Exit Function
    '
    On Error Resume Next
    m_localTimers.Remove sID
    m_localTimers.Add lpTimerFunc, sID 'Dispatch will need the TimerProc later
    On Error GoTo 0
    '
    SetTimer = result
End Function

'*******************************************************************************
'The only TimerProc called remotely
'*******************************************************************************
Private Sub TimerProc(ByVal hWnd As LongPtr _
                    , ByVal wMsg As Long _
                    , ByVal nIDEvent As LongPtr _
                    , ByVal wTime As Long)
    Dim oPtr As LongPtr: oPtr = ObjPtr(ThisWorkbook)
    Dim rHWnd As LongPtr: rHWnd = GetReadyHWnd()
    '
    KillTimerAPI rHWnd, oPtr 'Kill the only TimerProc
    If oPtr = nIDEvent Then Exit Sub
    '
    On Error Resume Next
    Dim tProc As LongPtr: tProc = m_localTimers(GetTimerID(hWnd, nIDEvent))
    On Error GoTo 0
    If tProc = 0 Then Exit Sub 'State was lost
    '
    Dim sDisp As New SafeDispatch 'Will dispatch msg on termination
    sDisp.Init hWnd, wMsg, nIDEvent, tProc, wTime, m_bookID
    '
    SetTimerAPI rHWnd, oPtr, USER_TIMER_MAXIMUM, AddressOf TimerProc
End Sub

'*******************************************************************************
'Utility for collection keys
'*******************************************************************************
Private Function GetTimerID(ByVal hWnd As LongPtr _
                          , ByVal nIDEvent As LongPtr) As String
    GetTimerID = hWnd & "_" & nIDEvent
End Function

'*******************************************************************************
'Safe wrapper around Win API
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-killtimer
'Parameters:
'   - hWnd: a handle to the window associated with the specified timer
'   - nIDEvent: the timer to be destroyed
'*******************************************************************************
Public Function KillTimer(ByVal hWnd As LongPtr _
                        , ByVal nIDEvent As LongPtr) As Long
    Dim sID As String: sID = GetTimerID(hWnd, nIDEvent)
    Dim remoteResult As Boolean
    '
    On Error Resume Next
    remoteResult = m_remoteTimers.DeleteTimer(sID)
    m_localTimers.Remove sID
    On Error GoTo 0
    '
    If remoteResult Then KillTimer = 1
End Function

'*******************************************************************************
'Removes all existing timers
'*******************************************************************************
Public Sub RemoveAllTimers()
    On Error Resume Next
    If m_remoteTimers.DeleteAllTimers() Then Set m_localTimers = New Collection
    On Error GoTo 0
End Sub

'*******************************************************************************
'Returns 'True' only if the object is set and still connected to the remote app
'*******************************************************************************
Private Function IsConnected(ByVal obj As Object) As Boolean
    If Not obj Is Nothing Then
        IsConnected = (TypeName(obj) <> "Object")
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
    Dim tProc As LongPtr: tProc = VBA.Int(AddressOf TimerProc)
    Dim oPtr As LongPtr:  oPtr = ObjPtr(ThisWorkbook)
    Dim rHWnd As LongPtr: rHWnd = GetReadyHWnd()
    m_bookID = CStr(oPtr)
    '
    Set m_localTimers = New Collection
    Set m_remoteTimers = app.Run("GetBookTimers", rHWnd, m_bookID, tProc)
    SetTimerAPI rHWnd, oPtr, USER_TIMER_MAXIMUM, tProc
    '
    InitTimers = True
End Function
Private Function IsFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    IsFile = ((GetAttr(filePath) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function
Private Function GetBookPath() As String
    Dim folderPath As String: folderPath = Environ$("temp")
    GetBookPath = folderPath & Application.PathSeparator & BOOK_NAME
End Function

'*******************************************************************************
'Returns the existing remote app or opens a new one if needed
'*******************************************************************************
Private Function GetRemoteApp() As Application
    Dim mainHWnd As LongPtr
    Dim remoteHWnd As LongPtr
    Dim app As Application
    Dim book As Workbook
    '
    Do
        Set app = GetNextApplication(mainHWnd)
        If Not app Is Nothing Then
            Set book = Nothing
            remoteHWnd = 0
            '
            On Error Resume Next
            Set book = app.Workbooks(BOOK_NAME)
            If Not book Is Nothing Then
                remoteHWnd = app.Run("GetReadyHWnd")
            End If
            On Error GoTo 0
            If remoteHWnd = GetReadyHWnd() Then Exit Do
            Set app = Nothing
        End If
    Loop Until mainHWnd = 0
    If app Is Nothing Then
        Set app = NewApp()
        app.Workbooks.Open GetBookPath(), False, False
    End If
    Set GetRemoteApp = app
End Function
Private Function GetNextApplication(ByRef mainHWnd As LongPtr) As Application
    mainHWnd = FindWindowEx(0, mainHWnd, "XLMAIN", vbNullString)
    If mainHWnd = 0 Then Exit Function
    '
    Dim w As Window
    For Each w In Application.Windows
        If w.hWnd = mainHWnd Then Exit Function
    Next w
    '
    Const OBJID_NATIVEOM As Long = &HFFFFFFF0
    Dim deskHWnd As LongPtr
    Dim excelHWnd As LongPtr
    Dim wnd As Window
    '
    deskHWnd = FindWindowEx(mainHWnd, 0, "XLDESK", vbNullString)
    If deskHWnd = 0 Then Exit Function
    excelHWnd = FindWindowEx(deskHWnd, 0, "EXCEL7", vbNullString)
    If excelHWnd = 0 Then Exit Function
    '
    AccessibleObjectFromWindow excelHWnd, OBJID_NATIVEOM, IDispGuid(), wnd
    If wnd Is Nothing Then Exit Function
    Set GetNextApplication = wnd.Application
End Function
Private Function IDispGuid() As GUID
    With IDispGuid 'IDispatch
        .data1 = &H20400
        .data4(0) = &HC0
        .data4(7) = &H46
    End With
End Function

'*******************************************************************************
'Creates a new app instance and sets certain properties
'*******************************************************************************
Private Function NewApp() As Application
    Set NewApp = New Application
    With NewApp
        .Visible = False
        .PrintCommunication = False
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Interactive = False
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
    If Not isVBOMOn Then
        book.Close False
        app.Quit
        Set app = Nothing
        EnableOfficeVBOM False
        Set app = NewApp()
        app.Workbooks.Open GetBookPath(), False, False
    End If
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
    Dim rKey As String
    rKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version _
         & "\" & Replace(Application.Name, "Microsoft ", vbNullString) _
         & "\Security\AccessVBOM"
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite rKey, i, "REG_DWORD"
    EnableOfficeVBOM = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Returns the handle for the main VB IDE window
'*******************************************************************************
Private Function GetVBIDEHWnd() As LongPtr
    If m_VBIDEHWnd = 0 Then
        If IsVBOMEnabled() Then
            m_VBIDEHWnd = Application.VBE.MainWindow.hWnd
        Else
            EnumThreadWindows GetCurrentThreadId, AddressOf EnumThreadWndProcVBIDE, 0
        End If
    End If
    GetVBIDEHWnd = m_VBIDEHWnd
End Function
Private Function EnumThreadWndProcVBIDE(ByVal hWnd As LongPtr _
                                      , ByVal lParam As LongPtr) As Long
    Const className As String = "wndclass_desked_gsk"
    Const bufferSize As Long = 260
    Dim cName As String * bufferSize
    '
    If Left$(cName, GetClassName(hWnd, cName, bufferSize)) = className Then
        m_VBIDEHWnd = hWnd
        Exit Function
    End If
    EnumThreadWndProcVBIDE = 1
End Function

'*******************************************************************************
'Returns the handle for the '<Ready>' window under the parent 'Locals' window
'*******************************************************************************
Private Function GetReadyHWnd() As LongPtr
    Static readyHWnd As LongPtr
    If readyHWnd = 0 Then
        Dim localsHWnd As LongPtr
        localsHWnd = FindWindowEx(GetVBIDEHWnd(), 0, vbNullString, "Locals")
        readyHWnd = FindWindowEx(localsHWnd, 0, "Edit", vbNullString)
    End If
    GetReadyHWnd = readyHWnd
End Function

'*******************************************************************************
'Code running 'on the other side'
'*******************************************************************************
Private Function LibRemoteCode() As String
Dim s As String
Const n As String = vbNewLine
s = s & "Option Explicit" & n
s = s & "" & n
s = s & "#If VBA7 = 0 Then" & n
s = s & "    Public Enum LongPtr" & n
s = s & "        [_]" & n
s = s & "    End Enum" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "#If VBA7 Then" & n
s = s & "    Private Declare PtrSafe Function IsWindow Lib ""user32"" (ByVal hWnd As LongPtr) As Long" & n
s = s & "    Private Declare PtrSafe Function SendMessage Lib ""user32"" Alias ""SendMessageA"" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr" & n
s = s & "    Public Declare PtrSafe Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)" & n
s = s & "#Else" & n
s = s & "    Private Declare Function IsWindow Lib ""user32"" (ByVal hWnd As Long) As Long" & n
s = s & "    Private Declare Function SendMessage Lib ""user32"" Alias ""SendMessageA"" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long" & n
s = s & "    Public Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "Private m_appTimers As AppTimers" & n
s = s & "Private m_readyHWnd As LongPtr" & n
s = s & "" & n
s = s & "Public Function GetReadyHWnd() As LongPtr" & n
s = s & "    GetReadyHWnd = m_readyHWnd" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function GetBookTimers(ByVal readyHWnd As LongPtr _" & n
s = s & "                            , ByVal bookID As String _" & n
s = s & "                            , ByVal tProc As LongPtr) As BookTimers" & n
s = s & "    If m_readyHWnd = 0 Then" & n
s = s & "        m_readyHWnd = readyHWnd" & n
s = s & "        Set m_appTimers = New AppTimers" & n
s = s & "        Application.OnTime Now(), ""MainLoop""" & n
s = s & "    End If" & n
s = s & "    With New BookTimers" & n
s = s & "        .Init bookID, tProc" & n
s = s & "        m_appTimers.Add .Self" & n
s = s & "        Set GetBookTimers = .Self" & n
s = s & "    End With" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Sub MainLoop()" & n
s = s & "    Do While IsWindow(m_readyHWnd)" & n
s = s & "        m_appTimers.CheckRefs" & n
s = s & "        If m_appTimers.Count > 0 And m_appTimers.CanPost Then" & n
s = s & "            If Not m_appTimers.PopIfNeeded Then Sleep 1" & n
s = s & "        Else" & n
s = s & "            Sleep 1" & n
s = s & "        End If" & n
s = s & "        DoEvents" & n
s = s & "    Loop" & n
s = s & "    Set m_appTimers = Nothing" & n
s = s & "    Application.Quit" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Public Function IsIDEReady() As Boolean" & n
s = s & "    Const readyLabelCurANSI As String = ""1758492059378.1308"" '<Ready>" & n
s = s & "    Static readyLabel As Currency" & n
s = s & "    Const WM_GETTEXT As Long = &HD" & n
s = s & "    Dim buff As Currency" & n
s = s & "    '" & n
s = s & "    If readyLabel = 0 Then readyLabel = CCur(readyLabelCurANSI)" & n
s = s & "    If SendMessage(m_readyHWnd, WM_GETTEXT, 8, VarPtr(buff)) = 0 Then Exit Function" & n
s = s & "    IsIDEReady = (buff = readyLabel)" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function NowMSec() As Date" & n
s = s & "    Const secondsPerDay As Long = 24& * 60& * 60&" & n
s = s & "    NowMSec = Date + Round(Timer, 3) / secondsPerDay" & n
s = s & "End Function"
LibRemoteCode = s
End Function
Private Function TimerContainerCode() As String
Dim s As String
Const n As String = vbNewLine
s = s & "Option Explicit" & n
s = s & "" & n
s = s & "Private m_hWnd As LongPtr" & n
s = s & "Private m_nIDEvent As LongPtr" & n
s = s & "Private m_id As String" & n
s = s & "Private m_delayMs As Long" & n
s = s & "Private m_earliestTime As Date" & n
s = s & "Private m_originalTime As Date" & n
s = s & "" & n
s = s & "Public Sub Init(ByRef hWnd As LongPtr _" & n
s = s & "              , ByRef nIDEvent As LongPtr _" & n
s = s & "              , ByRef sID As String _" & n
s = s & "              , ByRef delayMs As Long _" & n
s = s & "              , ByRef callTime As Date)" & n
s = s & "    m_hWnd = hWnd" & n
s = s & "    m_nIDEvent = nIDEvent" & n
s = s & "    m_id = sID" & n
s = s & "    m_delayMs = delayMs" & n
s = s & "    m_earliestTime = callTime" & n
s = s & "    m_originalTime = m_earliestTime" & n
s = s & "End Sub" & n
s = s & "Public Function Self() As TimerContainer" & n
s = s & "    Set Self = Me" & n
s = s & "End Function" & n
s = s & "Public Property Get hWnd() As LongPtr" & n
s = s & "    hWnd = m_hWnd" & n
s = s & "End Property" & n
s = s & "Public Property Get EventID() As LongPtr" & n
s = s & "    EventID = m_nIDEvent" & n
s = s & "End Property" & n
s = s & "Public Property Get ID() As String" & n
s = s & "    ID = m_id" & n
s = s & "End Property" & n
s = s & "Public Property Get Delay() As Long" & n
s = s & "    Delay = m_delayMs" & n
s = s & "End Property" & n
s = s & "Public Property Get EarliestTime() As Date" & n
s = s & "    EarliestTime = m_earliestTime" & n
s = s & "End Property" & n
s = s & "" & n
s = s & "Public Sub UpdateTime()" & n
s = s & "    Const msPerDay As Long = 24& * 60& * 60& * 1000&" & n
s = s & "    Dim daysDelay As Double" & n
s = s & "    Dim skipCount As Long" & n
s = s & "    '" & n
s = s & "    daysDelay = m_delayMs / msPerDay" & n
s = s & "    skipCount = Int((NowMSec - m_originalTime) / daysDelay)" & n
s = s & "    m_earliestTime = m_originalTime + (skipCount + 1) * daysDelay" & n
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
s = s & "Public Function CanPost() As Boolean" & n
s = s & "    Dim bt As BookTimers" & n
s = s & "    For Each bt In m_bookTimers" & n
s = s & "        If Not bt.CanPost Then Exit Function" & n
s = s & "    Next bt" & n
s = s & "    CanPost = True" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Sub CheckRefs()" & n
s = s & "    Dim bt As BookTimers" & n
s = s & "    For Each bt In m_bookTimers" & n
s = s & "        If bt.RefsCount = 3 Then" & n
s = s & "            m_bookTimers.Remove bt.ID" & n
s = s & "            bt.KillBookTimer" & n
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
s = s & "        If minBT Is Nothing Then" & n
s = s & "            If bt.Count > 0 Then Set minBT = bt" & n
s = s & "        Else" & n
s = s & "            If bt.Count > 0 Then" & n
s = s & "                If bt.EarliestTime < minBT.EarliestTime Then" & n
s = s & "                    Set minBT = bt" & n
s = s & "                End If" & n
s = s & "            End If" & n
s = s & "        End If" & n
s = s & "    Next bt" & n
s = s & "    If minBT Is Nothing Then" & n
s = s & "        PopIfNeeded = False" & n
s = s & "    Else" & n
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
s = s & "#If VBA7 Then" & n
s = s & "    Private Declare PtrSafe Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As LongPtr)" & n
s = s & "    Private Declare PtrSafe Function IsWindow Lib ""user32"" (ByVal hWnd As LongPtr) As Long" & n
s = s & "    Private Declare PtrSafe Function PostMessage Lib ""user32"" Alias ""PostMessageA"" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long" & n
s = s & "#Else" & n
s = s & "    Private Declare Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As Long)" & n
s = s & "    Private Declare Function IsWindow Lib ""user32"" (ByVal hWnd As Long) As Long" & n
s = s & "    Private Declare Function PostMessage Lib ""user32"" Alias ""PostMessageA"" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "#If Win64 Then" & n
s = s & "    Private Const PTR_SIZE As Long = 8" & n
s = s & "#Else" & n
s = s & "    Private Const PTR_SIZE As Long = 4" & n
s = s & "#End If" & n
s = s & "" & n
s = s & "Private m_canPost As Boolean" & n
s = s & "Private m_id As String" & n
s = s & "Private m_refCount As Variant" & n
s = s & "Private m_timers As Collection" & n
s = s & "Private m_tProc As LongPtr" & n
s = s & "" & n
s = s & "Public Sub Init(ByVal bookID As String, ByVal tProc As LongPtr)" & n
s = s & "    m_id = bookID" & n
s = s & "    m_tProc = tProc" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Sub Class_Initialize()" & n
s = s & "    Set m_timers = New Collection" & n
s = s & "    SetRefCount" & n
s = s & "    m_canPost = True" & n
s = s & "End Sub" & n
s = s & "" & n
s = s & "Private Sub Class_Terminate()" & n
s = s & "    Set m_timers = Nothing" & n
s = s & "    On Error Resume Next" & n
s = s & "    DeleteSetting ""SafeTimers"", m_id" & n
s = s & "    On Error GoTo 0" & n
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
s = s & "    RefsCount = GetLongByRef(m_refCount)" & n
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
s = s & "Public Property Get CanPost() As Boolean" & n
s = s & "    If Not m_canPost Then" & n
s = s & "        m_canPost = (GetSetting(""SafeTimers"", m_id, ""CanPost"") = ""True"")" & n
s = s & "        If m_canPost Then" & n
s = s & "            Dim lostID As String" & n
s = s & "            lostID = GetSetting(""SafeTimers"", m_id, ""LostID"")" & n
s = s & "            If LenB(lostID) > 0 Then" & n
s = s & "                DeleteTimer lostID" & n
s = s & "                DeleteSetting ""SafeTimers"", m_id, ""LostID""" & n
s = s & "            End If" & n
s = s & "        End If" & n
s = s & "    End If" & n
s = s & "    CanPost = m_canPost" & n
s = s & "End Property" & n
s = s & "" & n
s = s & "Public Property Get EarliestTime() As Date" & n
s = s & "    EarliestTime = m_timers(1).EarliestTime" & n
s = s & "End Property" & n
s = s & "" & n
s = s & "Public Function AddTimer(ByVal hWnd As LongPtr _" & n
s = s & "                       , ByVal nIDEvent As LongPtr _" & n
s = s & "                       , ByVal sID As String _" & n
s = s & "                       , ByVal delayMs As Long) As Boolean" & n
s = s & "    DeleteTimer sID" & n
s = s & "    With New TimerContainer" & n
s = s & "        Const msPerDay As Long = 24& * 60& * 60& * 1000&" & n
s = s & "        Dim nextRun As Date: nextRun = NowMSec() + delayMs / msPerDay" & n
s = s & "        '" & n
s = s & "        .Init hWnd, nIDEvent, sID, delayMs, nextRun" & n
s = s & "        InsertTimer .Self" & n
s = s & "    End With" & n
s = s & "    AddTimer = True" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Function DeleteTimer(ByVal sID As String) As Boolean" & n
s = s & "    On Error Resume Next" & n
s = s & "    m_timers.Remove sID" & n
s = s & "    DeleteTimer = (Err.Number = 0)" & n
s = s & "    On Error GoTo 0" & n
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
s = s & "    Const WM_TIMER As Long = &H113" & n
s = s & "    Dim tc As TimerContainer: Set tc = m_timers(1)" & n
s = s & "    '" & n
s = s & "    If tc.EarliestTime > NowMSec() Then Exit Function" & n
s = s & "    If Not IsIDEReady() Then Exit Function" & n
s = s & "    '" & n
s = s & "    m_timers.Remove 1" & n
s = s & "    If PostMessage(tc.hWnd, WM_TIMER, tc.EventID, m_tProc) = 0& Then" & n
s = s & "        If IsWindow(tc.hWnd) = 0& Then Exit Function" & n
s = s & "    End If" & n
s = s & "    m_canPost = False" & n
s = s & "    PopIfNeeded = True" & n
s = s & "    '" & n
s = s & "    tc.UpdateTime" & n
s = s & "    InsertTimer tc" & n
s = s & "End Function" & n
s = s & "" & n
s = s & "Public Sub KillBookTimer()" & n
s = s & "    Const WM_TIMER As Long = &H113" & n
s = s & "    Dim rHWnd As LongPtr: rHWnd = GetReadyHWnd()" & n
s = s & "    Dim tID As LongPtr: tID = VBA.Int(m_id)" & n
s = s & "    '" & n
s = s & "    Do While IsWindow(rHWnd)" & n
s = s & "        If IsIDEReady() Then" & n
s = s & "            If PostMessage(rHWnd, WM_TIMER, tID, m_tProc) <> 0& Then Exit Do" & n
s = s & "        End If" & n
s = s & "        Sleep 1" & n
s = s & "    Loop" & n
s = s & "End Sub"
BookTimersCode = s
End Function
