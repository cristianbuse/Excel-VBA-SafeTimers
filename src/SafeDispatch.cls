VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SafeDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type MSG
    hWnd As LongPtr
    wMsg As Long
    wParam As LongPtr
    lParam As LongPtr
    wTime As Long
    pt As POINTAPI
End Type

#If VBA7 Then
    Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As LongPtr
    Private Declare PtrSafe Function KillTimerAPI Lib "user32" Alias "KillTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function SetTimerAPI Lib "user32" Alias "SetTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
#Else
    Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
    Private Declare Function KillTimerAPI Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function SetTimerAPI Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
#End If

Private m_msg As MSG
Private m_bookID As String

Public Sub Init(ByVal hWnd As LongPtr _
              , ByVal wMsg As Long _
              , ByVal wParam As LongPtr _
              , ByVal lParam As LongPtr _
              , ByVal wTime As Long _
              , ByVal bookID As String)
    With m_msg
        .hWnd = hWnd
        .wMsg = wMsg
        .wParam = wParam
        .lParam = lParam
        .wTime = wTime
    End With
    m_bookID = bookID
End Sub

Private Sub Class_Terminate()
    If LenB(m_bookID) = 0 Then Exit Sub 'In case state was lost
    '
    Const USER_TIMER_MAXIMUM As Long = &H7FFFFFFF 'Around 25 days
    With m_msg
        If SetTimerAPI(.hWnd, .wParam, USER_TIMER_MAXIMUM, .lParam) <> 0 Then
            DispatchMessage m_msg
            KillTimerAPI .hWnd, .wParam
        Else
            SaveSetting "SafeTimers", m_bookID, "LostID", .hWnd & "_" & .wParam
        End If
    End With
    SaveSetting "SafeTimers", m_bookID, "CanPost", "True"
End Sub
