VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'@Exposed

Private m_id As String
Private m_callbackName As String
Private m_data As Variant

Public Sub Init(ByVal callbackName As String, ByVal argData As Variant)
    m_id = CStr(ObjPtr(Me))
    m_callbackName = callbackName
    If IsObject(argData) Then Set m_data = argData Else m_data = argData
End Sub

Public Function Self() As TimerCallback
    Set Self = Me
End Function

Public Property Get ID() As String
    ID = m_id
End Property

Public Function TimerProc() As Long
    Const errMissingArgument As Long = 449
    Const errNotAvailable As Long = 1004
    Const errTypeMismatch As Long = 13
    Dim errCode As Long
    '
    On Error Resume Next
    Application.Run m_callbackName, m_id, m_data
    errCode = Err.Number
    On Error GoTo 0
    '
    Select Case errCode
    Case errMissingArgument
        Debug.Print m_id & " - Missing argument for: " & m_callbackName
    Case errNotAvailable
        Debug.Print m_id & " - Invalid method name: " & m_callbackName
    Case errTypeMismatch
        Debug.Print m_id & " - Wrong argument type for: " & m_callbackName
    End Select
    '
    TimerProc = errCode
End Function
