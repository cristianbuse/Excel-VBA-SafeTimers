Attribute VB_Name = "CodeGenerator"
Option Explicit
Option Private Module

Private Function GetCodeModule(ByVal cmpName As String) As Object
    Const n As String = vbNewLine
    Dim cmp As Object
    Dim i As Long
    '
    If Not IsVBOMEnabled() Then
        MsgBox "Please enable 'Trust access to the VBA project object model'" _
             & " under File/Options/Trust Center/Trust Center Settings/" _
             & "Macro Settings/Developer Macro Settings" _
             , vbInformation, "Cancelled"
        Exit Function
    End If
    With ThisWorkbook.VBProject.VBComponents
        For i = 1 To .Count
            Set cmp = .Item(i)
            If cmp.Name = cmpName Then
                Set GetCodeModule = cmp.CodeModule
                Exit For
            End If
        Next i
    End With
End Function

Private Function IsVBOMEnabled() As Boolean
    On Error Resume Next
    IsVBOMEnabled = Not Application.VBE.ActiveVBProject Is Nothing
    On Error GoTo 0
End Function

Private Function GetLibRemote() As String
    Dim cm As Object: Set cm = GetCodeModule("LibRemote")
    If cm Is Nothing Then Exit Function
    '
    Const lineSplitter As String = """ & n" & vbNewLine & "s = s & """
    Const n As String = vbNewLine
    Dim fullCode As String
    '
    fullCode = cm.Lines(1, cm.CountOfLines)
    fullCode = Replace(fullCode, """", """""")
    '
    GetLibRemote = "Private Function LibRemoteCode() As String" & n _
                  & "Dim s As String" & n _
                  & "Const n As String = vbNewLine" & n _
                  & "s = s & """ _
                  & Replace(fullCode, n, lineSplitter) & """" & n _
                  & "LibRemoteCode = s" & n _
                  & "End Function"
End Function

Private Sub PrintLibRemote()
    Debug.Print GetLibRemote()
End Sub

Public Function GetTimerContainer() As String
    Dim cm As Object: Set cm = GetCodeModule("TimerContainer")
    If cm Is Nothing Then Exit Function
    '
    Const lineSplitter As String = """ & n" & vbNewLine & "s = s & """
    Const n As String = vbNewLine
    Dim fullCode As String
    '
    fullCode = cm.Lines(1, cm.CountOfLines)
    fullCode = Replace(fullCode, """", """""")
    '
    GetTimerContainer = "Private Function TimerContainerCode() As String" & n _
                      & "Dim s As String" & n _
                      & "Const n As String = vbNewLine" & n _
                      & "s = s & """ _
                      & Replace(fullCode, n, lineSplitter) & """" & n _
                      & "TimerContainerCode = s" & n _
                      & "End Function"
End Function

Private Sub PrintTimerContainer()
    Debug.Print GetTimerContainer()
End Sub

Public Function GetAppTimers() As String
    Dim cm As Object: Set cm = GetCodeModule("AppTimers")
    If cm Is Nothing Then Exit Function
    '
    Const lineSplitter As String = """ & n" & vbNewLine & "s = s & """
    Const n As String = vbNewLine
    Dim fullCode As String
    '
    fullCode = cm.Lines(1, cm.CountOfLines)
    fullCode = Replace(fullCode, """", """""")
    '
    GetAppTimers = "Private Function AppTimersCode() As String" & n _
              & "Dim s As String" & n _
              & "Const n As String = vbNewLine" & n _
              & "s = s & """ _
              & Replace(fullCode, n, lineSplitter) & """" & n _
              & "AppTimersCode = s" & n _
              & "End Function"
End Function

Private Sub PrintAppTimers()
    Debug.Print GetAppTimers()
End Sub

Public Function GetBookTimers() As String
    Dim cm As Object: Set cm = GetCodeModule("BookTimers")
    If cm Is Nothing Then Exit Function
    '
    Const lineSplitter As String = """ & n" & vbNewLine & "s = s & """
    Const n As String = vbNewLine
    Dim fullCode As String
    '
    fullCode = cm.Lines(1, cm.CountOfLines)
    fullCode = Replace(fullCode, """", """""")
    '
    GetBookTimers = "Private Function BookTimersCode() As String" & n _
              & "Dim s As String" & n _
              & "Const n As String = vbNewLine" & n _
              & "s = s & """ _
              & Replace(fullCode, n, lineSplitter) & """" & n _
              & "BookTimersCode = s" & n _
              & "End Function"
End Function

Private Sub PrintBookTimers()
    Debug.Print GetBookTimers()
End Sub
