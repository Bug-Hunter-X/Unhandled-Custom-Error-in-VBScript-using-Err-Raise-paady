Function MyFunction(param)
  On Error Resume Next
  If IsEmpty(param) Then
    Err.Raise 9999, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function

Sub Main()
  On Error GoTo ErrorHandler

  Dim result
  result = MyFunction(Null)

  MsgBox "Function executed successfully: " & result
  Exit Sub

ErrorHandler:
  If Err.Number = 9999 Then
    MsgBox "Custom Error: " & Err.Description
  Else
    MsgBox "An unexpected error occurred: " & Err.Description
  End If
  Err.Clear
End Sub