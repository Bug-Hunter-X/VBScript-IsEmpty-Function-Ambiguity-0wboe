Function MyFunction(param1, param2)
  If param1 = vbNullString Then
    'Handle empty string specifically
    param1 = ""
  ElseIf IsEmpty(param1) Then
    'Handle uninitialized variable
    param1 = "DefaultValue"
  End If
  If IsEmpty(param2) Then
    ' Handle uninitialized numeric parameter
    param2 = 0
  End If
  ' ... rest of the function ...
End Function