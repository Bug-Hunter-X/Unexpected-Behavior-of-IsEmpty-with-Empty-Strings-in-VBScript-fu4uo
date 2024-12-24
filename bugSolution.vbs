Function MyFunction(param1)
  If param1 = "" Then
    ' Explicitly check for empty string
    ' Handle empty string parameter
    Exit Function
  ElseIf IsEmpty(param1) Then
    ' Handle other cases of empty variants
    Exit Function
  End If

  ' ... rest of the function ...
End Function