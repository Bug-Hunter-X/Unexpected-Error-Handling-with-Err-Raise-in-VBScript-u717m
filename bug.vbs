Function MyFunction(param)
  If IsEmpty(param) Then
    Err.Raise 9999, , "Parameter cannot be empty!" 'This line is problematic for certain scenarios.
  End If
  ' ... rest of your function code ...
End Function