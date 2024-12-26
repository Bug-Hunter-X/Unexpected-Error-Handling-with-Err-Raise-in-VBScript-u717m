Function MyFunction(param)
  On Error GoTo ErrHandler
  If IsEmpty(param) Then
    Err.Raise vbError + 1, , "Parameter cannot be empty!" ' Use vbError to avoid potential conflicts
  End If
  ' ... rest of your function code ...
  Exit Function
ErrHandler:
  'Handle the error appropriately, provide more context
  MsgBox "Error in MyFunction: " & Err.Description & " (Error Number: " & Err.Number & ")", vbCritical
  Err.Clear
End Function