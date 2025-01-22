Function GetObjectWithImprovedErrorHandling(objectPath)
  On Error Resume Next
  Dim obj
  Set obj = GetObject(objectPath)
  If Err.Number <> 0 Then
    Select Case Err.Number
      Case 429:
        'ActiveX component can't create object.
        MsgBox "Error: ActiveX component can't create object. Check if the object is registered correctly.", vbCritical
      Case 424:
        'Object required.
        MsgBox "Error: Object required. The specified object path may be invalid.", vbCritical
      Case Else
        MsgBox "Error: An unspecified error occurred while getting the object. Error number: " & Err.Number, vbCritical
    End Select
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObjectWithImprovedErrorHandling = obj
End Function