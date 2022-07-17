Attribute VB_Name = "Errores"
Public Sub LogError(ErrNumber As Long, ErrDesc As String, strTitle As String, iLineNu As Integer)
    On Error GoTo ShowError
    Dim fh As Integer
    fh = FreeFile
    Open App.path & "/error.log" For Append As fh
    Print #fh, Format(Now, "dd-mm-yyyy hh:nn:ss") _
   & ErrNumber & " " & ErrDesc & " - " & strTitle & " Line Number: " & iLineNu
    Close fh
    
    Screen.MousePointer = vbDefault
    'MsgBox "Error: " & ErrNumber & vbNewLine & ErrDesc, vbOKOnly + vbExclamation  ', strTitle
    Exit Sub

ShowError:
    MsgBox "Error: " & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbExclamation   ', "Show Error"
   
    Exit Sub
End Sub

