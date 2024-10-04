Option Compare Database

Public Function IsFileOpen(FileName)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    End Select
End Function