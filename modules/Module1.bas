Option Compare Database

Public Sub Evaluate()


If Not IsNull(Forms!Frm_Families_And_Objectives.Requirement_Special_Considerations_Score.Value) Then
        Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Enabled = False
        Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = False
        Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = False
        Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
        Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = True
        Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = True
        f = 1
        Set rs = Forms!Frm_Families_And_Objectives!SubFrm_Objectives.Form.RecordsetClone
        rs.MoveFirst
        Do While Not rs.EOF
        cbothereval = rs!Objective_Other_Than_Satisfied
        cbeval = rs!Objective_Satisfied
        NotDoneYet = cbeval & " " & cbothereval
      
        If cbothereval = "True" Then
            f = 2
            
                Forms!Frm_Families_And_Objectives.Requirement_Satisfied = False
                Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Enabled = False
                Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
                Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = True
                Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = True
                Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = False
                Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = False
                If Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = True Or Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied = True Then
                    Exit Do
                Else
                   If Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = False And Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied = False Then
                      x = MsgBox("You must manually select Other Than Satisfied for the full 5 point deduction, or Other than Satisfied with Special Considerations for the 3 point deduction", vbOKOnly)
                      f = 2
                      Exit Do
                   End If
                End If
        End If
        
        If NotDoneYet = "False False" Then
            NDY = "Not Set"
        End If
        If NDY = "Not Set" Then
            Forms!Frm_Families_And_Objectives.Requirement_Satisfied = False
            Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Enabled = False
            Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
            Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = True
            Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = True
            Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = False
            Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = False
            f = 0
            
        Else
            If Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = True Or Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied = True Then
                ' Exit Do
                
            Else
                Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = False
                Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = False
                Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
                Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = True
                Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = True
            End If
        End If
rs.MoveNext
Loop
If f = 1 Then
    Forms!Frm_Families_And_Objectives.Requirement_Satisfied = True
    Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = False
    Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied = False
    Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = False
    Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = False
    Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
    Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = True
    Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = True
    
    
Else
    If f = 0 Then
       Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Enabled = False
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = False
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied = False
       Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = True
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = True
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = False
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = False
    Else
       Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Enabled = False
       Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = True
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = True
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = False
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = False
    End If
End If

Else
Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Locked = True
Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Locked = True
Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Locked = True
Set rs = Forms!Frm_Families_And_Objectives!SubFrm_Objectives.Form.RecordsetClone
rs.MoveFirst
Do While Not rs.EOF
    cbothereval = rs!Objective_Other_Than_Satisfied
    cbeval = rs!Objective_Satisfied
    NotDoneYet = cbeval & " " & cbothereval
    
    If cbothereval = "True" Then
        Forms!Frm_Families_And_Objectives.Requirement_Satisfied = False
        Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = True
        Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied = False
        Forms!Frm_Families_And_Objectives.Requirement_Satisfied.Enabled = False
        Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied.Enabled = False
        Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = False
        Exit Do
    End If
    
    If NotDoneYet = "False False" Then
         NDY = "Not Set"
    End If
    
    If NDY = "Not Set" Then
       Forms!Frm_Families_And_Objectives.Requirement_Satisfied = False
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = False
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = False
    Else
       Forms!Frm_Families_And_Objectives.Requirement_Satisfied = True
       Forms!Frm_Families_And_Objectives.Requirement_Not_Satisfied = False
       Forms!Frm_Families_And_Objectives.Requirement_SC_Satisfied.Enabled = True
    End If
rs.MoveNext
Loop

Set cbothereval = Nothing
Set cbeval = Nothing
Set rs = Nothing
Set NDY = Nothing
End If

End Sub