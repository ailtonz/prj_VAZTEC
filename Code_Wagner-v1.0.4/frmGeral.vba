Option Compare Database

Private Sub cmdCadastro_Click()
On Error GoTo Err_cmdCadastro_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frmCadastro"
    
    stLinkCriteria = "[codCadastro]=" & Me![codCadastro]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCadastro_Click:
    Exit Sub

Err_cmdCadastro_Click:
    MsgBox Err.Description
    Resume Exit_cmdCadastro_Click
    
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub
Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub
Private Sub cmdCliente_Click()
On Error GoTo Err_cmdCliente_Click


    DoCmd.GoToRecord , , acNewRec

Exit_cmdCliente_Click:
    Exit Sub

Err_cmdCliente_Click:
    MsgBox Err.Description
    Resume Exit_cmdCliente_Click
    
End Sub
