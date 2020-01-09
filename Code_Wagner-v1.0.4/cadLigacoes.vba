Option Compare Database
Option Explicit

Private Sub Cliente_Click()
    Me.codCadastro = Me.Cliente.Column(1)
End Sub

Private Sub Cliente_Enter()
Dim strObras As String: strObras = "SELECT cadGeralEnderecos.Endereco, cadGeralEnderecos.codEndereco, cadGeralEnderecos.codCadastro, cadGeral.Nome, cadGeralEnderecos.NF, cadGeralEnderecos.Valor FROM cadGeral INNER JOIN cadGeralEnderecos ON cadGeral.codCadastro = cadGeralEnderecos.codCadastro WHERE (((cadGeral.codCadastro)=Forms.Calendario!Ambiente.form!codCadastro))"

    Me.Cliente.Requery
    Me.Obra.RowSource = strObras
    Me.Obra.Requery
End Sub

Private Sub Cliente_Exit(Cancel As Integer)
    Me.Obra.Requery
    Me.Contato.Requery
End Sub

Private Sub Cliente_NotInList(NewData As String, Response As Integer)
'Permite adicionar a editora à lista
Dim DB As DAO.Database
Dim rst As DAO.Recordset

On Error GoTo ErrHandler

'Pergunta se deseja acrescentar o novo item
If Confirmar("O Cliente: " & NewData & "  não faz parte da lista." & vbCrLf & "Deseja acrescentá-lo?") = True Then
    Set DB = CurrentDb()
    'Abre a tabela, adiciona o novo item e atualiza a combo
    Set rst = DB.OpenRecordset("cadGeral")
    With rst
        .AddNew
        !codCadastro = NovoCodigo("cadGeral", "codCadastro")
        !Nome = NewData
        .Update
        Response = acDataErrAdded
        .Close
    End With
        
    DoCmd.OpenForm "cadGeral", , , "Nome = '" & NewData & "'"
    
Else
    Response = acDataErrDisplay
End If

ExitHere:
Set rst = Nothing
Set DB = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description & vbCrLf & Err.Number & _
vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
Resume ExitHere

End Sub

Private Sub Contato_Exit(Cancel As Integer)
    Me.Contato = UCase(Me.Contato)
End Sub

Private Sub Form_AfterInsert()
    Me.Data = [Forms]![Calendario]![Cal].[Value]
End Sub

Private Sub Obra_Click()
    
    Me.codObra = Me.Obra.Column(1)
    Me.codCadastro = Me.Obra.Column(2)
    Me.Cliente = Me.Obra.Column(3)
    Me.NF = Me.Obra.Column(4)
    Me.Valor = Me.Obra.Column(5)
    
End Sub

Private Sub Obra_Enter()
Dim strObras As String: strObras = "SELECT cadGeralEnderecos.Endereco, cadGeralEnderecos.codEndereco, cadGeralEnderecos.codCadastro, cadGeral.Nome, cadGeralEnderecos.NF FROM cadGeral INNER JOIN cadGeralEnderecos ON cadGeral.codCadastro = cadGeralEnderecos.codCadastro"

If IsNull(Me.Cliente.Value) Then
    Me.Obra.RowSource = strObras
    Me.Obra.Requery
End If

End Sub

Private Sub OBS_Exit(Cancel As Integer)
    
    Me.OBS = UCase(Me.OBS)

End Sub

Private Sub C_Exit(Cancel As Integer)
    
    If Me.C <> 0 Then
        Me.DT_C = Forms!Calendario!Cal.Value
        Me.Data = Forms!Calendario!Cal.Value
    End If

End Sub

Private Sub R_Exit(Cancel As Integer)
    
    If Me.R <> 0 Then
        Me.DT_R = Forms!Calendario!Cal.Value
        Me.Data = Forms!Calendario!Cal.Value
    End If

End Sub

Private Sub T_Exit(Cancel As Integer)
    
    If Me.T <> 0 Then
        Me.DT_T = Forms!Calendario!Cal.Value
        Me.Data = Forms!Calendario!Cal.Value
    End If

End Sub
