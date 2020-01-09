Option Compare Database
Option Explicit

Private Sub cmdAtualizar_Click()
'Atualiza o v�nculo das tabelas
Dim fd As Office.FileDialog
Dim strArq As String
Dim varItem As Variant
Dim strTabela As String
Dim Banco As String

Dim cnn As ADODB.Connection
Dim cat As Object
Dim tbl As Object
Dim MyApl As String

On Error GoTo ErrHandler

MyApl = Application.CurrentProject.Path


If VerificaExistenciaDeArquivo(MyApl & "\caminho.log") Then Kill MyApl & "\caminho.log"

Call cmdSelecionar_Click

Banco = LocalizarBanco(CaminhoDoBanco)

If Banco <> "" Then

    strArq = Banco
    
    'Banco de dados atual
    Set cnn = CurrentProject.Connection
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = cnn
    
    'Percorre os itens da listbox
    For Each varItem In Me.lstTabelas.ItemsSelected
        strTabela = Me.lstTabelas.Column(1, varItem)
        On Error Resume Next
        'Define o novo v�nculo
        Set tbl = cat.Tables(strTabela)
        tbl.Properties("Jet OLEDB:Link Datasource") = strArq
        
        'Se houver erro, avisa
        If Not Err = 0 Then
            MsgBox "Erro ao vincular " & tbl.Name
            Err.Clear
        End If
        
    Next varItem
    
    'Atualiza a listbox
    Call PreencheLista
    MsgBox "Conex�o estabelecida com sucesso!!!", vbOKOnly + vbInformation, "Conex�o com banco de dados"
    
Else

    'Di�logo de selecionar arquivo - Office
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.Filters.Add "BDs do Access", "*.MDB;*.MDE"
    fd.TITLE = "Localize a fonte de dados"
    fd.AllowMultiSelect = False
    
    ''''''''''''''''''''''''''''
    'CAMINHO DO BANCO DE DADOS
    ''''''''''''''''''''''''''''
    
    If fd.Show = -1 Then
        strArq = fd.SelectedItems(1)
        GerarSaida strArq, "caminho.log"
    End If
    
    'Se selecionou arquivo, atualiza os v�nculos
    If strArq <> "" Then
    
        'Banco de dados atual
        Set cnn = CurrentProject.Connection
        Set cat = CreateObject("ADOX.Catalog")
        cat.ActiveConnection = cnn
       
        'Percorre os itens da listbox
        For Each varItem In Me.lstTabelas.ItemsSelected
            strTabela = Me.lstTabelas.Column(1, varItem)
            On Error Resume Next
            'Define o novo v�nculo
            Set tbl = cat.Tables(strTabela)
            tbl.Properties("Jet OLEDB:Link Datasource") = _
            strArq
            'Se houver erro, avisa
            If Not Err = 0 Then
                MsgBox "Erro ao vincular " & tbl.Name
                Err.Clear
            End If
        Next varItem
        'Atualiza a listbox
        Call PreencheLista
        MsgBox "Conex�o estabelecida com sucesso!!!", vbOKOnly + vbInformation, "Conex�o com banco de dados"
        
    Else
        MsgBox "ATEN��O: N�o foi informado o caminho do banco de dados !!!", vbOKOnly + vbExclamation, "Conex�o com banco de dados"
        GoTo ExitHere
        Exit Sub
    
    End If


End If

ExitHere:
'Libera a mem�ria
Set tbl = Nothing
Set cat = Nothing
Set cnn = Nothing

DoCmd.Close

Exit Sub

ErrHandler:
MsgBox Err.Description
Resume ExitHere

End Sub

Private Sub cmdLimpar_Click()
'Limpa a sele��o
Dim I As Integer

    For I = 0 To Me.lstTabelas.ListCount
        Me.lstTabelas.Selected(I) = False
    Next I
    Call AtivaBotao
    
End Sub

Private Sub cmdSelecionar_Click()
'Seleciona todos os itens da listbox
Dim I As Integer

    For I = 0 To Me.lstTabelas.ListCount
        Me.lstTabelas.Selected(I) = True
    Next I
    Call AtivaBotao

End Sub

Private Sub cmdVerificar_Click()
    Call PreencheLista
End Sub

Private Sub Form_Open(Cancel As Integer)
    Call PreencheLista
    cmdAtualizar_Click
End Sub

Private Sub lstTabelas_Click()
    Call AtivaBotao
End Sub

Sub AtivaBotao()
'Ativa ou desativa bot�o de atualizar v�nculos
If Me.lstTabelas.ItemsSelected.Count > 0 Then
    Me.cmdAtualizar.Enabled = True
Else
    Me.cmdVerificar.SetFocus
    Me.cmdAtualizar.Enabled = False
End If
End Sub

Sub PreencheLista()
Dim cnn As ADODB.Connection
Dim cat As Object
Dim tbl As Object

Dim strLista As String 'origem de linha
Dim strSource As String 'path do BD
Dim strStatus As String 'OK ou !

On Error GoTo ErrHandler

'Banco de dados atual
Set cnn = CurrentProject.Connection
Set cat = CreateObject("ADOX.Catalog")
cat.ActiveConnection = cnn

'Percorre todas as tabelas
For Each tbl In cat.Tables
    'Se for vinculada, inclui na listbox
    If tbl.Type = "LINK" Then
        'Armazena o data source
        strSource = _
        tbl.Properties("Jet OLEDB:Link Datasource")
        'Verifica se o link est� OK
        On Error Resume Next
        tbl.Properties("Jet OLEDB:Link Datasource") = _
        strSource
        'Verifica se ocorreu erro
        If Err = 0 Then
            'Se n�o ocorreu erro, OK
            strStatus = "OK"
        Else
            'Se ocorreu erro, exclama��o "!"
            strStatus = "!"
            Err.Clear
        End If
        'Tr�s colunas: status, nome da tabela, endere�o
        strLista = strLista & strStatus & ";" & tbl.Name _
        & ";" & strSource & ";"
    End If
Next tbl

'Origem da listbox
Me.lstTabelas.RowSource = strLista

Call AtivaBotao

ExitHere:
'Libera a mem�ria
Set tbl = Nothing
Set cat = Nothing
Set cnn = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description
Resume ExitHere
End Sub

