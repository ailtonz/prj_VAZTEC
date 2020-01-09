Option Compare Database

Public Function AbrirArquivo(sTitulo As String, sDecricao As String, sTipo As String, SelecaoMultipla As Boolean) As String
Dim fd As Office.FileDialog

'Diálogo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'Título
fd.TITLE = sTitulo

'Filtros e descrição dos mesmos
fd.Filters.Add sDecricao, sTipo

'Premissões de selação
fd.AllowMultiSelect = SelecaoMultipla

If fd.Show = -1 Then
    AbrirArquivo = fd.SelectedItems(1)
End If

End Function

Public Function Confirmar(sMensagem As String) As _
Boolean
'Faz uma pergunta ao usuário e retorma True se a
'resposta for SIM, e false se a resposta for NÃO
Dim intResp As Integer

intResp = MsgBox(sMensagem, vbYesNo + vbQuestion, _
"Confirmação")

If intResp = vbYes Then
    Confirmar = True
Else
    Confirmar = False
End If
End Function


