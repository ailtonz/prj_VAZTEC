Option Compare Database
Option Explicit

Function ImpressoraPadrao(prtDefault As String)

Dim XPrint     As Printer
Dim n          As Integer
  
'Busca o numero da impresora
For Each XPrint In Printers
    If XPrint.DeviceName = prtDefault Then
       Exit For
    End If
    n = n + 1
Next
 
'Efetiva a impressora como padrao
Set Application.Printer = Application.Printers(n)
 
End Function

Public Function ExecutarSQL(strSQL As String)
'Objetivo: Executar comandos SQL sem mostrar msg's do access.

'Desabilitar menssagens de execução de comando do access
DoCmd.SetWarnings False

'Saida strSQL

'Executar a instrução SQL
DoCmd.RunSQL strSQL

'Abilitar menssagens de execução de comando do access
DoCmd.SetWarnings True

End Function

Public Function GerarSaida(strConteudo As String, strArquivo As String)

Open Application.CurrentProject.Path & "\" & strArquivo For Append As #1

Print #1, strConteudo

Close #1

End Function

Public Function CaminhoDoBanco() As String
Dim Arq As String
Dim Caminho As String

Arq = "caminho.log"
Caminho = Application.CurrentProject.Path & "\" & Arq

'Verifica a existencia do caminho do banco de dados
If VerificaExistenciaDeArquivo(Caminho) Then
    CaminhoDoBanco = getCaminho(Application.CurrentProject.Path & "\" & Arq)
Else
    MsgBox "ATENÇÃO: Não é possível localizar o caminho do Banco de dados.", vbExclamation + vbOKOnly, "Caminho do Banco de Dados"
    CaminhoDoBanco = ""
End If

End Function

Public Function LocalizarBanco(Banco As String) As String

    'Verifica a existencia do banco de dados no caminho informado
    If VerificaExistenciaDeArquivo(Banco) Then
        LocalizarBanco = Banco
    Else
        MsgBox "ATENÇÃO: Não é possível localizar o Banco de dados.", vbExclamation + vbOKOnly, "Localiza Banco De Dados"
        LocalizarBanco = ""
    End If

End Function

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function

Public Function Categoria(strCategoria As String) As String

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("Select * from Categorias where Principal = 0 and Categoria = '" & strCategoria & "'")
If Not rstTabela.EOF Then
    Categoria = rstTabela.Fields("Descricao01")
Else
   Categoria = ""
End If
rstTabela.Close

End Function

Public Function Zebrar(rpt As Report)
Static fCinza As Boolean
Const conCinza = 15198183
Const conBranco = 16777215

On Error Resume Next

    rpt.Section(0).BackColor = IIf(fCinza, conCinza, conBranco)
    fCinza = Not fCinza

End Function

Public Sub testEtiqueta()

MsgBox right("admAilton", Len("admAilton") - 3)

End Sub
