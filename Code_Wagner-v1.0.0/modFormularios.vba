Option Compare Database

Public strTabela As String

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Function RedimencionaControle(frm As Form, ctl As Control)

Dim intAjuste As Integer
On Error Resume Next

intAjuste = frm.Section(acHeader).Height * frm.Section(acHeader).Visible

intAjuste = intAjuste + frm.Section(acFooter).Height * frm.Section(acFooter).Visible

On Error GoTo 0

intAjuste = Abs(intAjuste) + ctl.top

If intAjuste < frm.InsideHeight Then
    ctl.Height = frm.InsideHeight - intAjuste
'    ctl.Width = frm.InsideHeight + (intAjuste + intAjuste)
End If

End Function

Public Function EstaAberto(strName As String) As Boolean
On Error GoTo EstaAberto_Err
' Testa se o formulário está aberto

   Dim obj As AccessObject, dbs As Object
   Set dbs = Application.CurrentProject
   ' Procurar objetos AccessObject abertos na coleção AllForms.
   
   EstaAberto = False
   For Each obj In dbs.AllForms
        If obj.IsLoaded = True And obj.Name = strName Then
            ' Imprimir nome do obj.
            EstaAberto = True
            Exit For
        End If
   Next obj
    
EstaAberto_Fim:
  Exit Function
EstaAberto_Err:
  Resume EstaAberto_Fim
End Function

Public Function IsFormView(frm As Form) As Boolean
On Error GoTo IsFormView_Err
' Testa se o formulário está aberto em
' modo formulário (form view)

 IsFormView = False
 If frm.CurrentView = 1 Then
    IsFormView = True
 End If

IsFormView_Fim:
  Exit Function
IsFormView_Err:
  Resume IsFormView_Fim
End Function

