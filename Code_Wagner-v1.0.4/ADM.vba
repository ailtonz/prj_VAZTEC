Option Compare Database
Option Explicit

Public Inicio As String
Public Final As String


'Function GerarParcelamento(codPedido As Long, dtEmissao As Date, ValParcelado As Currency, Parcelamento As String)
'
'Dim matriz As Variant
'Dim x As Integer
'Dim Parcelas As DAO.Recordset
'
'Set Parcelas = CurrentDb.OpenRecordset("Select * from PedidosPagamentos")
'
'matriz = Array()
'matriz = Split(Parcelamento, ";")
'
'BeginTrans
'
'For x = 0 To UBound(matriz)
'    Parcelas.AddNew
'    Parcelas.Fields("codPedido") = codPedido
'    Parcelas.Fields("Vencimento") = CalcularVencimento2(dtEmissao, CInt(matriz(x)))
'    Parcelas.Fields("Valor") = ValParcelado / (UBound(matriz) + 1)
'    Parcelas.Update
'Next
'
'CommitTrans
'
'Parcelas.Close
'
'End Function

