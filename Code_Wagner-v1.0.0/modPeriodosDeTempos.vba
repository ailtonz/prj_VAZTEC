Option Compare Database

Public Function CalcularVencimento(Dia As Integer, Optional MES As Integer, Optional Ano As Integer) As Date

If Month(Now) = 2 Then
    If Dia = 29 Or Dia = 30 Or Dia = 31 Then
        Dia = 1
        MES = MES + 1
    End If
End If

If MES > 0 And Ano > 0 Then
    CalcularVencimento = Format((DateSerial(Ano, MES, Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And Ano > 0 Then
    CalcularVencimento = Format((DateSerial(Ano, Month(Now), Dia)), "dd/mm/yyyy")
ElseIf MES = 0 And Ano = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), Dia)), "dd/mm/yyyy")
End If

End Function

Public Function CalcularVencimento2(dtInicio As Date, qtdDias As Integer, Optional ForaMes As Boolean) As Date

    If ForaMes Then
        CalcularVencimento2 = Format((DateSerial(Year(dtInicio), Month(dtInicio) + 1, qtdDias)), "dd/mm/yyyy")
    Else
        CalcularVencimento2 = Format((DateSerial(Year(dtInicio), Month(dtInicio), Day(dtInicio) + qtdDias)), "dd/mm/yyyy")
    End If

End Function
