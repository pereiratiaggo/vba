'Utilio estas funcoes em meu excel para preencher rapidamente as horas com virugulas, converte-las para decimais, somar e ai converter para horas para lançamento em sistemas contabeis de folha

'cdh = Converter Decimal em Horas
Function cdh(valor_decimal)

    negativo = 0:
    If valor_decimal < 0 Then
        negativo = 1:
        valor_decimal = valor_decimal * -1:
    End If

    valor_dec = valor_decimal:
    valor_int = Application.WorksheetFunction.RoundDown(valor_decimal, 0):
    valor = valor_int + ((valor_dec - valor_int) * 60 / 100):

    If negativo = 1 Then
        valor = valor * -1:
    End If

    cdh = valor:

End Function

'chd = Converter Horas Em Decimal
Function chd(valor_hora_seprado_por_virgura)

    negativo = 0:
    If valor_decimal < 0 Then
        negativo = 1:
        valor_hora_seprado_por_virgura = valor_hora_seprado_por_virgura * -1:
    End If

    valor_hora = valor_hora_seprado_por_virgura:
    valor_int = Int(valor_hora):
    valor = (((valor_hora - valor_int) * 100) / 60) + valor_int:

    If negativo = 1 Then
        valor = valor * -1:
    End If

    chd = valor:

End Function

Function converterDecimalEmHoras(valor_decimal)
    converterDecimalEmHoras = cdh(valor_decimal)
End Function

Function ConverterDecimalEmHoras(valor_decimal)
    ConverterDecimalEmHoras = cdh(valor_decimal)
End Function

Function converterHorasEmDecimal(valor_decimal)
    converterHorasEmDecimal = chd(valor_decimal)
End Function

Function ConverterHorasEmDecimal(valor_decimal)
    ConverterHorasEmDecimal = chd(valor_decimal)
End Function