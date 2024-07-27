Attribute VB_Name = "Módulo2"
' Calcula a diferença de tempo entre a coluna B e a coluna C. Considera os Sabados e Domingos no calculo. O Resultado está no formato [h]:mm:ss
Sub CalcularDiferencaTempoComHorarioModulo2()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim startTime As Date
    Dim endTime As Date
    Dim diferenca As Double
    Dim startDay As Date
    Dim endDay As Date
    Dim workStart As Date
    Dim workEnd As Date
    Dim totalSegundos As Double
    Dim diferencaFormatada As String
    
    ' Definir a planilha que está ativa
    Set ws = ThisWorkbook.Sheets("page1") ' Altere "pagina1" para o nome da sua planilha
    
    ' Encontrar a última linha com dados na coluna B
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Loop através de cada linha e calcular a diferença de tempo
    For i = 2 To ultimaLinha
        startTime = ws.Cells(i, 2).Value ' Coluna B: "Interaction start"
        endTime = ws.Cells(i, 3).Value ' Coluna C: "Resolved"
        
        totalSegundos = 0
        
        ' Ajustar o horário de início e término para dentro do horário de trabalho
        Do While startTime < endTime
            startDay = DateValue(startTime)
            workStart = startDay + TimeValue("07:00:00")
            workEnd = startDay + TimeValue("17:00:00")
            
            If startTime < workStart Then
                startTime = workStart
            End If
            
            If startTime >= workEnd Then
                startTime = workStart + 1 ' Avançar para o próximo dia útil
            End If
            
            If endTime > workEnd Then
                diferenca = workEnd - startTime
                startTime = workStart + 1 ' Avançar para o próximo dia útil
            Else
                diferenca = endTime - startTime
                startTime = endTime
            End If
            
            ' Calcular a diferença de tempo em segundos
            If diferenca > 0 Then
                totalSegundos = totalSegundos + diferenca * 86400 ' Converter dias para segundos
            End If
        Loop
        
        ' Formatar a diferença em horas, minutos e segundos
        diferencaFormatada = Format(Int(totalSegundos / 3600), "00") & ":" & Format(Int((totalSegundos Mod 3600) / 60), "00") & ":" & Format(Int(totalSegundos Mod 60), "00")
        
        ws.Cells(i, 4).Value = diferencaFormatada ' Coluna D: Resultado
    Next i
    
    MsgBox "Diferença de tempo calculada com sucesso!", vbInformation
End Sub
