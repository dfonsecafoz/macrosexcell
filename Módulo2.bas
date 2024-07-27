Attribute VB_Name = "M�dulo2"
' Calcula a diferen�a de tempo entre a coluna B e a coluna C. Considera os Sabados e Domingos no calculo. O Resultado est� no formato [h]:mm:ss
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
    
    ' Definir a planilha que est� ativa
    Set ws = ThisWorkbook.Sheets("page1") ' Altere "pagina1" para o nome da sua planilha
    
    ' Encontrar a �ltima linha com dados na coluna B
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Loop atrav�s de cada linha e calcular a diferen�a de tempo
    For i = 2 To ultimaLinha
        startTime = ws.Cells(i, 2).Value ' Coluna B: "Interaction start"
        endTime = ws.Cells(i, 3).Value ' Coluna C: "Resolved"
        
        totalSegundos = 0
        
        ' Ajustar o hor�rio de in�cio e t�rmino para dentro do hor�rio de trabalho
        Do While startTime < endTime
            startDay = DateValue(startTime)
            workStart = startDay + TimeValue("07:00:00")
            workEnd = startDay + TimeValue("17:00:00")
            
            If startTime < workStart Then
                startTime = workStart
            End If
            
            If startTime >= workEnd Then
                startTime = workStart + 1 ' Avan�ar para o pr�ximo dia �til
            End If
            
            If endTime > workEnd Then
                diferenca = workEnd - startTime
                startTime = workStart + 1 ' Avan�ar para o pr�ximo dia �til
            Else
                diferenca = endTime - startTime
                startTime = endTime
            End If
            
            ' Calcular a diferen�a de tempo em segundos
            If diferenca > 0 Then
                totalSegundos = totalSegundos + diferenca * 86400 ' Converter dias para segundos
            End If
        Loop
        
        ' Formatar a diferen�a em horas, minutos e segundos
        diferencaFormatada = Format(Int(totalSegundos / 3600), "00") & ":" & Format(Int((totalSegundos Mod 3600) / 60), "00") & ":" & Format(Int(totalSegundos Mod 60), "00")
        
        ws.Cells(i, 4).Value = diferencaFormatada ' Coluna D: Resultado
    Next i
    
    MsgBox "Diferen�a de tempo calculada com sucesso!", vbInformation
End Sub
