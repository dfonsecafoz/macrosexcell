Attribute VB_Name = "M�dulo1"
' Calcula a diferen�a de tempo entre a coluna B e a coluna C. O Resultado est� no formato amig�vel por ex: 5 d, 21 h, 39 m, 01 s
' Possibilitade de inserir os dias de feriados na coluna E
Sub CalcularDiferencaTempoComHorarioModulo1()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim startTime As Date
    Dim endTime As Date
    Dim diferenca As Double
    Dim startDay As Date
    Dim workStart As Date
    Dim workEnd As Date
    Dim totalSegundos As Double
    Dim diferencaFormatada As String
    Dim feriados As Range
    Dim celula As Range
    Dim ehFeriado As Boolean
    Dim totalHoras As Double
    Dim dias As Long
    Dim horas As Long
    Dim minutos As Long
    Dim segundos As Long
    
    ' Definir a planilha que est� ativa
    Set ws = ThisWorkbook.Sheets("Page1") ' Altere "Planilha3" para o nome da sua planilha
    
    ' Encontrar a �ltima linha com dados na coluna B
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Definir o intervalo de feriados (ajuste conforme necess�rio)
    Set feriados = ws.Range("E2:E20") ' Supondo que os feriados est�o na coluna E, de E2 at� E20
    
    ' Loop atrav�s de cada linha e calcular a diferen�a de tempo
    For i = 2 To ultimaLinha
        startTime = ws.Cells(i, 2).Value ' Coluna B: "Interaction start"
        endTime = ws.Cells(i, 3).Value ' Coluna C: "Resolved"
        
        totalSegundos = 0
        
        ' Ajustar o hor�rio de in�cio e t�rmino para dentro do hor�rio de trabalho31
        Do While startTime < endTime
            startDay = DateValue(startTime)
            workStart = startDay + TimeValue("07:00:00")
            workEnd = startDay + TimeValue("17:00:00")
            
            ' Verificar se � final de semana ou feriado
            If Weekday(startDay, vbMonday) >= 6 Then
                ' Se for s�bado (6) ou domingo (7), pular para o pr�ximo dia �til
                startTime = startTime + 1
                GoTo ProximoDia
            End If
            
            ehFeriado = False
            For Each celula In feriados
                If celula.Value = startDay Then
                    ehFeriado = True
                    Exit For
                End If
            Next celula
            
            If ehFeriado Then
                ' Se for feriado, pular para o pr�ximo dia �til
                startTime = startTime + 1
                GoTo ProximoDia
            End If
            
            ' Ajustar o hor�rio de in�cio e t�rmino para o hor�rio de trabalho
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
            
ProximoDia:
        Loop
        
        ' Converter totalSegundos em dias, horas, minutos e segundos
        dias = Int(totalSegundos / 86400)
        totalSegundos = totalSegundos Mod 86400
        horas = Int(totalSegundos / 3600)
        totalSegundos = totalSegundos Mod 3600
        minutos = Int(totalSegundos / 60)
        segundos = totalSegundos Mod 60
        
        ' Formatar a diferen�a em dias, horas, minutos e segundos
        diferencaFormatada = dias & " d, " & Format(horas, "00") & " h, " & Format(minutos, "00") & " m, " & Format(segundos, "00") & " s"
        
        ws.Cells(i, 4).Value = diferencaFormatada ' Coluna D: Resultado
    Next i
    
    MsgBox "Diferen�a de tempo calculada com sucesso!", vbInformation
End Sub
