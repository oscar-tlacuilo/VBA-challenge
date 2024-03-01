Sub stock_analysis()

    ' Desactivar actualizaciones automáticas y cálculos automáticos para mejorar el rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Variables para el análisis
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim uniqueTickers As Collection
    Dim ticker As Variant
    Dim yearlyChange As Double
    Dim openPrice As Double
    Dim totalVolume As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Inicializar la colección para almacenar tickers únicos
    Set uniqueTickers = New Collection

    ' Bucle a través de todas las hojas de trabajo
    For Each ws In ThisWorkbook.Sheets
        ' Seleccionar la hoja de trabajo actual
        ws.Activate

        ' Encontrar la última fila con datos en la columna A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Bucle a través de los datos y agregar tickers únicos a la colección
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            On Error Resume Next
            uniqueTickers.Add ticker, CStr(ticker)
            On Error GoTo 0
        Next i

        ' Copiar tickers únicos a la columna I
        For i = 1 To uniqueTickers.Count
            ws.Range("I" & i + 1).Value = uniqueTickers.Item(i)
        Next i

        ' Agregar título a la columna I
        ws.Range("I1").Value = "Ticker"

        ' Establecer encabezados en la fila 1
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Establecer encabezados para la tabla en las columnas O:Q
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Inicializar variables para seguimiento de los mayores valores
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""

        ' Calcular y formatear Yearly Change, Percent Change y Total Stock Volume
        For i = 2 To uniqueTickers.Count + 1
            ' Calcular Yearly Change
            yearlyChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value ' Suponiendo close - open

            ' Calcular Precio de apertura
            openPrice = ws.Cells(i, 3).Value ' Precio de apertura

            ' Calcular Volumen total
            totalVolume = ws.Cells(i, 7).Value ' Obtener el volumen directamente

            ' Calcular Cambio porcentual
            If openPrice <> 0 Then
                percentChange = yearlyChange / Abs(openPrice) * 100
            Else
                percentChange = 0
            End If

            ' Agregar valores a las celdas correspondientes
            ws.Range("J" & i).Value = yearlyChange
            ws.Range("K" & i).Value = percentChange
            ws.Range("L" & i).Value = totalVolume

            ' Formatear color en la columna J según el cambio anual
            If yearlyChange < 0 Then
                ws.Range("J" & i).Interior.Color = RGB(255, 0, 0) ' Rojo
            ElseIf yearlyChange > 0 Then
                ws.Range("J" & i).Interior.Color = RGB(0, 255, 0) ' Verde
            Else
                ws.Range("J" & i).Interior.Color = RGB(255, 255, 255) ' Blanco
            End If

            ' Formatear la columna K como porcentaje con dos decimales
            ws.Range("K" & i).NumberFormat = "0.00%"

            ' Verificar si es el mayor % de aumento
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ws.Cells(i, 1).Value
            End If

            ' Verificar si es el mayor % de disminución
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ws.Cells(i, 1).Value
            End If

            ' Verificar si es el mayor volumen total
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ws.Cells(i, 1).Value
            End If
        Next i

        ' Agregar resultados para las columnas O:Q
        ws.Range("P2").Value = "Ticker"
        ws.Range("Q2").Value = "Value"
        
        ' Loop para los resultados de "Greatest % Increase", "Greatest % Decrease", y "Greatest Total Volume"
        For j = 2 To 4
            If j = 2 Then
                ws.Range("P" & j).Value = greatestIncreaseTicker
                ws.Range("Q" & j).Value = greatestIncrease
            ElseIf j = 3 Then
                ws.Range("P" & j).Value = greatestDecreaseTicker
                ws.Range("Q" & j).Value = greatestDecrease
            ElseIf j = 4 Then
                ws.Range("P" & j).Value = greatestVolumeTicker
                ws.Range("Q" & j).Value = greatestVolume
            End If
        Next j

        ' Reactivar actualizaciones automáticas y cálculos automáticos para la siguiente hoja de trabajo
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    Next ws

End Sub

