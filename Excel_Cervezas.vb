Private Sub CommandButton1_Click()

    Dim ppAPP As PowerPoint.Application
    Dim ppPresentacion As PowerPoint.Presentation
    Dim rngTexto As Range
    On Error GoTo ErrorHandler
    Dim strTexto As String
    Dim rutaPlantilla As String
    Dim rutaActualExcel As String
    Dim rutaActualPowerP As String
    Dim nomActual As String
    Dim exAPP As Workbook
    Dim archivoCompleto As String
    Dim chartDataWorkbook As Object
    Dim chartWorksheet As Object
    Dim pptTable As Object
    Dim pptShape As Object
    Dim pptSlide As Object ' Objeto de la diapositiva de PowerPoint
    Dim rngDatos As Range ' Rango de datos en Excel
    Dim sldIndex As Integer ' Índice de la diapositiva
    Dim i As Integer ' Variable de iteración
    Dim excelData As Variant
    Dim cell As Range
    Dim number As Integer
    Dim chart As ChartObject
    Dim serie As Series
    Dim point As point
    
    
    
    
    rutaActualExcel = Application.GetOpenFilename(Title:="Seleccione el Excel de los Datos")
    
    If rutaActualExcel = "False" Then
        Exit Sub
        
    End If
    
    rutaActualPowerP = Application.GetOpenFilename(Title:="Seleccione el Template de PowerPoint")
    
     If rutaActualPowerP = "False" Then
        Exit Sub
        
    End If
    
   Set exAPP = Workbooks.Open(rutaActualExcel)

If ppAPP Is Nothing Then Set ppAPP = New PowerPoint.Application

Set ppPresentacion = ppAPP.Presentations.Open(rutaActualPowerP)

If ppPresentacion Is Nothing Then
   MsgBox exAPP.Name & ppPresentacion.Name & "No se encuentra en la ruta."
Else
   
   'ESTA ES LA HOJA 3

   'Actualizar la diapositiva 9 con datos de la hoja "Perfil de imagen" en Excel
        ' Actualizar textos

        Set rngTexto = exAPP.Sheets("Tipología").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(3).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar texto "Base"

        Set rngTexto = exAPP.Sheets("Tipología").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(3).Shapes("Valor_Base").TextFrame.TextRange.Text = strTexto

   'Graficos

    sldIndex = 3
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 15")

    If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Tipología").Range("B6:C9")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
            For j = 1 To UBound(excelData, 2)
                chartWorksheet.Cells(i, j).Value = excelData(i, j)
            Next j
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    'ESTA ES LA HOJA 4
    'Actualizar textos


    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("A1")

    strTexto = rngTexto.Value

    ppPresentacion.Slides(4).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

    'Actualizar texto "Base"

    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("C6")

    strTexto = rngTexto.Value

    ppPresentacion.Slides(4).Shapes("TextBox 2").TextFrame.TextRange.Text = strTexto

    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("C6")

    number = Int(rngTexto.Value)

    strTexto = number


    'Textos en tabla

    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("B7:B15")

    excelData = rngTexto.Value

    sldIndex = 4
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 16").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'Graficos

    'Chart 1
    sldIndex = 4
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 5")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Conocimiento de marca").Range("C6:C15")

    On Error GoTo 0

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
                chartWorksheet.Cells(i, 2).Value = excelData(i, 1)

        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

        End If

        'Chart 2

    sldIndex = 4
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 9")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Conocimiento de marca").Range("D6:D15")

    On Error GoTo 0

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
                chartWorksheet.Cells(i, 2).Value = excelData(i, 1)

        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

        End If

        'Chart 3

    sldIndex = 4
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 10")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Conocimiento de marca").Range("E6:E15")

    On Error GoTo 0

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
                chartWorksheet.Cells(i, 2).Value = excelData(i, 1)

        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

        End If

    'ESTA ES LA PAGINA 5

    ' Actualizar la diapositiva 5 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Funnels").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Funnels").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 2").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 3").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("F6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 4").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("G6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("H6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 6").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("I6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 7").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("J6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 8").TextFrame.TextRange.Text = strTexto

        'Actualizar base
        Set rngTexto = exAPP.Sheets("Funnels").Range("D8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("TextBox 3").TextFrame.TextRange.Text = strTexto

        'Table 1

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 25").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'Table 2

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 27").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO 1

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 8")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("C9:C12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 2

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 28")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("D9:D12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 29")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("E9:E12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 4

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 30")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("F9:F12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 31")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("G9:G12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 32")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("H9:H12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 33")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("I9:I12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 8

    sldIndex = 5
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 34")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("J9:J12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'ESTA ES LA PAGINA 6

    ' Actualizar la diapositiva 6 con datos de la hoja "Funnels" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Funnels").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Funnels").Range("K6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("L6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 2").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("M6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 3").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("N6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 4").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("O6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("P6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 6").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("Q6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 7").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("R6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Brand 8").TextFrame.TextRange.Text = strTexto

        'Actualizar base
        Set rngTexto = exAPP.Sheets("Funnels").Range("D8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 3").TextFrame.TextRange.Text = strTexto

        'Table 1

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 25").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'Table 2

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 27").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO 1

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 8")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("K9:K12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 2

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 28")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("L9:L12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 29")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("M9:M12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 4

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 30")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("N9:N12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 31")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("O9:O12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 32")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("P9:P12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 33")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("Q9:Q12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 8

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 34")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("R9:R12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'ESTA ES LA HOJA 7

    ' Actualizar la diapositiva 7 con datos de la hoja "Funnels" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Funnels").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Funnels").Range("S6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("T6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 2").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("U6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 3").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("V6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 4").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("W6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("X6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 6").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("Y6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 7").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("Z6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Brand 8").TextFrame.TextRange.Text = strTexto

        'Actualizar base
        Set rngTexto = exAPP.Sheets("Funnels").Range("D8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 3").TextFrame.TextRange.Text = strTexto

        'Table 1

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 25").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'Table 2

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 27").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO 1

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 8")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("S9:S12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 2

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 28")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("T9:T12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 29")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("U9:U12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 4

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 30")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("V9:V12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 31")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("W9:W12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 32")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("X9:X12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 33")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("Y9:Y12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 8

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 34")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("Z9:Z12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'ESTA ES LA HOJA 8

    ' Actualizar la diapositiva 8 con datos de la hoja "Funnels" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Funnels").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Funnels").Range("AA6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("AB6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 2").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("AC6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 3").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("AD6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 4").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("AE6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("AF6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 6").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("AG6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("Brand 7").TextFrame.TextRange.Text = strTexto

        'Actualizar base
        Set rngTexto = exAPP.Sheets("Funnels").Range("D8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(8).Shapes("TextBox 3").TextFrame.TextRange.Text = strTexto

        'Table 1

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 25").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'Table 2

    Set rngTexto = exAPP.Sheets("Funnels").Range("B9:B12")

    excelData = rngTexto.Value

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 27").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO 1

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 8")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AA9:AA12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 2

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 28")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AB9:AB12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 29")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AC9:AC12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 4

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 30")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AD9:AD12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 31")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AE9:AE12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 32")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AF9:AF12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 8
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 33")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Funnels").Range("AG9:AG12")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(2).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If
    
   ' ESTA LA HOJA 9


     'Actualizar la diapositiva 9 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 26").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 36").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 37").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("F6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 38").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("G6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 39").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("H6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 40").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("I6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 41").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("J6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(9).Shapes("TextBox 42").TextFrame.TextRange.Text = strTexto

        'Textos en tabla
        'TABLA 1 "BASE"

    Set rngTexto = exAPP.Sheets("Consideración de compra").Range("B8:J8")

    excelData = rngTexto.Value

    sldIndex = 9
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO

    sldIndex = 9
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 49")

    If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("B9:J14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
            For j = 1 To UBound(excelData, 2)
                chartWorksheet.Cells(i + 1, j).Value = excelData(i, j)
            Next j
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
    ' ESTA LA HOJA 10


     'Actualizar la diapositiva 10 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("K6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 26").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("L6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 36").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("M6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 37").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("N6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 38").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("O6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 39").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("P6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 40").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("Q6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 41").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("R6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(10).Shapes("TextBox 42").TextFrame.TextRange.Text = strTexto

        'Textos en tabla
        'TABLA 1 "BASE"

    Set rngTexto = exAPP.Sheets("Consideración de compra").Range("K8:R8")

    excelData = rngTexto.Value

    sldIndex = 10
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO

     sldIndex = 10
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 49")

    If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("B9:B14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
            For j = 1 To UBound(excelData, 2)
                chartWorksheet.Cells(i + 1, j).Value = excelData(i, j)
            Next j
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If


     sldIndex = 10
Set pptSlide = ppPresentacion.Slides(sldIndex)
Set pptShape = pptSlide.Shapes("Chart 49")

If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("K9:R14")

    excelData = rngDatos.Value

    Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
    chartDataWorkbook.Activate

    Set chartWorksheet = chartDataWorkbook.Worksheets(1)
    chartWorksheet.Columns("B:I").ClearContents

    For i = 1 To UBound(excelData, 1)
        For j = 1 To UBound(excelData, 2)
            chartWorksheet.Cells(i + 1, j + 1).Value = excelData(i, j) ' +1 en j para empezar en la columna B
        Next j
    Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    
      ESTA LA HOJA 11


     'Actualizar la diapositiva 11 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("S6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 26").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("T6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 36").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("U6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 37").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("V6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 38").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("W6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 39").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("X6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 40").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("Y6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 41").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("Z6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(11).Shapes("TextBox 42").TextFrame.TextRange.Text = strTexto

        'Textos en tabla
        'TABLA 1 "BASE"

    Set rngTexto = exAPP.Sheets("Consideración de compra").Range("S8:Z8")

    excelData = rngTexto.Value

    sldIndex = 11
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO

     sldIndex = 11
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 49")

    If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("B9:B14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
            For j = 1 To UBound(excelData, 2)
                chartWorksheet.Cells(i + 1, j).Value = excelData(i, j)
            Next j
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If


     sldIndex = 11
Set pptSlide = ppPresentacion.Slides(sldIndex)
Set pptShape = pptSlide.Shapes("Chart 49")

If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("S9:Z14")

    excelData = rngDatos.Value

    Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
    chartDataWorkbook.Activate

    Set chartWorksheet = chartDataWorkbook.Worksheets(1)
    chartWorksheet.Columns("B:S").ClearContents

    For i = 1 To UBound(excelData, 1)
        For j = 1 To UBound(excelData, 2)
            chartWorksheet.Cells(i + 1, j + 1).Value = excelData(i, j) ' +1 en j para empezar en la columna B
        Next j
    Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    ESTA ES LA HOJA 12
    
    'Actualizar la diapositiva 12 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AA6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 26").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AB6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 36").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AC6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 37").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AD6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 38").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AE6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 39").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AF6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 40").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AG6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(12).Shapes("TextBox 41").TextFrame.TextRange.Text = strTexto

        'Textos en tabla
        'TABLA 1 "BASE"

    Set rngTexto = exAPP.Sheets("Consideración de compra").Range("AA8:AG8")

    excelData = rngTexto.Value

    sldIndex = 12
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICO

     sldIndex = 12
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 49")

    If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("B9:B14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

        For i = 1 To UBound(excelData, 1)
            For j = 1 To UBound(excelData, 2)
                chartWorksheet.Cells(i + 1, j).Value = excelData(i, j)
            Next j
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If


     sldIndex = 12
Set pptSlide = ppPresentacion.Slides(sldIndex)
Set pptShape = pptSlide.Shapes("Chart 49")

If pptShape.HasChart Then

    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("AA9:AG14")

    excelData = rngDatos.Value

    Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
    chartDataWorkbook.Activate

    Set chartWorksheet = chartDataWorkbook.Worksheets(1)
    chartWorksheet.Columns("B:Y").ClearContents

    For i = 1 To UBound(excelData, 1)
        For j = 1 To UBound(excelData, 2)
            chartWorksheet.Cells(i + 1, j + 1).Value = excelData(i, j) ' +1 en j para empezar en la columna B
        Next j
    Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
    'ESTA ES LA HOJA 13

     ' Actualizar la diapositiva 13 con datos de la hoja "Perfil de imagen" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("F6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 15").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("G6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 16").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("H6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 17").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("I6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 18").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("J6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(13).Shapes("TextBox 19").TextFrame.TextRange.Text = strTexto

         ' Actualizar tablas en la diapositiva 13
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("B9:B23")
        excelData = rngTexto.Value
        sldIndex = 13
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 16").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

         'TABLA 2 "BASE"

    Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("C8:J8")

    excelData = rngTexto.Value

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICOS

    'GRAFICO 1

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 9")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("C9:C23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

     'GRAFICO 2

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 10")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("D9:D23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 11")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("E9:E23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

     'GRAFICO 4

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 12")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("F9:F23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 13")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("G9:G23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 20")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("H9:H23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 21")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("I9:I23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 8

    sldIndex = 13
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 22")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("J9:J23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If
    
    'ESTA ES LA HOJA 14

     ' Actualizar la diapositiva 14 con datos de la hoja "Perfil de imagen" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("K6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("L6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("M6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("N6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 15").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("O6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 16").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("P6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 17").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("Q6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 18").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("R6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(14).Shapes("TextBox 19").TextFrame.TextRange.Text = strTexto

         ' Actualizar tablas en la diapositiva 14
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("B9:B23")
        excelData = rngTexto.Value
        sldIndex = 14
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 16").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

         'TABLA 2 "BASE"

    Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("K8:R8")

    excelData = rngTexto.Value

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICOS

    'GRAFICO 1

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 9")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("K9:K23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

     'GRAFICO 2

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 10")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("L9:L23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 11")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("M9:M23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

     'GRAFICO 4

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 12")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("N9:N23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 13")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("O9:O23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 20")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("P9:P23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 21")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("Q9:Q23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 8

    sldIndex = 14
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 22")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("R9:R23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'ESTA ES LA HOJA 15

     ' Actualizar la diapositiva 15 con datos de la hoja "Perfil de imagen" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("S6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("T6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("U6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("V6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 15").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("W6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 16").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("X6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 17").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("Y6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 18").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("Z6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(15).Shapes("TextBox 19").TextFrame.TextRange.Text = strTexto

         ' Actualizar tablas en la diapositiva 15
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("B9:B23")
        excelData = rngTexto.Value
        sldIndex = 15
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 16").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

         'TABLA 2 "BASE"

    Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("S8:Z8")

    excelData = rngTexto.Value

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 5").Table

    If Not pptTable Is Nothing Then

        With pptTable

            Do While .Rows.Count < UBound(excelData, 1)
                .Rows.Add
            Loop
            Do While .Rows.Count > UBound(excelData, 1)
                .Rows(.Rows.Count).Delete
            Loop
            Do While .Columns.Count < UBound(excelData, 2)
                .Columns.Add
            Loop
            Do While .Columns.Count > UBound(excelData, 2)
                .Columns(.Columns.Count).Delete
            Loop

            For i = 1 To UBound(excelData, 1)
                For j = 1 To UBound(excelData, 2)
                    .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                Next j
            Next i
        End With

    Else

        MsgBox "El shape seleccionado no contiene una tabla."

    End If

    'GRAFICOS

    'GRAFICO 1

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 9")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("S9:S23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

     'GRAFICO 2

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 10")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("T9:T23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 3

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 11")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("U9:U23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

     'GRAFICO 4

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 12")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("V9:V23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 5

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 13")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("W9:W23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 6

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 20")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("X9:X23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 7

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 21")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("Y9:Y23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If

    'GRAFICO 8

    sldIndex = 15
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 22")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("Z9:Z23")

    On Error GoTo 0

    excelData = rngDatos.Value

        ' Referencia al workbook y worksheet del gráfico en PowerPoint
        Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)

        ' Limpiar solo las celdas de la columna B
        chartWorksheet.Columns(3).ClearContents

        ' Copiar los datos de Excel a la columna B del worksheet del gráfico
        For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close
    Else
        MsgBox "El shape especificado no contiene un gráfico.", vbExclamation
    End If
    
    'ESTA ES LA HOJA 17

     ' Actualizar la diapositiva 17 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos

        Set rngTexto = exAPP.Sheets("Branding").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(17).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de ADS

        Set rngTexto = exAPP.Sheets("Branding").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(17).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Branding").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(17).Shapes("TextBox 15").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Branding").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(17).Shapes("TextBox 16").TextFrame.TextRange.Text = strTexto


        ' Actualizar tablas en la diapositiva 16
        Set rngTexto = exAPP.Sheets("Branding").Range("B10:B14")
        excelData = rngTexto.Value

        sldIndex = 17
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 4").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

     'TABLA "BASE"

     Set rngTexto = exAPP.Sheets("Branding").Range("B8:E8")
        excelData = rngTexto.Value

        sldIndex = 17
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 26").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

     'GRAFICOS DEL 17
     'GRAFICO 1

      sldIndex = 17
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 14")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Branding").Range("C10:C14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

     'GRAFICO 2
        sldIndex = 17
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 15")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Branding").Range("D10:D14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    'GRAFICO 3

       sldIndex = 17
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 16")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Branding").Range("E10:E14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
        ESTA ES LA HOJA 18

     ' Actualizar la diapositiva 18 con datos de la hoja "Disfrute en Excel
        ' Actualizar textos

        Set rngTexto = exAPP.Sheets("Disfrute").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(18).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de ADS

        Set rngTexto = exAPP.Sheets("Disfrute").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(18).Shapes("TextBox 42").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Disfrute").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(18).Shapes("TextBox 43").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Disfrute").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(18).Shapes("TextBox 44").TextFrame.TextRange.Text = strTexto


        ' Actualizar tablas en la diapositiva 18
        Set rngTexto = exAPP.Sheets("Disfrute").Range("B10:B14")
        excelData = rngTexto.Value

        sldIndex = 18
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 4").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

     'TABLA "BASE"

     Set rngTexto = exAPP.Sheets("Disfrute").Range("B8:E8")
        excelData = rngTexto.Value

        sldIndex = 18
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 26").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

     'GRAFICOS DEL 18
     'GRAFICO 1

      sldIndex = 18
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 14")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Disfrute").Range("C10:C14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

     'GRAFICO 2
        sldIndex = 18
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 15")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Disfrute").Range("D10:D14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    'GRAFICO 3

       sldIndex = 18
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 16")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Disfrute").Range("E10:E14")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
    'ESTA ES LA HOJA 19

    'GRAFICO DE 19

    sldIndex = 19
Set pptSlide = ppPresentacion.Slides(sldIndex)
Set pptShape = pptSlide.Shapes("Chart 16")

If pptShape.HasChart Then

    On Error Resume Next

    Set rngDatos = exAPP.Sheets("Involucramiento").Range("B9:E20")

    On Error GoTo 0

    excelData = rngDatos.Value

    Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
    chartDataWorkbook.Activate

    Set chartWorksheet = chartDataWorkbook.Worksheets(1)
    chartWorksheet.Cells.Clear

    For i = 1 To UBound(excelData, 1)
        For j = 1 To UBound(excelData, 2)
            chartWorksheet.Cells(i + 1, j).Value = excelData(i, j) ' A partir de la segunda línea
        Next j
    Next i

    pptShape.chart.Refresh

    chartDataWorkbook.Close

Else
    MsgBox "El shape seleccionado no contiene un gráfico."
End If

  'Tabla de hoja 19

    Set rngTexto = exAPP.Sheets("Involucramiento").Range("B6:E20")
        excelData = rngTexto.Value
        sldIndex = 19
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 31").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

        'TABLA "BASE"

     Set rngTexto = exAPP.Sheets("Involucramiento").Range("C8:E8")
        excelData = rngTexto.Value
        sldIndex = 19
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 26").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If
    
     'ESTA ES LA HOJA 20

     ' Actualizar la diapositiva 20 con datos de la hoja "Disfrute" en Excel
        ' Actualizar textos

        Set rngTexto = exAPP.Sheets("Diferente").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(20).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de ADS

        Set rngTexto = exAPP.Sheets("Diferente").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(20).Shapes("TextBox 12").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Diferente").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(20).Shapes("TextBox 13").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Diferente").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(20).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto


     'Actualizar tabla 4

     Set rngTexto = exAPP.Sheets("Diferente").Range("B10:B13")
        excelData = rngTexto.Value
        sldIndex = 20
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 4").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If

        'TABLA 2 "BASE"

     Set rngTexto = exAPP.Sheets("Diferente").Range("B8:E8")
        excelData = rngTexto.Value
        sldIndex = 20
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 26").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If


     'Graficos de la Hojas 20
     'GRAFICO 1

     sldIndex = 20
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 14")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Diferente").Range("C10:C13")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    'GRAFICO 2

    sldIndex = 20
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 15")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Diferente").Range("D10:D13")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If

    'GRAFICO 3

    sldIndex = 20
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 16")

    If pptShape.HasChart Then

   On Error Resume Next

        Set rngDatos = exAPP.Sheets("Diferente").Range("E10:E13")

    excelData = rngDatos.Value

     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate

        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear

         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i

        pptShape.chart.Refresh

        chartDataWorkbook.Close

             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
      ESTA ES LA HOJA 21
     
      Actualizar la diapositiva 21 con datos de la hoja "Relevancia" en Excel
         Actualizar textos
        
        Set rngTexto = exAPP.Sheets("Relevancia").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(21).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

         Actualizar textos de ADS

        Set rngTexto = exAPP.Sheets("Relevancia").Range("C6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(21).Shapes("TextBox 12").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Relevancia").Range("D6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(21).Shapes("TextBox 13").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Relevancia").Range("E6")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(21).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto

     
     Actualizar tabla 4
     
     Set rngTexto = exAPP.Sheets("Relevancia").Range("B10:B14")
        excelData = rngTexto.Value
        sldIndex = 21
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 4").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If
        
        TABLA 2 "BASE"
        
     Set rngTexto = exAPP.Sheets("Relevancia").Range("B8:E8")
        excelData = rngTexto.Value
        sldIndex = 21
        Set pptSlide = ppPresentacion.Slides(sldIndex)
        Set pptTable = pptSlide.Shapes("Table 26").Table

        If Not pptTable Is Nothing Then
            With pptTable
                Do While .Rows.Count < UBound(excelData, 1)
                    .Rows.Add
                Loop
                Do While .Rows.Count > UBound(excelData, 1)
                    .Rows(.Rows.Count).Delete
                Loop
                Do While .Columns.Count < UBound(excelData, 2)
                    .Columns.Add
                Loop
                Do While .Columns.Count > UBound(excelData, 2)
                    .Columns(.Columns.Count).Delete
                Loop

                For i = 1 To UBound(excelData, 1)
                    For j = 1 To UBound(excelData, 2)
                        .cell(i, j).Shape.TextFrame.TextRange.Text = excelData(i, j)
                    Next j
                Next i
            End With
        Else
            MsgBox "El shape seleccionado no contiene una tabla."
        End If
     
     
     Graficos de la Hojas 21
     GRAFICO 1
     
     sldIndex = 21
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 14")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Relevancia").Range("C10:C14")
   
    excelData = rngDatos.Value
    
     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear
        
         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i
        
        pptShape.chart.Refresh
        
        chartDataWorkbook.Close
        
             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
    GRAFICO 2
    
    sldIndex = 21
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 15")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Relevancia").Range("D10:D14")
   
    excelData = rngDatos.Value
    
     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear
        
         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i
        
        pptShape.chart.Refresh
        
        chartDataWorkbook.Close
        
             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
    GRAFICO 3
    
    sldIndex = 21
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Content Placeholder 16")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Relevancia").Range("E10:E14")
   
    excelData = rngDatos.Value
    
     Set chartDataWorkbook = pptShape.chart.ChartData.Workbook
        chartDataWorkbook.Activate
        
        Set chartWorksheet = chartDataWorkbook.Worksheets(1)
        chartWorksheet.Cells.Clear
        
         For i = 1 To UBound(excelData, 1)
            chartWorksheet.Cells(i + 1, 2).Value = excelData(i, 1)
        Next i
        
        pptShape.chart.Refresh
        
        chartDataWorkbook.Close
        
             Else
        MsgBox "El shape seleccionado no contiene un gráfico."
    End If
    
    
    
   'ppPresentacion.Save
   'ppPresentacion.Close
    
     'ThisWorkbook.Close
    exAPP.Close SaveChanges:=False
    
    MsgBox "Datos automatizados correctamente."
    End If
    
    Exit Sub
    
     ppPresentacion.Close
ErrorHandler:
    MsgBox "Ha cerrado el programa: " & Err.Description, vbExclamation

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub