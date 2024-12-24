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
   
   
   'Textos
   
    Set rngTexto = exAPP.Sheets("Tipología").Range("A1")
    
    strTexto = rngTexto.Value
   
    ppPresentacion.Slides(3).Shapes("Title 2").TextFrame.TextRange.Text = strTexto
    
    Set rngTexto = exAPP.Sheets("Tipología").Range("C6")
    
    strTexto = rngTexto.Value
    
    ppPresentacion.Slides(3).Shapes("Valor_Base").TextFrame.TextRange.Text = strTexto
    
    Set rngTexto = exAPP.Sheets("Tipología").Range("C6")
    
    strTexto = rngTexto.Value
    
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
        
        'Textos
    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("A1")
    
    strTexto = rngTexto.Value
    
    ppPresentacion.Slides(4).Shapes("Title 2").TextFrame.TextRange.Text = strTexto
    
    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("C6")
    
    strTexto = rngTexto.Value
    
    ppPresentacion.Slides(4).Shapes("TextBox 2").TextFrame.TextRange.Text = strTexto
    
    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("C6")

    number = Int(rngTexto.Value)
    
    strTexto = number
    
    
    'Textos en tabla
    
    Set rngTexto = exAPP.Sheets("Conocimiento de marca").Range("B7:B12")
    
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
   
        Set rngDatos = exAPP.Sheets("Conocimiento de marca").Range("C6:C12")
        
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
   
        Set rngDatos = exAPP.Sheets("Conocimiento de marca").Range("D6:D12")
        
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
   
        Set rngDatos = exAPP.Sheets("Conocimiento de marca").Range("E6:E12")
        
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
    
    
    'ESTA ES LA HOJA 5
    
    ' Actualizar la diapositiva 5 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Funnels").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Funnels").Range("C7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("D7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 2").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("E7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 3").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("F7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 4").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("G7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("H7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 6").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Funnels").Range("I7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("Brand 7").TextFrame.TextRange.Text = strTexto

        'Actualizar base
        Set rngTexto = exAPP.Sheets("Funnels").Range("C8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(5).Shapes("TextBox 3").TextFrame.TextRange.Text = strTexto
    
    'Textos en tabla
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
    
    'GRAFICOS
    
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
    
    'ESTA ES LA HOJA 6
    

    'DATOS HOJA 6
    
     ' Actualizar la diapositiva 6 con datos de la hoja "Consideración de compra" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("Title 2").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas
        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("C7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 26").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("D7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 36").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("E7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 37").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("F7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 38").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("G7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 39").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("H7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 40").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Consideración de compra").Range("I7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(6).Shapes("TextBox 41").TextFrame.TextRange.Text = strTexto
        
    'Textos en tabla
    
    Set rngTexto = exAPP.Sheets("Consideración de compra").Range("B8:I8")
    
    excelData = rngTexto.Value
    
    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptTable = pptSlide.Shapes("Table 1").Table
    
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
    

    sldIndex = 6
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 49")
    
    If pptShape.HasChart Then
    
    Set rngDatos = exAPP.Sheets("Consideración de compra").Range("B9:I12")
   
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
    
    
    'ESTA ES LA HOJA 7
    
     ' Actualizar la diapositiva 7 con datos de la hoja "Perfil de imagen" en Excel
        ' Actualizar textos
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("A1")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("Title 2").TextFrame.TextRange.Text = strTexto
        'Actualizar bases para cada marca

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("C8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 35").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("D8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 24").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("E8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 25").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("F8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 26").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("G8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 27").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("H8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 28").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("I8")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 29").TextFrame.TextRange.Text = strTexto

        ' Actualizar textos de marcas

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("C7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 1").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("D7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 5").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("E7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 14").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("F7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 15").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("G7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 16").TextFrame.TextRange.Text = strTexto

        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("H7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 17").TextFrame.TextRange.Text = strTexto
        
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("I7")
        strTexto = rngTexto.Value
        ppPresentacion.Slides(7).Shapes("TextBox 18").TextFrame.TextRange.Text = strTexto


         ' Actualizar tablas en la diapositiva 7
        Set rngTexto = exAPP.Sheets("Perfil de imagen").Range("B9:B21")
        excelData = rngTexto.Value
        sldIndex = 7
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

    
    'GRAFICOS

    'GRAFICO 1

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 9")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("C9:C21")
        
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

     sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 10")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("D9:D21")
        
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

     sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 11")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("E9:E21")
        
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
    
     sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 12")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("F9:F21")
        
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

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 13")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("G9:G21")
        
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

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 20")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("H9:H21")
        
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

    sldIndex = 7
    Set pptSlide = ppPresentacion.Slides(sldIndex)
    Set pptShape = pptSlide.Shapes("Chart 21")
    
    If pptShape.HasChart Then
   
   On Error Resume Next
   
        Set rngDatos = exAPP.Sheets("Perfil de imagen").Range("I9:I21")
        
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