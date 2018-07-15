Attribute VB_Name = "mdMain"
Sub Main()

    Dim sourceWb As Variant
    Dim wbDest As Workbook
    Dim aWb As Workbook
    Dim ws As Worksheet
    Dim lrow As Integer
    
    sourceWb = Application.GetOpenFilename(Title:="Seleccionar archivo")
    If sourceWb = False Then     'si no selecciona algun archivo que de un mensage y salga
        MsgBox "No seleccionó ningún archivo"
        Exit Sub
    End If
    
    Set aWb = ThisWorkbook
    Set ws = aWb.Sheets(1)
    
    'copiamos los datos
    Set wbDest = Workbooks.Open(fileName:=sourceWb, Local:=True)
    wbDest.Sheets(1).Range("A1").CurrentRegion.Copy Destination:=aWb.Sheets(1).Range("A4")
    wbDest.Close
    
    
    order_formatData
    
    Application.ScreenUpdating = False
    groupby_acc_div
    pintarDeAmarillo
    Application.ScreenUpdating = True
    
    computeSubtotals
    
    ws.Cells.AutoFilter
    
    printConfiguration
    
End Sub
'borramos las columnas y ordenamos los datos
Sub order_formatData()

    Dim ws As Worksheet
    Dim lrow As Integer
    
    Set ws = ThisWorkbook.Sheets(1)
    
    ws.Activate
    
    With ActiveWindow
        .Zoom = 75
        .DisplayGridlines = False
    End With
    
    
    With ws
    
        .Range("A:B, D:E, G:I, K:K, O:U, W:W").Delete
    
      'Ordenamos los datos
        lrow = .Range("A" & Rows.Count).End(xlUp).Row
        
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("B5:B" & lrow) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("E5:E" & lrow) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("F5:F" & lrow) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("H5:H" & lrow) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range("A4:M" & lrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    
        .Cells.EntireColumn.AutoFit

        .Range("B:C, E:F, H:H, J:J, N:N").Font.Bold = True
        .Range("D:D, I:K, N:N").Style = "Comma"
        
    End With

    
    
End Sub

'Agrupamos según las cuentas depósito y divisa y hacemos los subtotales y las cuadrículas
Sub groupby_acc_div()

    Dim ws As Worksheet
    Dim lrow As Integer
    Dim dataRng As Range                            'rango con toda la tabla
    Dim accsRng As Range                            'rango que tomamos para las cuentas
    Dim divRng As Range                             'rango que tomamos para las divisas
    Dim uniqueDIV() As Variant
    Dim uniqueACC() As Variant
    Dim rTable As Range                             'Rango resultante del filtrado por cuenta y divisa
    Dim lHeadersRows As Long
    Dim subtotalRng As Range                        'rango en el que calculamos el subtotal
    Dim fsumrow As Integer                         'primera y última fila para las sumas
    
    Set ws = ThisWorkbook.Sheets(1)
    lrow = ws.Range("A" & Rows.Count).End(xlUp).Row
    Set dataRng = ws.Range("A4").CurrentRegion
    Set accsRng = ws.Range("F5:F" & lrow)
    Set divRng = ws.Range("E5:E" & lrow)
    
    'Tomamos las cuentas para filtrar por cada una. Luego filtramos por la divisa, y hacemos el
    'subtotal y marcamos los bordes
    
    uniqueACC = getUniqueValues(ws, accsRng)
    uniqueDIV = getUniqueValues(ws, divRng)
    
    dataRng.AutoFilter
    For i = 1 To UBound(uniqueACC, 1)
        For j = 1 To UBound(uniqueDIV, 1)
        
            With dataRng        'filtro por cuenta y divisa
                .AutoFilter Field:=6, Criteria1:=uniqueACC(i) & ""
                .AutoFilter Field:=5, Criteria1:=uniqueDIV(j) & ""
            End With
            
            'si hay datos en el filtro, agregamos los bordes y el subtotal
            lrow = ws.Range("A" & Rows.Count).End(xlUp).Row
            If lrow <> 4 Then
            
                fsumrow = getFirstFilterRow(5, 10)
                ws.Cells(fsumrow, 14).Formula = _
                    "=SUM(J" & fsumrow & ":J" & lrow & ")"
                
                'definimos el rango filtrado
                Set rTable = ws.Range("A4").CurrentRegion
                lHeadersRows = rTable.ListHeaderRows
                Set rTable = rTable.Resize(rTable.Rows.Count - lHeadersRows, rTable.Columns.Count)
                Set rTable = rTable.Offset(1)
                
                borders rTable
                
                
            End If
            
        Next j
    Next i
    
    dataRng.AutoFilter
    total_borders dataRng.Resize(dataRng.Rows.Count, 14)
    total_borders Range("A4:N4")
    
    
    
    
End Sub
'función que devuelve array con los valores únicos de un rango dado en una ws dada
Function getUniqueValues(wks As Worksheet, valueRng As Range)
    
    Dim arrUniqueValues() As Variant
    Dim lrow As Long            'vemos si hay una sola divisa
    
    valueRng.Copy Destination:=wks.Cells(1000000, Columns.Count)
    wks.Range(Cells(1000000, Columns.Count), Cells(Rows.Count, Columns.Count).End(xlUp)).RemoveDuplicates _
        Columns:=1, Header:=xlNo
    
    lrow = wks.Cells(Rows.Count, Columns.Count).End(xlUp).Row
    If lrow = 1000000 Then
        ReDim arrUniqueValues(1 To 1)
        arrUniqueValues(1) = wks.Cells(1000000, Columns.Count).Value
    Else
        arrUniqueValues = Application.WorksheetFunction.Transpose _
            (wks.Range(Cells(1000000, Columns.Count), Cells(Rows.Count, Columns.Count).End(xlUp)).Value)
    End If
    
    wks.Range(Cells(1000000, Columns.Count), Cells(Rows.Count, Columns.Count).End(xlUp)).ClearContents
    
    getUniqueValues = arrUniqueValues
    
End Function
'marcamos solo el borde de abajo para los filtros
Sub borders(rng As Range)
    
    With rng.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    
End Sub


'marcamos el borde del título y de toda la tabla
Sub total_borders(rng As Range)
    
    With rng.borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With rng.borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With rng.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With rng.borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    
End Sub
'Función para obtener la primer fila dentro del rango filtrado
Function getFirstFilterRow(startRow As Integer, col As Integer)
    
    Dim c As Range
    Dim ws As Worksheet
    
    Set ws = Sheets(1)
    Set c = ws.Cells(startRow, col)
    
    c.Activate
    
    Do Until c.EntireRow.Hidden = False
  
       Set c = c.Offset(1, 0)
  
    Loop
  
  getFirstFilterRow = c.Row


End Function

'pintamos de amarillo algunas filas particulares según los brokers(ver mail Macro gestoras con detalle):
    'para todas las divisas:  CD, DEX y MS
    'para USD: DWS, GSIE,
    'para EUR: BR, DEXIAFR
    'para JPY: PAR, JPM, JPMLIQ
    'para NOK: PTEMP
    'para el broker ARL marcamos las emisiones IE
'tambien pintamos los valores "DELEGATED" de la columna cuenta depósito
    
''Sacarlas de los subtotales por divisa!!!
Sub pintarDeAmarillo()

    
    Dim broker As String
    Dim lrow As Integer
    Dim divCol As Integer, brokerCol As Integer, isinCol As Integer
    Dim i As Integer
    
    ThisWorkbook.Sheets(1).Activate
    
    lrow = Range("A" & Rows.Count).End(xlUp).Row
    brokerCol = 2
    divCol = 5
    isinCol = 3
    
    For i = 5 To lrow
        
        'si dice delegated
        If Cells(i, 6).Value = "DELEGATED" Then
            pintar i
        End If
        
        broker = Cells(i, brokerCol).Value
        Select Case broker
            
            'Todas las divisas
            Case Is = "CD", "DEX", "MS"
                pintar i
            
            'USD
            Case Is = "DWS", "GSIE"
                If Cells(i, divCol).Value = "USD" Then
                    pintar i
                End If
                
            'EUR
            Case Is = "BR", "DEXIAFR"
                If Cells(i, divCol).Value = "EUR" Then
                    pintar i
                End If
            
            'JPY
            Case Is = "PAR", "JPM", "JPMLIQ"
                If Cells(i, divCol).Value = "JPY" Then
                    pintar i
                End If
            
            'NOK
            Case Is = "PTEMP"
                If Cells(i, divCol).Value = "NOK" Then
                    pintar i
                End If
            
            'marcamos las IE
            Case Is = "ALR"
                If Left(Cells(i, isinCol).Value, 2) = "IE" Then
                    pintar i
                End If
        
        End Select
        
    Next

End Sub
'pintamos la fila desde la "A" a la "N"
Sub pintar(fila As Integer)

    Dim rng As Range
    
    Set rng = Range("A" & fila & ":N" & fila)

    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub
'hacemos el subtotal por cada divisa, excluyendo aquellas celdas pintadas
Sub computeSubtotals()

    Dim lrow As Integer, ws As Worksheet
    Dim uniqueDIV() As Variant
    Dim divRng As Range                             'Columna con las divisas
    Dim dataRng As Range                            'tabla completa de datos
    Dim subtotalRng As Range                        'rango con los subtotales
    Dim c As Range
    Dim divisa As String
    Dim coldivisa As Integer
    Dim arrEUR() As Variant, arrUSD() As Variant, arrJPY() As Variant, arrGBP() As Variant, arrNOK() As Variant
    Dim counter As Integer                          'contador para ver cuantos subtotales pintados hay
    
    Set ws = ThisWorkbook.Sheets(1)
    lrow = ws.Range("A" & Rows.Count).End(xlUp).Row
    Set divRng = ws.Range("E5:E" & lrow)
    uniqueDIV = getUniqueValues(ws, divRng)
    coldivisa = 5
    
    
    'hacemos un array con los address de los subtotales pintados para restarlos del sumif.
    'Primero filtramos por color y aquellos valores no vacios y depues los agregamos a un array
    'dependiendo de la divisa
    Set dataRng = ws.Range("A5").CurrentRegion
    dataRng.AutoFilter Field:=10, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor
    dataRng.AutoFilter Field:=14, Criteria1:="<>"
    
    Set subtotalRng = ws.Range("N5:N" & lrow)
    For Each c In subtotalRng.SpecialCells(xlCellTypeVisible)
        
        counter = counter + 1
        
        c.Activate
        divisa = ws.Cells(c.Row, coldivisa).Value
        Select Case divisa
            Case Is = "EUR"
                i = i + 1
                ReDim Preserve arrEUR(1 To i)
                arrEUR(i) = c.Address
            Case Is = "USD"
                j = j + 1
                ReDim Preserve arrUSD(1 To j)
                arrUSD(j) = c.Address
            Case Is = "JPY"
                k = k + 1
                ReDim Preserve arrUSD(1 To k)
                arrUSD(k) = c.Address
            Case Is = "GBP"
                m = m + 1
                ReDim Preserve arrGBP(1 To m)
                arrGBP(m) = c.Address
             Case Is = "NOK"
                n = n + 1
                ReDim Preserve arrNOK(1 To n)
                arrNOK(n) = c.Address
                
        End Select
            
    Next
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add Key:="EUR", Item:=listToSubstract(arrEUR)
    dict.Add Key:="USD", Item:=listToSubstract(arrUSD)
    dict.Add Key:="JPY", Item:=listToSubstract(arrJPY)
    dict.Add Key:="GBP", Item:=listToSubstract(arrGBP)
    dict.Add Key:="NOK", Item:=listToSubstract(arrNOK)
    
    'subtotales por divisa, separamos entre si hay o no celdas a restar
    For i = 1 To UBound(uniqueDIV, 1)
        ws.Range("M" & lrow + i + 2).Value = "Total " & uniqueDIV(i)
        If dict(uniqueDIV(i)) <> "" Then
            ws.Range("N" & lrow + i + 2).Formula = _
                "=+SUMIF(E:E," & Chr(34) & uniqueDIV(i) & Chr(34) & ",J:J) + SUM(" & dict(uniqueDIV(i)) & ")"
        Else
            ws.Range("N" & lrow + i + 2).Formula = _
                "=+SUMIF(E:E," & Chr(34) & uniqueDIV(i) & Chr(34) & ",J:J)"
        End If

    Next
    
    total_borders ws.Range("M" & lrow + 3).CurrentRegion
    ws.Range("N:N").EntireColumn.AutoFit
    
    
    
    'Contamos la cantidad de subtotales no pintados y ponemos el valor en la celda "N2". Para ello contamos el total
    'de valores en la columna N y le restamos la cantidad de elementos de cada array según divisa
    Dim cellsNonYellow As Integer
    
    cellsNonYellow = Application.WorksheetFunction.CountA(Range(Range("N5"), Range("N" & lrow))) - counter
    With Range("N2")
        .Value = cellsNonYellow
        .NumberFormat = "General"
    End With
    
    
    
    'contamos el total de filas con datos y lo ponemos en la celda A2
    With Range("A2")
        .Value = Application.WorksheetFunction.CountA(Range(Range("A4"), Range("A" & lrow)))
        .NumberFormat = "General"
    End With
    
    ajustarYcentrar
    
    
    Range("L2").Value = Range("L5").Value
    
End Sub

'ponemos los valores del array en una lista para restarlos
Function listToSubstract(arrdiv() As Variant) As String
    
    On Error Resume Next
    
    For i = 1 To UBound(arrdiv, 1)
        listToSubstract = listToSubstract & "-" & arrdiv(i)
    Next
    
End Function


Sub ajustarYcentrar()

    Range("A1:A3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
End Sub

Sub printConfiguration()

    Application.PrintCommunication = False
    With ThisWorkbook.Sheets(1).PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ThisWorkbook.Sheets(1).PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ThisWorkbook.Sheets(1).PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.PrintCommunication = False
    With ThisWorkbook.Sheets(1).PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ThisWorkbook.Sheets(1).PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ThisWorkbook.Sheets(1).PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.PrintCommunication = False
    With ThisWorkbook.Sheets(1).PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ThisWorkbook.Sheets(1).PageSetup.PrintArea = Sheets(1).UsedRange.Address     'area de impresión
    Application.PrintCommunication = False
    With ThisWorkbook.Sheets(1).PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True

End Sub
