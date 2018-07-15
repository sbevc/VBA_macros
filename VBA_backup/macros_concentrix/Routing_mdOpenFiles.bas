Attribute VB_Name = "mdOpenFiles"
'Pedimos al usuario que seleccione hasta 3 archivos excel de los cuales copiamos algunos datos
Sub OpenFiles()

Dim Awb As Workbook, wb As Workbook
Dim wb1 As Variant
Dim wb2 As Variant
Dim wb3 As Variant
Dim lrow As Integer
Dim dataRng As Range

Set Awb = ThisWorkbook

wb1 = Application.GetOpenFilename(Title:="Seleccionar archivo")
wb2 = Application.GetOpenFilename(Title:="Seleccionar archivo")
wb3 = Application.GetOpenFilename(Title:="Seleccionar archivo")

'Si no selecciona ningún archivo exit sub
If wb1 = False And wb2 = False And wb3 = False Then
    MsgBox "No seleccionó ningún archivo"
    Exit Sub
End If


colNames

'Copiamos los datos
Set wb = Workbooks.Open(fileName:=wb1, Local:=True)
CopyData wb, Awb
wb.Close

If wb2 <> False Then
    Set wb = Workbooks.Open(fileName:=wb2, Local:=True)
    CopyData wb, Awb
    wb.Close
End If

If wb3 <> False Then
    Set wb = Workbooks.Open(fileName:=wb3, Local:=True)
    CopyData wb, Awb
    wb.Close
End If


'Ponemos ceros en la columna R
lrow = Awb.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Awb.Sheets(1).Range("R5:R" & lrow).Value = 0
Awb.Sheets(1).Range("O6").Copy
Awb.Sheets(1).Range("R5:R" & lrow).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False


'Chequeamos si en la columna P (total neto) hay algún valor cero, y si lo hay avisamos
For i = 5 To lrow
    If Awb.Sheets(1).Range("P" & i).Value = 0 Then
        MsgBox "Se encontró algún valor total neto cero"
        Exit Sub
    End If
Next

Set dataRng = Awb.Sheets(1).Range("A4").CurrentRegion

'Ordenamos por ISIN
Awb.Sheets(1).Sort.SortFields.Add Key:=Range("E5:E" & lrow), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
Awb.Sheets(1).Sort.SortFields.Add Key:=Range("C5:C" & lrow), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
Awb.Sheets(1).Sort.SortFields.Add Key:=Range("A5:A" & lrow), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With Awb.Sheets(1).Sort
    .SetRange Range("A4:X" & lrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


Awb.Sheets(1).Range("H5:J" & lrow).Style = "Comma"

'ponemos el conteo de filas con datos a partir de la A5
With Awb.Sheets(1).Range("A2")
        .Value = Application.WorksheetFunction.CountA _
            (Awb.Sheets(1).Range(Range("A5"), Range("A5").End(xlDown)))
        .Font.Size = 18
        With .borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
    End With
End With

printConfiguration

'para las operaciones en mercado CCY y divisa GBP, dividimos el Trade Price entre 100
caseGBP

End Sub
'Ponemos los nombres de las columnas y les damos formato
Sub colNames()

Dim arrColName(1 To 24) As String

arrColName(1) = "B/S"
arrColName(2) = "Mkt CCY"
arrColName(3) = "Leg Curr"
arrColName(4) = "Security"
arrColName(5) = "Isin Code"
arrColName(6) = "Trade Date"
arrColName(7) = "Settle Date"
arrColName(8) = "Quantity"
arrColName(9) = "Trade Price"
arrColName(10) = "All in Net Price"
arrColName(11) = "Consideration"
arrColName(12) = "Commission"
arrColName(13) = "Local Charges"
arrColName(14) = "Stamp"
arrColName(15) = "Fee3"
arrColName(16) = "Total Net"
arrColName(17) = "Sub a/c Name"
arrColName(18) = ""
arrColName(19) = "Matched"
arrColName(20) = "Trade Time"
arrColName(21) = "Ref"
arrColName(22) = "Term"
arrColName(23) = "Status"
arrColName(24) = "Av Price"



Sheets(1).Range("F2").Value = "Banco Bilbao Vizcaya Argentaria, S.A"
Sheets(1).Range("A4:X4").Value = arrColName


End Sub

Sub CopyData(sourceWb As Workbook, destWb As Workbook)

Dim lrowSource As Integer
Dim lrowDest As Integer

lrowSource = sourceWb.Sheets("SUMMARY").Range("A" & Rows.Count).End(xlUp).Row
lrowDest = destWb.Sheets(1).Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row

'Copiamos el formato
sourceWb.Sheets("SUMMARY").Range("A13:x13").Copy
destWb.Sheets(1).Range("A4:X4").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'Copiamos los datos
sourceWb.Sheets("SUMMARY").Range("A14:O" & lrowSource).Copy _
    Destination:=destWb.Sheets(1).Range("A" & lrowDest)
sourceWb.Sheets("SUMMARY").Range("Q14:R" & lrowSource).Copy _
    Destination:=destWb.Sheets(1).Range("P" & lrowDest)
sourceWb.Sheets("SUMMARY").Range("S14:X" & lrowSource).Copy _
    Destination:=destWb.Sheets(1).Range("S" & lrowDest)

End Sub

Sub caseGBP()

    Dim Awb As Workbook
    Dim ws As Worksheet
    Dim tradeRng As Range, c As Range
    
    Set Awb = ThisWorkbook
    Set ws = Awb.Sheets(1)
    

    'filtramos por mercado UKE y divisa GBP
    ws.Range("A4").CurrentRegion.AutoFilter Field:=2, Criteria1:="UKE"
    ws.Range("A4").CurrentRegion.AutoFilter Field:=3, Criteria1:="GBP"
    
    'nos fijamos si hay datos en el filtro y loopeamos por los valores de la columna Trade price
    'dividiendo el valor actual entre 100
    If ws.Range("A" & Rows.Count).End(xlUp).Row <> 4 Then
        
        Set tradeRng = ws.Range(Range("I5"), Range("I5").End(xlDown))
        
        For Each c In tradeRng.SpecialCells(xlCellTypeVisible)
            
            c.Value = c.Value / 100
            
            'chequeamos que los valores de All in Net Price y Trade Price sean idénticos hasta el primer decimal.
            'si difieren en mas de 0,0999999 que marque la linea de rojo y no la agregue al CSV. Agregamos un filtro
            'al csv por color de fuente para que no agregue los casos marcados
            Dim diff As Double
            diff = Abs(c.Value - c.Offset(0, 1).Value)
            If diff > 0.09999999 Then
                Range(Cells(c.Row, 1), Cells(c.Row, 25)).Font.Color = -16776961
            End If
            
        Next c
        
    End If
    
    
    ws.Cells.AutoFilter
    
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
    ThisWorkbook.Sheets(1).PageSetup.PrintArea = _
        Range(Range("A4"), Range("P4").End(xlDown)).Address     'area de impresión
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

