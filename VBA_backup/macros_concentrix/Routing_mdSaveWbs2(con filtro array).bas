Attribute VB_Name = "mdSaveWbs2"
'Public Const filePath As String = "U:\prueba\"

Sub SaveAsTxt()

Dim Awb As Workbook         'this workbook
Dim wbUSA As Workbook       'wb con emisiones de USA y europa que guardaremos como CSV
Dim wbASIA As Workbook      'wb con emisiones de ASIA que guardaremos como CSV
Dim wbLiq As Workbook       'wb con el que trabaja liquidación
Dim sourceRow As Integer
Dim rngMarket As Range      'Rango con los mercados
Dim UniqueMkts() As Variant  'Mercados, sacados con la función uniquevalues
Dim USAmkts() As String     'mercados USA
Dim ASIAmkts() As String    'mercados ASIA
Dim xTradeDate As Date      'fecha tradedate

Set Awb = ThisWorkbook
sourceRow = Awb.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
xTradeDate = Awb.Sheets(1).Range("f5").Value

'Sacamos los valores únicos de la columna de los mercados y los separamos en dos arrays, uno con
'los mercados asiáticos y otro con el resto
Set rngMarket = Awb.Sheets(1).Range("B5:B" & sourceRow)
UniqueMkts = getUniqueValues(Sheets(1), rngMarket)
For i = 1 To UBound(UniqueMkts, 1)
    Select Case UniqueMkts(i)
        Case Is = "AUE", "HKE", "NZE"
            j = j + 1
            ReDim Preserve ASIAmkts(1 To j)
            ASIAmkts(j) = UniqueMkts(i)
        Case Else
            k = k + 1
            ReDim Preserve USAmkts(1 To k)
            USAmkts(k) = UniqueMkts(i)
    End Select
Next


'Filtramos según los arrays y copiamos los datos filtrados a un libro nuevo y los guardamos
'con la funcion guardado CSV
If Len(Join(ASIAmkts)) > 0 Then
    Awb.Sheets(1).Range("A5").CurrentRegion.AutoFilter _
        Field:=2, Criteria1:=(ASIAmkts), Operator:=xlFilterValues
    Awb.Sheets(1).Range("A5").CurrentRegion.AutoFilter _
        Field:=4, Operator:=xlFilterAutomaticFontColor      'sacamos las filas con font roja
    Awb.Sheets(1).Range(Range("A4"), Range("P4").End(xlDown)).Copy
    Set wbASIA = Workbooks.Add
    wbASIA.Sheets(1).Range("A3").PasteSpecial Paste:=xlPasteValues
    wbASIA.Sheets(1).Range("F1").Value = "Banco Bilbao Vizcaya Argentaria, S.A"
    wbASIA.Sheets(1).Range("F:G").NumberFormat = "dd/mm/yyyy"
    saveAsCSV wbASIA, xTradeDate, " 0001"
End If

If IsEmpty(USAmkts) = False Then
    Awb.Sheets(1).Range("A5").CurrentRegion.AutoFilter _
        Field:=2, Criteria1:=(USAmkts), Operator:=xlFilterValues
    Awb.Sheets(1).Range("A5").CurrentRegion.AutoFilter _
        Field:=4, Operator:=xlFilterAutomaticFontColor      'sacamos las filas con font roja
    Awb.Sheets(1).Range(Range("A4"), Range("P4").End(xlDown)).Copy
    Set wbUSA = Workbooks.Add
    wbUSA.Sheets(1).Range("A3").PasteSpecial Paste:=xlPasteValues
    wbUSA.Sheets(1).Range("F1").Value = "Banco Bilbao Vizcaya Argentaria, S.A"
    wbUSA.Sheets(1).Range("F:G").NumberFormat = "dd/mm/yyyy"
    saveAsCSV wbUSA, xTradeDate, " 0145"
End If


'Guardamos el archivo que usa liquidación
Dim strMonth As String, strDay As String, stryear As String
strDay = Format(day(xTradeDate), "00")
strMonth = Format(Month(xTradeDate), "00")
stryear = Format(Year(xTradeDate), "0000")
wbName = strDay & "." & strMonth & "." & stryear

Awb.Sheets(1).Cells.AutoFilter
Awb.Sheets(1).Copy
ActiveWorkbook.SaveAs fileName:=filePath(xTradeDate, False) & "\" & wbName & ".xls", FileFormat:=xlOpenXMLWorkbook
Awb.Close savechanges:=False

End Sub

Sub saveAsCSV(wb As Workbook, xdate As Date, xname As String)

Dim wbName As String
Dim strMonth As String, strDay As String, stryear As String
Dim lrow As Integer
Dim wb2 As Workbook

strDay = Format(day(xdate), "00")
strMonth = Format(Month(xdate), "00")
stryear = Format(Year(xdate), "0000")
wbName = strDay & "." & strMonth & "." & stryear

wb.Activate

'Nos fijamos si los datos superan la fila 104 para dividir el libro en 2
If IsEmpty(wb.Sheets(1).Range("A104")) = False Then
    lrow = wb.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Set wb2 = Workbooks.Add
    wb.Sheets(1).Rows("105:" & lrow).Cut Destination:=wb2.Sheets(1).Range("A4") 'cortamos a partir de la fila 104
    wb.Sheets(1).Rows(3).EntireRow.Copy Destination:=wb2.Sheets(1).Range("A3")
    wb2.SaveAs fileName:=filePath(xdate, True) & "\" & wbName & xname & "_2.csv", _
            FileFormat:=xlCSV, Local:=True
    wb2.Close savechanges:=False
End If

    wb.SaveAs fileName:=filePath(xdate, True) & "\" & wbName & xname & ".csv", _
            FileFormat:=xlCSV, Local:=True
    wb.Close savechanges:=False

End Sub

Function getUniqueValues(ws As Worksheet, valueRng As Range)
    
    Dim arrUniqueValues() As Variant
    
    valueRng.Copy Destination:=ws.Cells(1, Columns.Count)
    ws.Range(Cells(1, Columns.Count), Cells(Rows.Count, Columns.Count).End(xlUp)).RemoveDuplicates _
        Columns:=1, Header:=xlNo
    arrUniqueValues = Application.WorksheetFunction.Transpose _
        (ws.Range(Cells(1, Columns.Count), Cells(Rows.Count, Columns.Count).End(xlUp)).Value)
    ws.Range(Cells(1, Columns.Count), Cells(Rows.Count, Columns.Count).End(xlUp)).ClearContents
    
    getUniqueValues = arrUniqueValues
    
End Function


'Función para ver la ruta dónde guardar los archivos en función del tradedate (xdate).
'Si la ruta no existe la crea. Distinguimos entre los CSV y XLS que se gueardan en rutas distintas.
'Si el fileType es True toma la ruta para CVS y si es False para XLS
Function filePath(xdate As Date, fileType As Boolean) As String

    Const staticPath As String = "H:\SC000068\INSTINET\"        'parte constante de la ruta
    Dim varPath1 As String      'parte INSITINET & AÑO de la ruta
    Dim varpath2 As String      'parte dd-mes de la ruta
    Const staticPath2 As String = "GESTORA"
    Dim strMonth As String
    Dim CSVPath As String
    Dim XLSPath As String
    
    varPath1 = "INSTINET " & Year(xdate)
    
    strMonth = Format(Month(xdate), "00")
    varpath2 = strMonth & " - " & UCase(MonthName(Month(xdate)))
    
    CSVPath = staticPath & varPath1 & "\" & varpath2 & "\" & staticPath2
    XLSPath = staticPath & varPath1 & "\" & varpath2
    
    'vamos chequeando si las carpetas existen y si no las creamos
    If Len(Dir(staticPath & varPath1, vbDirectory)) = 0 Then
        MkDir staticPath & varPath1
        MkDir staticPath & varPath1 & "\" & varpath2
        MkDir CSVPath
    ElseIf Len(Dir(staticPath & varPath1 & "\" & varpath2, vbDirectory)) = 0 Then
        MkDir staticPath & varPath1 & "\" & varpath2
        MkDir CSVPath
    ElseIf Len(Dir(CSVPath, vbDirectory)) = 0 Then
        MkDir CSVPath
    End If
        
    'Por último devolvermos la ruta según sea un archivo CSV o un excel normal.
    'Para los CSV la ruta será H:\SC000068\INSTINET\INSTINET (año)\dd-mes\GESTORA
    'Para los xls la ruta será H:\SC000068\INSTINET\INSTINET (año)\dd-mes\
    If fileType = True Then
        filePath = CSVPath
    Else
        filePath = XLSPath
    End If
    
    
End Function
