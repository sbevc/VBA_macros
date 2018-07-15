Attribute VB_Name = "mdSaveWbs"
'en este modulo pasamos la info haciendo un loop celda por celda. lo cambié haciendo un filtro que anda
'más rápido en el módulo mdSaveWbs2

Public Const filePath As String = "U:\prueba\"

Sub SaveAsTxtViejo()

Dim Awb As Workbook         'this workbook
Dim wbUSA As Workbook       'wb con emisiones de USA y europa que guardaremos como CSV
Dim wbASIA As Workbook      'wb con emisiones de ASIA que guardaremos como CSV
Dim wbLiq As Workbook       'wb con el que trabaja liquidación
Dim sourceRow As Integer
Dim destRow As Integer
Dim market As String


Set Awb = ThisWorkbook
sourceRow = Awb.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row


'Creamos dos archivos, en uno pasamos los trades asiáticos y en el otro el resto
Set wbUSA = Workbooks.Add
Set wbASIA = Workbooks.Add
For i = 5 To sourceRow
    market = Awb.Sheets(1).Cells(i, 2).Value
    Select Case market
        Case Is = "AUE", "JPE", "HKE", "NZE"
            destRow = wbASIA.Sheets(1).Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
            wbASIA.Sheets(1).Rows(destRow).EntireRow.Value = Awb.Sheets(1).Rows(i).EntireRow.Value
        Case Else
            destRow = wbUSA.Sheets(1).Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Row
            wbUSA.Sheets(1).Rows(destRow).EntireRow.Value = Awb.Sheets(1).Rows(i).EntireRow.Value
    End Select
Next i
        
'Guardamos los archivos
Dim xTradeDate As Date

xTradeDate = Awb.Sheets(1).Range("f5").Value
saveAsCSV wbUSA, xTradeDate, " 0145"
saveAsCSV wbASIA, xTradeDate, " 0001"

'Guardamos éste libro como xlsx
Dim strMonth As String, strDay As String, stryear As String
strDay = Format(day(xTradeDate), "00")
strMonth = Format(Month(xTradeDate), "00")
stryear = Format(Year(xTradeDate), "0000")
wbName = strDay & "." & strMonth & "." & stryear

Awb.Sheets(1).Copy
ActiveWorkbook.SaveAs fileName:=filePath & wbName & ".xls", FileFormat:=xlOpenXMLWorkbook
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

With wb
    With .Sheets(1)     'copiamos los nombres de las columnas
        If .Range("A" & Rows.Count).End(xlUp).Row <> 1 Then
            .Rows("1:2").EntireRow.Insert
            mdSaveWbs.colNames
        Else
            wb.Close
        End If
    End With
End With


'Nos fijamos si los datos superan la fila 104 para dividir el libro en 2
If IsEmpty(wb.Sheets(1).Range("A104")) = False Then
    lrow = wb.Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Set wb2 = Workbooks.Add
    wb.Sheets(1).Rows("105:" & lrow).Cut Destination:=wb2.Sheets(1).Range("A4")
    mdSaveWbs.colNames
    wb2.SaveAs fileName:=filePath & wbName & xname & "_2.csv", _
            FileFormat:=xlCSV, Local:=True
    wb2.Close savechanges:=False
End If

    wb.SaveAs fileName:=filePath & wbName & xname & ".csv", _
            FileFormat:=xlCSV, Local:=True
    wb.Close savechanges:=False

End Sub

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



Sheets(1).Range("F1").Value = "Banco Bilbao Vizcaya Argentaria, S.A"
Sheets(1).Range("A3:X3").Value = arrColName


End Sub

