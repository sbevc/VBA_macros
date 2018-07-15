Attribute VB_Name = "mdCopyData"
Sub Main()

Dim xEntrada As Variant, xsalida As Variant, lrow As Integer
Dim wb As Workbook, Awb As Workbook, destWs As Worksheet

Set Awb = ThisWorkbook
Set destWs = Awb.Sheets(wsDay)

'--------Copiamos los datos--------'
xEntrada = Application.GetOpenFilename( _
    Title:="Seleccionar Horarios de Entrada")

xsalida = Application.GetOpenFilename( _
    Title:="Seleccionar Horarios de Salida")

If xEntrada = False Or xsalida = False Then     'si no selecciona algun archivo que de un mensage y salga
    MsgBox "Falta seleccionar archivos"
    Exit Sub
End If
    
If IsEmpty(destWs.Range("D3")) = False Then     'si hay datos en la hoja destino, que pregunte si desea sobreescribirlos
    Dim answer As Integer
    answer = MsgBox("El día seleccionado ya contiene datos. Desea sobreescribirlos?", vbYesNo, "Sobreescribir datos?")
    
    If answer = vbYes Then
        lrow = destWs.Range("C2").End(xlDown).Row
        destWs.Range("D3:T" & lrow).ClearContents
        lrow = destWs.Range("B60").End(xlDown).Row
        destWs.Range("B61:F" & lrow).ClearContents
    Else
        Exit Sub
    End If
End If


Set wb = Workbooks.Open(FileName:=xEntrada, Local:=True)
copyFrom wb, destWs  'Tomamos los datos de la celda B2, ver función wsDay
wb.Close

Set wb = Workbooks.Open(FileName:=xsalida, Local:=True)
copyFrom wb, destWs   'Tomamos los datos de la celda B2, ver función wsDay
wb.Close


'--------Les damos formato--------'
destWs.Activate
formatData

'Resume los datos
ResumeData

With destWs.Tab
    .ThemeColor = xlThemeColorAccent3
    .TintAndShade = 0.399975585192419
End With

End Sub

'Copiamos los datos de los horarios al archivo de la macro
Sub copyFrom(wbFrom As Workbook, wsTo As Worksheet)

Dim wsFrom As Worksheet, lrow As Integer, destRow As Integer
Set wsFrom = wbFrom.Sheets("APN5")
lrow = wsFrom.Range("E1").End(xlDown).Row

destRow = wsTo.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Row

wsFrom.Range("B2:E" & lrow).Copy Destination:=wsTo.Range("B" & destRow)

End Sub

Function wsDay()

Dim dateInput As Date, dia As String, mes As String     'al dar formato string permite que quede dd/mm
cellInput = ThisWorkbook.Sheets("Resumen").Range("B2")

dia = Format(day(cellInput), "00")
mes = Format(month(cellInput), "00")

wsDay = dia & "-" & mes

End Function

Sub formatData()

Dim lrow As Integer, ws As Worksheet, xData As Range

Set ws = ActiveSheet
lrow = ws.Range("B60").End(xlDown).Row

Range("E61:E" & lrow).Insert Shift:=xlToRight
Range("D61:D" & lrow).TextToColumns Destination:=Range("D61"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
Range("B61:B" & lrow).NumberFormat = "[$-F400]h:mm:ss AM/PM"

Range("B60:F" & lrow).Sort _
    key1:=Range("D61:D" & lrow), _
    key2:=Range("B61:B" & lrow), _
    Header:=xlYes

End Sub

