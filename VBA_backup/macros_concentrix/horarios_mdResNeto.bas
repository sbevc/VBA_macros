Attribute VB_Name = "mdResNeto"
Public ws As Worksheet

Sub Resumen_neto()

Set ws = Sheets("Neto")

If IsEmpty(ws.Range("d5")) = True Then
    Neto mdResNeto.lastWs  'Vemos hasta qué hoja hay datos cargados, y corremos los resúmenes hasta esa ws
Else
    Dim answer As Integer
    answer = MsgBox("Ya existen datos, desea sobreescribirlos?", vbYesNo, "Sobreescribir datos?")
    
    If answer = vbYes Then
        mdResTarde.clearData ws
        Neto mdResNeto.lastWs
    Else
        Exit Sub
    End If
End If


End Sub


Sub Neto(lastWsWithData As Integer)

Dim id_arr() As Variant, ws As Worksheet, sumArr() As Variant

Set ws = Sheets("Neto")
id_arr = Application.WorksheetFunction.Transpose(ws.Range("B5", Range("B5").End(xlDown)))
ReDim sumArr(4 To Sheets.Count, 1 To UBound(id_arr))

'----------COPIAMOS LLEGADAS TARDE----------''loopeamos por las ws y ponemos los días de las llegadas tarde y la hora de llegada
For i = 4 To lastWsWithData
    Sheets(i).Activate
    For j = 1 To UBound(id_arr)
        'Ponemos los datos del día y tiempo que no se cumplió
        If Range("H" & idRow(id_arr(j))).Value = "No cumple" Then       'buscamos el ID según la función
            With ws.Range("B" & j + 4).End(xlToRight).Offset(0, 1)
                .Value = DateSerial(Year(Date), Right(ActiveSheet.Name, 2), Left(ActiveSheet.Name, 2))
                .NumberFormat = "[$-C0A]d-mmm;@"
            End With
            With ws.Range("B" & j + 4).End(xlToRight).Offset(0, 1)
                .Value = ActiveSheet.Range("F" & idRow(id_arr(j))).Value
                .NumberFormat = "[$-F400]h:mm:ss AM/PM"
            End With
        End If
        
        'Array con horas trabajadas
        If Range("F" & idRow(id_arr(j))).Value <> "NO MARCO" Then
            sumArr(i, j) = Sheets(i).Range("F" & idRow(id_arr(j))).Value
        Else
            sumArr(i, j) = 0
        End If
    Next j
Next i

ws.Activate
ws.Range("D:I").Insert Shift:=xlToRight


'------RESUMIMOS Y CALCULAMOS DATOS SOBRE LAS LLEGADAS TARDE-----'
Dim dSum As Double, col_total_horas As Integer, col_días_con_datos As Integer
Dim col_NO_MARCO As Integer, col_neto_días As Integer, col_hs_prom As Integer
Dim col_régimen As Integer, col_dif_tiempo As Integer

col_total_horas = 4
col_días_con_datos = 5
col_NO_MARCO = 6
col_neto_días = 7
col_hs_prom = 8
col_régimen = 9
col_dif_tiempo = 10

On Error Resume Next
For i = 1 To UBound(sumArr, 2)
    'En la columna D ponemos la suma de todas las horas del mes
    With Application.WorksheetFunction
        dSum = .Sum(.Index(sumArr, 0, i))
    End With
    ws.Cells(i + 4, col_total_horas).Value = dSum * 24
    ws.Cells(i + 4, col_total_horas).Style = "Comma"
    ws.Cells(i + 4, col_días_con_datos).Value = mdResNeto.lastWs - 3 'contamos los días con datos menos las 3 1ras pestañas
    ws.Cells(i + 4, col_NO_MARCO).Value = Sheets("Tarde").Range("E" & i + 4).Value  'Traemos los NO MARCO de la pestaña Tarde
    ws.Cells(i + 4, col_neto_días).Value = _
        ws.Cells(i + 4, col_días_con_datos).Value - ws.Cells(i + 4, col_NO_MARCO).Value
    With ws.Cells(i + 4, col_hs_prom)       'distinguimos el caso que el neto de días sea 0, para evitar el error de división /0
        If ws.Cells(i + 4, col_neto_días).Value = 0 Then
            .Value = 0
        Else
            .Value = (ws.Cells(i + 4, col_total_horas).Value) / (ws.Cells(i + 4, col_neto_días).Value) / 24
            .NumberFormat = "[$-F400]h:mm:ss AM/PM"
        End If
    End With
    ws.Cells(i + 4, col_régimen).Value = Sheets("Resumen").Cells(i + 7, 5)
    With ws.Cells(i + 4, col_dif_tiempo)    'como excel no muestra tiempos negativos, hay que formatearlo "a mano"
        If ws.Cells(i + 4, col_hs_prom).Value < ws.Cells(i + 4, col_régimen).Value / 24 Then
            .Value = "-" & Format(Abs(Cells(i + 4, col_hs_prom).Value - Cells(i + 4, col_régimen).Value / 24), "h:mm:ss")
            .NumberFormat = "[$-F400]h:mm:ss AM/PM"
            .Font.Color = -11489280
        Else
            .Value = ws.Cells(i + 4, col_hs_prom).Value - ws.Cells(i + 4, col_régimen).Value / 24
            .NumberFormat = "[$-F400]h:mm:ss AM/PM"
        End If
    End With
Next


ws.Cells(4, col_total_horas).Value = "Total Horas Mes"
ws.Cells(4, col_total_días).Value = "Días Mes"
ws.Cells(4, col_NO_MARCO).Value = "Días sin marcar"
ws.Cells(4, col_neto_días).Value = "Neto días"
ws.Cells(4, col_hs_prom).Value = "Hs promedio"
ws.Cells(4, col_régimen).Value = "Régimen"
ws.Cells(4, col_dif_tiempo).Value = "Dif Tiempo"


mdResNeto.formatData

ws.Cells.EntireColumn.AutoFit
ws.Outline.ShowLevels ColumnLevels:=1

End Sub

Sub formatData()

lcol = mdResTarde.LastCol(Sheets("Neto"))

With Range(Cells(4, 11), Cells(4, lcol))
    .EntireColumn.Group
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
    .Font.Bold = True
    .Value = "Resumen Horarios"
    With .Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End With

End Sub
'Vemos hasta qué hoja hay datos cargados para ver hasta qué ws correr la macro
Function lastWs()

x = 3

For i = 4 To Sheets.Count
    If IsEmpty(Sheets(i).Range("F3")) = False Then
        x = x + 1
    Else
        GoTo Result
    End If
Next

Result:
lastWs = x

End Function
