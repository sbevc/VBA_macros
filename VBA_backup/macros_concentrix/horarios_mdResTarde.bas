Attribute VB_Name = "mdResTarde"
Sub Resumen_llegadasTarde()

Dim ws As Worksheet
Set ws = Sheets("Tarde")

If IsEmpty(ws.Range("d5")) = True Then
    Llegadas_tarde mdResNeto.lastWs
Else
    Dim answer As Integer
    answer = MsgBox("Ya existen datos, desea sobreescribirlos?", vbYesNo, "Sobreescribir datos?")
    
    If answer = vbYes Then
        clearData ws
        Llegadas_tarde mdResNeto.lastWs
    Else
        Exit Sub
    End If
End If


End Sub


'Resumen de llegadas tarde hasta la última pestaña con datos
Sub Llegadas_tarde(lastWsWithData As Integer)

Dim id_arr() As Variant, ws As Worksheet

Set ws = Sheets("Tarde")
id_arr = Application.WorksheetFunction.Transpose(ws.Range("B5", Range("B5").End(xlDown)))

'----------RESUMEN LLEGADAS TARDE----------''loopeamos por las ws y ponemos los días de las llegadas tarde y la hora de llegada
For i = 4 To lastWsWithData
    Sheets(i).Activate
    For j = 1 To UBound(id_arr)
        If Range("E" & idRow(id_arr(j))).Value = "Llegada tarde" Then       'buscamos el ID según la función
            With ws.Range("B" & j + 4).End(xlToRight).Offset(0, 1)
                .Value = DateSerial(Year(Date), Right(ActiveSheet.Name, 2), Left(ActiveSheet.Name, 2))
                .NumberFormat = "[$-C0A]d-mmm;@"
            End With
            With ws.Range("B" & j + 4).End(xlToRight).Offset(0, 1)
                .Value = ActiveSheet.Range("I" & idRow(id_arr(j))).Value
                .NumberFormat = "[$-F400]h:mm:ss AM/PM"
            End With
        End If
    Next j
Next i

'----------CONTEO LLEGADAS TARDE----------'Insertamos una columna y contamos las llegadas tarde
ws.Range("D:D").EntireColumn.Insert Shift:=xlToRight
Dim lcol As Integer
For i = 1 To UBound(id_arr)
    If IsEmpty(ws.Range("E" & i + 4)) = True Then
        ws.Range("D" & i + 4).Value = 0
    Else
        lcol = ws.Range("E" & i + 4).End(xlToRight).Column
        ws.Range("D" & i + 4).Value = (lcol - 4) / 2
    End If
Next i


'----------RESUMEN NO MARCÓ----------'Ponemos los no marcados luego de las llegadas tarde
Dim lcol2 As Integer
lcol2 = LastCol(ws)

For i = 4 To lastWsWithData
    Sheets(i).Activate
    For j = 1 To UBound(id_arr)
        If Range("F" & idRow(id_arr(j))).Value = "NO MARCO" Then       'buscamos el ID según la función
            If IsEmpty(ws.Cells(j + 4, lcol2 + 1)) = True Then
                With ws.Cells(j + 4, lcol2 + 1)
                    .Value = DateSerial(Year(Date), Right(ActiveSheet.Name, 2), Left(ActiveSheet.Name, 2))
                    .NumberFormat = "[$-C0A]d-mmm;@"
                End With
            Else
                With ws.Cells(j + 4, Columns.Count).End(xlToLeft).Offset(0, 1)
                    .Value = DateSerial(Year(Date), Right(ActiveSheet.Name, 2), Left(ActiveSheet.Name, 2))
                    .NumberFormat = "[$-C0A]d-mmm;@"
                End With
            End If
        End If
    Next j
Next i


'----------CONTEO NO MARCÓ----------'Insertamos una columna y contamos las llegadas tarde
ws.Activate
ws.Range("E:E").EntireColumn.Insert Shift:=xlToRight
For i = 1 To UBound(id_arr)
    If IsEmpty(ws.Cells(i + 4, lcol2 + 2)) = True Then
        ws.Range("E" & i + 4).Value = 0
    Else
        lcol = ws.Cells(i + 4, Columns.Count).End(xlToLeft).Column
         ws.Range(Cells(i + 4, lcol2 + 2), Cells(i + 4, lcol)).Select
         ws.Range("E" & i + 4).Value = _
            Application.WorksheetFunction.CountA(ws.Range(Cells(i + 4, lcol2 + 2), Cells(i + 4, lcol)))
    End If
Next i

mdResTarde.formatData lcol2 + 1

ws.Cells(4, 4).Value = "#Llegadas tarde"
ws.Cells(4, 5).Value = "#NO MARCO"

ws.Cells.EntireColumn.AutoFit
ws.Outline.ShowLevels ColumnLevels:=1

End Sub

'Hallamos el número de fila en que está el id buscado
Function idRow(ID) As Integer

Dim foundCell As Range, lookRange As Range
Set foundCell = Range("B3:B53").Find(What:=ID)

idRow = foundCell.Row

End Function

'Hallamos la última columna con datos para agregar los datos de "NO MARCO" LUEGO
Function LastCol(ws As Worksheet) As Integer

Dim rLastCell As Range

Set rLastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

LastCol = rLastCell.Column

End Function
'Pasamos la última columna con llegadas tarde. Agrupamos las llegadas tarde y los NO MARCO
Sub formatData(colTarde As Integer)

Dim ws As Worksheet
Set ws = Sheets("Tarde")

ws.Cells(1, colTarde + 1).EntireColumn.Insert Shift:=xlToRight
ws.Range(Cells(1, 6), Cells(1, colTarde)).EntireColumn.Group

Dim lcol As Integer
lcol = LastCol(ws)
ws.Range(Cells(1, colTarde + 2), Cells(1, lcol)).EntireColumn.Group

With Range(Cells(4, 6), Cells(4, colTarde))
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
    .Value = "Resumen llegadas tarde"
    With .Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
End With


With Range(Cells(4, colTarde + 2), Cells(4, lcol))
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
    .Value = "Resumen NO MARCO"
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
End With

End Sub

'Cuando ya hay datos en el resumen, borramos los ya existentes y creamos uno nuevo
Sub clearData(clearws As Worksheet)

Dim clearrng As Range
Set clearrng = clearws.Range("D:D", Range("D:D").End(xlToRight))

On Error Resume Next
With clearrng

    .ClearContents
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext

    .UnMerge
    
    .Columns.Ungroup
   
    .EntireColumn.Hidden = False
    With .Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End With

End Sub
