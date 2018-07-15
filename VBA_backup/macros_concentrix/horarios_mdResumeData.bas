Attribute VB_Name = "mdResumeData"
Sub ResumeData()

Dim arrRng As Range, id_arr() As Variant
Dim ws As Worksheet
Dim lrow As Long, lcol As Integer
Dim fixedH As Date, hEntrada As Date        'fixedH horario de entrada +16, hEntrada = hora que marcó
Dim regimen As Double, timeWorked As Double

id_col = 3
data_Row = 60
Title = "B60:F60"

Set ws = ActiveSheet

id_arr = Application.WorksheetFunction.Transpose(Range("B3", Range("B3").End(xlDown)))


'Filtramos por cada ID del regristro, si hay datos los pegamos y si no, "NO MARCO"
For i = 1 To UBound(id_arr)

    ws.Range(Title).AutoFilter Field:=id_col, Criteria1:=id_arr(i) & ""     'Filtramos por cada uno de los nombres
    lrow = ws.Cells(Rows.Count, 2).End(xlUp).Row
    
    If lrow = 60 Then
        Cells(i + 2, 6).Value = "NO MARCO"
        
    Else
        'separamos entre la noche y los que entran "normal" para el cálculo del tiempo/llegadas tarde
        Select Case id_arr(i)
            Case Is = 231376, 160085
                turno_noche i + 2       'Pasamos la fila al subprocedure
            Case Else
                'copia los tiempos filtrados
                ws.Range("B" & lrow, Range("B60").Offset(1, 0)).Copy
                Cells(i + 2, 9).PasteSpecial Paste:=xlPasteAll, _
                    Operation:=xlNone, SkipBlanks:=False, Transpose:=True
                'Tiempo total trabajado
                lcol = Cells(i + 2, 9).End(xlToRight).Column
                Cells(i + 2, 6).Value = Cells(i + 2, lcol).Value - Cells(i + 2, 9).Value
        End Select
        
        'En hora?
        Cells(i + 2, 4).Value = Application.WorksheetFunction.VLookup _
                                (id_arr(i), Sheets(1).Cells(Rows.Count, 2).End(xlUp).CurrentRegion, 3, 0)  'Traemos la hora de entrada
        
        fixedH = DateAdd("n", 16, Cells(i + 2, 4).Value)        'DateAdd("n", 16, Cells(i + 2, 3)) suma 16 minutos
        hEntrada = Cells(i + 2, 9).Value
        If hEntrada <= fixedH Then
            Cells(i + 2, 5).Value = "En hora"
        Else
            Cells(i + 2, 5).Value = "Llegada tarde"
        End If
        
        'Cumple régimen?
        Cells(i + 2, 7).Value = Application.WorksheetFunction.VLookup _
                                (id_arr(i), Sheets(1).Cells(Rows.Count, 2).End(xlUp).CurrentRegion, 4, 0)  'Traemos la hora de entrada
        
        regimen = (Cells(i + 2, 7).Value - 10 / 60) / 24    'Al régimen le restamos 10 minutos de margen para ver si cumple o no
        timeWorked = Cells(i + 2, 6).Value
        If timeWorked >= regimen Then
            Cells(i + 2, 8).Value = "Cumple"
        Else
            Cells(i + 2, 8).Value = "No cumple"
        End If
        
    End If
    
Next

Cells.AutoFilter

FormatData2

End Sub

Sub turno_noche(id_row As Integer)
    
Dim ws As Worksheet, lrow As Integer
Set ws = ActiveSheet
lrow = ws.Cells(Rows.Count, 2).End(xlUp).Row

'copia los tiempos filtrados
ws.Range("B" & lrow, Range("B60").Offset(1, 0)).Copy
Cells(id_row, 9).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    

'Hacemos dos arrays con los tiempos marcados, uno con aquellos mayores a 0.5 y otro con los menores
Dim arr_max() As Double, arr_min() As Double, colData As Integer
Dim hEntrada As Double, hSalida As Double, timeWorked As Double

colData = 9
j = 1
k = 1
Do Until IsEmpty(Cells(id_row, colData)) = True
    If Cells(id_row, colData).Value > 0.5 Then
        ReDim Preserve arr_max(1 To j)
        arr_max(j) = Cells(id_row, colData).Value
        j = j + 1
    Else
        ReDim Preserve arr_min(1 To k)
        arr_min(k) = Cells(id_row, colData).Value
        k = k + 1
    End If
    colData = colData + 1
Loop

On Error Resume Next
hEntrada = Application.WorksheetFunction.Min(arr_max)
hSalida = Application.WorksheetFunction.Max(arr_min)

timeWorked = 1 - hEntrada + hSalida
With ws.Cells(id_row, 6)
    .Value = timeWorked
    .NumberFormat = "[$-F400]h:mm:ss AM/PM"
End With

With ws.Cells(id_row, 9)
    .Value = hEntrada
    .NumberFormat = "[$-F400]h:mm:ss AM/PM"
End With

End Sub

Sub FormatData2()

Dim lrow As Integer, tbl As Range

lrow = Range("B60").End(xlUp).Row

Range("D3:D" & lrow).NumberFormat = "[$-F400]h:mm:ss AM/PM"
Range("F3:F" & lrow).NumberFormat = "[$-F400]h:mm:ss AM/PM"
Range("F3:F" & lrow).HorizontalAlignment = xlRight

Columns("B:B").ColumnWidth = 7.67
Columns("C:C").ColumnWidth = 22.89
Columns("E:E").ColumnWidth = 12
Columns("F:F").ColumnWidth = 12

ActiveWindow.DisplayGridlines = False


Set tbl = Range("B2:H" & lrow)

    With tbl.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tbl.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tbl.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tbl.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tbl.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With tbl.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

Range("B2:H2").Font.Bold = True

End Sub

Sub asda()

Dim arr_hs() As Variant
arr_hs = Range(Cells(42, 9), Cells(42, 9).End(xlToRight))

'For i = 1 To UBound(arr_hs)
'    Debug.Print arr_hs(i)
'Next

Debug.Print Application.WorksheetFunction.Max(arr_hs)

End Sub

Sub alsdkjas()

Dim arr_max() As Double, arr_min() As Double, colData As Integer

colData = 9
j = 1
k = 1
Do Until IsEmpty(Cells(57, colData)) = True
    If Cells(57, colData).Value > 0.5 Then
        ReDim Preserve arr_max(1 To j)
        arr_max(j) = Cells(57, colData).Value
        j = j + 1
    Else
        ReDim Preserve arr_min(1 To k)
        arr_min(k) = Cells(57, colData).Value
        k = k + 1
    End If
    colData = colData + 1
Loop

End Sub
