Attribute VB_Name = "mdSwifts"
Sub swifts()

Dim ColWidth(1 To 15) As Variant
Dim MyArr() As Variant, Lrow As Long

ColWidth(1) = 19
ColWidth(2) = 2.43
ColWidth(3) = 14.14
ColWidth(4) = 9.71
ColWidth(5) = 9.71
ColWidth(6) = 9.71
ColWidth(7) = 9.71
ColWidth(8) = 10.71
ColWidth(9) = 10.71
ColWidth(10) = 10.71
ColWidth(11) = 10.71
ColWidth(12) = 10.71
ColWidth(13) = 10.71
ColWidth(14) = 3.57
ColWidth(15) = 9.71


Sheets(1).Activate

Range("B:B, D:D, F:F, K:K, O:O, Q:Q, S:S, V:V, W:W, X:X").Delete shift:=xlToLeft 'Borrar columnas
Range("A1", Range("A1").End(xlToRight)).Font.Bold = True
Columns("L:M").NumberFormat = "#,##0.00"

'Cuadrícula
With Range("A1").CurrentRegion.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("A1").CurrentRegion.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("A1").CurrentRegion.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("A1").CurrentRegion.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("A1").CurrentRegion.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("A1").CurrentRegion.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Orden
Lrow = Range("A1").End(xlDown).Row

With Sheets(1).Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("E2:E" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    .SortFields.Add Key:=Range("C2:C" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    .SortFields.Add Key:=Range("D2:D" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    .SetRange Range("A1:O" & Lrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'ColWidths
For i = 1 To 15
    Columns(i).ColumnWidth = ColWidth(i)
Next i

'Pintar 109803, JPY, FECHAS DE HOY, Borrar BRL y ARG
For i = 2 To Lrow
    If Left(Range("H" & i), 6) = "109803" Then
        With Range("A" & i & ":O" & i).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    ElseIf Range("E" & i).Value = Date Then
        With Range("A" & i & ":O" & i).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    ElseIf Left(Range("N" & i), 3) = "JPY" Then
        With Range("A" & i & ":O" & i).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    ElseIf Left(Range("N" & i), 3) = "BRL" Then
        Rows(i).Delete
    ElseIf Left(Range("N" & i), 3) = "ARG" Then
        Rows(i).Delete
    End If
Next i

End Sub

