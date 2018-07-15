Attribute VB_Name = "mdAnálisis"
Option Explicit

Dim ws As Worksheet
Dim lrow As Long, i As Integer


Sub AnalizeData()

    Sumarceldas

    Results

    FormatData

End Sub



Sub Sumarceldas()

Dim ws As Worksheet
Dim lrow As Long

Set ws = Sheets(1)

    ws.Activate
    
    ws.Range("D1").Value = "S/Settle"
    ws.Range("E1").Value = "Dif"
    
    ws.Range("J1").Value = "S/Pending"
    ws.Range("K1").Value = "Dif"
    
    
    'Hacemos los sumifs
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lrow
        ws.Cells(i, 4).Value = CalcSumifs(9, 7, Cells(i, 1), 8, Cells(i, 2))
        ws.Cells(i, 5).Value = Cells(i, 3).Value - Cells(i, 4).Value
    Next i
        
    lrow = ws.Cells(Rows.Count, 7).End(xlUp).Row
    For i = 2 To lrow
        ws.Cells(i, 10).Value = CalcSumifs(3, 1, Cells(i, 7), 2, Cells(i, 8))
        ws.Cells(i, 11).Value = Cells(i, 9).Value - Cells(i, 10).Value
    Next i
    
End Sub

'Hacemos una función sumif que facilite los rangos
Function CalcSumifs(ColSumCrit As Integer, ColCrit1 As Integer, Crit1 As Variant, Colcrit2 As Integer, Crit2 As Variant)

Dim SumRng As Range, Crit1rng As Range, Crit2rng As Range

lrow = Cells(Rows.Count, ColSumCrit).End(xlUp).Row

Set SumRng = Range(Cells(2, ColSumCrit), Cells(lrow, ColSumCrit))
Set Crit1rng = Range(Cells(2, ColCrit1), Cells(lrow, ColCrit1))
Set Crit2rng = Range(Cells(2, Colcrit2), Cells(lrow, Colcrit2))


CalcSumifs = Application.WorksheetFunction.Sumifs(SumRng, Crit1rng, Crit1, Crit2rng, Crit2)

End Function

Sub Results()

Dim wsr As Worksheet
Set ws = Sheets("Datos")

    
    Set wsr = Sheets.Add(Before:=ws)
    wsr.Name = "Resultados"
    
    
    'Movemos los datos del pending con diferencias
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("A1:E1").Copy Destination:=wsr.Range("A1")
    
    For i = 2 To lrow
    
        If ws.Cells(i, 5).Value <> 0 Then
            ws.Range("A" & i & ":E" & i).Copy Destination:=wsr.Range("A" & Rows.Count).End(xlUp).Offset(1)
        End If
        
    Next i
    
    
    lrow = ws.Cells(Rows.Count, 7).End(xlUp).Row
    ws.Range("G1:K1").Copy Destination:=wsr.Range("G1")
    
    For i = 2 To lrow
    
        If ws.Cells(i, 11).Value <> 0 Then
            ws.Range("G" & i & ":K" & i).Copy Destination:=wsr.Range("G" & Rows.Count).End(xlUp).Offset(1)
        End If
        
    Next i
    
    
End Sub

Sub FormatData()
    
Dim wsr As Worksheet
Dim rng As Range

Set wsr = Sheets("Resultados")

    wsr.Rows(1).Insert xlDown
    wsr.Range("A1").Value = "Arreglos / Sin Confirmación"
    With Range("A1:E1")
    
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
        
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
    End With
    


   wsr.Range("G1").Value = "Arreglos / No previsadas"
   With Range("G1:K1")
    
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
        
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
    End With
    
  
  wsr.Columns("A:K").AutoFit
    
   
  Set rng = Union(wsr.Range("A1").CurrentRegion, wsr.Range("G1").CurrentRegion)
      
      rng.Borders(xlDiagonalDown).LineStyle = xlNone
      rng.Borders(xlDiagonalUp).LineStyle = xlNone
      With rng.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With rng.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With rng.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With rng.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With rng.Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With rng.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      
   wsr.Range("G1").CurrentRegion.Cut Destination:=wsr.Range("A1").End(xlDown).Offset(2)
   
   wsr.Rows(1).Insert xlDown
   wsr.Columns(1).Insert xlToRight
   
   wsr.Columns("A:K").AutoFit
   wsr.Columns(1).ColumnWidth = 2.14
     
   ActiveWindow.DisplayGridlines = False
     
End Sub
