Attribute VB_Name = "mdOrderData"
Option Explicit

Dim ws As Worksheet, i As Integer

Sub Format()

Set ws = ThisWorkbook.Sheets(1)

    On Error Resume Next
    
        ws.Activate
        
        DeleteCols
        
        OrderCols
        
        SortFields
        
        MoveRows 1, "14923"
        
        ws.ShowAllData
        
        For i = 1 To Sheets.Count
            Sheets(i).Activate
            Formatws
        Next i

ws.Activate

End Sub


Sub DeleteCols()

Range("B:B, I:I, K:L, P:Q, X:X, Z:AE, AG:AG").Delete shift:=xlToLeft 'Borrar columnas


End Sub


Sub OrderCols() 'Ordena las columnas de las planillas y renombra los títulos
    
Dim ArrOrdenCols(1 To 19) As Integer

'ArrOrdenCols: array con orden de columnas según planilla
ArrOrdenCols(1) = 2
ArrOrdenCols(2) = 3
ArrOrdenCols(3) = 4
ArrOrdenCols(4) = 1
ArrOrdenCols(5) = 6
ArrOrdenCols(6) = 7
ArrOrdenCols(7) = 9
ArrOrdenCols(8) = 8
ArrOrdenCols(9) = 11
ArrOrdenCols(10) = 12
ArrOrdenCols(11) = 15
ArrOrdenCols(12) = 16
ArrOrdenCols(13) = 14
ArrOrdenCols(14) = 13
ArrOrdenCols(15) = 17
ArrOrdenCols(16) = 18
ArrOrdenCols(17) = 19
ArrOrdenCols(18) = 5
ArrOrdenCols(19) = 10



    'Para ordenar las columnas, insertamos una fila al inicio, pegamos los array con el orden y nombre de
    'las columnas y luego las ordenamos.

    With ActiveSheet
    
          Range("A1").EntireRow.Insert
          Range("A1:S1").Value = ArrOrdenCols
          .Sort.SortFields.Clear
          .Sort.SortFields.Add Key:= _
          Range("A1:S1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
          
             With .Sort
              .SetRange Range("A1").CurrentRegion
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlLeftToRight
              .SortMethod = xlPinYin
              .Apply
             End With
             
          Range("A1").EntireRow.Delete
          Range("T1").Value = "Comentario"
                  
    End With

End Sub


Sub Formatws()


Dim colw(1 To 19) As Variant, i As Integer

    colw(1) = 8
    colw(2) = 4.71
    colw(3) = 20.86
    colw(4) = 12
    colw(5) = 6.57
    colw(6) = 12.43
    colw(7) = 7.29
    colw(8) = 11.29
    colw(9) = 14.43
    colw(10) = 11.71
    colw(11) = 18.71
    colw(12) = 11.71
    colw(13) = 13.86
    colw(14) = 4.86
    colw(15) = 15.86
    colw(16) = 14.86
    colw(17) = 16.29
    colw(18) = 18.29
    colw(19) = 17

    
    For i = 1 To 19
        Columns(i).ColumnWidth = colw(i)
    Next

    
    ActiveWindow.Zoom = 80
            
    Range("A1:T1").Font.Bold = True
    Range("A1").CurrentRegion.Font.Name = "Arial"
    Range("A1").CurrentRegion.Font.Size = 10
      
     'Freezeamos la primer fila
     With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
    Range("K:K").NumberFormat = "0"
    Columns("K:K").EntireColumn.AutoFit
    
End Sub


Sub SortFields()

Dim MyTable As Range, Lrow As Long

Set MyTable = Range("A1").CurrentRegion
Lrow = MyTable.Rows.Count


    With Sheets(1)
    
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range _
            ("A2:A" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortTextAsNumbers
        .Sort.SortFields.Add Key:=Range _
            ("I2:I" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        .Sort.SortFields.Add Key:=Range _
            ("O2:O" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            
        With .Sort
            .SetRange Range("A1:T" & Lrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    End With

End Sub


'////Sub para borrar filas de una tabla que empieza en el rango A1, según el criterio "MoveCriteria" en
' la columna "ColCriteria\\\\

Sub MoveRows(ColCriteria As Integer, MoveCriteria As Variant)


Dim cl As Range, rng As Range, Lrow As Long
Set ws = Sheets(1)

    Range("A1").CurrentRegion.AutoFilter Field:=ColCriteria, Criteria1:=MoveCriteria
    

    If Range("A" & Rows.Count).End(xlUp).Row = 1 Then
        
        Exit Sub
        
    Else
        
        'Agregamos una hoja con el nombre del campo a mover
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = MoveCriteria & ""
        
        'Filtramos según el criterio y cortamos y pegamos los datos
        ws.Range("A1").CurrentRegion.Copy Destination:=Sheets(MoveCriteria & "").Range("A1")
        ws.Range("A1").CurrentRegion.Offset(1, 0).EntireRow.Delete
        
    End If


ws.Activate
ws.ShowAllData

End Sub

