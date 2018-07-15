Attribute VB_Name = "AAMacroThisWorkbook"
Option Explicit

Dim ws As Worksheet, i As Integer

Sub Format()

Set ws = ThisWorkbook.Sheets(1)

    On Error Resume Next
    
        ws.Activate
        
        DeleteCols
        
        OrderCols
        
        SortFields
        
        DeleteRows 4, "CORPORATE ACTION"
        
        MoveRows 1, "20845"
        MoveRows 1, "109866"
        MoveRows 9, "0,00"
        
        Sheets("0,00").Name = "Administración"
        
        MoveRowsByDate 12
    
        ws.ShowAllData
        
        Pintar
        
        For i = 1 To Sheets.Count
            Sheets(i).Activate
            Formatws
        Next i

ws.Activate

End Sub


Sub DeleteCols()

Range("B:E, G:K, M:M, O:O, R:S, U:W, AA:AA, AC:AE, AH:AK, AM:AP, AR:AR, AT:AV , AX:CN").Delete shift:=xlToLeft 'Borrar columnas


End Sub


Sub OrderCols() 'Ordena las columnas de las planillas y renombra los títulos
    
Dim ArrOrdenCols(1 To 17) As Integer

'ArrOrdenCols: array con orden de columnas según planilla
ArrOrdenCols(1) = 1
ArrOrdenCols(2) = 6
ArrOrdenCols(3) = 7
ArrOrdenCols(4) = 10
ArrOrdenCols(5) = 2
ArrOrdenCols(6) = 3
ArrOrdenCols(7) = 4
ArrOrdenCols(8) = 13
ArrOrdenCols(9) = 14
ArrOrdenCols(10) = 15
ArrOrdenCols(11) = 16
ArrOrdenCols(12) = 11
ArrOrdenCols(13) = 12
ArrOrdenCols(14) = 17
ArrOrdenCols(15) = 8
ArrOrdenCols(16) = 9
ArrOrdenCols(17) = 5


    'Para ordenar las columnas, insertamos una fila al inicio, pegamos los array con el orden y nombre de
    'las columnas y luego las ordenamos.

    With ActiveSheet
    
          Range("A1").EntireRow.Insert
          Range("A1:Q1").Value = ArrOrdenCols
          .Sort.SortFields.Clear
          .Sort.SortFields.Add Key:= _
          Range("A1:Q1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
          
             With .Sort
              .SetRange Range("A1").CurrentRegion
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlLeftToRight
              .SortMethod = xlPinYin
              .Apply
             End With
             
          Range("A1").EntireRow.Delete
          Range("R1").Value = "Comentario"
                  
        
    End With

End Sub


Sub Formatws()


Dim ColW(1 To 18) As Variant, i As Integer

    ColW(1) = 10
    ColW(2) = 14.43
    ColW(3) = 20.29
    ColW(4) = 10
    ColW(5) = 16.14
    ColW(6) = 12.14
    ColW(7) = 10
    ColW(8) = 13.71
    ColW(9) = 13.57
    ColW(10) = 10
    ColW(11) = 10
    ColW(12) = 10
    ColW(13) = 10
    ColW(14) = 10
    ColW(15) = 10
    ColW(16) = 31
    ColW(17) = 22.57
    ColW(18) = 13.29

    
    For i = 1 To 18
        Columns(i).ColumnWidth = ColW(i)
    Next

    Rows(1).RowHeight = 38.25
    Rows(1).WrapText = True
    
    ActiveWindow.Zoom = 80
            
    Range("A1:R1").Font.Bold = True
    Range("A1").CurrentRegion.Font.Name = "Arial"
    Range("A1").CurrentRegion.Font.Size = 10
      
     'Freezeamos la primer fila
     With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

    
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
            ("L2:L" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        .Sort.SortFields.Add Key:=Range _
            ("I2:I" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        .Sort.SortFields.Add Key:=Range _
            ("G2:G" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        .Sort.SortFields.Add Key:=Range _
            ("E2:E" & Lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            
        With .Sort
            .SetRange Range("A1:R" & Lrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    End With

End Sub



'////Sub para borrar filas de una tabla que empieza en el rango A1, según el criterio "deletecriteria" en
' la columna "ColCriteria\\\\

Sub DeleteRows(ColCriteria As Integer, Deletecriteria As String)


Dim cl As Range, rng As Range, Lrow As Long


    Range("A1").CurrentRegion.AutoFilter Field:=ColCriteria, Criteria1:=Deletecriteria
    

    If Range("A" & Rows.Count).End(xlUp).Row = 1 Then
        
        Exit Sub
        
    Else
        
        Range("A1").CurrentRegion.Offset(1, 0).EntireRow.Delete
        
    End If

ws.ShowAllData

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

'Movemos las filas que tengan fcha mayor a hoy(según la fecha en la columna "colcriteria")

Sub MoveRowsByDate(ColCriteria As Integer)

Dim dDate As Date, lDate As Long
Dim cl As Range, rng As Range, Lrow As Long, Hoy As Date
Set ws = Sheets(1)

    
    dDate = DateSerial(Year(Date), Month(Date), Day(Date))
    lDate = dDate


    Range("A1").CurrentRegion.AutoFilter Field:=ColCriteria, Criteria1:=">" & lDate


    If Range("A" & Rows.Count).End(xlUp).Row = 1 Then
        Exit Sub
    Else
        'Agregamos una hoja con el nombre "Operaciones Futuras"
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "Operaciones Futuras"
        
        'Filtramos según el criterio y cortamos y pegamos los datos
        ws.Range("A1").CurrentRegion.Copy Destination:=Sheets("Operaciones Futuras").Range("A1")
        Worksheets("Operaciones Futuras").Columns(ColCriteria).NumberFormat = "dd/mm/yyyy"
        ws.Range("A1").CurrentRegion.Offset(1, 0).EntireRow.Delete
    End If

    ws.Activate
    ws.ShowAllData


End Sub


'Pintamos las filas con mercado FRB

Sub Pintar()

Dim Lrow As Integer
Set ws = Sheets(1)

Lrow = ws.Range("A1").End(xlDown).Row

'Pintar FRB
For i = 2 To Lrow

    If Left(Range("G" & i), 3) = "FRB" Then
        With Range("A" & i & ":R" & i).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If
    
Next i
    
    
    
End Sub
