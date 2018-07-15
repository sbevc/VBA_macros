Attribute VB_Name = "mdControl"
Option Explicit

Public ws As Worksheet
Public lrow As Long, i As Integer

Sub Separar_Datos()

    CopyRows 5, "BNP"
    CopyRows 5, "BONY"
    CopyRows 5, "CLEARSTREAM"
    CopyRows 5, "SOCIETE"
    CopyRows 5, "BNP"
    FilterOthers

End Sub

'Pintamos las cancelaciones de confirmación de rojo
Sub Canc_Conf()

Set ws = Sheets(1)
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Primero dejamos todos los MT566 en negrita para borrar los rojos anteriores
    With ws.Range("A1").CurrentRegion.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With


    For i = 2 To lrow
        
        If Range("C" & i).Value = "CANC. CONFIRMACION" Then
            With Range("A" & i & ":S" & i).Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        End If
    
    Next i
    

End Sub


Sub DeleteCols()
    
Set ws = Sheets(1)
    
    ws.Range("B:C, F:F, H:H, K:K, P:S").Delete Shift:=xlToLeft 'Borrar columnas
    ws.Range("A1").End(xlToRight).Offset(0, 1).Value = "Comentario"
    ws.Columns("A:K").EntireColumn.AutoFit
    
    
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Orden según ISIN y Tipo de operación
    ws.Range("A1").CurrentRegion.AutoFilter
    With ws.AutoFilter.Sort
    
        .SortFields.Clear
        .SortFields.Add _
            Key:=Range("C2:C" & lrow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SortFields.Add _
            Key:=Range("B2:B" & lrow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
    
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    
    
End Sub

Sub Listado_Comentarios()

Set ws = Sheets(1)
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'Para la columna comentario insertamos una lista predeterminada. La lista está en la celda "A1" _
    'de la pestaña "Comentario"(oculta)
    
        With ws.Range("K2:K" & lrow).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=Comentarios"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    
End Sub

Public Sub CopyRows(ColCriteria As Integer, MoveCriteria As Variant)


Dim rTable As Range, lHeadersRows As Long

Set ws = Sheets(1)

    'Primero filtamos por las <> ok liquidadas, y desp según el criterio
    ws.Range("A1").CurrentRegion.AutoFilter Field:=11, Criteria1:="Pendiente (de gestión)", Operator:=xlAnd
    ws.Range("A1").CurrentRegion.AutoFilter Field:=ColCriteria, Criteria1:="*" & MoveCriteria & "*"


    If Cells(Rows.Count, 1).End(xlUp).Row = 1 Then
        
        Exit Sub
        
    Else
            
        Set rTable = ws.Range("A1").CurrentRegion
        lHeadersRows = rTable.ListHeaderRows
        
        Set rTable = rTable.Resize(rTable.Rows.Count - lHeadersRows)
        Set rTable = rTable.Offset(1)

        rTable.Copy Destination:=Sheets(MoveCriteria & "").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
        
    End If


ws.Activate
ws.ShowAllData

End Sub

'Sub para filtrar otros custodios
Sub FilterOthers()

Dim rng As Range, c As Range
Dim rTable As Range, lHeadersRows As Long

Set ws = Sheets(1)

    On Error Resume Next
    
    ws.Cells.AutoFilter
    
    'Agregamos una columna al principio para discriminar cuales son "otros cusodios" _
    'y luego filtramos y copiamos la data según la columna agregada
    ws.Columns("A:A").Insert Shift:=xlToRight
    Set rng = ws.Range("F1", Range("F2").End(xlDown))
    
    For Each c In rng
    
        Select Case c
        
            Case "BONY", "SOCIETE PARIS", "CLEARSTREAM", "BNP MILAN", "BNP PARIS", "Custodio"
                Cells(c.Row, 1).Value = c.Value
            
            Case Else
                Cells(c.Row, 1).Value = "OTROS"
        
        End Select
      
    Next
    
    
    Set rTable = ws.Range("A1").CurrentRegion
    
    rTable.AutoFilter Field:=12, Criteria1:="Pendiente (de gestión)", Operator:=xlAnd
    rTable.AutoFilter Field:=1, Criteria1:="OTROS"
    
    If ws.Cells(Rows.Count, 1).End(xlUp).Row = 1 Then
        
        ws.Columns("A:A").Delete Shift:=xlToLeft
        ws.ShowAllData
        Exit Sub
    
    Else
    
        lHeadersRows = rTable.ListHeaderRows
        
        Set rTable = rTable.Resize(rTable.Rows.Count - lHeadersRows)
        Set rTable = rTable.Offset(1, 1)
    
        rTable.Copy Destination:=Sheets("OTROS CUSTODIOS").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
        
        ws.Columns("A:A").Delete Shift:=xlToLeft

    End If
    
    ws.ShowAllData

End Sub



