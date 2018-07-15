Attribute VB_Name = "mdPend"

Sub OrderPend()

Dim ws3 As Worksheet, ws10 As Worksheet
Dim lrow3 As Integer, lrow10 As Integer

Set ws3 = Sheets("Pend > 3 días")
Set ws10 = Sheets("Pend > 10 días")

lrow3 = ws3.Cells(Rows.Count, 1).End(xlUp).Row + 1
lrow10 = ws10.Cells(Rows.Count, 1).End(xlUp).Row + 1

ws3.Rows("2:" & lrow3).Delete
ws10.Rows("2:" & lrow10).Delete


    For i = 2 To 6
        Sheets(i).Activate
        ThreeDays
        TenDays
    Next

End Sub


'Filtrados de pendientes de entre 3 y 10 días

Sub ThreeDays()

Dim ws As Worksheet
Dim dd0 As Date, dd3 As Date, dd10 As Date
Dim d0 As Long, d3 As Long, d10 As Long

dd0 = Date
dd3 = Date - 3
dd10 = Date - 10

dd0 = DateSerial(Year(dd0), Month(dd0), Day(dd0))
dd3 = DateSerial(Year(dd3), Month(dd3), Day(dd3))
dd10 = DateSerial(Year(dd10), Month(dd10), Day(dd10))

d0 = dd0
d3 = dd3
d10 = dd10


Set ws = ActiveSheet

    'Primero filtamos los MT con fechas de captura entre 4 y 10 días
    ws.Range("A1").CurrentRegion.AutoFilter Field:=10, Criteria1:="<" & d3, Operator:=xlAnd
    ws.Range("A1").CurrentRegion.AutoFilter Field:=10, Criteria1:=">" & d10
    
    'Filtro por MT566 y por pendiente de gestión
    ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:="MT566"
    ws.Range("A1").CurrentRegion.AutoFilter Field:=11, Criteria1:="Pendiente (de gestión)"
    

    'Copiamos los datos a la pestaña
    If Cells(Rows.Count, 1).End(xlUp).Row = 1 Then

        Exit Sub

    Else
    
        Dim rTable As Range, lHeadersRows As Long
        Set rTable = ws.Range("A1").CurrentRegion
        lHeadersRows = rTable.ListHeaderRows

        Set rTable = rTable.Resize(rTable.Rows.Count - lHeadersRows)
        Set rTable = rTable.Offset(1)

        rTable.Copy Destination:=Sheets("Pend > 3 días").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)

    End If


ws.Activate

    ws.ShowAllData
    ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:="MT566"
    

End Sub

'Filtrados de pendientes de entre 3 y 10 días

Sub TenDays()

Dim ws As Worksheet
Dim dd0 As Date, dd3 As Date, dd10 As Date
Dim d0 As Long, d3 As Long, d10 As Long

dd0 = Date
dd3 = Date - 3
dd10 = Date - 10

dd0 = DateSerial(Year(dd0), Month(dd0), Day(dd0))
dd3 = DateSerial(Year(dd3), Month(dd3), Day(dd3))
dd10 = DateSerial(Year(dd10), Month(dd10), Day(dd10))

d0 = dd0
d3 = dd3
d10 = dd10


Set ws = ActiveSheet

    'Primero filtamos los MT con fechas de captura entre 4 y 10 días
    ws.Range("A1").CurrentRegion.AutoFilter Field:=10, Criteria1:="<" & d10
    
    'Filtro por MT566 y por pendiente de gestión
    ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:="MT566"
    ws.Range("A1").CurrentRegion.AutoFilter Field:=11, Criteria1:="Pendiente (de gestión)"
    

    'Copiamos los datos a la pestaña
    If Cells(Rows.Count, 1).End(xlUp).Row = 1 Then

        Exit Sub

    Else
    
        Dim rTable As Range, lHeadersRows As Long
        Set rTable = ws.Range("A1").CurrentRegion
        lHeadersRows = rTable.ListHeaderRows

        Set rTable = rTable.Resize(rTable.Rows.Count - lHeadersRows)
        Set rTable = rTable.Offset(1)

        rTable.Copy Destination:=Sheets("Pend > 10 días").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)

    End If


ws.Activate

    ws.ShowAllData
    ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:="MT566"
    

End Sub

