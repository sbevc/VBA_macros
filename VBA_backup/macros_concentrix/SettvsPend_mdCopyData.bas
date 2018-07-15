Attribute VB_Name = "mdCopyData"
Option Explicit

Sub CopyData()

Dim wsData As Worksheet, wsPending As Worksheet, wsSett As Worksheet
Dim PendIsin As Integer, PendAcc As Integer, PendNet As Integer, SettIsin As Integer, SettAcc As Integer, SettNet As Integer

Set wsData = Sheets(1)
Set wsPending = Sheets(2)
Set wsSett = Sheets(3)

'Columnas a copiar
PendIsin = 48
PendAcc = 36
PendNet = 31

SettIsin = 61
SettAcc = 1
SettNet = 30

    
    'Reemplazamos . por . en la columna de netammount(reconoce el segundo punto como comma)
    wsPending.Columns(PendNet).Replace What:=".", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Ordenamos los datos
    OrderData wsPending, PendIsin, PendAcc
    OrderData wsSett, SettIsin, SettAcc
    
    
    'Copiamos los datos
    wsPending.Columns(PendIsin).Copy Destination:=wsData.Columns(1)
    wsPending.Columns(PendAcc).Copy Destination:=wsData.Columns(2)
    wsPending.Columns(PendNet).Copy Destination:=wsData.Columns(3)

    wsSett.Columns(SettIsin).Copy Destination:=wsData.Columns(7)
    wsSett.Columns(SettAcc).Copy Destination:=wsData.Columns(8)
    wsSett.Columns(SettNet).Copy Destination:=wsData.Columns(9)
    
    
    wsData.Columns("A:G").AutoFit
    
    Range("C:C, G:G").Style = "Comma"
    
End Sub

'Ordenamos datos en la worksheet ws y según dos columnas, de menor a mayor
Sub OrderData(ws As Worksheet, ColToOrder1 As Integer, ColToOrder2 As Integer)

    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Columns(ColToOrder1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=Columns(ColToOrder2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ws.Sort
        .SetRange ws.Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

