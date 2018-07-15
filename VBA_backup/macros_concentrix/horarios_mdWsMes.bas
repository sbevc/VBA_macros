Attribute VB_Name = "mdWsMes"
'Crea una pestaña por cada día hábil del mes(lunes a viernes) según el mes del combobox
Sub CreatewsMes()

'-----Primero vemos si hay pestañas creadas y preguntamos si desea borrarlas-----'
Dim answer As Integer

If Sheets.Count > 3 Then

    answer = MsgBox("Desea borrar las pestañas actuales y crear nuevas?", vbYesNo)
    If answer = vbYes Then
        For i = Sheets.Count To 1 Step -1
            If Sheets(i).Name <> "Resumen" And Sheets(i).Name <> "Tarde" And Sheets(i).Name <> "Neto" Then
            Application.DisplayAlerts = False
                Sheets(i).Delete
            Application.DisplayAlerts = True
            End If
        Next
    Else
        Exit Sub
    End If
    
End If

'-----Creamos las pestañas en función del mes del combobox-----'
Dim arrDays(1 To 2, 1 To 12) As Variant, xday As Date, xdaystr As String, xmonthstr As String, xMonth As Integer
Dim startCol As Integer, startRow As Integer
Dim ws As Worksheet, tbl As Range

Set ws = Sheets("Resumen")
Set tbl = ws.Cells(Rows.Count, 2).End(xlUp).CurrentRegion

startCol = 2
startRow = 2

On Error Resume Next

For i = 1 To 31
    
    xMonth = Sheets(1).cmbMonths.Value
    xday = DateSerial(Year(Date), xMonth, i)
    xdaystr = Format(day(xday), "00")
    xmonthstr = Format(xMonth, "00")
    
    If Weekday(xday, vbMonday) <> 6 And Weekday(xday, vbMonday) <> 7 Then
    
        Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = xdaystr & "-" & xmonthstr
        
        'Ya agregamos los títulos de los datos y los ID y nombres de la pestaña resumen
        ActiveSheet.Range("B60") = "Time"
        ActiveSheet.Range("C60") = "Event"
        ActiveSheet.Range("D60") = "ID"
        ActiveSheet.Range("E60") = "Nombre"
        ActiveSheet.Range("F60") = "Device"
        
        tbl.Resize(tbl.Rows.Count, tbl.Columns.Count - 2).Copy Destination:=ActiveSheet.Range("B2")
        
        Cells(startRow, startCol).Value = "ID"
        Cells(startRow, startCol + 1).Value = "Nombre"
        Cells(startRow, startCol + 2).Value = "Hora Entrada"
        Cells(startRow, startCol + 3).Value = "En hora?"
        Cells(startRow, startCol + 4).Value = "Tiempo total"
        Cells(startRow, startCol + 5).Value = "Régimen"
        Cells(startRow, startCol + 6).Value = "Cumple?"
        
        Range("I2, K2, M2, O2, Q2, S2").Value = "Entrada"
        Range("J2, L2, N2, P2, R2, T2").Value = "Salida"
        
    End If
    
Next


End Sub
