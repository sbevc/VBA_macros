Attribute VB_Name = "PLanillasSOCIETE"
Option Explicit

Dim Lrow As Long, i As Long, j As Long, k As Long, MyTable As Range
Public Ws As Worksheet


'===================================================MAIN==========================================================
'=================================================================================================================

Sub CreateWsSOCIETE()

Set Ws = Sheets("Datos")

    DataCheck
    ClearNoSOCIETE
    DeleteBBVA
    
    
    'Chequeamos si está el cliente de ZURICH para cambiar los datos
    Dim txt As String, FindRng As Range, FoundRng As Range
        txt = "W0072130H"
        Set FindRng = Ws.Range("P:P")
        Set FoundRng = FindRng.Find(txt)

        If Not FoundRng Is Nothing Then
            EditZURICH
        End If
        
    
    Concatenate
    Parse_Data_Cuentas
    
    For j = 3 To Sheets.Count
        Sheets(j).Activate
        PDOrder
        FillData
    Next j
    
    For k = 3 To Sheets.Count
        Sheets(k).Activate
        FormatSOCIETE
    Next k


End Sub

'Chequeamos que hayan datos en la celda A1 y si hay mas de un wb abierto, que active este.
Sub DataCheck()

Set Ws = Sheets("Datos")

    If IsEmpty(Ws.Range("A1")) = True Then
        MsgBox "Insertar datos antes de ejecutar la macro"
        End
    ElseIf Workbooks.Count > 1 Then
        ThisWorkbook.Sheets("Datos").Activate
    End If
    
End Sub
    
    
'Dejamos solo las cuentas que son de bony, el resto las borramos
Sub ClearNoSOCIETE()

Set Ws = Sheets("Datos")

    Ws.Range("A1").CurrentRegion.AutoFilter Field:=10, Criteria1:="<>BBVA/*", Operator:=xlAnd
    Lrow = Ws.Range("A1").End(xlDown).Row
    Ws.Rows("2:" & Lrow).Delete
    Ws.Range("X:X").Delete
    Ws.AutoFilterMode = False

End Sub

Sub DeleteBBVA()

Dim rng As Range, c As Range
Dim ArrAcc() As Variant

Set Ws = Sheets("Datos")

Lrow = Ws.Range("A1").End(xlDown).Row
Set rng = Ws.Range("J2:J" & Lrow)

    'Sacamos los BBVA de las cuentas
    rng.Replace What:="BBVA/", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False

    'Sacamos los espacios del nombre de las cuentas
    ArrAcc = rng.Value
    For i = 1 To UBound(ArrAcc, 1)
        ArrAcc(i, 1) = Application.WorksheetFunction.Trim(ArrAcc(i, 1))
    Next i

    rng.Value = ArrAcc
    
    
End Sub

Sub Concatenate()

Dim ISINCol As Integer, AccCol As Integer, PayDateCol As Integer

Set Ws = Sheets("Datos")

    Lrow = Ws.Range("A" & Rows.Count).End(xlUp).Row
    ISINCol = 4
    AccCol = 10
    Ws.Cells(1, 1).Value = "Nombre"
    
    'Concatenar
    For i = 2 To Lrow
        Ws.Cells(i, 1) = Cells(i, ISINCol) & " " & Cells(i, AccCol)
    Next i

End Sub


Sub Parse_Data_Cuentas()

Dim lr As Long
Dim Ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim MyArr As Variant
Dim Title As String
Dim TitleRow As Integer

    vcol = 1
    Set Ws = Sheets("Datos")
    lr = Ws.Cells(Ws.Rows.Count, vcol).End(xlUp).Row
    Title = "A1:W1"
    TitleRow = Ws.Range(Title).Cells(1).Row
    icol = Ws.Columns.Count
    Ws.Cells(1, icol) = "Unique"
    
    For i = 2 To lr
    On Error Resume Next
    If Ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(Ws.Cells(i, vcol), Ws.Columns(icol), 0) = 0 Then
    Ws.Cells(Ws.Rows.Count, icol).End(xlUp).Offset(1) = Ws.Cells(i, vcol)
    End If
    Next
    MyArr = Application.WorksheetFunction.Transpose(Ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    Ws.Columns(icol).Clear
    For i = 2 To UBound(MyArr)
    Ws.Range(Title).AutoFilter Field:=vcol, Criteria1:=MyArr(i) & ""
    If Not Evaluate("=ISREF('" & MyArr(i) & "'!A1)") Then
    Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = MyArr(i) & ""
    Else
    Sheets(MyArr(i) & "").Move After:=Worksheets(Worksheets.Count)
    End If
    Ws.Range("A" & TitleRow & ":A" & lr).EntireRow.Copy Sheets(MyArr(i) & "").Range("A1")
    Sheets(MyArr(i) & "").Columns.AutoFit
    Next
    Ws.AutoFilterMode = False
    Ws.Activate
    
    Erase MyArr
        
End Sub

'Ordena las columnas de las planillas y renombra los títulos planilla 15%
Sub Reorder_15()
    
Dim ArrOrdenCols(1 To 26) As Integer, ArrNombreCols(1 To 10) As String

    'ArrOrdenCols: array con orden de columnas según planilla
    ArrOrdenCols(1) = 13
    ArrOrdenCols(2) = 14
    ArrOrdenCols(3) = 3
    ArrOrdenCols(4) = 1
    ArrOrdenCols(5) = 2
    ArrOrdenCols(6) = 16
    ArrOrdenCols(7) = 17
    ArrOrdenCols(8) = 15
    ArrOrdenCols(9) = 19
    ArrOrdenCols(10) = 20
    ArrOrdenCols(11) = 5
    ArrOrdenCols(12) = 21
    ArrOrdenCols(13) = 22
    ArrOrdenCols(14) = 23
    ArrOrdenCols(15) = 26
    ArrOrdenCols(16) = 8
    ArrOrdenCols(17) = 7
    ArrOrdenCols(18) = 4
    ArrOrdenCols(19) = 9
    ArrOrdenCols(20) = 11
    ArrOrdenCols(21) = 12
    ArrOrdenCols(22) = 25
    ArrOrdenCols(23) = 24
    ArrOrdenCols(24) = 18
    ArrOrdenCols(25) = 6
    ArrOrdenCols(26) = 10


    'ArrNombreCols: array con nombre de columnas según planilla
    ArrNombreCols(1) = "ISIN"
    ArrNombreCols(2) = "Security name"
    ArrNombreCols(3) = "PayDate: example yyyymmdd"
    ArrNombreCols(4) = "Legal status 1 = Individual 2 = Corporation 3 = CIV 4 = Pension fund 5 = other"
    ArrNombreCols(5) = "Position"
    ArrNombreCols(6) = "Country code:example:ES(for Spain)"
    ArrNombreCols(7) = "Beneficial owner's name"
    ArrNombreCols(8) = "Tax identification number"
    ArrNombreCols(9) = "Address (column 1):Street X nbr Y…"
    ArrNombreCols(10) = "Address (column 2):ZIP code City name"

    'Reordemanos las columnas según las planillas

    With ActiveSheet
          Range("A1").EntireRow.Insert
          Range("A1:Z1").Value = ArrOrdenCols
          .Sort.SortFields.Clear
          .Sort.SortFields.Add Key:= _
          Range("A1:Z1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
          xlSortNormal
             With .Sort
              .SetRange Range("A1").CurrentRegion
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlLeftToRight
              .SortMethod = xlPinYin
              .Apply
             End With
          Range("A1").EntireRow.Delete
          Range("A1:j1").Value = ArrNombreCols  'Dejamos los titulos como en la planilla
    End With

End Sub

'Ordena las columnas de las planillas y renombra los títulos planilla CIV
Sub Reorder_CIV()
    
Dim ArrOrdenCols(1 To 27) As Integer, ArrNombreCols(1 To 9) As String

    'ArrOrdenCols: array con orden de columnas según planilla
    ArrOrdenCols(1) = 10
    ArrOrdenCols(2) = 11
    ArrOrdenCols(3) = 3
    ArrOrdenCols(4) = 1
    ArrOrdenCols(5) = 2
    ArrOrdenCols(6) = 13
    ArrOrdenCols(7) = 14
    ArrOrdenCols(8) = 12
    ArrOrdenCols(9) = 19
    ArrOrdenCols(10) = 20
    ArrOrdenCols(11) = 4
    ArrOrdenCols(12) = 18
    ArrOrdenCols(13) = 16
    ArrOrdenCols(14) = 17
    ArrOrdenCols(15) = 27
    ArrOrdenCols(16) = 21
    ArrOrdenCols(17) = 6
    ArrOrdenCols(18) = 22
    ArrOrdenCols(19) = 8
    ArrOrdenCols(20) = 23
    ArrOrdenCols(21) = 24
    ArrOrdenCols(22) = 25
    ArrOrdenCols(23) = 26
    ArrOrdenCols(24) = 15
    ArrOrdenCols(25) = 5
    ArrOrdenCols(26) = 7
    ArrOrdenCols(27) = 9


    'ArrNombreCols: array con nombre de columnas según planilla
    ArrNombreCols(1) = "ISIN"
    ArrNombreCols(2) = "Security name"
    ArrNombreCols(3) = "PayDate: example yyyymmdd"
    ArrNombreCols(4) = "Position"
    ArrNombreCols(5) = "Country code: example ES(for Spain)"
    ArrNombreCols(6) = "CIV 's name"
    ArrNombreCols(7) = "CIV 's ISIN***"
    ArrNombreCols(8) = "Address (column 1):Street X nbr Y…"
    ArrNombreCols(9) = "Address (column 2): ZIP code City name"


    'Reordemanos las columnas según las planillas

    With ActiveSheet
          Range("A1").EntireRow.Insert
          Range("A1:AA1").Value = ArrOrdenCols
          .Sort.SortFields.Clear
          .Sort.SortFields.Add Key:= _
          Range("A1:AA1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
          xlSortNormal
             With .Sort
              .SetRange Range("A1").CurrentRegion
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlLeftToRight
              .SortMethod = xlPinYin
              .Apply
             End With
          Range("A1").EntireRow.Delete
          Range("A1:I1").Value = ArrNombreCols  'Dejamos los titulos como en la planilla
    End With

End Sub

'Insertamos Paydates con inputboxes
Sub PayDate()

Dim PayDate As Variant, ISIN As String, Lrow As Long
Lrow = Range("A1").End(xlDown).Row
ISIN = Left(ActiveSheet.Name, 12)
        
    'Evaluamos si el ISIN de la hoja actual es igual al de la hoja anterior. Si son diferentes,
    'pedimos los nuevos datos, sino copiamos los de la hoja anterior.
    If ISIN <> Left(Sheets(ActiveSheet.Index - 1).Name, 12) Then
        
        PayDate = InputBox("PayDate for ISIN " & ISIN & " (AAAA/MM/DD)")
        
        'Evaluamos si hay una sola fila con datos, para determinar dónde pegamos los valores
        'de las inputboxes
        If IsEmpty(Cells(3, 1)) = True Then
            Range("C2").Value = PayDate
            Range("C2").NumberFormat = "yyyy/mm/dd"
        Else
            Range("C2:C" & Lrow).Value = PayDate
            Range("C2:C" & Lrow).NumberFormat = "yyyy/mm/dd"
        End If
        
    Else
    
        If IsEmpty(Cells(3, 1)) = True Then
            Sheets(ActiveSheet.Index - 1).Range("C2").Copy Range("C2")
        Else
            Sheets(ActiveSheet.Index - 1).Range("C2").Copy Range("C2:C" & Lrow)
        End If
    
    End If

End Sub

'Damos formato a las planillas del 15%
Sub FormatPlanillas_15()
        
Dim TotalRow As Integer
        
    With ActiveSheet
        
        .Rows("1:1").RowHeight = 51
        .Range("K:R, U:Y").Delete
        .Columns("K:M").Hidden = True
        
        With .Range("A1").End(xlDown).CurrentRegion
            .Font.Name = "Arial"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        With Range("A1:J1")
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
            .Interior.Color = 12632256
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            .WrapText = True
            .EntireColumn.AutoFit
            
        End With
        
        'Insertamos la suma de la posición
        TotalRow = Range("E1").End(xlDown).Offset(1, 0).Row
        With Range("E" & TotalRow)
            .Formula = "=SUM(E1:E" & TotalRow - 1 & ")"
            .NumberFormat = "#,##0"
            .Font.Bold = True
        End With
        
    End With

End Sub

'Formato planillas CIV
Sub FormatPlanillas_CIV()
        
Dim TotalRow As Integer

    With ActiveSheet
        
        .Rows("1:1").RowHeight = 51
        .Range("J:R, V:Z").Delete
        .Columns("J:M").Hidden = True
        

        
        With .Range("A1").End(xlDown).CurrentRegion
            .Font.Name = "Arial"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        With Range("A1:I1")
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
            .Interior.Color = 12632256
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            .WrapText = True
            .EntireColumn.AutoFit
        End With
        
        'Insertamos la suma de la posición
        TotalRow = Range("D1").End(xlDown).Offset(1, 0).Row
        With Range("D" & TotalRow)
            .Formula = "=SUM(D1:D" & TotalRow - 1 & ")"
            .NumberFormat = "#,##0"
            .Font.Bold = True
        End With
        
      
        
        
        .Range("C" & TotalRow).Value = "Gross position ( :"
        .Range("C" & TotalRow).Font.Bold = True
        
        .Range("C" & TotalRow + 3).Value = "Gross dividend (total amount of the dividends )*"
        
        
        
        .Rows("1:3").Insert shift:=xlDown
        .Range("A1").Value = "Identity and complete address of the  account manager abroad"
        .Range("A2").Value = "Unit value of the coupon"
        .Range("B1").Value = "BANCO BILBAO VIZCAYA ARGENTARIA S.A. PLAZA SAN NICOLAS 4 48005 BILBAO SPAIN"
        .Range("B1").Font.Name = "Arial"
        .Range("B1").Font.Bold = True
            
        'Insertamos la multiplicación títulos*unitario
        TotalRow = Range("D1").End(xlDown).End(xlDown).Offset(1, 0).Row
        With Range("D" & TotalRow + 2)
            .Formula = "=D" & TotalRow - 1 & " * B2"
            .NumberFormat = "#,##0"
            .Font.Bold = True
        End With
        
            
    End With

End Sub

Sub PDOrder()

Dim ACC As Integer

'Tomamos los últimos 4 dígitosde la cuenta para diferenciar entre que formato aplicamos (15% o CIV)
ACC = Right(Range("I2").Value, 4)

    Select Case ACC
    
        Case Is = 3312, 3395, 3502, 3510, 3445, 3478, 3528, 3379, 3437, 3494, 3551
            
            PayDate
            Reorder_CIV
            
        Case Else
            
            PayDate
            Reorder_15

    End Select
        
        
End Sub

Sub FormatSOCIETE()

Dim ACC As Integer

'Tomamos los últimos 4 dígitosde la cuenta para diferenciar entre que formato aplicamos (15% o CIV)
ACC = Right(Range("S2").Value, 4)

    Select Case ACC
    
        Case Is = 3312, 3395, 3502, 3510, 3445, 3478, 3528, 3379, 3437, 3494, 3551
            
            FormatPlanillas_CIV
            
        Case Else
            
            FormatPlanillas_15

    End Select

End Sub

'Llenamos los datos del country Code y address(ZIP CODE & CITY)
Sub FillData()

Dim Frow As Integer, Lrow As Integer, LrISINs As Integer
Dim ACC As Integer
Dim ArrISINs() As Variant

ACC = Right(Range("S2").Value, 4)

Frow = Range("A" & Rows.Count).End(xlUp).Row

    'Separamos entre si hay una única fila de datos o màs filas
    If Frow = 2 Then
    
        Select Case ACC
        
            Case Is = 3312, 3395, 3502, 3510, 3445, 3478, 3528, 3379, 3437, 3494, 3551
            
                LrISINs = Sheets("ISIN CIVs").Range("A" & Rows.Count).End(xlUp).Row
                ArrISINs = Sheets("ISIN CIVs").Range("A1:B" & LrISINs).Value
                
                Range("E2").Value = Left(Range("Y2").Value, 2)
                Range("I2").Value = Range("W2") & " " & Range("X2")
                
                'Vlookup de los ISIN par alas CIV
                    On Error Resume Next
                    Range("G2").Value = Application.WorksheetFunction.VLookup(Trim(Range("U2")), ArrISINs, 2, 0)
                    If Err <> 0 Then
                        Range("G2").Value = "PEDIR ISIN"
                        With Range("G2").Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
            
            Case Else
    
                Range("F2").Value = Left(Range("Y2").Value, 2)
                Range("J2").Value = Range("K2") & " " & Range("L2")
                Range("D2").Value = 4
            End Select
    
    Else
    
        Lrow = Range("A" & Rows.Count).End(xlUp).Row
        
        Select Case ACC
        
            Case Is = 3312, 3395, 3502, 3510, 3445, 3478, 3528, 3379, 3437, 3494, 3551
            
            LrISINs = Sheets("ISIN CIVs").Range("A" & Rows.Count).End(xlUp).Row
            ArrISINs = Sheets("ISIN CIVs").Range("A1:B" & LrISINs).Value
                
                For i = 2 To Lrow
                
                    Range("E" & i).Value = Left(Range("Y" & i).Value, 2)
                    Range("I" & i).Value = Range("W" & i) & " " & Range("X" & i)
                    
                    'Vlookup de los ISIN par alas CIV
                    On Error Resume Next
                    Range("G" & i).Value = Application.WorksheetFunction.VLookup(Trim(Range("U" & i)), ArrISINs, 2, 0)
                    If Err <> 0 Then
                        Range("G" & i).Value = "PEDIR ISIN"
                        With Range("G" & i).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                        
                Next
            
                        
            Case Else
        
                For i = 2 To Lrow
                    Range("F" & i).Value = Left(Range("Y" & i).Value, 2)
                    Range("J" & i).Value = Range("K" & i) & " " & Range("L" & i)
                    Range("D" & i).Value = 4
                Next
    
            End Select
            
    End If
    


End Sub

'Editamos los datos de los clientes zurich
Sub EditZURICH()

Set Ws = Sheets("Datos")

'Filtramos por el TaxID del cliente a cambiar
Ws.Range("A1").CurrentRegion.AutoFilter Field:=16, Criteria1:="W0072130H"

Dim ColAddress As Integer, ColZipCode As Integer, ColCity As Integer, ColCountry As Integer

ColAddress = 19
ColZipCode = 20
ColCity = 21
ColCountry = 22

Dim cl As Range, rng As Range, x As Integer
Lrow = Ws.Range("A1").End(xlDown).Row
Set rng = Ws.Range("P2:P" & Lrow)

        For Each cl In rng.SpecialCells(xlCellTypeVisible)
            x = cl.Row
            Cells(x, ColAddress).Value = "ZURICH HOUSE, BALLSBRIGE PARK"
            Cells(x, ColZipCode).Value = ""
            Cells(x, ColCity).Value = "DUBLIN"
            Cells(x, ColCountry).Value = "IRELAND"
        Next cl

Ws.Cells.AutoFilter

End Sub

