Attribute VB_Name = "btnCrearPlanillas"
Option Explicit

Dim LRow As Long, i As Long, MyTable As Range
Public Ws As Worksheet


'===================================================MAIN==========================================================
'=================================================================================================================

Sub CreateWS()


Dim txt As String, FindRng As Range, FoundRng As Range, Emision As String


    DataCheck '()
    ClearNoBNY '()
    
    
        'Chequeamos antes de seguir con la macro que luego de borrar las cuentas que no son de BNY, sigan quedando cuentas
        If IsEmpty(Cells(2, 1)) = True Then
            MsgBox "No se encontraron cuentas de BONY"
            Exit Sub
        End If
        
    
    Concatenate '()
    WriteInvType '()
    BuscarCLientRefs '()
        
        'Chequeamos que hayan "PEDIR CLIENTREF" antes de correr las macros de las nuevas refs, sino
        'hacemos GoTo creando las pestañas de las cuentas.
        txt = "PEDIR CLIENTREF"
        LRow = Ws.Range("A1").End(xlDown).Row
        Set FindRng = Sheets(1).Range("W2:W" & LRow)
        Set FoundRng = FindRng.find(txt)

        If Not FoundRng Is Nothing Then
            PestañaNuevasRefs '()
            If Sheets("Nuevas Refs").Range("A" & Rows.Count).End(xlUp).Row <> 1 Then
                Parse_Data_NuevasRefs '() Verificamos antes que halla alguna referencia ESPAÑOLA que crear
            End If
        End If
        
        
    Parse_Data_Cuentas '()
    
    
        For i = 2 To Sheets.Count
        
            Sheets(i).Activate
            Emision = Left(ActiveSheet.Name, 2)
            
            Select Case Emision
            
                Case Is = "FR", "IT", "IE", "FI", "SE", "NO", "PT"
                    ReorderPlanillas
                    FillDataPlanillas
                    TaxRates
                    EventID_PayDate
                    FormatPlanillas
                    
                Case Is = "BO"
                    ReorderBOs
                    FillDataBOs
                    FormatBOs
                    
            End Select
                
        Next i


End Sub

'Chequeamos que hayan datos en la celda A1 y si hay mas de un wb abierto, que active este.
Sub DataCheck()

    If IsEmpty(Sheets(1).Range("A1")) = True Then
        MsgBox "Insertar datos antes de ejecutar la macro"
        End
    ElseIf Workbooks.Count > 1 Then
        ThisWorkbook.Sheets(1).Activate
    End If
    
End Sub
    
    
'Dejamos solo las cuentas que son de bony, el resto las borramos
Sub ClearNoBNY()

Dim ACCsBNY(1 To 14, 1 To 1) As Variant, ArrBNY() As Variant, ArrACC() As Variant

    'Listado de cuentas de bony que llevan planilla
    ACCsBNY(1, 1) = 109020
    ACCsBNY(2, 1) = 109798
    ACCsBNY(3, 1) = 109799
    ACCsBNY(4, 1) = 109803
    ACCsBNY(5, 1) = 109860
    ACCsBNY(6, 1) = 109861
    ACCsBNY(7, 1) = 109862
    ACCsBNY(8, 1) = 109863
    ACCsBNY(9, 1) = 109874
    ACCsBNY(10, 1) = 109882
    ACCsBNY(11, 1) = 498593
    ACCsBNY(12, 1) = 498652
    ACCsBNY(13, 1) = 188604
    ACCsBNY(14, 1) = 109872

LRow = Range("A1").End(xlDown).Row
ArrACC = Range("I2:I" & LRow).Value

    ReDim ArrBNY(1 To UBound(ArrACC, 1), 1 To 1)
    For i = 1 To UBound(ArrACC, 1)
        On Error Resume Next
        ArrBNY(i, 1) = Application.WorksheetFunction.VLookup(ArrACC(i, 1), ACCsBNY, 1, 0)
        If Err <> 0 Then
        ArrBNY(i, 1) = "NO BONY"
        End If
    Next i
    
    Range("X2:X" & LRow).Value = ArrBNY

Erase ACCsBNY
Erase ArrACC
Erase ArrBNY

    Range("A1").CurrentRegion.AutoFilter Field:=24, Criteria1:="NO BONY"
    LRow = Range("A1").End(xlDown).Row
    Rows("2:" & LRow).Delete
    Range("X:X").Delete
    Sheets(1).AutoFilterMode = False

End Sub

Sub Concatenate()

Dim ISINCol As Integer, AccCol As Integer, PayDateCol As Integer
Set Ws = Sheets(1)

    LRow = Range("A" & Rows.Count).End(xlUp).Row
    ISINCol = 4
    AccCol = 9
    Ws.Cells(1, 1).Value = "Nombre"
    PayDateCol = 8

    'Reemplazamos "/" con "." en la columna PayDate
    Ws.Columns(PayDateCol).Replace What:="/", Replacement:=".", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
    
    'Concatenar
    For i = 2 To LRow
        Cells(i, 1) = Cells(i, ISINCol) & " " & Cells(i, AccCol)
    Next i

End Sub

Sub WriteInvType()

Dim VlookupArray(1 To 9, 1 To 2) As Variant, ArrayNum1() As Variant, ArrayInvType() As Variant
Set Ws = Sheets(1)


    LRow = Ws.Range("A1").End(xlDown).Row
    
    VlookupArray(1, 1) = 109020
    VlookupArray(2, 1) = 109798
    VlookupArray(3, 1) = 109799
    VlookupArray(4, 1) = 109860
    VlookupArray(5, 1) = 109861
    VlookupArray(6, 1) = 109863
    VlookupArray(7, 1) = 109872
    VlookupArray(8, 1) = 109882
    VlookupArray(9, 1) = ""
    
    VlookupArray(1, 2) = "INDIVIDUAL"
    VlookupArray(2, 2) = "INVEST FUND"
    VlookupArray(3, 2) = "PENSION FUND"
    VlookupArray(4, 2) = "INVEST FUND"
    VlookupArray(5, 2) = "PENSION FUND"
    VlookupArray(6, 2) = "SICAV"
    VlookupArray(7, 2) = "PENSION FUND"
    VlookupArray(8, 2) = "PENSION FUND"
    VlookupArray(9, 2) = ""
    
    ArrayNum1 = Ws.Range("I2:I" & LRow).Value
    
    ReDim ArrayInvType(1 To UBound(ArrayNum1, 1), 1 To 1)
    For i = 1 To UBound(ArrayNum1)
        On Error Resume Next
        ArrayInvType(i, 1) = Application.WorksheetFunction.VLookup(ArrayNum1(i, 1), VlookupArray, 2, 0)
        If Err <> 0 Then
        ArrayInvType(i, 1) = "Investor Type not Found"
        End If
    Next i
    
    Ws.Range("R2:R" & LRow).Value = ArrayInvType

End Sub

'Buscamos los clientref en el archivo "CLIENT REFERENCE(ULTIMO)". Si no los encuentra escribimos "PEDIR CLIENTREF"

Sub BuscarCLientRefs()

'===================PARTE I - Trimm los TxID===================
'Definimos la matriz txID y la le sacamos los espacios con la función trim.

Dim TxID() As Variant
Set Ws = Sheets(1)
LRow = Range("A1").End(xlDown).Row

    
    TxID = Ws.Range("P2:P" & LRow).Value
    For i = 1 To UBound(TxID, 1)
        TxID(i, 1) = Application.WorksheetFunction.Trim(TxID(i, 1))
    Next i
    

'===================PARTE II - VLOOKUP===================
'(ACTUALIZADA): En vez de hacer VLookUp, definimos un Dictionary y buscamos los valores en el diccionario.

Dim ClientRef() As Variant
Dim Name As String, Path As String, FileName As String

    Path = "H:\SC000068\OPERACIONES FINANCIERAS\INTERESES Y DIVIDENDOS\IMPUESTOS\INTERNACIONAL\BONY\EVENTOS\EVENTOS 2017\"
    Name = "CLIENT REFERENCE(ULTIMO).xlsx"
    Workbooks.Open (Path & Name)
    Workbooks("CLIENT REFERENCE(ULTIMO)").Sheets(1).Activate
    
'Definimos el diccionario
Dim dict As Dictionary, rng As Range, c As Range
Set dict = New Dictionary
LRow = Range("A" & Rows.Count).End(xlUp).Row
Set rng = Range("A2:A" & LRow)
    
     For Each c In rng
        dict.Add Key:=(c.Value), Item:=(c.Offset(0, 1).Value)
    Next
    
    Workbooks("CLIENT REFERENCE(ULTIMO)").Close
    
    
'Buscamos los ClientRefs en el diccionario

    ReDim ClientRef(1 To UBound(TxID, 1), 1 To 1)
    For i = 1 To UBound(TxID, 1)
        'Cuando el TxID no esté vacio, que haga busque en el dictionary
        If TxID(i, 1) <> "" Then
            If dict.Exists(TxID(i, 1)) Then
                ClientRef(i, 1) = dict(TxID(i, 1))
            Else
                ClientRef(i, 1) = "PEDIR CLIENTREF"
            End If
        'Cuando el TxID SI esté vacio, que el client ref lo deje vacío tamb.
        ElseIf TxID(i, 1) = "" Then
            ClientRef(i, 1) = ""
        End If
    Next i
    
    Ws.Activate
    Range("W2:W" & LRow).Value = ClientRef

    Erase TxID
    Erase ClientRef
    
    
End Sub


Sub PestañaNuevasRefs()

Set Ws = Sheets(1)

    'Filtramos y copiamos las filas con "pedir client refs" a una hoja nueva
    Ws.Range("A1").CurrentRegion.AutoFilter Field:=23, Criteria1:="PEDIR CLIENTREF"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Nuevas Refs"
    Ws.Range("A1").CurrentRegion.Copy Sheets("Nuevas Refs").Range("A1")
    
    'Sacamos de la hoja "Nuevas Refs" aquellos clientes que no son españoles ya que nose mandan.
    ActiveSheet.Cells(2, 22).Activate
    Do Until IsEmpty(ActiveCell) = True
        If Left(ActiveCell.Value, 4) <> "ESPA" Then
        Rows(ActiveCell.Row).Delete
        ActiveCell.Offset(-1, 0).Select
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
    
    'Como los datos quedaron filtrados en la hoja 1, loopeamos por el filtro (por las celdas visibles) y _
    cambiamos "Pedir client ref" por los taxID
    Ws.Activate
    
    Dim cl As Range, rng As Range, x As Integer
    LRow = Range("A1").End(xlDown).Row
    Set rng = Range("W2:W" & LRow)
    
        For Each cl In rng.SpecialCells(xlCellTypeVisible)
            x = cl.Row
            cl.Value = Range("P" & x).Value
        Next cl
        
        Ws.AutoFilterMode = False
    

End Sub

'Separamos las nuevas refs en pestañas nuevas
Sub Parse_Data_NuevasRefs()

Dim lr As Long
Dim Ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim Title As String
Dim TitleRow As Integer
    
    vcol = 1
    Set Ws = Sheets("Nuevas Refs")
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
    myarr = Application.WorksheetFunction.Transpose(Ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    Ws.Columns(icol).Clear
    For i = 2 To UBound(myarr)
    Ws.Range(Title).AutoFilter Field:=vcol, Criteria1:=myarr(i) & ""
    If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
    Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "BO " & myarr(i)
    Else
    Sheets("BO " & myarr(i)).Move After:=Worksheets(Worksheets.Count)
    End If
    Ws.Range("A" & TitleRow & ":A" & lr).EntireRow.Copy Sheets("BO " & myarr(i)).Range("A1")
    Sheets("BO " & myarr(i)).Columns.AutoFit
    Next
    Ws.AutoFilterMode = False
    Ws.Activate
    
    Erase myarr
    
End Sub

Sub Parse_Data_Cuentas()

Dim lr As Long
Dim Ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim Title As String
Dim TitleRow As Integer

    vcol = 1
    Set Ws = Sheets(1)
    lr = Ws.Cells(Ws.Rows.Count, vcol).End(xlUp).Row
    Title = "A1:X1"
    TitleRow = Ws.Range(Title).Cells(1).Row
    icol = Ws.Columns.Count
    Ws.Cells(1, icol) = "Unique"
    
    For i = 2 To lr
    On Error Resume Next
    If Ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(Ws.Cells(i, vcol), Ws.Columns(icol), 0) = 0 Then
    Ws.Cells(Ws.Rows.Count, icol).End(xlUp).Offset(1) = Ws.Cells(i, vcol)
    End If
    Next
    myarr = Application.WorksheetFunction.Transpose(Ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    Ws.Columns(icol).Clear
    For i = 2 To UBound(myarr)
    Ws.Range(Title).AutoFilter Field:=vcol, Criteria1:=myarr(i) & ""
    If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
    Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
    Else
    Sheets(myarr(i) & "").Move After:=Worksheets(Worksheets.Count)
    End If
    Ws.Range("A" & TitleRow & ":A" & lr).EntireRow.Copy Sheets(myarr(i) & "").Range("A1")
    Sheets(myarr(i) & "").Columns.AutoFit
    Next
    Ws.AutoFilterMode = False
    Ws.Activate
    
    Erase myarr
        
End Sub

Sub ReorderPlanillas() 'Ordena las columnas de las planillas y renombra los títulos
    
Dim ArrOrdenCols(1 To 28) As Integer, ArrNombreCols(1 To 19) As String

    'ArrOrdenCols: array con orden de columnas según planilla
    ArrOrdenCols(1) = 28
    ArrOrdenCols(2) = 21
    ArrOrdenCols(3) = 22
    ArrOrdenCols(4) = 2
    ArrOrdenCols(5) = 1
    ArrOrdenCols(6) = 4
    ArrOrdenCols(7) = 23
    ArrOrdenCols(8) = 3
    ArrOrdenCols(9) = 5
    ArrOrdenCols(10) = 24
    ArrOrdenCols(11) = 7
    ArrOrdenCols(12) = 25
    ArrOrdenCols(13) = 26
    ArrOrdenCols(14) = 27
    ArrOrdenCols(15) = 20
    ArrOrdenCols(16) = 8
    ArrOrdenCols(17) = 6
    ArrOrdenCols(18) = 9
    ArrOrdenCols(19) = 10
    ArrOrdenCols(20) = 11
    ArrOrdenCols(21) = 12
    ArrOrdenCols(22) = 13
    ArrOrdenCols(23) = 16
    ArrOrdenCols(24) = 14
    ArrOrdenCols(25) = 15
    ArrOrdenCols(26) = 17
    ArrOrdenCols(27) = 18
    ArrOrdenCols(28) = 19

    'ArrNombreCols: array con nombre de columnas según planilla
    ArrNombreCols(1) = "Security Name"
    ArrNombreCols(2) = "Security ID"
    ArrNombreCols(3) = "Pay Date(MM-DD-YYYY)"
    ArrNombreCols(4) = "Security Location"
    ArrNombreCols(5) = "Account Number"
    ArrNombreCols(6) = "Underlying Owner Name"
    ArrNombreCols(7) = "Share Holdings"
    ArrNombreCols(8) = "Tax ID"
    ArrNombreCols(9) = "Investor Type"
    ArrNombreCols(10) = "Address"
    ArrNombreCols(11) = "Zip Code"
    ArrNombreCols(12) = "City of Residence"
    ArrNombreCols(13) = "Country of Residence"
    ArrNombreCols(14) = "Withholding Rate (0.00)"
    ArrNombreCols(15) = "Currency Code"
    ArrNombreCols(16) = "Cln - Ref - ID"
    ArrNombreCols(17) = "Pool Flag(Y/N)"
    ArrNombreCols(18) = "Option Number"
    ArrNombreCols(19) = "Notification/Event ID"

    'Para ordenar las columnas, insertamos una fila al inicio, pegamos los array con el orden y nombre de
    'las columnas y luego las ordenamos.

    With ActiveSheet
          Range("A1").EntireRow.insert
          Range("A1:AB1").Value = ArrOrdenCols
          .Sort.SortFields.Clear
          .Sort.SortFields.Add Key:= _
          Range("A1:AB1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
          xlSortNormal
             With Sheets(i).Sort
              .SetRange Range("A1").CurrentRegion
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlLeftToRight
              .SortMethod = xlPinYin
              .Apply
             End With
          Range("A1").EntireRow.Delete
          Range("A1:S1").Value = ArrNombreCols  'Dejamos los titulos como en la planilla
    End With

End Sub

'Llenamos las columnas de currency, poolflag y optionnumber
'Borramos la última fila y dejamos solo el subtotal de títulos

Sub FillDataPlanillas()

Dim PoolFlag As String, OptionNumber As String, ArrVlookuP(1 To 7, 1 To 3) As Variant
PoolFlag = "N"
OptionNumber = "'001"
LRow = ActiveSheet.Range("A1").End(xlDown).Row

'Array con Security location(2) y currency(3) por cada país
ArrVlookuP(1, 1) = "FR"
ArrVlookuP(2, 1) = "IE"
ArrVlookuP(3, 1) = "IT"
ArrVlookuP(4, 1) = "FI"
ArrVlookuP(5, 1) = "SE"
ArrVlookuP(6, 1) = "NO"
ArrVlookuP(7, 1) = "PT"

ArrVlookuP(1, 2) = "BPR"
ArrVlookuP(2, 2) = "DTC"
ArrVlookuP(3, 2) = "BCI"
ArrVlookuP(4, 2) = "SEF"
ArrVlookuP(5, 2) = "SEB"
ArrVlookuP(6, 2) = "SEN"
ArrVlookuP(7, 2) = "CGA"

ArrVlookuP(1, 3) = "EUR"
ArrVlookuP(2, 3) = "USD"
ArrVlookuP(3, 3) = "EUR"
ArrVlookuP(4, 3) = "EUR"
ArrVlookuP(5, 3) = "SEK"
ArrVlookuP(6, 3) = "NOK"
ArrVlookuP(7, 3) = "EUR"
  
        With ActiveSheet
            
            '//Separamos según si hay una o más filas de datos para rellenar los poolflags y optionnumbers
            If IsEmpty(Range("A3")) = False Then
            
                On Error Resume Next 'Insertamos currency
                .Range("O2:O" & LRow).Value = Application.WorksheetFunction.VLookup(Left(Sheets(i).Name, 2), ArrVlookuP, 3, 0)
                
                If Err <> 0 Then
                    MsgBox "Currency no encontrada para emisión " & Left(Sheets(i).Name, 2)
                End If
                
                On Error Resume Next 'Insertamos Security location
                .Range("D2:D" & LRow).Value = Application.WorksheetFunction.VLookup(Left(Sheets(i).Name, 2), ArrVlookuP, 2, 0)
                
                If Err <> 0 Then
                    MsgBox "Security location no encontrada para emisión " & Left(Sheets(i).Name, 2)
                End If
                
                .Range("Q2:Q" & LRow).Value = PoolFlag
                .Range("R2:R" & LRow).Value = OptionNumber
                .Rows(LRow + 1).Delete
                
                    With .Range("G" & LRow + 1)
                        .Formula = "=sum(G2:G" & LRow & ")"
                        '.Value = Application.WorksheetFunction.Sum(Sheets(i).Range("G2:G" & LRow))
                        .NumberFormat = "#,##0"
                        .Font.Bold = True
                    End With
                    
            Else
            
                On Error Resume Next 'Insertamos currency
                .Range("O2").Value = Application.WorksheetFunction.VLookup(Left(Sheets(i).Name, 2), ArrVlookuP, 3, 0)
                
                If Err <> 0 Then
                    MsgBox "Currency no encontrada para emisión " & Left(Sheets(i).Name, 2)
                End If
                
                On Error Resume Next 'Insertamos Security location
                .Range("D2").Value = Application.WorksheetFunction.VLookup(Left(Sheets(i).Name, 2), ArrVlookuP, 2, 0)
                
                If Err <> 0 Then
                    MsgBox "Security location no encontrada para emisión " & Left(Sheets(i).Name, 2)
                End If
                
                .Range("Q2").Value = PoolFlag
                .Range("R2").Value = OptionNumber
                .Rows(3).Delete
                
                    With .Range("G3")
                        .Value = Range("G2").Value
                        .NumberFormat = "#,##0"
                        .Font.Bold = True
                    End With
                    
            End If
            
        End With

End Sub


Sub TaxRates()     'Pone los tax solo para planillas con más de una línea

Dim ArrMaxRate As Variant, ArrTxRate As Variant
Dim MaxFR As Double, MinFR As Double, MaxIE As Double, MinIE As Double, MaxIT As Double, MinIT As Double
Dim MaxNO As Double, MinNO As Double
Dim Emision As String, j As Integer

MaxFR = 0.3
MinFR = 0.15
MaxIE = 0.2
MinIE = 0
MaxIT = 0.26
MinIT = 0.012
MaxNO = 0.25
MinNO = 0.15


LRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Emision = Left(ActiveSheet.Name, 2)
    
    
    On Error Resume Next
    
    'Definimos el array de maxrates. Distinguimos el caso de si el array es de una sola celda o más. Si llega a ser un solo valor, forzamos un array
    'de una sola celda.
    ArrMaxRate = ActiveSheet.Range("AA2:AA" & LRow).Value
    
    If Not IsArray(ArrMaxRate) Then
        ArrMaxRate = ActiveSheet.Range("AA2:AB2").Value
      ReDim Preserve ArrMaxRate(1 To 1, 1 To 1)
    End If
    
    ReDim ArrTxRate(1 To UBound(ArrMaxRate, 1), 1 To 1)

        If Emision = "FR" Or Emision = "FI" Or Emision = "SE" Then
        
            For j = 1 To UBound(ArrMaxRate, 1)
                If ArrMaxRate(j, 1) = 0 Then
                ArrTxRate(j, 1) = MinFR
                Else
                ArrTxRate(j, 1) = MaxFR
                End If
            Next j
            Range("N2:N" & LRow).Value = ArrTxRate
            
        ElseIf Emision = "IE" Then
            For j = 1 To UBound(ArrMaxRate, 1)
                If ArrMaxRate(j, 1) = 0 Then
                ArrTxRate(j, 1) = MinIE
                Else
                ArrTxRate(j, 1) = MaxIE
                End If
            Next j
            Range("N2:N" & LRow).Value = ArrTxRate
            
        ElseIf Emision = "IT" Then
            For j = 1 To UBound(ArrMaxRate, 1)
                If ArrMaxRate(j, 1) = 0 Then
                ArrTxRate(j, 1) = MinIT
                Else
                ArrTxRate(j, 1) = MaxIT
                End If
            Next j
            Range("N2:N" & LRow).Value = ArrTxRate
            
        ElseIf Emision = "NO" Then
            For j = 1 To UBound(ArrMaxRate, 1)
                If ArrMaxRate(j, 1) = 0 Then
                ArrTxRate(j, 1) = MinNO
                Else
                ArrTxRate(j, 1) = MaxNO
                End If
            Next j
            Range("N2:N" & LRow).Value = ArrTxRate
            
            
        End If

    Erase ArrMaxRate, ArrTxRate

End Sub

'Insertamos los eventID y Paydates con inputboxes
Sub EventID_PayDate()

Dim EventID As Variant, PayDate As Variant, ISIN As String, LRow As Long
LRow = Range("A1").End(xlDown).Row
ISIN = Left(ActiveSheet.Name, 12)
        
    'Evaluamos si el ISIN de la hoja actual es igual al de la hoja anterior. Si son diferentes,
    'pedimos los nuevos datos, sino copiamos los de la hoja anterior.
    If ISIN <> Left(Sheets(ActiveSheet.Index - 1).Name, 12) Then
        
        EventID = InputBox("EventID for ISIN " & ISIN)
        PayDate = InputBox("PayDate for ISIN " & ISIN & " (MM/DD/AAAA)")
        
        'Evaluamos si hay una sola fila con datos, para determinar dónde pegamos los valores
        'de las inputboxes
        If IsEmpty(Cells(3, 1)) = True Then
            Range("S2").Value = EventID
            Range("C2").Value = PayDate
            Range("C2").NumberFormat = "mm/dd/yyyy"
        Else
            Range("S2:S" & LRow).Value = EventID
            Range("C2:C" & LRow).Value = PayDate
            Range("C2:C" & LRow).NumberFormat = "mm/dd/yyyy"
        End If
        
    Else
    
         If IsEmpty(Cells(3, 1)) = True Then
            Sheets(ActiveSheet.Index - 1).Range("S6").Copy Range("S2")
            Sheets(ActiveSheet.Index - 1).Range("C6").Copy Range("C2")
        Else
            Sheets(ActiveSheet.Index - 1).Range("S6").Copy Range("S2:S" & LRow)
            Sheets(ActiveSheet.Index - 1).Range("C6").Copy Range("C2:C" & LRow)
        End If
    
    End If

End Sub

'Damos formato a las planillas
Sub FormatPlanillas()

Dim ColW(1 To 19) As Double, RowHTitle As Double, RowHAverage As Double, RowH1 As Double, NroFilas As Integer
Dim j As Integer
Dim fechaG1 As Date

    fechaG1 = Date
    RowHTitle = 25.5
    RowHAverage = 14.25
    RowH1 = 18.75
    
    'Array con ancho columnas de la planilla francesa
    ColW(1) = 13.71
    ColW(2) = 13.57
    ColW(3) = 13.71
    ColW(4) = 9.43
    ColW(5) = 8.71
    ColW(6) = 39.86
    ColW(7) = 13.43
    ColW(8) = 12.86
    ColW(9) = 13.57
    ColW(10) = 10.71
    ColW(11) = 11.57
    ColW(12) = 13.14
    ColW(13) = 14.43
    ColW(14) = 17.71
    ColW(15) = 9.43
    ColW(16) = 10.29
    ColW(17) = 10.14
    ColW(18) = 8.43
    ColW(19) = 19.14

        'Ajustar ancho columnas
        For j = 1 To 19
        ActiveSheet.Columns(j).ColumnWidth = ColW(j)
        Next j
        
        'Ajustar alto de filas y formato
        ActiveWindow.Zoom = 76
        
            With ActiveSheet
            
                .Rows("1:4").EntireRow.insert
                .Rows.RowHeight = RowHAverage       'altura filas
                .Rows(5).RowHeight = RowHTitle      'altura fila titulos
                .Rows(1).RowHeight = RowH1          'altura fila 1
                .Columns("U:AB").Delete
                .Columns("T:T").Hidden = True
                
                With .Range("A1").End(xlDown).CurrentRegion
                    .Font.Name = "Arial"
                    .Font.Size = 10
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                With Range("A1:C1")
                    .Font.Name = "Arial"
                    .Font.Size = 14
                    .Cells(1, 1).Value = "Breakdown Submittal Format"
                    .Font.Bold = True
                    .Merge
                    .HorizontalAlignment = xlCenter
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    
                End With
                
                With Range("E1:F1")
                    .Font.Name = "Arial"
                    .Font.Size = 10
                    .Cells(1, 1).Value = "Submittal Date:"
                    .Font.Bold = True
                    .Merge
                    .HorizontalAlignment = xlLeft
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                End With
                
                With Range("G1")
                    .Font.Bold = True
                    .NumberFormat = "@"
                    .Value = fechaG1
                    .Interior.Color = 13434828
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                End With
                
                With .Range("A5:S5")
                    .HorizontalAlignment = xlGeneral
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .Font.Bold = True
                    .Interior.Color = 10092543
                End With
                
            End With

End Sub

Sub ReorderBOs() 'Ordena las columnas de las planillas y renombra los títulos
    
Dim ArrOrdenCols(1 To 26) As Integer, ArrNombreCols(1 To 19) As String

ActiveWindow.DisplayGridlines = False 'Sacamos las grindlines



    'ArrOrdenCols: array con orden de columnas según planilla
    ArrOrdenCols(1) = 2
    ArrOrdenCols(2) = 13
    ArrOrdenCols(3) = 14
    ArrOrdenCols(4) = 15
    ArrOrdenCols(5) = 16
    ArrOrdenCols(6) = 17
    ArrOrdenCols(7) = 18
    ArrOrdenCols(8) = 19
    ArrOrdenCols(9) = 3
    ArrOrdenCols(10) = 20
    ArrOrdenCols(11) = 21
    ArrOrdenCols(12) = 22
    ArrOrdenCols(13) = 23
    ArrOrdenCols(14) = 24
    ArrOrdenCols(15) = 25
    ArrOrdenCols(16) = 5
    ArrOrdenCols(17) = 4
    ArrOrdenCols(18) = 6
    ArrOrdenCols(19) = 7
    ArrOrdenCols(20) = 8
    ArrOrdenCols(21) = 9
    ArrOrdenCols(22) = 10
    ArrOrdenCols(23) = 26
    ArrOrdenCols(24) = 1
    ArrOrdenCols(25) = 11
    ArrOrdenCols(26) = 12


    'ArrNombreCols: array con nombre de columnas según planilla
    ArrNombreCols(1) = ""
    ArrNombreCols(2) = "Action"
    ArrNombreCols(3) = "Account Number"
    ArrNombreCols(4) = "Underlying Owner Name"
    ArrNombreCols(5) = " Tax ID"
    ArrNombreCols(6) = " Investor Type"
    ArrNombreCols(7) = " Address"
    ArrNombreCols(8) = " Zip Code"
    ArrNombreCols(9) = " City of Residence"
    ArrNombreCols(10) = " Country of Residence"
    ArrNombreCols(11) = "Tax Exempt"
    ArrNombreCols(12) = "Cln-Ref-ID(Leave blank in case of creation)"


    'Para ordenar las columnas, insertamos una fila al inicio, pegamos los array con el orden y nombre de
    'las columnas y luego las ordenamos.

    With ActiveSheet
          Range("A1").EntireRow.insert
          Range("A1:z1").Value = ArrOrdenCols
          .Sort.SortFields.Clear
          .Sort.SortFields.Add Key:= _
          Range("A1:z1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
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
          Range("A1:S1").Value = ArrNombreCols  'Dejamos los titulos como en la planilla
    End With

End Sub


Sub FillDataBOs()

Dim Action As String, TaxExempt As String
Action = "Create"
TaxExempt = "No"
LRow = Range("B1").End(xlDown).Row
    
    
    '//Llenamos los datos de las columnas Action y TaxExempt, discriminando según si hay una o más de una fila.
    With ActiveSheet
    
        If IsEmpty(ActiveSheet.Cells(3, 2)) = True Then
            Range("B2").Value = Action
            Range("K2").Value = TaxExempt
        Else
            Range("B2:B" & LRow).Value = Action
            Range("K2:K" & LRow).Value = TaxExempt
        End If
    
    End With



End Sub


'Damos formato a los BOs
Sub FormatBOs()

Dim ColW(1 To 12) As Double, RowHAverage As Double, RowH1to6 As Double, NroFilas As Integer, RowHTitle As Double
Dim j As Integer, TitleRow As Integer
Dim fechaG1 As String, Title As String

'    fechaG1 = "05 - 12 - 2016"
    RowHTitle = 48.75
    RowHAverage = 21
    RowH1to6 = 27.75
    Title = "B7:L7"
    TitleRow = Range(Title).Row
    
    'Array con ancho columnas de la planilla francesa
    ColW(1) = 2.86
    ColW(2) = 16.29
    ColW(3) = 22.43
    ColW(4) = 34.43
    ColW(5) = 14.86
    ColW(6) = 15.86
    ColW(7) = 23.86
    ColW(8) = 11
    ColW(9) = 17.57
    ColW(10) = 21.14
    ColW(11) = 13.57
    ColW(12) = 30.29

        'Ajustar ancho columnas
        For j = 1 To 12
        ActiveSheet.Columns(j).ColumnWidth = ColW(j)
        Next j
        
        'Ajustar alto de filas y formato
            ActiveWindow.Zoom = 65
        
            With ActiveSheet
            
                .Rows("1:6").EntireRow.insert
                .Rows.RowHeight = RowHAverage
                .Rows("1:6").RowHeight = RowH1to6                                    'altura filas
                .Rows(TitleRow).RowHeight = RowHTitle      'altura fila titulos
                .Columns("M:Z").Delete
                
                With .Range("B1").End(xlDown).CurrentRegion
                    .Font.Name = "Arial"
                    .Font.Size = 12
                End With
                
                With Range(Title)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                With Range("C1")
                    .Value = "BENEFICIAL OWNER SETUP FORM"
                    .Font.Name = "Arial"
                    .Font.Size = 12
                    .Font.Bold = True
                End With
                
                With Range("C3:D4")
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Font.Name = "Arial"
                    .Font.Size = 14
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlBottom
                End With
                
                Range("C3").Value = "Submission date: "
                Range("C4").Value = "Submitted by:"
                
                With Range("D3")
                    .Value = Date
                    .Font.Bold = True
                End With
                
                With Range("D4")
                    .Value = "SEBASTIAN BEVC"
                    .Font.Bold = True
                End With
                
               Range("B7").Interior.Color = 10921638
               Range("C7:K7").Interior.Color = 49407
               Range("L7").Interior.Color = 15773696
               
               Range("B8:B100").Interior.Color = 15921906
               Range("C8:K100").Interior.Color = 14281213
               Range("L8:L100").Interior.Color = 15986394
               
               With Range("B7:L100").Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
               End With
               
               
            End With

End Sub

