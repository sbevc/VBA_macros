Attribute VB_Name = "CrearPlanillas"
Option Explicit

Dim ws As Worksheet, DestWs As Worksheet, Lrow As Long, i As Integer

Sub Crear_Planillas()

Dim Emisión As String, ISIN As String

Set ws = Sheets("Datos")
ws.Activate

    DataCheck
    Concatenate
    Parse_Data_Cuentas
    
      
        For i = 5 To Sheets.Count
            
            Sheets(i).Activate
            Emisión = Left(ActiveSheet.Name, 2)
            ISIN = Left(ActiveSheet.Name, 12)
            
            Select Case Emisión
            
                Case Is = "XS"
                    DoXS
                    
                Case Is = "ES"
                    If Left(ISIN, 5) = "ES000" Then
                        DoES_Pub
                    Else
                        DoES_Priv
                    End If
                    
                Case Is = "PT"
                    DOPT
                    
            End Select
            
            
        Next i

    ReorderWs

End Sub

'Chequeamos que hayan datos en la celda A1 y si hay mas de un wb abierto, que active este.

Sub DataCheck()

Set ws = Sheets("Datos")

    If IsEmpty(Range("A1")) = True Then
        MsgBox "Insertar datos antes de ejecutar la macro"
        End
    ElseIf Workbooks.Count > 1 Then
        ThisWorkbook.Sheets("Datos").Activate
    End If
    
End Sub


Sub Concatenate()

Dim ISINCol As Integer, AccCol As Integer, PayDateCol As Integer


    Lrow = Range("A" & Rows.Count).End(xlUp).Row
    ISINCol = 4
    AccCol = 9
    ws.Cells(1, 1).Value = "Nombre"
    PayDateCol = 8

    'Reemplazamos "/" con "." en la columna PayDate
    'ws.Columns(PayDateCol).Replace What:="/", Replacement:=".", LookAt:=xlPart, _
           ' SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            'ReplaceFormat:=False
    
    'Concatenar
    For i = 2 To Lrow
        Cells(i, 1) = Cells(i, ISINCol) & " " & Cells(i, AccCol)
    Next i

End Sub

Sub Parse_Data_Cuentas()

Dim lr As Long
Dim ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim Title As String
Dim TitleRow As Integer

    vcol = 1
    Set ws = Sheets("Datos")
    lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
    Title = "A1:X1"
    TitleRow = ws.Range(Title).Cells(1).Row
    icol = ws.Columns.Count
    ws.Cells(1, icol) = "Unique"
    
    For i = 2 To lr
    On Error Resume Next
    If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
    ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
    End If
    Next
    myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    ws.Columns(icol).Clear
    For i = 2 To UBound(myarr)
    ws.Range(Title).AutoFilter Field:=vcol, Criteria1:=myarr(i) & ""
    If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
    Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
    Else
    Sheets(myarr(i) & "").Move After:=Worksheets(Worksheets.Count)
    End If
    ws.Range("A" & TitleRow & ":A" & lr).EntireRow.Copy Sheets(myarr(i) & "").Range("A1")
    Sheets(myarr(i) & "").Columns.AutoFit
    Next
    ws.AutoFilterMode = False
    ws.Activate
    
    Erase myarr
        
End Sub

Sub DoXS()


Dim ISIN As String, ACC As String, Desc As String, PD, Total As Long, Wsname As String
Set ws = ActiveSheet

ISIN = Range("D2").Value
ACC = Range("I2").Value
Desc = Range("E2").Value
PD = Range("H2").Value
Total = Range("K2").End(xlDown).Value
Wsname = ISIN & " " & ACC & " Planilla"


    'Copiamos la pestaña XS y le pegamos los datos generales (ISIN, etc)
    
    Sheets("Modelo XS").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = Wsname
    Sheets(Sheets.Count).Visible = True


    With Sheets(Wsname & "")
        .Tab.ThemeColor = xlThemeColorAccent3
        .Tab.TintAndShade = 0.399975585192419
        .Range("A3").Value = ACC
        .Range("B3").Value = ISIN
        .Range("C3").Value = Desc
        .Range("F3").Value = PD
        .Range("C4").Value = Total
    End With


    'Filtramos por ">1" para saber si hay personas jurídicas
    
    ws.Activate
    Range("A1").CurrentRegion.AutoFilter Field:=18, Criteria1:=">1"
    

    'Separamos entre si el filtrado tiene o no resultados.
    'Sino tiene resultados no hacemos nada ya que las PF no se ponen. Si tiene resultados,
    'copiamos los datos a la hoja nueva.
    
    If Cells(Rows.Count, 1).End(xlUp).Row <> 1 Then
        
        Cells.AutoFilter
        CopyDataXS
        
    End If
        
    
End Sub


Sub CopyDataXS()

Dim SrchRng As Range, cel As Range
Dim ColNif As Integer, ColCLIENT_NAME As Integer, ColREF As Integer, ColPOSITION As Integer, ColFISCAL_ADDRESS _
As Integer, ColCP As Integer, ColPLACE As Integer, ColCOUNTRY As Integer
Dim x As Integer, y As Integer
Dim R_Res As Integer, R_NoRes As Integer
Set ws = ActiveSheet
Set DestWs = Sheets(Sheets.Count)

Dim col
x = 0
y = 0

'Columnas de los respectivos datos
ColNif = 16
ColCLIENT_NAME = 17
ColREF = 15
ColPOSITION = 11
ColFISCAL_ADDRESS = 19
ColCP = 20
ColPLACE = 21
ColCOUNTRY = 22

'Filas de los títulos de los residentes y no residentes
R_Res = 20
R_NoRes = 12


    '//Loopeamos por la columna de KIND OF PERSON, si es distinta de 1 (osea si no es persona física) _
    '//la incluimos en la lista de las XS
    
    If IsEmpty(Cells(3, 18)) = True Then
        Set SrchRng = Range("R2")
    Else
        Set SrchRng = Range("R2", Range("R2").End(xlDown))
    End If
    
    For Each cel In SrchRng
    
        If cel.Value <> 1 Then
                    
            'Si es residente ("ESPA") lo mandamos a la fila (R_Res + x) y si no es residente, a la fila (R_NoRes + y)
            'Siempre va a ser la ùltima hoja la de destino (Sheets(sheets.count))
            
            If Left(Range("V" & cel.Row), 4) = "ESPA" Then
                x = x + 1
                DestWs.Cells(R_Res + x, 1).Value = Cells(cel.Row, ColNif).Value
                DestWs.Cells(R_Res + x, 2).Value = Cells(cel.Row, ColCLIENT_NAME).Value
                DestWs.Cells(R_Res + x, 3).Value = Cells(cel.Row, ColREF).Value
                DestWs.Cells(R_Res + x, 6).Value = Cells(cel.Row, ColPOSITION).Value
                DestWs.Cells(R_Res + x, 7).Value = Cells(cel.Row, ColFISCAL_ADDRESS).Value
                DestWs.Cells(R_Res + x, 8).Value = Cells(cel.Row, ColCP).Value
                DestWs.Cells(R_Res + x, 9).Value = Cells(cel.Row, ColPLACE).Value
                DestWs.Cells(R_Res + x, 10).Value = Cells(cel.Row, ColCOUNTRY).Value
                
                
            Else
                y = y + 1
                DestWs.Cells(R_NoRes + y, 1).Value = Cells(cel.Row, ColNif).Value
                DestWs.Cells(R_NoRes + y, 2).Value = Cells(cel.Row, ColCLIENT_NAME).Value
                DestWs.Cells(R_NoRes + y, 3).Value = Cells(cel.Row, ColREF).Value
                DestWs.Cells(R_NoRes + y, 6).Value = Cells(cel.Row, ColPOSITION).Value
                DestWs.Cells(R_NoRes + y, 7).Value = Cells(cel.Row, ColFISCAL_ADDRESS).Value
                DestWs.Cells(R_NoRes + y, 8).Value = Cells(cel.Row, ColCP).Value
                DestWs.Cells(R_NoRes + y, 9).Value = Cells(cel.Row, ColPLACE).Value
                DestWs.Cells(R_NoRes + y, 10).Value = Cells(cel.Row, ColCOUNTRY).Value
                
                
            End If
            
        End If
        
    Next cel

End Sub


Sub DoES_Pub()


Dim ISIN As String, ACC As String, Desc As String, PD, Wsname As String
Set ws = ActiveSheet


    'Copiamos los datos Generales
    ISIN = Range("D2").Value
    ACC = Range("I2").Value
    Desc = Range("E2").Value
    PD = Range("H2").Value
    
    'Copiamos la hoja "Modelo ES Pub"
    Wsname = ISIN & " " & ACC & " Planilla"
    Sheets("Modelo ES Pub").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = Wsname
    Sheets(Sheets.Count).Visible = True

    With Sheets(Wsname & "")
        .Tab.ThemeColor = xlThemeColorAccent3
        .Tab.TintAndShade = 0.399975585192419
        .Range("B1").Value = ACC
        .Range("B2").Value = ISIN
        .Range("B3").Value = Desc
        .Range("B4").Value = PD
    End With
    
    ws.Activate
    
    CopyES_Pub
    

End Sub

Sub CopyES_Pub()

Dim RDatos As Integer, x As Integer
Dim ColNif As Integer, ColCLIENT_NAME As Integer, ColKindOfPerson As Integer, ColPOSITION As Integer, _
ColFISCAL_ADDRESS As Integer, ColCP As Integer, ColPLACE As Integer, ColCOUNTRY As Integer
Dim FiscalAddress As String, CP As String, PLACE As String, COUNTRY As String, FullAddress As String
Dim ARow As Integer

Set ws = ActiveSheet
Set DestWs = Sheets(Sheets.Count)

'Columnas de los respectivos datos
ColNif = 16
ColCLIENT_NAME = 17
ColKindOfPerson = 18
ColPOSITION = 11
ColFISCAL_ADDRESS = 19
ColCP = 20
ColPLACE = 21
ColCOUNTRY = 22

RDatos = 6 'Fila a partir de la cual pegamos los datos
x = 0


    'Copiamos los datos específicos (Nombre del cliente, NIF, etc) a la hoja copiada "Modelo ES Pub"
    'que va a ser la sheets(sheets.count)
    
    Range("Q2").Activate
    
    Do Until IsEmpty(ActiveCell) = True
            
            x = x + 1
            ARow = ActiveCell.Row
            
            DestWs.Cells(RDatos + x, 1).Value = "FINAL BO"
            DestWs.Cells(RDatos + x, 2).Value = Cells(ARow, ColCLIENT_NAME)
            DestWs.Cells(RDatos + x, 3).Value = Cells(ARow, ColNif)
            
            'Individual o corporation segùn el nro de kind of person
            If Cells(ARow, ColKindOfPerson) = 1 Then
                DestWs.Cells(RDatos + x, 4).Value = "INDIVIDUAL"
            Else
                DestWs.Cells(RDatos + x, 4).Value = "CORPORATION"
            End If
            
            ' Trimeamos y concatenamos el address y demás
            FiscalAddress = Application.WorksheetFunction.Trim(Cells(ARow, ColFISCAL_ADDRESS))
            CP = Application.WorksheetFunction.Trim(Cells(ARow, ColCP))
            PLACE = Application.WorksheetFunction.Trim(Cells(ARow, ColPLACE))
            COUNTRY = Application.WorksheetFunction.Trim(Cells(ARow, ColCOUNTRY))
            
            FullAddress = FiscalAddress & ", " & CP & ", " & PLACE & ", " & COUNTRY
            
            DestWs.Cells(RDatos + x, 5).Value = FullAddress
            DestWs.Cells(RDatos + x, 6).Value = Cells(ARow, ColPOSITION)
            
            
            ActiveCell.Offset(1, 0).Activate
            
    Loop
            
End Sub


Sub DoES_Priv()

Dim ISIN As String, ACC As String, Desc As String, PD, Wsname As String
Set ws = ActiveSheet


    'Copiamos los datos Generales
    ISIN = Range("D2").Value
    ACC = Range("I2").Value
    Desc = Range("E2").Value
    PD = Range("H2").Value
    
    'Copiamos la hoja "Modelo ES Pub"
    Wsname = ISIN & " " & ACC & " Planilla"
    Sheets("Modelo ES Priv").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = Wsname
    Sheets(Sheets.Count).Visible = True

    With Sheets(Wsname & "")
        .Tab.ThemeColor = xlThemeColorAccent3
        .Tab.TintAndShade = 0.399975585192419
        .Range("B1").Value = ACC
        .Range("B2").Value = ISIN
        .Range("B3").Value = Desc
        .Range("B4").Value = PD
    End With

    ws.Activate
    
    CopyES_Priv
    

End Sub


Sub CopyES_Priv()

Dim Rng As Range, c As Range, TotalPJNoRes As Double, TotalPJRes As Double, TotalPF As Double
Dim KindOfPersonCol As Integer, CountryCol As Integer

Set DestWs = Sheets(Sheets.Count)

KindOfPersonCol = 18
CountryCol = 22

TotalPF = 0
TotalPJNoRes = 0
TotalPJRes = 0

    'Separamos el análisis entre si hay una sola linea o más de una linea
    If IsEmpty(Cells(4, 1)) = True Then
    
            If Range("R2").Value = 1 Then
                DestWs.Range("G10").Value = Range("M2").Value
            Else
                If Left(Range("V2").Value, 4) = "ESPA" Then
                    DestWs.Range("G8").Value = Range("M2").Value
                Else
                    DestWs.Range("G9").Value = Range("M2").Value
                End If
            End If


    Else
            
              Set Rng = Range("M2", Range("M2").End(xlDown).Offset(-1, 0))
              
              For Each c In Rng
                            
                    If Cells(c.Row, KindOfPersonCol) = 1 Then
                        TotalPF = TotalPF + c.Value
                    Else
                        If Left(Cells(c.Row, CountryCol), 4) = "ESPA" Then
                            TotalPJRes = TotalPJRes + c.Value
                        Else
                            TotalPJNoRes = TotalPJNoRes + c.Value
                        End If
                    End If
                    
              Next
              
              DestWs.Range("G8").Value = TotalPJRes
              DestWs.Range("G9").Value = TotalPJNoRes
              DestWs.Range("G10").Value = TotalPF
              
    End If

End Sub


Sub ReorderWs()


Dim i As Integer
Dim j As Integer


   For i = 5 To Sheets.Count
   
      For j = 5 To Sheets.Count - 1

            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
            
      Next j
      
   Next i
   

End Sub


Sub DOPT()

Dim ISIN As String, ACC As String, Desc As String, PD, Wsname As String
Set ws = ActiveSheet


    'Copiamos los datos Generales
    ISIN = Range("D2").Value
    ACC = Range("I2").Value
    Desc = Range("E2").Value
    PD = Range("H2").Value
    
    'Copiamos la hoja "Modelo ES Pub"
    Wsname = ISIN & " " & ACC & " Planilla"
    Sheets("Modelo PT").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = Wsname
    Sheets(Sheets.Count).Visible = True

    With Sheets(Wsname & "")
        .Tab.ThemeColor = xlThemeColorAccent3
        .Tab.TintAndShade = 0.399975585192419
        .Range("B10").Value = ACC
        .Range("B8").Value = ISIN
        .Range("B9").Value = Desc
        .Range("B7").Value = PD
    End With

    ws.Activate
    
    CopyPT


End Sub

Sub CopyPT()

Dim SrchRng As Range, cel As Range
Dim ColNif As Integer, ColCLIENT_NAME As Integer, ColPOSITION As Integer, ColFISCAL_ADDRESS _
As Integer, ColCP As Integer, ColPLACE As Integer, ColCOUNTRY As Integer
Dim x As Integer
Dim RDatos As Integer, ARow As Integer
Dim FiscalAddress As String, CP As String, PLACE As String, COUNTRY As String, FullAddress As String

Set ws = ActiveSheet
Set DestWs = Sheets(Sheets.Count)


x = 0


'Columnas de los respectivos datos
ColNif = 16 'va
ColCLIENT_NAME = 17 'va
ColPOSITION = 11 'va
ColFISCAL_ADDRESS = 19 'va
ColCP = 20 'va
ColPLACE = 21 'va
ColCOUNTRY = 22 'va

'Filas a partir de la cual pegamos los datos
RDatos = 17



    '//Loopeamos por la columna de KIND OF PERSON y separamos el rango de loopeo srchrng
    'entre si hay una línea de datos o mas de una
     
    If IsEmpty(Cells(3, 18)) = True Then
        Set SrchRng = Range("R2")
    Else
        Set SrchRng = Range("R2", Range("R2").End(xlDown))
    End If
    
    'Copiamos los datos a la hoja destino DestWs
    For Each cel In SrchRng

        x = x + 1
        DestWs.Cells(RDatos + x, 2).Value = Cells(cel.Row, ColNif).Value
        DestWs.Cells(RDatos + x, 1).Value = Cells(cel.Row, ColCLIENT_NAME).Value
        DestWs.Cells(RDatos + x, 4).Value = Cells(cel.Row, ColPOSITION).Value
        
        ARow = cel.Row
        
        FiscalAddress = Application.WorksheetFunction.Trim(Cells(ARow, ColFISCAL_ADDRESS))
        CP = Application.WorksheetFunction.Trim(Cells(ARow, ColCP))
        PLACE = Application.WorksheetFunction.Trim(Cells(ARow, ColPLACE))
        COUNTRY = Application.WorksheetFunction.Trim(Cells(ARow, ColCOUNTRY))
        
        FullAddress = FiscalAddress & ", " & CP & ", " & PLACE & ", " & COUNTRY
        
        DestWs.Cells(RDatos + x, 3).Value = FullAddress
               
    Next cel


End Sub


