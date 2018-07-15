Attribute VB_Name = "mdFuncionesVarias"

'Arreglamos el formato de los numeros de CORP
Sub FixNums()
    

    Dim c As Range
    
    
    For i = 1 To Selection.Columns.Count
    
        With Selection.Columns(1)
            .Replace What:=",", Replacement:="", LookAt:=xlPart, _
           SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
           ReplaceFormat:=False
        
            .Replace What:=".", Replacement:=",", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
            
            .TextToColumns Destination:=Selection.Cells(1, 1), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="|", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        End With
    
    Next i
    
    
    On Error Resume Next
    For Each c In Selection
        c.value = Application.WorksheetFunction.Trim(c)
    Next c

End Sub

Sub newWb()

Workbooks.Add

End Sub
'Saca los signos de = y " de gamma
Sub FixGamma()

    With Selection
        .Replace What:="=", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        .Replace What:="""", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End With
    
End Sub

Sub copyUnique()
Attribute copyUnique.VB_ProcData.VB_Invoke_Func = "a\n14"
    
    Dim MSForms_DataObject As Object
    Dim arr As String, coll As New Collection, c As Range
    
    If Selection.Cells.Count = 1 Then
        Set c = Selection
        coll.Add item:=CStr(Trim(c.value)), Key:=CStr(Trim(c.value))
    Else
        For Each c In Selection.SpecialCells(xlCellTypeVisible)
            On Error Resume Next
            coll.Add item:=CStr(Trim(c.value)), Key:=CStr(Trim(c.value))
        Next c
    End If
    
    For i = 1 To coll.Count
        If arr = "" Then
            arr = coll(i)
        Else
            arr = arr & vbNewLine & coll(i)
        End If
    Next i
    
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    MSForms_DataObject.SetText arr
    MSForms_DataObject.PutInClipboard
    

End Sub

'Copiamos los VAN cambiándoles los separadores para buscar!!
Sub copyVANs()

    Dim MSForms_DataObject As Object
    Dim arr As String, coll As New Collection, c As Range
    
    If Selection.Cells.Count = 1 Then
        Set c = Selection
        coll.Add item:=CStr(Trim(c.value)), Key:=CStr(Trim(c.value))
    Else
        For Each c In Selection.SpecialCells(xlCellTypeVisible)
            On Error Resume Next
            cleanedValue =
            coll.Add item:=CStr(Trim(c.value)), Key:=CStr(Trim(c.value))
        Next c
    End If
    
    For i = 1 To coll.Count
        If arr = "" Then
            arr = coll(i)
        Else
            arr = arr & vbNewLine & coll(i)
        End If
    Next i
    
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    MSForms_DataObject.SetText arr
    MSForms_DataObject.PutInClipboard


End Sub

'sacar los separadores de los VAN para buscarlos luego
Function cleanVAN(value)
    
    If InStr(value, " ") Then
        cleanVAN = Replace(value, " ", "*")
    
    

End Function

'partiendo de los VAN hace la variante y el grouping
'asumimos que la parte igual de los VAN tiene el mismo largo para todos los items
Sub getVariantByVAN()

    Dim VANrng As Range 'rango de VANs
    Dim c As Range, wb As Workbook
    Set wb = ActiveWorkbook
    Set VANrng = wb.Sheets(1).Range(Cells(18, VAN_s), Cells(Rows.Count, VAN_s).End(xlUp))
    
    Dim strLengh As Integer 'hasta que caracter son iguales los vans
    strLengh = lenghtEqual(VANrng.Cells(1, 1))
    
    Dim VANs As New Scripting.Dictionary
    For Each c In VANrng
        Dim van As String
        van = Left(c.value, strLengh)
        If VANs.Exists(van) Then
            VANs(van).Add c.Row
        Else
            Set coll = New Collection
            coll.Add c.Row
            VANs.Add item:=coll, Key:=van
        End If
    Next c
    
    
    Dim coll2, item, grouping
    Dim itemID: itemID = 1000
    Dim groupingCol: groupingCol = 16
    
    For Each coll2 In VANs
        itemID = itemID + 1
        grouping = 1
        For Each item In VANs(coll2)
            With wb.Sheets(1)
                .Cells(item, groupingCol).value = grouping
                .Cells(item, ID_s).value = itemID & "¦" & Format(grouping, "000")
            End With
            grouping = grouping + 1
        Next item
    Next coll2
    
    
End Sub
'compara los VAN y se fija hasta qué caracter son iguales (si hay un espacio para el conteo), devuelve la posición de ese caracter para luego definir el grouping por VAN
Function lenghtEqual(c As Range)
    
    Dim equal As Boolean
    Dim counter As Integer: counter = 1
    
    
    Do Until equal
        If Left(c.value, counter) = Left(c.Offset(1, 0).value, counter) And Mid(c.value, counter + 1, 1) <> " " Then   'si los valores son iguales y si el último dígito no es un espacio
            counter = counter + 1
        Else
            equal = True
        End If
    Loop
    
   
    lenghtEqual = counter
    
End Function

Sub CopyCondRecNo()

    Dim c As Range, txt As String
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    
    For Each c In Selection
        If txt = "" Then
            txt = Format(c.value, "0000000000")
        Else
            txt = txt & vbNewLine & Format(c.value, "0000000000")
        End If
    Next c
    
    MSForms_DataObject.SetText txt
    MSForms_DataObject.PutInClipboard

End Sub

Sub getVariantByGrouping()
    
    Dim gRng As Range   'grouping range
    Dim c As Range, wb As Workbook
    Dim lr  'ultima fila con datos
    
    Set wb = ActiveWorkbook
    lr = wb.Sheets(1).Range("A17").End(xlDown).Row
    Set gRng = wb.Sheets(1).Range("P18:P" & lr)
    
    'chequeamos que no haya ningún grouping vacío
    For Each c In gRng
        If isEmpty(c) Or c.value = "" Then
            MsgBox "Falta grouping en la celda " & c.Address
            Exit Sub
        End If
    Next c
    
    Dim itemID As Integer: itemID = 1000
    
    For Each c In gRng
        If c.value = 1 Then: itemID = itemID + 1
        wb.Sheets(1).Cells(c.Row, ID_s).value = itemID & "¦" & Format(c.value, "000")
    Next c
    
    
End Sub

Sub formatEANs()
Attribute formatEANs.VB_ProcData.VB_Invoke_Func = "w\n14"
    
    Selection.NumberFormat = "0"

End Sub
'traemos los GAMMA sites de los SAP sites seleccionados
