Attribute VB_Name = "Módulo1"
Sub asdas()

    Set RE = CreateObject("vbscript.regexp")

    RE.Global = True: RE.IgnoreCase = True
    'RE.Pattern = "\d*(\.|,)?\d+"
    RE.Pattern = "[A-Z]{2}\d{2}"
    
    Dim c As Range
    
    For Each c In Range("A4:A510")
        If RE.Test(c.value) Then: Debug.Print c.value
    Next c

End Sub
Sub asdas2()

    Set RE = CreateObject("vbscript.regexp")

    RE.Global = True: RE.IgnoreCase = True
    'RE.Pattern = "\d*(\.|,)?\d+"
    RE.Pattern = "[A-Z]{2}T\d{3}"
    
    Dim c As Range
    
    For Each c In Range("G512:G614")
        If RE.Test(c.value) Then: Debug.Print c.value
    Next c

End Sub

Sub fdfd()

Dim sWb As Workbook, dWb As Workbook, lrow As Integer
Set sWb = ActiveWorkbook
Set dWb = Workbooks.Add

For i = 1 To sWb.Sheets.Count
    lrow = sWb.Sheets(i).Range("A" & Rows.Count).End(xlUp).Row
    sWb.Sheets(i).Range("A5:B" & lrow).Copy Destination:=dWb.Sheets(1).Range("A" & Range("A" & Rows.Count).End(xlUp).Offset(1).Row)
Next i

End Sub

Public Sub PerformCopy()
    CopyFiles "C:\VIS\", "C:\Users\sebbev\Desktop\VIS\"
End Sub


Public Sub CopyFiles(ByVal strPath As String, ByVal strTarget As String)

    Set fso = CreateObject("scripting.filesystemobject")
    DoFolder fso.GetFolder(strPath), strTarget
    
    
End Sub

Sub DoFolder(folder, strTarget)

    Dim SubFolder
    For Each SubFolder In folder.SubFolders
        DoFolder SubFolder, strTarget
    Next
    
    Dim file
    For Each file In folder.Files
        file.Copy strTarget
    Next
    
End Sub

'agregar sufijos a archivos en una carpeta dada
Sub AddSuffix()

    Set fso = CreateObject("scripting.filesystemobject")
    Dim folder
    Set folder = fso.GetFolder("C:\Users\sebbev\Desktop\VIS\")
    Dim file, oldName, newName
    
    For Each file In folder.Files
        oldName = folder & "\" & file.Name
        newName = folder & "\" & Left(file.Name, Len(file.Name) - 5) & " - SB" & ".xlsx"
        Name oldName As newName
    Next file
    
End Sub

Sub asdasdas()
    
    Dim c As Range
    x = 1001
    
    
    For Each c In Selection.SpecialCells(xlCellTypeVisible)
        c.value = x
        x = x + 1
    Next c
    
End Sub
Sub asda()


Select Case Application.International(XlApplicationInternational.xlCountryCode)
   Case 1: Call MsgBox("English")
   Case 33: Call MsgBox("French")
   Case 49: Call MsgBox("German")
   Case 81: Call MsgBox("Japanese")
End Select

End Sub

Sub asda1510s()

    Dim rng As Range, c As Range
    Set rng = Range("C18:C378")
    
    For Each c In rng
        If Application.WorksheetFunction.IsNA(c) Then: Debug.Print c.Address
    Next c

End Sub

Sub asdasda()

    Dim c As Range, txt As String
    Dim objData As New MsForms.DataObject
    
    For Each c In Selection.SpecialCells(xlCellTypeVisible)
        If txt = "" Then
            txt = c.value
        Else
            txt = txt & ", " & c.value
        End If
    Next c
    
    objData.SetText txt
    objData.PutInClipboard
    
End Sub
'pintamos del mismo color(aleatorio) los mismos valores (duplicados)
Sub same()

    Dim c As Range, dict As New Scripting.Dictionary, col As Long, rng As Range
    Dim xRed As Byte
    Dim xGreen As Byte
    Dim xBule As Byte
    
    Set rng = Selection
    
    For Each c In rng
        If Not Application.WorksheetFunction.IsNA(c) Then
        If Not isEmpty(c) And c.value <> "" Then
            
            Count = 0
            For Each area In rng.Areas  'hay que dividir el countif por area de la selección
                Count = Count + Application.WorksheetFunction.CountIf(area, c.value)
            Next area
            
            If Count > 1 Then
            
                If dict.Exists(c.value) Then
                    With c.Interior
                        .Color = dict(c.value)
                        .Pattern = xlSolid
                        .PatternColor = xlAutomatic
                    End With
                Else
                    xRed = Application.WorksheetFunction.RandBetween(0, 255)
                    xGreen = Application.WorksheetFunction.RandBetween(0, 255)
                    xBule = Application.WorksheetFunction.RandBetween(0, 255)
                    col = VBA.RGB(xRed, xGreen, xblue)
                    dict.Add Key:=c.value, item:=col
                    With c.Interior
                        .Color = dict(c.value)
                        .Pattern = xlSolid
                        .PatternColor = xlAutomatic
                    End With
                End If
            
            End If
            
        End If
        End If
    Next c
    
 End Sub
 
 
'Dada la selección art|site, arma el VIS para delisting
Sub Delisting()

    Dim wbDest As Workbook, rng As Range, c As Range
    
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    
    
    
    
    For Each c In Selection.Columns(1).SpecialCells(xlCellTypeVisible)
        
    Next c
    
    Set wbDest = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Areas de trabajo\Final Template_NEW_VIS_V2.1 (plantilla).xltx", local:=True)

End Sub

'seleccionando 3 columnas hacemos un buscarV de TODOS los valores y los pegamos a la derecha del valor buscado
    'columna 1: valor a buscar
    'columna 2: donde buscar
    'columna 3: valores buscados
Sub lookupALL()

    Dim valuesToLook As Range, whereTolook As Range, returnValues As Range
    Dim c As Range
    Dim results As New Scripting.Dictionary
    
    'verificamos que se hayan seleccionado 3 columnas o 3 áreas
    If Selection.Areas.Count <> 3 Then: MsgBox "Seleccione 3 áreas": Exit Sub
    
    'si seleccionan una columna entera tomamos solo los valores con datos para que sea más rapido
    If Selection.Areas(1).Cells.Count = 1048576 Then
        Set valuesToLook = usedArea(Selection.Areas(1))
    Else
        Set valuesToLook = Selection.Areas(1)
    End If
    If Selection.Areas(2).Cells.Count = 1048576 Then
        Set whereTolook = usedArea(Selection.Areas(2))
    Else
        Set whereTolook = Selection.Areas(2)
    End If
    If Selection.Areas(3).Cells.Count = 1048576 Then
        Set returnValues = usedArea(Selection.Areas(3))
    Else
        Set returnValues = Selection.Areas(3)
    End If
    
    'armamos un diccionario con las columnas whereToLook y returnValues para agrupar los datos
    Dim coll As Collection, d_whereToLook As New Scripting.Dictionary
    For Each c In whereTolook
        If Not isEmpty(c) Then
            value = Trim(c.value)
            If Not d_whereToLook.Exists(value) Then
                Set coll = New Collection
                On Error Resume Next
                coll.Add Key:=Trim(Cells(c.Row, returnValues.Column).value), item:=Trim(Cells(c.Row, returnValues.Column).value)
                d_whereToLook.Add Key:=value, item:=coll
            Else
                On Error Resume Next
                d_whereToLook(value).Add Key:=Trim(Cells(c.Row, returnValues.Column).value), item:=Trim(Cells(c.Row, returnValues.Column).value)
            End If
        End If
    Next c
    
    Dim result As Variant, offsetCol As Integer
    For Each c In valuesToLook.SpecialCells(xlCellTypeVisible)
        If Not isEmpty(c) Then
            value = Trim(c.value)
            If d_whereToLook.Exists(value) Then
                offsetCol = 0
                For Each result In d_whereToLook(value)
                    offsetCol = offsetCol + 1
                    c.Offset(0, offsetCol).value = result
                Next result
            Else
                c.Offset(0, 1).value = "#N/A"
            End If
        End If
    Next c
    
End Sub
 
Function usedArea(rng As Range) As Range
    
    Dim fcell As Range, lcell As Range, lrow As Integer, col As Integer
    
    Set fcell = rng.Cells(1)
    Set lcell = rng.Cells(rng.Cells.Count).End(xlUp)

    Set usedArea = Range(fcell, lcell)

End Function

'Arreglar cuando mandan un delisting con los sites separados por coma
'por ejemplo:
' item | sites
' 1234 | CO01,CO02,COBG,COBH,COBC

'Al seleccionar los articulos toma también los sites de al lado
'y los pone en un libro nuevo

Sub fixDelisting()

    Dim c1 As Range, rngArts As Range
    Dim wbDest As Workbook, site, counter
    
    Set rngArts = Selection
    
    Set wbDest = Workbooks.Add
    wbDest.Sheets(1).Range("A1").value = "Art #"
    wbDest.Sheets(1).Range("B1").value = "Site"
    
    counter = 2
    
    For Each c1 In rngArts.SpecialCells(xlCellTypeVisible)
        For Each site In Split(c1.Offset(0, 1).value, ",")
            wbDest.Sheets(1).Cells(counter, 1).value = Trim(c1.value)
            wbDest.Sheets(1).Cells(counter, 2).value = Trim(site)
            counter = counter + 1
        Next site
    Next c1
    
End Sub


Sub copyVAN()


    Dim c As Range
    alphaNum = "abcdefghijklmnopqrstuvwxyz0123456789"
    
    
    For Each c In Selection.SpecialCells(xlCellTypeVisible)
        
        
        
    Next c

End Sub










