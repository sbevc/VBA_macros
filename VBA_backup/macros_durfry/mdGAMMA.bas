Attribute VB_Name = "mdGAMMA"
'-----------------------------------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------------------------------'

Sub getGammaSites()

    Dim wbGAMMA As Workbook, estGAMMA() As Variant
    Set wbGAMMA = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Estructura Gamma-Sap.xlsx", _
                               UpdateLinks:=False, ReadOnly:=True)
    wbGAMMA.Sheets("Enterprise Struct in SAP Corp").Activate
    estGAMMA = Range("A6").CurrentRegion.value  'array con estructura GAMMA
    wbGAMMA.Close savechanges:=False
    
    Dim c As Range, v As Variant, collNotFound As Collection
    
    For Each c In Selection.SpecialCells(xlCellTypeVisible)
        v = Application.VLookup(c.value, estGAMMA, 3, 0)
        If IsError(v) Then
            If collNotFound Is Nothing Then
                Set collNotFound = New Collection
            Else
                On Error Resume Next
                collNotFound.Add item:=CStr(c.value), Key:=CStr(c.value)
            End If
            c.Offset(0, 1).value = "Site not found"
        Else
            c.Offset(0, 1).value = v
        End If
    Next c
    
    If Not collNotFound Is Nothing Then
    
        Dim errorSites As String
        For i = 1 To collNotFound.Count
            If errorSites = "" Then
                errorSites = collNotFound(i)
            Else
                errorSites = errorSites & vbNewLine & collNotFound(i)
            End If
        Next i
        
        MsgBox "No se encontraron los sites:" & vbNewLine & errorSites
        
    End If
    
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------------------------------'

Private Sub cmdDesproteger_Click()
    
    Dim lastRow As Integer
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    

    For x = 5 To lastRow
        'Idioma
            With Range("K" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = "01"
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = "0" & .value
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
        
        'Estado Artículo
            With Range("N" & x)
                .NumberFormat = "@"
                If .value = "" Then
                    .value = "N"
                Else
                    .value = UCase(.value)
                End If
                .Font.Color = vbRed
                .HorizontalAlignment = xlCenter
            End With
            
        'Fecha Afectación
            With Range("O" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = Format(Now(), "YYYYMMDD")
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = CStr(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
        
        'Alto
            With Range("Y" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "@"
                    .value = 1
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    .value = .value
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
            
         'Largo
            With Range("Z" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "@"
                    .value = 1
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    .value = .value
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
            
         'Profundo
            With Range("AA" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "@"
                    .value = 1
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    .value = .value
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
            
        'Peso Bruto
            With Range("AD" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "0.00"
                    .value = 1
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "0.00"
                    .value = .value
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
            
        'Peso Neto
            With Range("AC" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "0.00"
                    If CDbl(Range("AD" & x)) < 1 Then
                        .value = CDbl(Range("AD" & x).value)
                    Else
                        .value = 1
                    End If
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "0.00"
                    If CDbl(Range("AD" & x)) < CDbl(Range("AC" & x)) Then
                        .value = CDbl(Range("AD" & x).value)
                    Else
                        If (.value = "" Or CDbl(.value) = 0) Then
                            .value = 1
                        End If
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
        
        'Fecha lanzamiento
            With Range("AF" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = Format(Now(), "YYYYMMDD")
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = CStr(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
            
        'Fecha disponibilidad
            With Range("AG" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = Format(Now(), "YYYYMMDD")
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = CStr(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
    
        'Denom
            With Range("AP" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "0.00"
                    .value = 1
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "0.00"
                    If CDbl(.value) = 0 Then
                        .value = 1
                    Else
                        .value = CDbl(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With

       'Numer
            With Range("AQ" & x)
                If (.value = "" Or CDbl(.value) = 0) Then
                    .NumberFormat = "0"
                    .value = 1
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "0"
                    If CDbl(.value) = 0 Then
                        .value = 1
                    Else
                        .value = CDbl(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
    
         'Largo
            With Range("AR" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = "N"
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    .value = UCase(.value)
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
    
    
        'Fecha PC
            With Range("AW" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = Format(Now(), "YYYYMMDD")
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = CStr(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
    
        'Fecha Vigor
            With Range("BC" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = Format(Now(), "YYYYMMDD")
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = CStr(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
    
    
        'Fecha Vigor PVP
            With Range("BG" & x)
                If (.value = "") Then
                    .NumberFormat = "@"
                    .value = Format(Now(), "YYYYMMDD")
                    .Font.Color = vbRed
                Else
                    .NumberFormat = "@"
                    If Len(.value) = 1 Then
                        .value = CStr(.value)
                    End If
                    .Font.Color = vbRed
                End If
                .HorizontalAlignment = xlCenter
            End With
            
       'UoM Compra
            With Range("BJ" & x)
                .NumberFormat = "@"
                If .value = "" Then
                    .value = "UN"
                Else
                    .value = UCase(.value)
                End If
                .Font.Color = vbRed
                .HorizontalAlignment = xlCenter
            End With
            
      'UoM Precio
            With Range("BK" & x)
                .NumberFormat = "@"
                If .value = "" Then
                    .value = "UN"
                Else
                    .value = UCase(.value)
                End If
                .Font.Color = vbRed
                .HorizontalAlignment = xlCenter
            End With
    Next
    
    
    'Formato MACROGLO (seba)
    'arreglamos las columnas
    
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range(Cells(4, 1), Cells(lastRow, 1)).ClearContents
    
    'resaltamos si falta algún dato
    Dim colObligatorias() As Variant, col As Variant
    colObligatorias = Array(2, 3, 5, 6, 8, 9, 10, 11, 12, 14, 15, 17, 18, 19, 20, 21, 24, 25, 26, 27, 29, 30, 32, 33, 35, 36, 37, 41, 42, 43, 44, 45, 47, 48, 49, 54, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 76, 77, 98)
    
    Dim c As Range
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    For Each col In colObligatorias
        For Each c In Range(Cells(4, col), Cells(lastRow, col))
            If isEmpty(c) Then
                With c.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next c
    Next col
    
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------------------------------'

'armamos los desafich dados los artículos GAMMA y los GAMMA sites, los guardamos en la carpeta C:GAMMAF\ e imprimimos los sites en un txt
' input: range de dos columnas, ordenadas: art GAMMA | GAMMA site

Sub buildDesafich()

    Dim c As Range
    Dim colArt As Integer, colSite As Integer: colArt = 1: colSite = 2
    
    'Primero chequeamos si el libro de partida está guardado
    If Not ActiveWorkbook.Saved Then
        MsgBox "Guardar el libro antes de ejecutar la macro"
        Exit Sub
    End If
    
    Dim d_UN As New Scripting.Dictionary            '"UN" : d_sites
    Dim d_sites As Scripting.Dictionary             '"site" : d_art_col
    Dim d_art_col As Scripting.Dictionary           '"site" : col_art
    Dim coll_art As Collection                      'artículos
    
    Dim site As String, article As String, UN As String
    
    'agrupamos los artículos por UN y site
    For Each c In Selection.Columns(colSite).SpecialCells(xlCellTypeVisible)
    
        site = c.value
        UN = Left(site, 2)
        article = CStr(c.Offset(0, -1).value)
     
        If Not d_UN.Exists(UN) Then
            
            Set coll_art = New Collection
            coll_art.Add Key:=article, item:=article
            
            Set d_art_col = New Scripting.Dictionary
            d_art_col.Add Key:=site, item:=coll_art
            
            Set d_sites = New Scripting.Dictionary
            d_sites.Add Key:=site, item:=d_art_col
            
            d_UN.Add Key:=UN, item:=d_sites
            
        Else
            
            If d_UN(UN).Exists(site) Then
                d_UN(UN)(site)(site).Add Key:=article, item:=article
            Else
                
                Set coll_art = New Collection
                coll_art.Add Key:=article, item:=article
                
                Set d_art_col = New Scripting.Dictionary
                d_art_col.Add Key:=site, item:=coll_art
                
                d_UN(UN).Add Key:=site, item:=d_art_col
                
            End If
            
        End If
    
    Next c
    
   

    'Creamos los wb separados por UN y creamos un txt con los sites a pegar en el DESAFICH en GAMMA
    Set aWb = ActiveWorkbook
    
    Dim key1, key2, art, wb As Workbook, x As Integer
    Dim textToPrint As String   'sites en el TXT según UN

    For Each key1 In d_UN
'        agregamos la UN al string a imprimir en el txt
        If textToPrint = "" Then
            textToPrint = key1 & ": "
        Else
            textToPrint = textToPrint & vbNewLine & key1 & ": "
        End If

        Set wb = Workbooks.Add
        x = 1

        For Each key2 In d_UN(key1)
'            agregamos los sites al string a imprimir en el txt
            textToPrint = textToPrint & Right(key2, 4) & ","

            For Each art In d_UN(key1)(key2)(key2)
                With wb.Sheets(1)
                    .Cells(x, 1).value = key2
                    .Cells(x, 3).value = art
                    x = x + 1
                End With
            Next art
        Next key2

        wb.SaveAs Filename:=getFileName(aWb.Name) & " - " & key1 & ".xlsx"
        wb.Close

    Next key1

    writeTo_txt textToPrint


End Sub


Function getFileName(originalName As String)

    Dim position As Integer, length As Integer, extension As String, Path As String
    
    Path = "C:\GAMMAF\"
    
    position = InStrRev(StrConv(originalName, vbLowerCase), ".xl")
    extension = Mid(originalName, position, Len(originalName) - (position - 1))
    
    getFileName = Path & "DESAFICH - " & Replace(originalName, extension, "")
    
End Function

    
Sub writeTo_txt(text As String)
    
    'we need this, even though it doesn't make much sense!
    Dim fso As New FileSystemObject
    
    'the file we're going to write to
    Dim ts As TextStream
    
    'open this file to write to it
    Set ts = fso.CreateTextFile("C:\GAMMAF\info.txt", True)
    
    'write out a couple of lines
    ts.WriteLine (text)
    
    Workbooks.OpenText "C:\GAMMAF\info.txt"
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------------------------------------------'

Sub MACROGLO_title()
Attribute MACROGLO_title.VB_ProcData.VB_Invoke_Func = "e\n14"

    Dim arr(1 To 100) As String
    
    arr(1) = ""
    arr(2) = "U.N"
    arr(3) = "Centro"
    arr(4) = ""
    arr(5) = "Articulo SAP"
    arr(6) = "Proveedor"
    arr(7) = ""
    arr(8) = "VAN"
    arr(9) = "Codigo EAN"
    arr(10) = "Descripción artículo"
    arr(11) = "Idioma"
    arr(12) = "Desc.POS Art."
    arr(13) = "Def. producto"
    arr(14) = "Estado Articulo"
    arr(15) = "Fecha Afectación/De."
    arr(16) = ""
    arr(17) = "Cat. Dufry"
    arr(18) = "Agrupación"
    arr(19) = "Marca Dufry"
    arr(20) = "Marca"
    arr(21) = "Línea Dufry"
    arr(22) = ""
    arr(23) = ""
    arr(24) = "Fabricante"
    arr(25) = "Alto"
    arr(26) = "Largo"
    arr(27) = "Profundo"
    arr(28) = ""
    arr(29) = "Peso neto"
    arr(30) = "Peso Bruto"
    arr(31) = ""
    arr(32) = "Fecha de lanzamiento"
    arr(33) = "Fecha disponibilidad"
    arr(34) = ""
    arr(35) = "Temporada"
    arr(36) = "Año colección"
    arr(37) = "Posición estadística"
    arr(38) = ""
    arr(39) = ""
    arr(40) = ""
    arr(41) = "Unidad estadística"
    arr(42) = "Denom."
    arr(43) = "Numer."
    arr(44) = "Tercer país"
    arr(45) = "Pais origen"
    arr(46) = ""
    arr(47) = "Precio de compra"
    arr(48) = "Divisa PC"
    arr(49) = "Fecha  PC"
    arr(50) = ""
    arr(51) = ""
    arr(52) = ""
    arr(53) = ""
    arr(54) = "Divisa IC"
    arr(55) = ""
    arr(56) = "Descuento IC"
    arr(57) = "PVP"
    arr(58) = "Divisa PVP"
    arr(59) = "Fecha vigor PVP"
    arr(60) = "Caducidad"
    arr(61) = "%Grado plato"
    arr(62) = "UoM Compra"
    arr(63) = "UoM Precio"
    arr(64) = "UoM Venta"
    arr(65) = "UoM Packing"
    arr(66) = "Cant. Compra-Venta"
    arr(67) = "Cant. Compra-Venta"
    arr(68) = "Cant. Compra-Venta"
    arr(69) = "Cant. Compra-Venta"
    arr(70) = "Cant.packing"
    arr(71) = ""
    arr(72) = ""
    arr(73) = ""
    arr(74) = ""
    arr(75) = ""
    arr(76) = "Reaprov SAP F&R"
    arr(77) = "Tipo de aprov."
    arr(78) = ""
    arr(79) = ""
    arr(80) = ""
    arr(81) = ""
    arr(82) = ""
    arr(83) = ""
    arr(84) = ""
    arr(85) = ""
    arr(86) = ""
    arr(87) = ""
    arr(88) = ""
    arr(89) = ""
    arr(90) = ""
    arr(91) = ""
    arr(92) = ""
    arr(93) = ""
    arr(94) = ""
    arr(95) = ""
    arr(96) = ""
    arr(97) = ""
    arr(98) = "Usuario SAP"
    arr(99) = ""
    arr(100) = ""


    ActiveSheet.Rows(1).EntireRow.Insert
    
    For i = 1 To 100
        ActiveSheet.Cells(1, i).value = arr(i)
    Next
    
End Sub

