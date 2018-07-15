Attribute VB_Name = "mdListing"
Option Explicit

Sub getlisting()

    Dim wbS As Workbook, wsS As Worksheet   'template/hoja listing
    Set wbS = ActiveWorkbook
    Set wsS = wbS.Sheets("Store Listing")
    
    If wsS.Range("A" & Rows.Count).End(xlUp).Row = 24 Then: MsgBox "No hay artículos": Exit Sub
    
    'estructura GAMMA
    Dim wbGAMMA As Workbook, estGAMMA() As Variant
    Set wbGAMMA = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Estructura Gamma-Sap.xlsx", _
                               UpdateLinks:=False, ReadOnly:=True)
    wbGAMMA.Sheets("Enterprise Struct in SAP Corp").Activate
    estGAMMA = Range("A6").CurrentRegion.value  'array con estructura GAMMA
    wbGAMMA.Close savechanges:=False
    
    
    Dim colSites As New Collection  'columnas con datos en el listing
    Dim artRng As Range, a As Range, centralizado As Boolean
    Set colSites = getColsList(wsS)
    wsS.Activate
    Set artRng = wsS.Range(Range("A25"), Range("A" & Rows.Count).End(xlUp))
    If wbS.Worksheets(1).Range("E9").value = "Yes" Then: centralizado = True
    
    Dim wbDest As Workbook, counter As Integer, i As Integer
    Set wbDest = Workbooks.Add
    
    With wbDest.Sheets(1)
        .Range("A1").value = "Art #"
        .Range("B1").value = "Site SAP"
        .Range("C1").value = "Site GAMMA"
        If centralizado Then: .Range("D1").value = "Cliente IOS"
    End With
    
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")

    RE.Global = True: RE.IgnoreCase = True
    RE.Pattern = "Z\w{3}"
    counter = 2
    For Each a In artRng
        Dim sites() As Variant
        sites = getSites(wsS, a.Row, estGAMMA, colSites)
        For i = 1 To UBound(sites)
        If Not RE.Test(sites(i)) Then
            With wbDest.Sheets(1)
                .Range("A" & counter).value = a.value
                .Range("B" & counter).value = sites(i)
                On Error Resume Next
                .Range("C" & counter).value = Application.WorksheetFunction.VLookup(sites(i), estGAMMA, 3, 0)
                If centralizado Then: .Range("D" & counter).value = Application.WorksheetFunction.VLookup(sites(i), estGAMMA, 10, 0)
                counter = counter + 1
            End With
        End If
        Next i
    Next a
    
End Sub
'devuelve las columnas con datos del listing
Function getColsList(ByVal wsListing As Worksheet) As Collection
    
    Dim coll_sites As New Collection, i As Integer
    
    For i = 9 To 1500
        If Not wsListing.Columns(i).Hidden Then
            If wsListing.Cells(Rows.Count, i).End(xlUp).Row > 24 Then
                coll_sites.Add i
            End If
        End If
    Next i
    
    Set getColsList = coll_sites

End Function

'Parámetros:
    'ws: worksheet listing,
    'r: fila del artículo
    'arr: estructuraGAMMA
    'coll_sites: columnas con datos
Function getSites(ByVal ws As Worksheet, r As Integer, arr, coll_sites As Collection) As Variant
    
    Dim f_sites() As Variant, col, site As String
    Dim row_sites As Integer: row_sites = 23
    Dim coll_listing As New Collection
    Dim arr_sites() As Variant, prefSite As String
    arr_sites = arr
    
    
    For Each col In coll_sites
        If ws.Cells(r, col) <> "" Then
        site = ws.Cells(row_sites, col).value
        If site = "UYMA" Or site = "UYMB" Or site = "ECGA" Then: site = cambioSite(site)
            coll_listing.Add Key:=site, item:=site
            On Error Resume Next
            prefSite = Application.WorksheetFunction.VLookup(site, arr, 9, 0)
            Do While prefSite <> "" And prefSite <> "stop"
                coll_listing.Add Key:=prefSite, item:=prefSite
                If Application.WorksheetFunction.VLookup(prefSite, arr, 9, 0) = prefSite Then
                    prefSite = "stop"
                Else
                    prefSite = Application.WorksheetFunction.VLookup(prefSite, arr, 9, 0)
                End If
            Loop
        End If
    Next col

    
    If coll_listing.Count = 0 Then
        ReDim f_sites(1 To 1)
        f_sites(1) = "Sin listing"
        getSites = f_sites
    Else
        Dim i As Integer
        ReDim f_sites(1 To coll_listing.Count)
        For i = 1 To coll_listing.Count
            f_sites(i) = coll_listing(i)
        Next i
        getSites = f_sites
    End If
    

End Function

Function cambioSite(site As String) As String

    Select Case site
        Case Is = "UYMA"
            cambioSite = "UY10"
        Case Is = "UYMB"
            cambioSite = "UY20"
        Case Is = "ECGA"
            cambioSite = "EC01"
    End Select
    
End Function

