Attribute VB_Name = "mdToVIS"
Sub show_ufrm()
    
    ufrmDataOptions.cbBasicData.value = True
    ufrmDataOptions.cbPurchData.value = True
    ufrmDataOptions.cbListing.value = True
    ufrmDataOptions.cbRetail.value = True
    
    ufrmDataOptions.Show vbModeless
    
End Sub

Sub tempToVIS()

    Dim wbS       'template
    Set wbS = ActiveWorkbook
    
    'seteamos las ws de donde sacar los datos
    Dim wsGD As Worksheet, wsListing As Worksheet, wsRetail As Worksheet
    Set wsGD = wbS.Sheets("General Data")
    Set wsListing = wbS.Sheets("Store Listing")
    Set wsRetail = wbS.Sheets("Retail Price")
    
    'chequeamos que hayan N° de artículos
    wsGD.Activate
    Dim ID_range As Range, c As Range
    If isEmpty(wsGD.Range("C18")) Then
        MsgBox "No hay articulos"
        Exit Sub
    Else
        Set ID_range = wsGD.Range(Range("C18"), Range("C" & Rows.Count).End(xlUp))
    End If
    
    'creamos una copia del template por cada "acción" (modificar/crear/extender)
    Dim wbCreate As Workbook, wbExtend As Workbook, wbModify As Workbook
    Dim cAction As Range
    wsGD.Activate
    For Each cAction In wbS.Sheets(1).Range(Range("A18"), Range("A18").End(xlDown))
        If wbCreate Is Nothing Then
            If cAction.value = "Create" Then: Set wbCreate = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Areas de trabajo\Final Template_NEW_VIS_V2.1 (plantilla).xltx", local:=True)
        End If
        If wbExtend Is Nothing Then
            If cAction.value = "Extend" Then: Set wbExtend = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Areas de trabajo\Final Template_NEW_VIS_V2.1 (plantilla).xltx", local:=True)
        End If
        If wbModify Is Nothing Then
            If cAction.value = "Modify" Then: Set wbModify = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Areas de trabajo\Final Template_NEW_VIS_V2.1 (plantilla).xltx", local:=True)
        End If
    Next
    
    
    'matriz con estructura GAMMA
    If ufrmDataOptions.cbListing.value = True Then
    
        Dim arrSites() As Variant
        Dim wbGAMMA As Workbook
        Set wbGAMMA = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\Accesos Directos\Files & Location\Estructura Gamma-Sap.xlsx", _
                                   UpdateLinks:=False, ReadOnly:=True)
        wbGAMMA.Sheets("Enterprise Struct in SAP Corp").Activate
        arrSites = Range("A6").CurrentRegion.value
        wbGAMMA.Close savechanges:=False
        
        Dim coll_sites As New Collection
        Set coll_sites = getColsList(wsListing)
        
        Dim centralizado As Boolean
        If wsGD.Range("E9").value = "Yes" Then: centralizado = True
        
    End If
    
    
    'levantamos los datos de los artículos dependiendo de que checkboxes se hayan marcado y los
    'agregamos a la colección coll_arts
    Dim coll_arts As New Collection
    Dim art As article
    
    For Each c In ID_range
    
        Set art = New article
            
            art.Action = c.Offset(0, -2).value
            
            'seteamos el VIS destino de cada artículo
            If art.Action = "Create" Then
                Set art.wbDest = wbCreate
            ElseIf art.Action = "Extend" Then
                Set art.wbDest = wbExtend
            ElseIf art.Action = "Modify" Then
                Set art.wbDest = wbModify
            End If
            
            art.ID = c.value
            
            'Basic Data
            If ufrmDataOptions.cbBasicData.value = True Then
                With wsGD
                    art.ArticleType = Left(.Cells(c.Row, ArticleType_s).value, 4)
                    art.Merch_Category = Left(.Cells(c.Row, Category_s).value, 7)
                    art.Desc = .Cells(c.Row, Description_s).value
                    art.SAPBrand = .Cells(c.Row, S_Brand_s).value
                    art.CO = .Cells(c.Row, CountryOfOr_s).value
                    art.EAN = .Cells(c.Row, EAN_s).value
                    art.GWeight = .Cells(c.Row, GrossW_s).value
                    art.NWeight = .Cells(c.Row, NetW_s).value
                    art.Lenght = .Cells(c.Row, Lenght_s).value
                    art.Width = .Cells(c.Row, Width_s).value
                    art.Height = .Cells(c.Row, Height_s).value
                    art.DCategory = .Cells(c.Row, D_Category_s).text
                    art.DGroup = .Cells(c.Row, D_Group_s).text
                    art.DSubGroup = .Cells(c.Row, D_SubGroup_s).text
                    art.DBrand = .Cells(c.Row, D_Brand_s).value
                    art.DLine = .Cells(c.Row, D_Line_s).value
                    art.DMan = .Cells(c.Row, D_Man_s).value
                    If art.DCategory = 70 Then
                        art.CharProfile = .Cells(c.Row, CharProfile_s).value
                        art.Color = .Cells(c.Row, Color_s).value
                        art.Talle = .Cells(c.Row, Size_s).value
                        art.Season = .Cells(c.Row, Season_s).value
                        art.artYear = .Cells(c.Row, Year_s).value
                    ElseIf art.DCategory = 30 Then
                        art.ShelfLife = .Cells(c.Row, shelfLife_s).value
                    End If
                End With
            End If
            
            'purch Data
            If ufrmDataOptions.cbPurchData.value = True Then
                Dim purchInfo As New Collection
                Set purchInfo = getPurchInfo(wsGD, c.Row)
                art.purchPrice = purchInfo(1)
                art.purchDIV = purchInfo(2)
                art.purchORG = purchInfo(3)
                With wsGD
                    art.purchGrp = .Cells(c.Row, purchGroup_s).value
                    art.Vendor = .Cells(c.Row, vendor_s).value
                    art.van = .Cells(c.Row, VAN_s).value
                    art.MinOrder = .Cells(c.Row, minOrderQty_s).value
                End With
            End If
            
            
            'listing
            If ufrmDataOptions.cbListing.value = True Then
                art.CommCode = wsGD.Cells(c.Row, commCode_s).value
                With wsListing
                    Dim sites() As Variant
                    sites = getSites(wsListing, c.Row, arrSites, coll_sites)
                    art.Listing = sites
                End With
            End If
        
            'retails
            If ufrmDataOptions.cbRetail.value = True Then
                art.TaxData = wsGD.Cells(c.Row, TaxData_s).value
                With wsRetail
                    Dim Retails As New Collection
                    Set Retails = getReails(wsRetail, c.Row, art.TaxData)
                    art.RetailDF = Retails(1)
                    art.RetailDP = Retails(2)
                    art.RetailDIV = Retails(3)
                    art.RetailORG = Retails(4)
                End With
            End If
        
        coll_arts.Add art
        
    Next c
    
    
    If ufrmDataOptions.cbListing.value = False Then
        art.toVIS coll_arts
    Else
        art.toVIS coll_arts, arrSites, centralizado
    End If
    
    
End Sub
'devuelve una colección de arrays con los valores purch price, purch div y purch org.
Function getPurchInfo(ByVal ws As Worksheet, r As Integer) As Collection

    Dim cols As New Collection
    Dim info As New Collection

    For i = 68 To 480 Step 4
        If Not ws.Columns(i).Hidden Then
            If ws.Cells(Rows.Count, i).End(xlUp).Row > 17 Then
                cols.Add i
            End If
        End If
    Next i
    
    If cols.Count = 0 Then
        Exit Function
    End If
    
    Dim prices() As Variant, DIVs() As Variant, ORGs() As Variant
    ReDim prices(1 To cols.Count)
    ReDim DIVs(1 To cols.Count)
    ReDim ORGs(1 To cols.Count)
    
    For i = 1 To cols.Count
        prices(i) = Round(ws.Cells(r, cols(i)).value, 2)
        DIVs(i) = ws.Cells(r, cols(i) + 1).value
        ORGs(i) = ws.Cells(16, cols(i) + 3).value
    Next i
    
    info.Add Key:="Price", item:=prices
    info.Add Key:="DIV", item:=DIVs
    info.Add Key:="ORG", item:=ORGs
        
    
    Set getPurchInfo = info

End Function

'devuelve los sites (incluyendo los preferentes y cambiando los UYMA, UYMB y ECGA) para cada item como un array
'sino tiene listing devuelve "sin listing"
'Parámetros:
    'ws: worksheet listing del template
    'r: fila del artículo para el cual sacamos los sites
    'arr: estructura gamma
    'coll_sites: columnas donde posiblemente hay un check del site, la cual sale con la función "getColsList"
    
Function getSites(ByVal ws As Worksheet, r As Integer, arr, coll_sites As Collection) As Variant
    
    Dim f_sites() As Variant, col, site As String
    Dim row_sites As Integer: row_sites = 23
    Dim coll_listing As New Collection
    Dim arr_sites() As Variant, prefSite As String
    arr_sites = arr
    
    For Each col In coll_sites
        If ws.Cells(r + 7, col) <> "" Then
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
        ReDim f_sites(1 To coll_listing.Count)
        For i = 1 To coll_listing.Count
            f_sites(i) = coll_listing(i)
        Next i
        getSites = f_sites
    End If

End Function

'devuleve una colección de arrays con los retails DP, DF las divisas y las sales ORG
'Parámetros:
    'ws: ws retails del template
    'r: row del item del cual queremos los datos
    'tax: si está especificado el tax en la ws "General Data", calcula el retail agregando el tax.
    '     el dato del tax lo tomamos como integer entre 1 y 100

Function getReails(ByVal ws As Worksheet, r As Integer, tax As Integer) As Collection

    Dim cols As New Collection
    Dim info As New Collection
    Dim df As Integer: df = 0
    Dim dp As Integer: dp = 0
    
    'tomamos las columnas con datos
    For i = 9 To 3404 Step 4
        If Not ws.Columns(i).Hidden Then
            If ws.Cells(Rows.Count, i).End(xlUp).Row > 19 Or ws.Cells(Rows.Count, i + 1).End(xlUp).Row > 19 Then
                cols.Add i
            End If
        End If
    Next i
    
    If cols.Count = 0 Then
        Exit Function
    End If
    
    
    Dim retailsDF() As Variant, DIVsDF() As Variant, retailsDP() As Variant, DIVsDP() As Variant, ORGs() As Variant
    ReDim retailsDF(1 To cols.Count)
    ReDim retailsDP(1 To cols.Count)
    ReDim DIVs(1 To cols.Count)
    ReDim ORGs(1 To cols.Count)
    
    For i = 1 To cols.Count
        If tax = 0 Then
            retailsDF(i) = Round(ws.Cells(r + 2, cols(i)).value, 2)
            retailsDP(i) = Round(ws.Cells(r + 2, cols(i) + 1).value, 2)
            DIVs(i) = ws.Cells(17, cols(i)).value
            ORGs(i) = ws.Cells(16, cols(i)).value
        Else
            retailsDF(i) = Round(ws.Cells(r + 2, cols(i)).value * (1 + tax / 100), 2)
            retailsDP(i) = Round(ws.Cells(r + 2, cols(i) + 1).value * (1 + tax / 100), 2)
            DIVs(i) = ws.Cells(17, cols(i)).value
            ORGs(i) = ws.Cells(16, cols(i)).value
        End If
    Next i
    
    info.Add Key:="DF", item:=retailsDF
    info.Add Key:="DP", item:=retailsDP
    info.Add Key:="DIV", item:=DIVs
    info.Add Key:="ORG", item:=ORGs
        
    Set getReails = info
    Set info = Nothing

End Function
'devuelve una colección con los nros de columna en donde hay algún check de listing. recorre solo las columnas visibles
Function getColsList(ByVal wslinting As Worksheet) As Collection
    
    Dim coll_sites As New Collection
    
    For i = 9 To 1500
        If Not wslinting.Columns(i).Hidden Then
            If wslinting.Cells(Rows.Count, i).End(xlUp).Row > 24 Then
                coll_sites.Add i
            End If
        End If
    Next i
    
    Set getColsList = coll_sites

End Function
'cambiamos el nombre de estos clientes para buscar en la estructura GAMMA
Function cambioSite(site As String)

    Select Case site
        Case Is = "UYMA"
            cambioSite = "UY10"
        Case Is = "UYMB"
            cambioSite = "UY20"
        Case Is = "ECGA"
            cambioSite = "EC01"
    End Select

End Function

