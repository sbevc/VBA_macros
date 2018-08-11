VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Article"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pAction, pID, pArticleType, pMerch_Category, pDesc, pSAPBrand, pCO, pEAN, pGWeight, pNWeight, pLenght, pWidth, pHeight, _
pDCategory, pDGroup, pDSubGroup, pDBrand, pDLine, pDMan, pPurchPrice, pPurchDIV, pPurchORG, pPurchGrp, pVendor, pVAN, _
pMinOrder, pUOfMeas, pDateFrom, pDateTo, pListing, pCommCode, pRetailDF, pRetailDP, pRetailDIV, pRetailORG, temp, pWDest, pTaxData As Integer, _
pColor, pTalle, pCharProfile, pShelfLife, pSeason, partYear

Dim allMatches As Object, RE As Object
'BASIC DATA
Public Property Get Action()
    Action = pAction
End Property
Public Property Let Action(value As Variant)
    pAction = value
End Property
Public Property Get wbDest()
    Set wbDest = pWDest
End Property
Public Property Set wbDest(ByVal wb As Workbook)
    Set pWDest = wb
End Property
Public Property Get ID()
    ID = pID
End Property
Public Property Let ID(value As Variant)
    pID = value
End Property
Public Property Get ArticleType()
    ArticleType = pArticleType
End Property
Public Property Let ArticleType(value As Variant)
    pArticleType = value
End Property
Public Property Get Merch_Category()
    Merch_Category = pMerch_Category
End Property
Public Property Let Merch_Category(value As Variant)
    pMerch_Category = value
End Property
Public Property Get Desc()
    Desc = pDesc
End Property
Public Property Let Desc(value As Variant)
    pDesc = Trim(CStr(value))
End Property
Public Property Get SAPBrand()
    SAPBrand = pSAPBrand
End Property
Public Property Let SAPBrand(value As Variant)
    If value = "" Then
        pSAPBrand = 99999
    Else
        pSAPBrand = value
    End If
End Property
Public Property Get CO()
    CO = pCO
End Property
Public Property Let CO(value As Variant)
    If Len(value) = 2 Then
        pCO = value
    Else
        pCO = Right(value, 2)
    End If
End Property
Public Property Get EAN()
    EAN = pEAN
End Property
Public Property Let EAN(value As Variant)
    pEAN = Trim(value)
End Property
Public Property Get GWeight()
    GWeight = pGWeight
End Property
Public Property Let GWeight(value As Variant)
    If value = 0 Or value = "" Then
        pGWeight = 1
    Else
        pGWeight = cleanValue(value)
    End If
End Property
Public Property Get NWeight()
    NWeight = pNWeight
End Property
Public Property Let NWeight(value As Variant)
    If value = 0 Or value = "" Then
        pNWeight = 1
    Else
        pNWeight = cleanValue(value)
    End If
End Property
Public Property Get Lenght()
    Lenght = pLenght
End Property
Public Property Let Lenght(value As Variant)
    If value = 0 Or value = "" Then
        pLenght = 1
    Else
        pLenght = cleanValue(value)
    End If
End Property
Public Property Get Width()
    Width = pWidth
End Property
Public Property Let Width(value As Variant)
    If value = 0 Or value = "" Then
        pWidth = 1
    Else
        pWidth = cleanValue(value)
    End If
End Property
Public Property Get Height()
    Height = pHeight
End Property
Public Property Let Height(value As Variant)
    If value = 0 Or value = "" Then
        pHeight = 1
    Else
        pHeight = cleanValue(value)
    End If
End Property
Public Property Get DCategory()
    DCategory = pDCategory
End Property
Public Property Let DCategory(value As Variant)
    If InStr(value, "-") Then
        temp = Split(value, "-")
        pDCategory = Trim(temp(0))
    Else
        pDCategory = value
    End If
End Property
Public Property Get DGroup()
    DGroup = pDGroup
End Property
Public Property Let DGroup(value As Variant)
    If InStr(value, "-") Then
        temp = Split(value, "-")
        pDGroup = Trim(temp(0))
    Else
        pDGroup = value
    End If
End Property
Public Property Get DSubGroup()
    DSubGroup = pDSubGroup
End Property
Public Property Let DSubGroup(value As Variant)
    If InStr(value, "-") Then
        temp = Split(value, "-")
        pDSubGroup = Trim(temp(0))
    Else
        pDSubGroup = value
    End If
End Property
Public Property Get DBrand()
    DBrand = pDBrand
End Property
Public Property Let DBrand(value As Variant)

    Set RE = CreateObject("vbscript.regexp")
    
    With RE
        .Global = True
        .IgnoreCase = True
        .Pattern = "(\d+)"
    End With
    
    Set allMatches = RE.Execute(CStr(value))
    If allMatches.Count <> 0 Then
        pDBrand = allMatches(0)
    Else
        pDBrand = value
    End If
    
End Property
Public Property Get DLine()
    DLine = pDLine
End Property
Public Property Let DLine(value As Variant)
    If InStr(value, "-") Then
        temp = Split(value, "-")
        pDLine = Trim(temp(0))
    ElseIf InStr(StrConv(value, vbLowerCase), "no line") Then
        pDLine = 0
    ElseIf value = "" Then
        pDLine = 0
    Else
        pDLine = Trim(value)
    End If
End Property
Public Property Get DMan()
    DMan = pDMan
End Property
Public Property Let DMan(value As Variant)
    If InStr(StrConv(value, vbLowerCase), "unknown") Then
        pDMan = 0
    Else
        pDMan = cleanValue(value)
    End If
End Property

'PURCH DATA
Public Property Get purchPrice() As Variant
    purchPrice = pPurchPrice
End Property
Public Property Let purchPrice(value As Variant)
    pPurchPrice = value
End Property
Public Property Get purchDIV()
    purchDIV = pPurchDIV
End Property
Public Property Let purchDIV(value As Variant)
    pPurchDIV = value
End Property
Public Property Get purchORG() As Variant
    purchORG = pPurchORG
End Property
Public Property Let purchORG(ByVal value As Variant)
    pPurchORG = value
End Property
Public Property Get purchGrp()
    purchGrp = pPurchGrp
End Property
Public Property Let purchGrp(value As Variant)
    If InStr(value, "-") Then
        temp = Split(value, "-")
        pPurchGrp = Trim(temp(0))
    Else
        pPurchGrp = Trim(value)
    End If
End Property
Public Property Get Vendor()
    Vendor = pVendor
End Property
Public Property Let Vendor(value As Variant)
    pVendor = value
End Property
Public Property Get van()
   van = pVAN
End Property
Public Property Let van(value As Variant)
    pVAN = Trim(value)
End Property
Public Property Get MinOrder()
   MinOrder = pMinOrder
End Property
Public Property Let MinOrder(value As Variant)
    pMinOrder = value
End Property
Public Property Get UOfMeas()
   UOfMeas = pUOfMeas
End Property
Public Property Let UOfMeas(value As Variant)
    pUOfMeas = value
End Property
Public Property Get Listing()
   Listing = pListing
End Property
Public Property Let Listing(value As Variant)
    pListing = value
End Property
Public Property Get CommCode()
   CommCode = pCommCode
End Property
Public Property Let CommCode(value As Variant)
    If InStr(value, ".") Then
        pCommCode = Replace(value, ".", "")
    Else
        pCommCode = value
    End If
End Property
Public Property Get RetailDF() As Variant
    RetailDF = pRetailDF
End Property
Public Property Let RetailDF(ByVal value As Variant)
    pRetailDF = value
End Property
Public Property Get RetailDP() As Variant
    RetailDP = pRetailDP
End Property
Public Property Let RetailDP(ByVal value As Variant)
    pRetailDP = value
End Property
Public Property Get RetailDIV()
    RetailDIV = pRetailDIV
End Property
Public Property Let RetailDIV(value As Variant)
    pRetailDIV = value
End Property
Public Property Get RetailORG() As Variant
    RetailORG = pRetailORG
End Property
Public Property Let RetailORG(ByVal value As Variant)
    pRetailORG = value
End Property
Public Property Get TaxData()
    TaxData = pTaxData
End Property
Public Property Let TaxData(value As Variant)
    If value < 1 And value > 0 Then
        pTaxData = Round(value * 100, 0)
    ElseIf value > 1 Then
        pTaxData = Round(value, 0)
    End If
End Property
Public Property Get Color() As Variant
    Color = pColor
End Property
Public Property Let Color(ByVal value As Variant)
    pColor = value
End Property
Public Property Get Talle() As Variant
    Talle = pTalle
End Property
Public Property Let Talle(ByVal value As Variant)
    pTalle = value
End Property
Public Property Get CharProfile() As Variant
    CharProfile = pCharProfile
End Property
Public Property Let CharProfile(ByVal value As Variant)
    pCharProfile = value
End Property
Public Property Get ShelfLife() As Variant
    ShelfLife = pShelfLife
End Property
Public Property Let ShelfLife(ByVal value As Variant)
    pShelfLife = value
End Property
Public Property Get Season() As Variant
    Season = pSeason
End Property
Public Property Let Season(ByVal value As Variant)
    pSeason = cleanValue(value)
End Property
Public Property Get artYear() As Variant
    artYear = partYear
End Property
Public Property Let artYear(ByVal value As Variant)
    partYear = cleanValue(value)
End Property



Sub toVIS(articles As Collection, Optional estGAMMA As Variant, Optional centralizado As Boolean)
    
    Dim errors As New Scripting.Dictionary
    Dim wbListErr As Workbook  'wb para el listing/errores

    If ufrmDataOptions.cbBasicData.value = True Then
        toVIS_Basic articles, errors
    End If
    If ufrmDataOptions.cbPurchData.value = True Then
        toVIS_purch articles, errors
    End If
    If ufrmDataOptions.cbListing.value = True Then
        Set wbListErr = Workbooks.Add
        toVIS_listing articles, errors, wbListErr
    End If
    If ufrmDataOptions.cbRetail.value = True Then
        toVIS_retail articles, errors
    End If
    
    If wbListErr Is Nothing Then: Set wbListErr = Workbooks.Add
    
    If ufrmDataOptions.cbListing.value = False Then
        printListing_Errors errors, wbListErr
    ElseIf ufrmDataOptions.cbListing.value = True Then
        printListing_Errors errors, wbListErr, estGAMMA, centralizado
    End If
        
End Sub

Private Sub toVIS_Basic(articles As Collection, errors As Scripting.Dictionary)

    Dim hierarchies As New Collection, h As String
    Dim ManBrands As New Collection, mb As String
    Dim EANs As New Collection, VANs As New Collection  'collections para ver si hay duplicados
    
    '--------------------COPY_DATA--------------------'
    Dim article, x
    Dim xCreate As Integer: xCreate = 3
    Dim xModify As Integer: xModify = 3
    
    For Each article In articles
        'With wbDest.Sheets(1)
        With article.wbDest.Sheets(1)
        
            If article.Action <> "Extend" Then
            
                If article.Action = "Create" Then
                    x = xCreate
                ElseIf article.Action = "Modify" Then
                    x = xModify
                End If
            
                pasteData .Cells(x, ID_d), article.ID, "ID", errors, article.ID
                pasteData .Cells(x, ArticleType_d), article.ArticleType, "Article Type", errors, article.ID
                pasteData .Cells(x, Category_d), article.Merch_Category, "Merchandise Category", errors, article.ID
                pasteData .Cells(x, Description_d), article.Desc, "Article description", errors, article.ID
                
                .Cells(x, CharProfile_d).value = article.CharProfile
                .Cells(x, Color_d).value = article.Color
                .Cells(x, Size_d).value = article.Talle
                .Cells(x, Season_d).value = article.Season
                .Cells(x, Year_d).value = article.artYear
                
                pasteData .Cells(x, shelfLife_d), article.ShelfLife, "Shelf Life", errors, article.ID, article
                
                pasteData .Cells(x, S_Brand_d), article.SAPBrand, "SAP brand", errors, article.ID
                pasteData .Cells(x, CountryOfOr_d), article.CO, "Country Of Origin", errors, article.ID
                
                pasteData .Cells(x, EAN_d), article.EAN, "EAN", errors, article.ID

                .Cells(x, articleStatus_d).value = "Z3"

                .Cells(x, GrossW_d).NumberFormat = "@"                     'formato para que quede con separador de puntos
                pasteData .Cells(x, GrossW_d), cleanValue(article.GWeight), "Gross Weight", errors, article.ID
                .Cells(x, NetW_d).NumberFormat = "@"
                pasteData .Cells(x, NetW_d), cleanValue(article.NWeight), "Net Weight", errors, article.ID
                .Cells(x, wUnit).value = "G"
                .Cells(x, Lenght_d).NumberFormat = "@"
                pasteData .Cells(x, Lenght_d), cleanValue(article.Lenght), "Lenght", errors, article.ID
                .Cells(x, Width_d).NumberFormat = "@"
                pasteData .Cells(x, Width_d), cleanValue(article.Width), "Width", errors, article.ID
                .Cells(x, Height_d).NumberFormat = "@"
                pasteData .Cells(x, Height_d), cleanValue(article.Height), "Height", errors, article.ID
                .Cells(x, dimUnit).value = "CM"

                pasteData .Cells(x, D_Category_d), article.DCategory, "Dufry Category", errors, article.ID
                pasteData .Cells(x, D_Group_d), article.DGroup, "Dufry Group", errors, article.ID
                pasteData .Cells(x, D_SubGroup_d), article.DSubGroup, "Dufry SubGroup", errors, article.ID
                pasteData .Cells(x, D_Brand_d), article.DBrand, "Dufry Brand", errors, article.ID
                .Cells(x, D_Line_d).value = article.DLine
                pasteData .Cells(x, D_Man_d), article.DMan, "Dufry Manufacturer", errors, article.ID
    
                
                If article.Action = "Create" Then
                    xCreate = xCreate + 1
                ElseIf article.Action = "Modify" Then
                    xModify = xModify + 1
                End If
                
                h = article.DCategory & "-" & article.DGroup & "-" & article.DSubGroup
                mb = article.DMan & "-" & article.DBrand
                On Error Resume Next
                hierarchies.Add item:=CStr(h), Key:=CStr(h)
                ManBrands.Add item:=CStr(mb), Key:=CStr(mb)
                
            End If
        End With
    Next article
    
    
    '--------------------VALIDATE_DATA--------------------'
    
    Dim wbH As Workbook, arr_hier() As Variant, hier
    Dim wbManBrand As Workbook, arr_ManBrand() As Variant, ManBrand
    
    'validate hierarchy
    Set wbH = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\SAP\JERARQUIAS - FINALES REVISADA 20-11.xls", _
                               UpdateLinks:=False, ReadOnly:=True)
    wbH.Sheets(1).Activate
    arr_hier = Range("E2", Range("E2").End(xlDown)).value
    wbH.Close savechanges:=False

    For Each hier In hierarchies
        If Application.WorksheetFunction.VLookup(hier, arr_hier, 1, 0) = "" Then
            With wbDest.Sheets(1).Cells(x + 2, D_Category_d)
                .value = hier
                .Offset(0, 1).value = "No existe en la jerarquía Dufry"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            x = x + 1
        End If
    Next hier
    
    'validate man-brand
    Set wbManBrand = Workbooks.Open(Filename:="I:\Departments\LOGISTICS\Master Data\SAP\07Feb14_Cat_Brand_Man_Report.xlsx", _
                               UpdateLinks:=False, ReadOnly:=True)
    wbManBrand.Sheets("Cat_Brand_Man_Mapping").Activate
    Dim lastCell As Long, colMAN_CODE As Integer
    lastCell = ActiveSheet.UsedRange.Rows.Count
    colMAN_CODE = Application.WorksheetFunction.Match("MAN_CODE", ActiveSheet.Rows(1), 0)
    arr_ManBrand = Range(Cells(2, colMAN_CODE), Cells(lastCell, colMAN_CODE)).value
    wbManBrand.Close savechanges:=False
    
    For Each ManBrand In ManBrands
        If Application.WorksheetFunction.VLookup(ManBrand, arr_ManBrand, 1, 0) = "" Then
            With wbDest.Sheets(1).Cells(x + 2, D_Category_d)
                .value = ManBrand
                .Offset(0, 1).value = "No existe relacion Man-Brand"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            x = x + 1
        End If
    Next ManBrand
    
    wbDest.Sheets(1).Activate
    formatduplicates Range(Cells(3, EAN_d), Cells(Rows.Count, EAN_d).End(xlUp))
    
    Set hierarchies = Nothing
    Set VANs = Nothing
    Set ManBrands = Nothing
    Set EANs = Nothing

End Sub

Private Sub toVIS_purch(articles As Collection, errors As Scripting.Dictionary)

    Dim article, i
    
    Dim x As Integer
    Dim xCreate As Integer: xCreate = 3 'counter de fila para las creaciones
    Dim xExtend As Integer: xExtend = 3 'counter de fila para las extenciones
    Dim xModify As Integer: xModify = 3 'counter de fila para las extenciones
    
        For Each article In articles
        
            Dim arrPrices() As Variant
            Dim arrDIVs() As Variant
            Dim arrPOrgs() As Variant
            
            arrPrices = article.purchPrice
            arrDIVs = article.purchDIV
            arrPOrgs = article.purchORG
                        
            For i = 1 To UBound(arrPrices)
            
                With article.wbDest.Sheets(2)
                If arrPrices(i) <> "0" Then
                
                    If article.Action = "Create" Then
                        x = xCreate
                    ElseIf article.Action = "Extend" Then
                        x = xExtend
                    ElseIf article.Action = "Modify" Then
                        x = xModify
                    End If
                    
                    .Cells(x, ID_d).value = article.ID
    
                    .Cells(x, purchORG_d).value = arrPOrgs(i)
                    .Cells(x, purchPrice_d).NumberFormat = "@"
                    .Cells(x, purchPrice_d).value = cleanValue(arrPrices(i))
                    .Cells(x, purchDIV_d).value = arrDIVs(i)
                    pasteData .Cells(x, purchGroup_d), article.purchGrp, "Purchasing Group", errors, article.ID
                    pasteData .Cells(x, vendor_d), article.Vendor, "Vendor", errors, article.ID
                    pasteData .Cells(x, VAN_d), article.van, "VAN", errors, article.ID
                    .Cells(x, minOrderQty_d).value = article.MinOrder
                    If article.DCategory = 10 Or article.DCategory = 20 Or article.DCategory = 30 Then
                        .Cells(x, POuOfM_d).value = "CS"
                    Else
                        .Cells(x, POuOfM_d).value = "PC"
                    End If
                    .Cells(x, numToBase_d).value = 1
                    .Cells(x, denToBase_d).value = 1
                    .Cells(x, pricingUn_d).value = 1
                    .Cells(x, pricingUnMeas_d).value = "PC"
                    .Cells(x, purchDateFrom_d).value = Format(CStr(Day(Date)), "00") & "." & Format(CStr(Month(Date)), "00") & "." & CStr(Year(Date))
                    .Cells(x, purchDateTo_d).value = "31.12.9999"
                    
                If article.Action = "Create" Then
                    xCreate = xCreate + 1
                ElseIf article.Action = "Extend" Then
                    xExtend = xExtend + 1
                ElseIf article.Action = "Modify" Then
                    xModify = xModify + 1
                End If
                
                End If
                End With
            Next i
        Next article
    
End Sub

Sub toVIS_listing(articles As Collection, errors As Scripting.Dictionary, ByVal wbListErr As Workbook)
    
    Dim article, i
    
    Dim x As Integer
    Dim xCreate As Integer: xCreate = 3 'counter de fila para las creaciones
    Dim xExtend As Integer: xExtend = 3 'counter de fila para las extenciones
    Dim xModify As Integer: xModify = 3 'counter de fila para las extenciones
    
    Dim arrSites() As Variant
    
    Dim regex As New RegExp
    With regex
        .Global = True
        .IgnoreCase = True
    End With
    
    For Each article In articles
        arrSites = article.Listing
        For i = 1 To UBound(arrSites)
        
        regex.Pattern = "Z\w{3}"    'para sacar las tiendas que empiezan con Z del listing
        If arrSites(i) = "Sin listing" Or regex.Test(arrSites(i)) Then
            If arrSites(i) = "Sin listing" Then: addToErrors "Falta Listing", errors, article.ID: GoTo NextArticle
            If regex.Test(arrSites(i)) Then: addToErrors "Site " & arrSites(i) & " no válido", errors, article.ID: GoTo Next_i
        Else
            
            With wbListErr.Sheets(1)
                .Range("A" & Rows.Count).End(xlUp).Offset(1, 0).value = article.ID
                .Range("B" & Rows.Count).End(xlUp).Offset(1, 0).value = arrSites(i)
            End With
            
        
            With article.wbDest.Sheets(3)
                
                If article.Action = "Create" Then
                    x = xCreate
                ElseIf article.Action = "Extend" Then
                    x = xExtend
                ElseIf article.Action = "Modify" Then
                    x = xModify
                End If
        
                .Cells(x, ID_d).value = article.ID
                .Cells(x, site_d).value = arrSites(i)
                .Cells(x, l_fromDate).value = Format(CStr(Day(Date)), "00") & "." & Format(CStr(Month(Date)), "00") & "." & CStr(Year(Date))
                .Cells(x, l_toDate).value = "31.12.9999"
                .Cells(x, l_artStatus).value = "Z3"
                pasteData .Cells(x, commCode_d), article.CommCode, "Commodity Code", errors, article.ID
                
                regex.Pattern = "[A-Z]{4}|[A-Z]{3}\d"  'Tiendas AAAA o AAA0
                If regex.Test(arrSites(i)) Then
                    .Cells(x, RepType_d).value = "ND"
                    .Cells(x, SourceOfSupp_d).value = 2
                Else
                    regex.Pattern = "([A-Z]{2})([0-9]{2})"  'Almacenes
                    If regex.Test(arrSites(i)) Then
                        .Cells(x, RepType_d).value = "RP"
                        .Cells(x, SourceOfSupp_d).value = 1
                    End If
                End If
                
                .Cells(x, GRProcTime_d).value = 1
                
                If article.Action = "Create" Then
                    xCreate = xCreate + 1
                ElseIf article.Action = "Extend" Then
                    xExtend = xExtend + 1
                ElseIf article.Action = "Modify" Then
                    xModify = xModify + 1
                End If
                
            End With
            
        End If
Next_i:
        Next i
NextArticle:
    Next article
    
End Sub


Sub toVIS_retail(articles As Collection, errors As Scripting.Dictionary)

    Dim article, i
    
    Dim x As Integer
    Dim xCreate As Integer: xCreate = 3 'counter de fila para las creaciones
    Dim xExtend As Integer: xExtend = 3 'counter de fila para las extenciones
    Dim xModify As Integer: xModify = 3 'counter de fila para las extenciones

    For Each article In articles
        
            Dim arrDF() As Variant
            Dim arrDP() As Variant
            Dim arrDIVs() As Variant
            Dim arrOrgs() As Variant
            
            arrDF = article.RetailDF
            arrDP = article.RetailDP
            arrDIVs = article.RetailDIV
            arrOrgs = article.RetailORG
            
            For i = 1 To UBound(arrOrgs)
            
                With article.wbDest.Sheets(4)
                If arrDF(i) <> 0 Or arrDP(i) <> 0 Then
                
                    If article.Action = "Create" Then
                        x = xCreate
                    ElseIf article.Action = "Extend" Then
                        x = xExtend
                    ElseIf article.Action = "Modify" Then
                        x = xModify
                    End If
                
                    .Cells(x, ID_d).value = article.ID
                    pasteData .Cells(x, salesOrg_d), arrOrgs(i), "Sales Org", errors, article.ID
                    .Cells(x, channDist).value = "01"
                    .Cells(x, r_fromDate).value = Format(CStr(Day(Date)), "00") & "." & Format(CStr(Month(Date)), "00") & "." & CStr(Year(Date))
                    .Cells(x, r_toDate).value = "31.12.9999"
                    If arrDF(i) <> 0 Then
                        .Cells(x, DFprice_d).NumberFormat = "@"
                        .Cells(x, DFprice_d).value = cleanValue(arrDF(i))
                        .Cells(x, DFcurr_d).value = arrDIVs(i)
                    End If
                    If arrDP(i) <> 0 Then
                        .Cells(x, DPprice_d).NumberFormat = "@"
                        .Cells(x, DPprice_d).value = cleanValue(arrDP(i))
                        .Cells(x, DPcurr_d).value = arrDIVs(i)
                    End If
                    
                    If article.Action = "Create" Then
                        xCreate = xCreate + 1
                    ElseIf article.Action = "Extend" Then
                        xExtend = xExtend + 1
                    ElseIf article.Action = "Modify" Then
                        xModify = xModify + 1
                    End If
                    
                End If
                End With
            Next i
        Next article



End Sub
  
Sub formatduplicates(rg As Range)
    Dim uv As UniqueValues
    Set uv = rg.FormatConditions.AddUniqueValues
    uv.DupeUnique = xlDuplicate
    uv.Interior.Color = vbRed
End Sub
'extrae solo los números, incluyendo los puntos y comas si los hay
Function cleanValue(value As Variant)
    Set RE = CreateObject("vbscript.regexp")

    RE.Global = True: RE.IgnoreCase = True
    RE.Pattern = "\d*(\.|,)?\d+"
    
    Set allMatches = RE.Execute(value)
    If allMatches.Count <> 0 Then
        cleanValue = allMatches.item(0).value
    Else
        cleanValue = value
    End If
    
End Function

'pegamos los datos en el VIS, y si falta el dato agregamos al dict de errores
Sub pasteData(dest As Range, value As Variant, prop As String, errors As Scripting.Dictionary, artID, Optional ByVal art As article)

    dest.value = value

    Select Case prop
        Case Is = "EAN"
            'dest.value = value
            If Len(value) <> 13 Then
                If Len(value) = 0 Then: dest.Offset(0, 1).value = "IE"
                addToErrors prop, errors, artID, dest
            End If
        Case Is = "Shelf Life"
            If art.DCategory = 30 Then
                If art.ShelfLife = "" Then
                    addToErrors "Shelf Life", errors, artID, dest
                    dest.value = 210
                End If
            End If
        Case Else
            If value = "" Then: addToErrors prop, errors, artID, dest
    End Select
    
End Sub

'agregamos al diccionario de errores y pintamos la celda destino de amarillo
Sub addToErrors(prop As String, errors As Scripting.Dictionary, artID, Optional dest As Range)
    
    If Not dest Is Nothing Then
        With dest.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
    
    If errors.Exists(prop) = False Then
        Dim coll_err As New Collection
        coll_err.Add artID
        errors.Add Key:=prop, item:=coll_err
    ElseIf errors.Exists(prop) Then
        errors(prop).Add artID
    End If
    
End Sub

'ponemos los errores y los demás sites(gamma y/o IOSC en función de si es centralizada o no la compra)
Sub printListing_Errors(errors As Scripting.Dictionary, wbListErr As Workbook, Optional estGAMMA As Variant, Optional centralizado As Boolean)

    Dim wsErrors As Worksheet, wsListing As Worksheet
    
    If isEmpty(wbListErr.Sheets(1).Range("A2")) Then
        Set wsErrors = wbListErr.Sheets(1)
    Else
        Set wsListing = wbListErr.Sheets(1)
        wsListing.Name = "Listing"
        wsListing.Range("A1").value = "Art #"
        wsListing.Range("B1").value = "Site"
        Set wsErrors = wbListErr.Sheets.Add(After:=wbListErr.Sheets(1))
        wsErrors.Name = "Errores"
    End If
    
    
    Dim counter As Integer: counter = 1
    Dim k As Variant, v
    Dim c As Range
    
    If Not IsMissing(estGAMMA) Then
        With wsListing
            .Range("C1").value = "Site GAMMA"
            If centralizado Then: .Range("D1").value = "Cliente IOSC"
            On Error Resume Next
            For Each c In .Range(.Range("B2"), .Range("B" & Rows.Count).End(xlUp))
                c.Offset(0, 1).value = Application.WorksheetFunction.VLookup(c.value, estGAMMA, 3, 0)
                If centralizado Then: c.Offset(0, 2).value = Application.WorksheetFunction.VLookup(c.value, estGAMMA, 10, 0)
            Next c
        End With
    End If
    With wsErrors
        For Each k In errors
            For Each v In errors(k)
                .Cells(counter, 1).value = k
                .Cells(counter, 2).value = v
                counter = counter + 1
            Next v
        Next k
    End With
        
End Sub
