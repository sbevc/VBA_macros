Attribute VB_Name = "mdGAMMA_Listing"
'Agregamos el replenishment type(RP para almacen y ND para tienda)
'y source of supply (1 para almacen y 2 para tienda) en la pestaña de listing
Sub wsListing()
    
    Dim regex As New RegExp
    
    Dim lrow As Integer
    lrow = Sheets("Listing").Range("A" & Rows.Count).End(xlUp).Row
    
    If lrow = 2 Then
        MsgBox "Falta agregar los items"
        Exit Sub
    End If
    
    Dim myRange As Range: Set myRange = Range("B3:B" & lrow)
    Dim c As Range
        
    With regex
        .Global = True
        .IgnoreCase = True
    End With
    
    For Each c In myRange
        regex.Pattern = "[A-Z]{4}"  'Tiendas
        If regex.Test(c.value) Then
            Cells(c.Row, 16).value = "ND"
            Cells(c.Row, 21).value = 2
        Else
            regex.Pattern = "([A-Z]{2})([0-9]{2})"  'Almacenes
            If regex.Test(c.value) Then
                Cells(c.Row, 16).value = "RP"
                Cells(c.Row, 21).value = 1
            Else
                x = x + 1
                With c.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next c
    
    If x > 0 Then: MsgBox "Verificar items resaltados!"
    
End Sub

