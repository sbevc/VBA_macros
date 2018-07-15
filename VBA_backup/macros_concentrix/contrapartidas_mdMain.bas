Attribute VB_Name = "mdMain"
Option Explicit

Sub main()

    Dim claim As clsClaim
    Dim c As Range
    'Dim coll As New Collection
    
    For Each c In Selection
    
        Set claim = New clsClaim
        
        claim.ISIN = Range("H" & c.Row).value 'Cells(c.Row, 2).value
        claim.Name = Range("I" & c.Row).value   'Cells(c.Row, 3).value
        claim.PayRec = Range("D" & c.Row).value 'Cells(c.Row, 4).value
        claim.OwnACC = Range("G" & c.Row).value
        claim.CUS = Range("F" & c.Row).value
        claim.CST = Range("S" & c.Row).value    'Cells(c.Row, 5).value
        claim.ACC = Range("T" & c.Row).value    'Cells(c.Row, 6).value
        claim.Nominal = Range("J" & c.Row).value    'Cells(c.Row, 7).value
        claim.Unitario = Range("K" & c.Row).value   'Cells(c.Row, 8).value
        claim.DIV = Range("M" & c.Row).value    'Cells(c.Row, 10).value
        claim.PD = Range("N" & c.Row).value 'Cells(c.Row, 11).value
        claim.RD = Range("O" & c.Row).value
        claim.TradeDate = Range("P" & c.Row).value 'Cells(c.Row, 13).value
        claim.SettDate = Range("Q" & c.Row).value   'Cells(c.Row, 14).value
        claim.cldate = Range("R" & c.Row).value 'Cells(c.Row, 15).value
        
        
        claim.fillData c.Row
        
        claim.mailToWord
        
    Next c
    
End Sub

