Attribute VB_Name = "mdSaveAs"

Sub saveWb()

    Const path As String = "U:\"
    
    Dim fechaLiquidaci�n As Date
    Dim wbName As String
    Dim strDay As String, strMonth As String
    Dim aWb As Workbook
    
    Set aWb = ThisWorkbook
    
    
    If Len(Dir(path)) = 0 Then
        MsgBox "No se encontr� el disco U, guardar el archivo manualmente"
    Else
        fechaLiquidaci�n = Range("L5").Value
        strDay = Format(Day(fechaLiquidaci�n), "00")
        strMonth = Format(Month(fechaLiquidaci�n), "00")
        wbName = "pagos " & strDay & "." & strMonth & ".xls"
        
        Sheets(1).Copy
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs fileName:=path & wbName, FileFormat:=xlWorkbookNormal
        Application.DisplayAlerts = True
        
        aWb.Close savechanges:=False
        
    End If


End Sub

