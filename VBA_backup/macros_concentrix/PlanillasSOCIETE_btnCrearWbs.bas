Attribute VB_Name = "btnCrearWbs"
Option Explicit

Sub Splitbook()

'Se guarda como .xlsx

Dim xPath As String, Wsname As String, ISIN As String, i As Integer, ColAcc As Integer, RowAcc As Integer
Dim ACCNumberCell As Range, ACCDescCell As Range, Emisión As String

    xPath = "H:\SC000068\OPERACIONES FINANCIERAS\INTERESES Y DIVIDENDOS\IMPUESTOS\INTERNACIONAL\SOCIETE\2017\Planillas MACRO\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        For i = 2 To Sheets.Count
        
            Sheets(i).Activate
            
            ColAcc = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).End(xlToRight).Column + 1
            RowAcc = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
            
            Set ACCNumberCell = ActiveSheet.Cells(RowAcc, ColAcc)
            Set ACCDescCell = ActiveSheet.Cells(RowAcc, ColAcc + 1)
            

            ISIN = Left(ActiveSheet.Name, 12)

            Wsname = ISIN & " " & ACCNumberCell.Value & " " & ACCDescCell.Value
            ActiveSheet.Copy
            Application.ActiveWorkbook.SaveAs FileName:=xPath & Wsname & ".xlsx"
            Application.ActiveWorkbook.Close False
                        

        Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
