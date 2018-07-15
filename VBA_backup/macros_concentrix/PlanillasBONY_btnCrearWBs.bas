Attribute VB_Name = "btnCrearWbs"
Option Explicit

Sub Splitbook()

'Se guarda como .xlsx

Dim xPath As String, Wsname As String, ISIN As String, ACC As String, PD As String, i As Integer
Dim ISINCell As Range, ACCCell As Range, PDCell As Range, Emisión As String, PositionCell As Range, Position As String

    xPath = "\\Datdpto1_01.ad.bbva.com\datos_datdpto1_01\SC000068\OPERACIONES FINANCIERAS\INTERESES Y DIVIDENDOS\IMPUESTOS\INTERNACIONAL\BONY\EVENTOS\EVENTOS 2017\Planillas Nuevas\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        For i = 2 To Sheets.Count
        
            Sheets(i).Activate
            Emisión = Left(ActiveSheet.Name, 2)
            
            Select Case Emisión
                
                Case Is = "FR", "IT", "IE", "JP", "FI", "SE", "NO", "PT"
                
                    Set ISINCell = ActiveSheet.Range("B6")
                    Set ACCCell = ActiveSheet.Range("E6")
                    Set PDCell = ActiveSheet.Range("C6")
                    Set PositionCell = ActiveSheet.Cells(Rows.Count, 7).End(xlUp)
                    ISIN = ISINCell.Value
                    ACC = ACCCell.Value
                    PD = Day(PDCell) & "." & Month(PDCell) & "." & Year(PDCell)
                    Position = PositionCell.Value
                    Wsname = ISIN & " " & ACC & " " & PD & " (" & Position & ")"
                    
                    ActiveSheet.Copy
                    ActiveSheet.Name = "Sheet1"
                    
                    Application.ActiveWorkbook.SaveAs FileName:=xPath & Wsname & ".xlsx"
                    Application.ActiveWorkbook.Close False
                        
                Case Is = "BO"
                    
                    ISIN = Mid(ActiveSheet.Name, 4, 12)
                    ACC = Right(ActiveSheet.Name, 6)
                    Wsname = "DOOR BO Setup Form - " & ISIN & " " & ACC
                    ActiveSheet.Copy
                    Application.ActiveWorkbook.SaveAs FileName:=xPath & Wsname & ".xlsx"
                    Application.ActiveWorkbook.Close False
            
            End Select
            
            
        Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
