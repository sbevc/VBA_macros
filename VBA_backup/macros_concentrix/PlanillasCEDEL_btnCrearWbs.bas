Attribute VB_Name = "btnCrearWbs"
Option Explicit

Sub Splitbook()

'Se guarda como .xlsx

Dim xPath As String, Wsname As String, ISIN As String, ACC As String, PD, i As Integer
Dim ISINCell As Range, ACCCell As Range, PDCell As Range, Emisión As String

    xPath = "\\Datdpto1_01.ad.bbva.com\datos_datdpto1_01\SC000068\OPERACIONES FINANCIERAS\INTERESES Y DIVIDENDOS\IMPUESTOS\INTERNACIONAL\Clearstream\CEDEL\EVENTOS CED 2017\Nuevas Planillas\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        For i = 1 To Sheets.Count
        
            Emisión = Left(Sheets(i).Name, 2)
            
            If Right(Sheets(i).Name, 8) = "Planilla" Then
                    
                    Select Case Emisión
                    
                        Case Is = "XS"
                        
                            Set ISINCell = Sheets(i).Range("B3")
                            Set ACCCell = Sheets(i).Range("A3")
                            Set PDCell = Sheets(i).Range("F3")
                            ISIN = ISINCell.Value
                            ACC = ACCCell.Value
                            PD = PDCell.Value
                            
                            Wsname = ISIN & " " & Day(PD) & "-" & Month(PD) & "-" & Year(PD) & " " & ACC
                            
                            Sheets(i).Copy
                            Application.ActiveWorkbook.SaveAs FileName:=xPath & Wsname & ".xls"
                            Application.ActiveWorkbook.Close False
                        
                        Case Is = "ES"
          
                            Set ISINCell = Sheets(i).Range("B2")
                            Set ACCCell = Sheets(i).Range("B1")
                            Set PDCell = Sheets(i).Range("B4")
                            ISIN = ISINCell.Value
                            ACC = ACCCell.Value
                            PD = PDCell.Value
                            
                            Wsname = ISIN & " " & Day(PD) & "-" & Month(PD) & "-" & Year(PD) & " " & ACC

                            
                            Sheets(i).Copy
                            Application.ActiveWorkbook.SaveAs FileName:=xPath & Wsname & ".xls"
                            Application.ActiveWorkbook.Close False
                            
                        Case Is = "PT"
                        
                            Set ISINCell = Sheets(i).Range("B8")
                            Set ACCCell = Sheets(i).Range("B10")
                            Set PDCell = Sheets(i).Range("B7")
                            ISIN = ISINCell.Value
                            ACC = ACCCell.Value
                            PD = PDCell.Value
                            
                            Wsname = ISIN & " " & Day(PD) & "-" & Month(PD) & "-" & Year(PD) & " " & ACC
                            
                            Sheets(i).Copy
                            Application.ActiveWorkbook.SaveAs FileName:=xPath & Wsname & ".xls"
                            Application.ActiveWorkbook.Close False
                            
                    End Select
                
            End If
            
        Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
