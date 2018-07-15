VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pISIN As String
Private pPD As Date
Private pRD As Date
Private pACC As String          'acc countrapartida
Private pOwnACC As String       'nuestra cuenta
Private pPayRec As String       'payable/receivable
Private pSubject As String      'asunto del mail
Private pCST As String          'CBL/ECL
Private pCUS As String          'CBL/ECL de nuestra cuenta
Private pNominal As Long
Private pUnitario As Double
Private pName As String         'Nombre de la emisión
Private pTradeDate As Date
Private pSettDate As Date
Private pClDate As Date
Private pDIV As String


'Calculamos el importe a reclamar y hacemos el asunto del mail
Public Sub fillData(crow As Integer)
    
    Dim pImporte As Double
    Dim lastLetter As String
    
    'vemos si es cobro o pago para poner R/P en el asunto del mail
    If pPayRec = "COBRO" Then
        lastLetter = "R"
    Else
        lastLetter = "P"
    End If
    
    
    pImporte = (pNominal / 1000) * pUnitario
    pSubject = _
        "BBVA-" & pISIN & "-PD" & _
        Format(Day(pPD), "00") & Format(Month(pPD), "00") & Format(Year(pPD), "0000") & _
        "-" & pCST & pACC & "-" & lastLetter
        
    Sheets(1).Range("B" & crow).value = pSubject
    Sheets(1).Range("L" & crow).value = pImporte
    
    'ponemos una lista en la columna "A" con el estado, pendiente o finalizado
    Dim MyList(1) As String
    MyList(0) = "Pendiente"
    MyList(1) = "Finalizado"
    
    With Sheets(1).Range("C" & crow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=Join(MyList, ",")
    End With
    
    
End Sub
'Escribimos el mail en un word
Sub mailToWord()

    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    
    
    Select Case pPayRec     'distinguimos la redaccion del mail según sea un cobro o pago de contrapartida
    
        Case Is = "COBRO"
        
        'copiamos las instrucciones como imagen según sea el pago en
        'euros o en dólares
        If pDIV = "EUR" Then
            Sheets("INX").Range("C11:C15").CopyPicture xlScreen, xlPicture
        ElseIf pDIV = "USD" Then
            Sheets("INX").Range("C5:C9").CopyPicture xlScreen, xlPicture
        End If
        
        
        With wdApp
        
            .Visible = True
            .Activate
            .Documents.Add
            
                With .Selection
                    
                    .Font.Name = "Trebuchet MS"
                    .Font.Bold = True
                    
                    .TypeText "Dear all," & vbNewLine & vbNewLine
                    
                    .TypeText "On checking our records, we have found out that there is an outstanding coupon with yourselves." & vbNewLine
                    .TypeText "Could you please check it and let us know if you agree?" & vbNewLine & vbNewLine
                    
                    .TypeText "TRADE DETAILS:" & vbNewLine
                    .TypeText "Name: " & pName & vbNewLine
                    .TypeText "ISIN: " & pISIN & vbNewLine
                    .TypeText "COUNTERPARTY ACCOUNT: " & pCST & " " & pACC & vbNewLine
                    .TypeText "OUR ACCOUNT: " & pCUS & " " & pOwnACC & vbNewLine & vbNewLine
                    .TypeText "Trade Date: " & Format(Day(pTradeDate), "00") & " / " & UCase(Left(MonthName(Month(pTradeDate)), 3)) & " / " & Format(Year(pTradeDate), "0000") & vbNewLine
                    .TypeText "Settlement date: " & Format(Day(pSettDate), "00") & " / " & UCase(Left(MonthName(Month(pSettDate)), 3)) & " / " & Format(Year(pSettDate), "0000") & vbNewLine
                    .TypeText "Clearing date: " & Format(Day(pClDate), "00") & " / " & UCase(Left(MonthName(Month(pClDate)), 3)) & " / " & Format(Year(pClDate), "0000") & vbNewLine
                    .TypeText "Nominal/Quantity: " & Format(pNominal, "#,##0") & vbNewLine & vbNewLine
                    .TypeText "COUPON DETAILS:" & vbNewLine
                    .TypeText "Payment Date: " & Format(Day(pPD), "00") & " / " & UCase(Left(MonthName(Month(pPD)), 3)) & " / " & Format(Year(pPD), "0000") & vbNewLine
                    .TypeText "Record Date: " & Format(Day(pRD), "00") & " / " & UCase(Left(MonthName(Month(pRD)), 3)) & " / " & Format(Year(pRD), "0000") & vbNewLine & vbNewLine
                    .TypeText "AMOUNT RECEIVABLE: " & pDIV & " " & Format((pNominal / 1000) * pUnitario, "#,##0.00") & vbNewLine & vbNewLine
                    .TypeText "If you agree, please pay us to the below SSI:" & vbNewLine
                    
                    .Paste
                    
                    .TypeParagraph
                    .TypeText pSubject
                    
                End With
                
        End With
    
        Case Is = "PAGO"
            
            With wdApp
        
            .Visible = True
            .Activate
            .Documents.Add
            
                With .Selection
                    
                    .Font.Name = "Trebuchet MS"
                    .Font.Bold = True
                    
                    .TypeText "Dear all," & vbNewLine & vbNewLine
                    
                    .TypeText "On checking our records, we have found out that there is an outstanding coupon with yourselves." & vbNewLine
                    .TypeText "Could you please check it and let us know if you agree?" & vbNewLine & vbNewLine
                    
                    .TypeText "TRADE DETAILS:" & vbNewLine
                    .TypeText "Name: " & pName & vbNewLine
                    .TypeText "ISIN: " & pISIN & vbNewLine
                    .TypeText "COUNTERPARTY ACCOUNT: " & pCST & " " & pACC & vbNewLine
                    .TypeText "OUR ACCOUNT: " & pCUS & " " & pOwnACC & vbNewLine & vbNewLine
                    .TypeText "Trade Date: " & Format(Day(pTradeDate), "00") & " / " & UCase(Left(MonthName(Month(pTradeDate)), 3)) & " / " & Format(Year(pTradeDate), "0000") & vbNewLine
                    .TypeText "Settlement date: " & Format(Day(pSettDate), "00") & " / " & UCase(Left(MonthName(Month(pSettDate)), 3)) & " / " & Format(Year(pSettDate), "0000") & vbNewLine
                    .TypeText "Clearing date: " & Format(Day(pClDate), "00") & " / " & UCase(Left(MonthName(Month(pClDate)), 3)) & " / " & Format(Year(pClDate), "0000") & vbNewLine
                    .TypeText "Nominal/Quantity: " & Format(pNominal, "#,##0") & vbNewLine & vbNewLine
                    .TypeText "COUPON DETAILS:" & vbNewLine
                    .TypeText "Payment Date: " & Format(Day(pPD), "00") & " / " & UCase(Left(MonthName(Month(pPD)), 3)) & " / " & Format(Year(pPD), "0000") & vbNewLine
                    .TypeText "Record Date: " & Format(Day(pRD), "00") & " / " & UCase(Left(MonthName(Month(pRD)), 3)) & " / " & Format(Year(pRD), "0000") & vbNewLine & vbNewLine
                    
                    .TypeText "AMOUNT PAYABLE: " & pDIV & " " & Format((pNominal / 1000) * pUnitario, "#,##0.00") & vbNewLine & vbNewLine
                    
                    .TypeText "If you agree, please send us your cash instructions." & vbNewLine & vbNewLine
                    .TypeText "Please do not hesitate to contact us you require further information." & vbNewLine
                    .TypeText "If the above claim is not for your department, please kindly forward it to the concerned department." & vbNewLine
                    .TypeText "Thanks in advance for your cooperation" & vbNewLine & vbNewLine
                    
                    .TypeText "Kind regards,"
                    .TypeParagraph
                    .TypeText pSubject
                    
                End With
                
        End With
            
        
        End Select
    

End Sub




Public Property Get ISIN() As String
    ISIN = pISIN
End Property
Public Property Let ISIN(value As String)
    pISIN = value
End Property
Public Property Get DIV() As String
    DIV = pDIV
End Property
Public Property Let DIV(value As String)
    pDIV = value
End Property
Public Property Get PD() As Date
    PD = pPD
End Property
Public Property Let PD(value As Date)
    pPD = value
End Property
Public Property Get RD() As Date
    RD = pRD
End Property
Public Property Let RD(value As Date)
    pRD = value
End Property
Public Property Get TradeDate() As Date
    TradeDate = pTradeDate
End Property
Public Property Let TradeDate(value As Date)
    pTradeDate = value
End Property
Public Property Get SettDate() As Date
    SettDate = pSettDate
End Property
Public Property Let SettDate(value As Date)
    pSettDate = value
End Property
Public Property Get cldate() As Date
    cldate = pClDate
End Property
Public Property Let cldate(value As Date)
    pClDate = value
End Property
Public Property Get ACC() As String
    ACC = pACC
End Property
Public Property Let ACC(value As String)
    pACC = value
End Property
Public Property Get OwnACC() As String
    OwnACC = pOwnACC
End Property
Public Property Let OwnACC(value As String)
    pOwnACC = value
End Property
Public Property Get PayRec() As String
    PayRec = pPayRec
End Property
Public Property Let PayRec(value As String)
    pPayRec = value
End Property
Public Property Get CST() As String
    CST = pCST
End Property
Public Property Let CST(value As String)
    pCST = value
End Property
Public Property Get CUS() As String
    CUS = pCUS
End Property
Public Property Let CUS(value As String)
    pCUS = value
End Property
Public Property Get Nominal() As Long
    Nominal = pNominal
End Property
Public Property Let Nominal(value As Long)
    pNominal = value
End Property
Public Property Get Unitario() As Double
    Unitario = pUnitario
End Property
Public Property Let Unitario(value As Double)
    pUnitario = value
End Property
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(value As String)
    pName = value
End Property

