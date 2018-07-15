Attribute VB_Name = "mdmain"
Option Explicit
Sub main()

    CopyWs
    ProcessData
    MySortMacro
    
    Dim c As Range, payRng As Range
    Dim payment As clsPayment
    Dim counter As Integer
    
    Set payRng = Sheets(1).Range(Range("A2"), Range("A2").End(xlDown))
    
    For Each c In payRng
    
        Set payment = New clsPayment
        
        counter = counter + 1
        
        payment.PayType = c.value
        payment.ISIN = c.Offset(0, 1).value
        payment.Name = c.Offset(0, 2).value
        
        payment.createCover counter
        
    Next

End Sub


Sub CopyWs()

Dim PendingPath As String
Dim PendingReportName As String
Dim xMonth As String, xPend As String
Dim wb As Workbook, Awb As Workbook
Dim xDate As String
Dim xDay As String

Set Awb = ThisWorkbook

PendingPath = "\\Datdpto1_01.ad.bbva.com\datos_datdpto1_01\SC001177\Bony_BBV\INFORMES_INFORM\Operaciones_financieras\Div_pdte\"

'Elegimos las primeras 3 letras del mes actual
xMonth = Choose(Month(Date), "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
xDay = Format(Now, "dd")

PendingReportName = "Pending Income csv_" & xDay & " " & xMonth & " " & Year(Date)

xPend = PendingPath & PendingReportName & ".csv"

'Abrimos el pending (Xpend) con local = true para que quede bien el archivo .csv
Set wb = Workbooks.Open(Filename:=xPend, Local:=True)
wb.Sheets(1).Copy after:=Awb.Sheets(Awb.Sheets.Count)
wb.Close


Sheets(Sheets.Count).Range("A2").CurrentRegion.Copy Destination:=Sheets(1).Range("A1")

Application.DisplayAlerts = False
Sheets(Sheets.Count).Delete
Application.DisplayAlerts = True


End Sub

'ponemos el texto en columnas
'dejamos solo el tipo de operación, ISIN, nombre y PD
Sub ProcessData()

    'TextoColumns
    Sheets(1).Activate
    Sheets(1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
        Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
        33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), _
        Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array( _
        46, 1), Array(47, 1), Array(48, 1), Array(49, 1), Array(50, 1), Array(51, 1), Array(52, 1), _
        Array(53, 1), Array(54, 1), Array(55, 1), Array(56, 1), Array(57, 1), Array(58, 1), Array( _
        59, 1), Array(60, 1), Array(61, 1), Array(62, 1), Array(63, 1), Array(64, 1), Array(65, 1), _
        Array(66, 1), Array(67, 1), Array(68, 1), Array(69, 1), Array(70, 1), Array(71, 1), Array( _
        72, 1), Array(73, 1), Array(74, 1), Array(75, 1), Array(76, 1), Array(77, 1)), _
        TrailingMinusNumbers:=True
        
     'Eliminar columnas
    Range("A:H,J:AU,AX:BY").Select
    Selection.Delete
     'Eliminar duplicados
    Sheets(1).Range("A1").CurrentRegion.Select
    Selection.RemoveDuplicates Columns:=2, Header:=xlNo
        
End Sub


Sub MySortMacro()

    Dim LastRow As Long
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("A2:C" & LastRow).Sort Key1:=Range("B3:B" & LastRow), _
       Order1:=xlAscending, Header:=xlNo

End Sub



