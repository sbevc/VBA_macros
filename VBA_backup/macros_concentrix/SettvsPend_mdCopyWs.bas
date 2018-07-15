Attribute VB_Name = "mdCopyWs"
Option Explicit

Sub CopyWs()

Dim xSett As Variant, xPend As Variant
Dim wb As Workbook, Awb As Workbook

Set Awb = ThisWorkbook

'--------Selección del pending y Settle-------------
'Con ChDrive y CgDir seleccionamos las carpetas por defecto

ChDrive "H:"
ChDir "H:\SC001177\Bony_BBV\INFORMES_INFORM\Operaciones_financieras\Div_liquidados"

xSett = Application.GetOpenFilename( _
    FileFilter:="Archivos csv(*.csv),*.csv", _
    Title:="Seleccionar Settle")
    
If xSett = "" Or xSett = False Then
    MsgBox "No se seleccionó ningún archivo"
    Exit Sub
End If


ChDir "H:\SC001177\Bony_BBV\INFORMES_INFORM\Operaciones_financieras\Div_pdte"

xPend = Application.GetOpenFilename( _
    FileFilter:="Archivos csv(*.csv),*.csv", _
    Title:="Seleccionar Pending")

If xPend = "" Or xSett = False Then
    MsgBox "No se seleccionó ningún archivo"
    Exit Sub
End If
'---------------------------------------------------

'Abrimos los libros y copiamos los datos
Set wb = Workbooks.Open(Filename:=xSett, Local:=True)
wb.Sheets(1).Copy After:=Awb.Sheets(1)
wb.Close

Set wb = Workbooks.Open(Filename:=xPend, Local:=True)
wb.Sheets(1).Copy After:=Awb.Sheets(1)
wb.Close

    'TextToColumns Pending
    Sheets(2).Range("A1").CurrentRegion.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
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

End Sub



