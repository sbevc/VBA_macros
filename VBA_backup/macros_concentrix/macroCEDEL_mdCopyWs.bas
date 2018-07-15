Attribute VB_Name = "mdCopyWs"
Option Explicit

Sub CopyWs()

Dim xReport As Variant, xPend As Variant
Dim wb As Workbook, Awb As Workbook

Set Awb = ThisWorkbook

'--------Selección del pending y Settle-------------
'Con ChDrive y CgDir seleccionamos las carpetas por defecto

ChDrive "U:"
ChDir "U:\Downloads"

xReport = Application.GetOpenFilename( _
    Title:="Seleccionar archivo")
    
If xReport = "" Or xReport = False Then
    MsgBox "No se seleccionó ningún archivo"
    Exit Sub
End If


'---------------------------------------------------

'Abrimos los libros y copiamos los datos
Set wb = Workbooks.Open(FileName:=xReport, Local:=True)
wb.Sheets(1).Range("A1").CurrentRegion.Copy Destination:=Awb.Sheets(1).Range("A1")
wb.Close

    'TextToColumns Data
    Sheets(1).Range("A1").CurrentRegion.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 2 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
        Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
        33, 1)), TrailingMinusNumbers:=True




End Sub



