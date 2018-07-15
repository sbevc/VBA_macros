Attribute VB_Name = "mdCopyWs"
Option Explicit

Sub CopyWs()

Dim xTxt As Variant
Dim wb As Workbook, Awb As Workbook

Set Awb = ThisWorkbook

'--------Selección del pending y Settle-------------
'Con ChDrive y CgDir seleccionamos las carpetas por defecto

    ChDrive "H:"
    ChDir "H:\TRANSMI\CR26G094\OPERACIONES_FINANCIERAS\DESGLOSES"
    
    xTxt = Application.GetOpenFilename( _
        FileFilter:="Archivos de texto(*.txt),*.txt", _
        Title:="Seleccionar TXT")
        
    If xTxt = "" Or xTxt = False Then
        MsgBox "No se seleccionó ningún archivo"
        Exit Sub
    End If

'---------------------------------------------------

'Abrimos los libros y copiamos los datos

    Workbooks.OpenText FileName:=xTxt _
        , Origin:=932, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=True, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), _
        Array(22, 1), Array(23, 1), Array(24, 1)), TrailingMinusNumbers:=True
        
    Set wb = Workbooks(Workbooks.Count)
    wb.Sheets(1).Range("A1").CurrentRegion.Copy Destination:=Awb.Sheets("Datos").Range("A1")
    wb.Close


End Sub



