Attribute VB_Name = "mdGuardado"
Sub Guardado()
Attribute Guardado.VB_ProcData.VB_Invoke_Func = " \n14"

Dim xName As String, MyArr(1 To 12, 1 To 2) As Variant, xFixedPath As String, xVarPath As String

MyArr(1, 1) = 1
MyArr(2, 1) = 2
MyArr(3, 1) = 3
MyArr(4, 1) = 4
MyArr(5, 1) = 5
MyArr(6, 1) = 6
MyArr(7, 1) = 7
MyArr(8, 1) = 8
MyArr(9, 1) = 9
MyArr(10, 1) = 10
MyArr(11, 1) = 11
MyArr(12, 1) = 12

MyArr(1, 2) = "01 - ENERO"
MyArr(2, 2) = "02 - FEBRERO"
MyArr(3, 2) = "03 - MARZO"
MyArr(4, 2) = "04 - ABRIL"
MyArr(5, 2) = "05 - MAYO"
MyArr(6, 2) = "06 - JUNIO"
MyArr(7, 2) = "07 - JULIO"
MyArr(8, 2) = "08 - AGOSTO"
MyArr(9, 2) = "09 - SEPTIEMBRE"
MyArr(10, 2) = "10 - OCTUBRE"
MyArr(11, 2) = "11 - NOVIEMBRE"
MyArr(12, 2) = "12 - DICIEMBRE"


xName = Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".xlsx"

xVarPath = Application.WorksheetFunction.VLookup(Month(Date), MyArr, 2, 0) & "\" 'Elije la carpeta del mes segùn el mes actual

xFixedPath = "U:\MACROS\MT564\"
'xFixedPath = "R:\COMUN\controlador_5\Operaciones Financieras\Listados MT564\Año 2017\"

Application.DisplayAlerts = False
Workbooks("CargaMT564").SaveAs Filename:=(xFixedPath & xName), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = True

Workbooks(xName).Close

End Sub
