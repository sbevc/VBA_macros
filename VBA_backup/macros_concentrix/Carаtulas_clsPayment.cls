VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pType As String
Private pISIN As String
Private pName As String


Public Sub createCover(counter As Integer)
    
    Dim ISINcell As String, Typecell As String, NameCell As String
    ISINcell = "B13"
    Typecell = "B7"
    NameCell = "B19"
    
    'Distinguimos entre si es la primer carátula (que solo pegamos los datos) o
    'si ya hay alguna carátula creada. Para distinguir sumamos 1 al counter
    If counter = 1 Then
        With Sheets(2)
            .Name = pISIN
            .Range(Typecell).value = pType
            .Range(ISINcell).value = pISIN
            .Range(NameCell).value = pName
        End With
    Else
        Sheets(2).Copy after:=Sheets(Sheets.Count)
        With ActiveSheet
             .Name = pISIN
            .Range(Typecell).value = pType
            .Range(ISINcell).value = pISIN
            .Range(NameCell).value = pName
        End With
    End If
    
End Sub




Public Property Get PayType() As String
    PayType = pType
End Property
Public Property Let PayType(value As String)
    pType = value
End Property
Public Property Get ISIN() As String
    ISIN = pISIN
End Property
Public Property Let ISIN(value As String)
    pISIN = value
End Property
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(value As String)
    pName = value
End Property
