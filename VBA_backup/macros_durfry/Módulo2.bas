Attribute VB_Name = "Módulo2"

'limpia los RE, RV, FV, etc de los subject
Sub cleanSubjet(rng As Range)
    
    Dim c As Range
    
    Set RE = CreateObject("vbscript.regexp")
    RE.Global = True: RE.IgnoreCase = True
    
    'Matchea al inicio (^) 2 o 3 letras seguidos por un espacio opcional (\s), : y otro espacio opcional
    RE.Pattern = "^\w{2,3}\s?:\s?"
    
    For Each c In rng
        Set allMatches = RE.Execute(c.value)
        If allMatches.Count <> 0 Then: c.value = Replace(c.value, allMatches.item(0).value, "")
    Next c
    

End Sub


'arregla las fechas de los mails al copiarlos desde outlook
Sub mail()

    Dim c As Range
    
    Set RE = CreateObject("vbscript.regexp")
    RE.Global = True: RE.IgnoreCase = True
    
    
    
    For Each c In Selection
        If TypeName(c.value) = "String" Then
            RE.Pattern = "^\d{1,2}:\d{2}\s(a.m.|p.m.)"
            Set allMatches = RE.Execute(c.value)
            If allMatches.Count <> 0 Then
                c.value = Date
            Else
                RE.Pattern = "\w{5,9}\s\d{1,2}:\d{2}\s(a.m.|p.m.)"
                Set allMatches = RE.Execute(c.value)
                If allMatches.Count <> 0 Then
                    c.value = get_date(Left(c.value, 3))
                Else
                    RE.Pattern = "\d{1,2}/\d{2}"
                    Set allMatches = RE.Execute(c.value)
                    c.value = allMatches.item(0)
                End If
            End If
        End If
    Next c


End Sub

Function get_date(day_) As Date

    Dim days As New Scripting.Dictionary, days_interval As Integer, m As Integer, d As Integer
    
    days.Add Key:="dom", item:=1
    days.Add Key:="lun", item:=2
    days.Add Key:="mar", item:=3
    days.Add Key:="mié", item:=4
    days.Add Key:="jue", item:=5
    days.Add Key:="vie", item:=6
    days.Add Key:="sab", item:=7
    
    
    days_interval = days(day_) - Weekday(Date)
    
    y = Year(Date)
    m = Month(DateAdd("d", days_interval, Date))
    d = Day(DateAdd("d", days_interval, Date))
    
    get_date = DateSerial(y, m, d)
    

End Function

Sub asdlkjsdfls()

    Dim c As Range
    
    For Each c In Selection
        Debug.Print TypeName(c.value)
    Next c
    
End Sub
