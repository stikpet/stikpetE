Attribute VB_Name = "eff_size_cohen_h2"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub es_cohen_h2_os_addHelp()
Application.MacroOptions _
    Macro:="es_cohen_h2", _
    Description:="Cohen's h2 for one-sample tests", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "range with data", _
        "optional range with the two codes to count", _
        "expected proportion (default 0.5)")
        
End Sub

Function es_cohen_h2_os(data As Range, Optional codes As Range, Optional p0 = 0.5)
'Function that determines Cohen's h for a one-sample.
'Input a list of scores, optional the two categories, the expected proportion from the null hypothesis, and the output

If codes Is Nothing Then

    k = 0
    nt = data.Rows.Count
    
    k1 = data.Cells(1, 1)
    i = 2
    If k1 = "" Then
        Do While k1 = ""
            k1 = data.Cells(i, 1)
            i = i + 1
        Loop
    End If
    
    k2 = data.Cells(i, 1)
    If k2 = "" Or k2 = k1 Then
        i = i + 1
        Do While k2 = "" Or k2 = k1
            k2 = data.Cells(i, 1)
            i = i + 1
        Loop
    End If

Else
    k1 = codes.Cells(1, 1)
    k2 = codes.Cells(2, 1)
End If
    
    
n1 = WorksheetFunction.CountIf(data, k1)
n2 = WorksheetFunction.CountIf(data, k2)
n = n1 + n2


p1 = n1 / n

phi1 = 2 * WorksheetFunction.Asin(p1 ^ 0.5)
phic = 2 * WorksheetFunction.Asin(p0 ^ 0.5)

h2 = phi1 - phic

es_cohen_h2_os = h2

End Function

