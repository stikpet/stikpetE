Attribute VB_Name = "eff_size_alt_ratio"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function es_alt_ratio(data As Range, Optional codes As Range, Optional p0 = 0.5, Optional category, Optional output = "all")
Attribute es_alt_ratio.VB_Description = "Alternative Ratio"
Attribute es_alt_ratio.VB_ProcData.VB_Invoke_Func = " \n14"
'Function that determines the Alternative Ratio a.k.a. Relative Risk
'Input a list of scores, the codes of two categories to compare, and the expected proportion from the null hypothesis
'Optional input the category of which to calculate the AR from


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

If Not IsMissing(category) Then
    If k2 = category Then
        k3 = k1
        k1 = k2
        k2 = k3
    End If
End If

n1 = WorksheetFunction.CountIf(data, k1)
n2 = WorksheetFunction.CountIf(data, k2)
n = n1 + n2

p1 = n1 / n

AR1 = p1 / p0

'OUTPUT
If output = "all" Then
    p2 = n2 / n
    AR2 = p2 / (1 - p0)
    Dim res(1 To 2, 1 To 2)
    res(1, 1) = "Alt. Ratio Cat. 1"
    res(1, 2) = "Alt. Ratio Cat. 2"
    res(2, 1) = AR1
    res(2, 2) = AR2
                
    es_alt_ratio = res
    
Else
    es_alt_ratio = AR1
End If


End Function
