Attribute VB_Name = "help_quartileIndexing"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function he_quartileIndexing(data, Optional method = "sas1")

    n = UBound(data, 1) - LBound(data, 1) + 1
    
    If method = "inclusive" Then
        If (n Mod 2) = 0 Then
            q1Index = (n + 2) / 4
            q3Index = (3 * n + 1) / 4
        Else
            q1Index = (n + 3) / 4
            q3Index = (3 * n + 1) / 4
        End If
    ElseIf method = "exclusive" Then
        If (n Mod 2) = 0 Then
            q1Index = (n + 2) / 4
            q3Index = (3 * n + 1) / 4
        Else
            q1Index = (n + 1) / 4
            q3Index = (3 * n + 3) / 4
        End If
            
    ElseIf method = "sas1" Then
        q1Index = n * 1 / 4
        q3Index = n * 3 / 4
    ElseIf method = "sas4" Then
        q1Index = (n + 1) * 1 / 4
        q3Index = (n + 1) * 3 / 4
    ElseIf method = "hl" Then
        q1Index = n * 1 / 4 + 1 / 2
        q3Index = n * 3 / 4 + 1 / 2
    ElseIf method = "excel" Then
        q1Index = (n - 1) * 1 / 4 + 1
        q3Index = (n - 1) * 3 / 4 + 1
    ElseIf method = "hf8" Then
        q1Index = (n + 1 / 3) * 1 / 4 + 1 / 3
        q3Index = (n + 1 / 3) * 3 / 4 + 1 / 3
    ElseIf method = "hf9" Then
        q1Index = (n + 1 / 4) * 1 / 4 + 3 / 8
        q3Index = (n + 1 / 4) * 3 / 4 + 3 / 8
    End If
    
    Dim results(0 To 1) As Double
    results(0) = q1Index
    results(1) = q3Index
    
    he_quartileIndexing = results

End Function
