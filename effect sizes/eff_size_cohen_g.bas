Attribute VB_Name = "eff_size_cohen_g"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function es_cohen_g(data As Range, Optional codes As Range)
Attribute es_cohen_g.VB_Description = "Cohen's g"
Attribute es_cohen_g.VB_ProcData.VB_Invoke_Func = " \n14"
'Function that determines Cohen's g.
'Note that this effect size is only used if the expected proportion is 0.5
'Input a list of scores, optional the two categories and output desired

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
    
    g = p1 - 0.5
    
    es_cohen_g = g

End Function
