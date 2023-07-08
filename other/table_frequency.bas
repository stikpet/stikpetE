Attribute VB_Name = "table_frequency"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function tab_frequency(data As Range, Optional order As Range)
Attribute tab_frequency.VB_Description = "Frequency Table"
Attribute tab_frequency.VB_ProcData.VB_Invoke_Func = " \n14"

    nMissing = 0
    
    'determine how many categories there are and the total sample size (n).
    'cats(i, 1) the label of the category i
    'cats(i, 2) the frequency of category i
    Dim cats As Variant
    Dim cats2 As Variant
    
    nr = data.Rows.Count
    'if no order is provided
    If order Is Nothing Then
        ReDim cats(1 To nr, 1 To 2)
            
        k = 0
        n = 0
        For i = 1 To nr
            If data(i, 1) <> "" Then
                n = n + 1
                newCat = True
                If k <> 0 Then
                    For j = 1 To k
                        If cats(j, 1) = data(i, 1) Then
                            cats(j, 2) = cats(j, 2) + 1
                            newCat = False
                        End If
                    Next j
                End If
                
                If newCat = True Then
                    k = k + 1
                    cats(k, 1) = data(i, 1)
                    cats(k, 2) = 1
                End If
            Else
                nMissing = nMissing + 1
            End If
        Next i
        
        cats2 = cats
        ReDim cats2(0 To k, 1 To 5)
        
        For i = 1 To k
            cats2(i, 1) = cats(i, 1)
            cats2(i, 2) = cats(i, 2)
        Next i
    
    'if order is provided
    Else
        n = 0
        k = order.Rows.Count
        ReDim cats2(0 To k, 1 To 5)
        For i = 1 To k
            cats2(i, 1) = order(i, 1)
            cats2(i, 2) = WorksheetFunction.CountIf(data, order(i, 1))
            n = n + cats2(i, 2)
        Next i
        
        nMissing = nr - n
        
    End If
        
    cats2(0, 1) = "category"
    cats2(0, 2) = "frequency"
    
    For i = 1 To k
        cats2(i, 3) = cats2(i, 2) / (n + nMissing) * 100
        cats2(i, 4) = cats2(i, 2) / n * 100
        cats2(i, 5) = cats2(i, 4) + cats2(i - 1, 5)
    Next i
    
    cats2(0, 3) = "percent"
    cats2(0, 4) = "valid percent"
    cats2(0, 5) = "cumulative percent"
    
    tab_frequency = cats2

End Function
