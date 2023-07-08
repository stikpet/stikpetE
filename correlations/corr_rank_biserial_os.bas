Attribute VB_Name = "corr_rank_biserial_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function r_rank_biserial_os(data As Range, _
                            Optional levels As Range, _
                            Optional mu = "none", _
                            Optional output = "all")
Attribute r_rank_biserial_os.VB_Description = "Rank biserial correlation coefficient (one-sample)"
Attribute r_rank_biserial_os.VB_ProcData.VB_Invoke_Func = " \n14"
                    
    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    If mu = "none" Then
        mu = (WorksheetFunction.Min(dataN) + WorksheetFunction.Max(dataN)) / 2
    End If
    
    Dim n, nr As Integer
    n = UBound(dataN, 1) - LBound(dataN, 1) + 1
    
    Dim i, k As Integer
    nEqMed = 0
    For i = 0 To n - 1
        If dataN(i) = mu Then
            nEqMed = nEqMed + 1
        End If
    Next i
    
    nr = n - nEqMed
    
    Dim absDiffs() As Double
    ReDim absDiffs(0 To nr - 1)
    Dim scores() As Double
    ReDim scores(0 To nr - 1)
    k = 0
    i = 1
    Do While k < nr
        If dataN(i - 1) <> mu Then
            absDiffs(k) = Abs(dataN(i - 1) - mu)
            scores(k) = dataN(i - 1)
            k = k + 1
        End If
        i = i + 1
    Loop
        
    'sort scores based on absolute differences
    changes = 1
    Do While changes <> 0
        changes = 0
        For i = 1 To nr - 1
            If absDiffs(i - 1) > absDiffs(i) Then
                ff1 = absDiffs(i)
                ff2 = scores(i)
                absDiffs(i) = absDiffs(i - 1)
                scores(i) = scores(i - 1)
                absDiffs(i - 1) = ff1
                scores(i - 1) = ff2
                
                changes = 1
            End If
        Next i
    Loop
    
    'we need the ranks for which we need the rank frequencies
    'store for each score how often it occurs
    Dim Rfreq As Variant
    ReDim Rfreq(1 To nr, 1 To 3)
    For i = 0 To nr - 1
        For j = 0 To nr - 1
            If absDiffs(j) = absDiffs(i) Then
                freq = freq + 1
            End If
        Next j
        
        Rfreq(i + 1, 1) = absDiffs(i)
        Rfreq(i + 1, 2) = freq
        freq = 0
    Next i
    
    'now for the ranks and sum of ranks
    nD0 = 0
    Rsum = 0
    RsumN = 0
    Rd0 = 0 'for sum of ranks of differences of 0
    ReDim r(1 To nr, 1 To 3)
    r(1, 1) = Rfreq(1, 1)
    r(1, 2) = Rfreq(1, 2)
    
    If Rfreq(1, 2) = 1 Then
        r(1, 3) = 1
        Else
        r(1, 3) = (1 + 1 + Rfreq(1, 2) - 1) / 2
    End If
    
    If scores(0) > mu Then
        Rsum = Rsum + r(1, 3)
    ElseIf scores(0) = mu Then
        nD0 = nD0 + 1
        Rd0 = Rd0 + r(1, 3)
    Else
        RsumN = RsumN + r(1, 3)
    End If
    
    
    For i = 2 To nr
        r(i, 1) = Rfreq(i, 1)
        r(i, 2) = Rfreq(i, 2)
        
        If Rfreq(i, 2) = 1 Then
            r(i, 3) = i
        ElseIf Rfreq(i, 1) <> Rfreq(i - 1, 1) Then
            r(i, 3) = (i + i + Rfreq(i, 2) - 1) / 2
        Else
            r(i, 3) = r(i - 1, 3)
        End If
            
        If scores(i - 1) > mu Then
            Rsum = Rsum + r(i, 3)
        ElseIf scores(i - 1) = mu Then
            nD0 = nD0 + 1
            Rd0 = Rd0 + r(i, 3)
        Else
            RsumN = RsumN + r(i, 3)
        End If
        
    Next i
    
    rb = Abs(Rsum - RsumN) / (Rsum + RsumN)
    
    If output = "mu" Then
        r_rank_biserial_os = mu
    ElseIf output = "value" Then
        r_rank_biserial_os = rb
    Else
        'Results
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "mu"
        res(1, 2) = "rb"
        res(2, 1) = mu
        res(2, 2) = rb
        
        r_rank_biserial_os = res
    End If

End Function

