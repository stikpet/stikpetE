Attribute VB_Name = "thumb_pearson_r"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function th_pearson_r(r, Optional qual = "bartz", Optional output = "all")
    
    'Rafter et al. (2003, p. 194).
    If qual = "rafter" Then
        ref = "Rafter et al. (2003, p. 194)"
        If Abs(r) < 0.25 Then
            qual = "weak"
        ElseIf Abs(r) < 0.75 Then
            qual = "moderate"
        Else
            qual = "strong"
        End If
    
    'Cohen (1988, p. 82).
    ElseIf qual = "cohen" Then
        ref = "Cohen (1988, p. 82)"
        If Abs(r) < 0.1 Then
            qual = "negligable"
        ElseIf Abs(r) < 0.3 Then
            qual = "small"
        ElseIf Abs(r) < 0.5 Then
            qual = "medium"
        Else
            qual = "large"
        End If
    
    'Rumsey (2011, p. 284).
    ElseIf qual = "rumsey" Then
        ref = "Rumsey (2011, p. 284)"
        If Abs(r) < 0.3 Then
        qual = "negligable"
        ElseIf Abs(r) < 0.5 Then
        qual = "weak"
        ElseIf Abs(r) < 0.7 Then
        qual = "moderate"
        Else
        qual = "strong"
        End If
    
    'Gignac and Szodorai (2016, p. 75); Hemphill (2003, p. 78)
    ElseIf qual = "gignac" Or qual = "hemphill" Then
        ref = "Gignac and Szodorai (2016, p. 75); Hemphill (2003, p. 78)"
        If Abs(r) < 0.1 Then
        qual = "negligable"
        ElseIf Abs(r) < 0.2 Then
        qual = "small"
        ElseIf Abs(r) < 0.3 Then
        qual = "medium"
        Else
        qual = "large"
        End If
    
    'Lovakov and Agadullina (2021, p. 514).
    ElseIf qual = "lovakov" Then
        ref = "Lovakov and Agadullina (2021, p. 514)"
        If Abs(r) < 0.12 Then
            qual = "negligable"
        ElseIf Abs(r) < 0.24 Then
            qual = "small"
        ElseIf Abs(r) < 0.41 Then
            qual = "medium"
        Else
            qual = "large"
        End If
    
    'Rosenthal (1996, p. 45).
    ElseIf qual = "rosenthal" Then
        ref = "Rosenthal (1996, p. 45)"
        If Abs(r) < 0.1 Then
            qual = "negligable"
        ElseIf Abs(r) < 0.3 Then
            qual = "small"
        ElseIf Abs(r) < 0.5 Then
            qual = "medium"
        ElseIf Abs(r) < 0.7 Then
            qual = "large"
        Else
            qual = "very large"
        End If
    
    'Agnes (2011).
    ElseIf qual = "agnes" Then
        ref = "Agnes (2011)"
        If Abs(r) < 0.2 Then
            qual = "negligable"
        ElseIf Abs(r) < 0.4 Then
            qual = "low"
        ElseIf Abs(r) < 0.6 Then
            qual = "moderate"
        ElseIf Abs(r) < 0.8 Then
            qual = "marked"
        Else
            qual = "high"
        End If
    
    'Bartz (1999, p. 184, as cited in Warmbrod 2001).
    ElseIf qual = "bartz" Then
        ref = "Bartz (1999, p. 184, as cited in Warmbrod 2001)"
        If Abs(r) < 0.2 Then
        qual = "very low"
        ElseIf Abs(r) < 0.4 Then
        qual = "low"
        ElseIf Abs(r) < 0.6 Then
        qual = "moderate"
        ElseIf Abs(r) < 0.8 Then
        qual = "strong"
        Else
        qual = "very high"
        End If
    
    'Disha (2016).
    ElseIf qual = "disha" Then
        ref = "Disha (2016)"
        If Abs(r) < 0.1 Then
        qual = "negligable"
        ElseIf Abs(r) < 0.3 Then
        qual = "very low"
        ElseIf Abs(r) < 0.5 Then
        qual = "low"
        ElseIf Abs(r) < 0.7 Then
        qual = "moderate"
        ElseIf Abs(r) < 0.9 Then
        qual = "high"
        Else
        qual = "very high"
        End If
    
    'Hopkins (1997, as cited in Warmbrod 2001).
    ElseIf qual = "hopkins" Then
        ref = "Hopkins (1997, as cited in Warmbrod 2001)"
        If Abs(r) < 0.1 Then
        qual = "trivial"
        ElseIf Abs(r) < 0.3 Then
        qual = "low"
        ElseIf Abs(r) < 0.5 Then
        qual = "moderate"
        ElseIf Abs(r) < 0.7 Then
        qual = "high"
        ElseIf Abs(r) < 0.9 Then
        qual = "very large"
        Else
        qual = "nearly perfect"
        End If
    
    'Funder and Ozer (2019, p. 166).
    ElseIf qual = "funder" Then
        ref = "Funder and Ozer (2019, p. 166)"
        If Abs(r) < 0.05 Then
        qual = "negligable"
        ElseIf Abs(r) < 0.1 Then
        qual = "very small"
        ElseIf Abs(r) < 0.2 Then
        qual = "small"
        ElseIf Abs(r) < 0.3 Then
        qual = "medium"
        ElseIf Abs(r) < 0.4 Then
        qual = "large"
        Else
        qual = "very large"
        End If
    End If
      
    'the output
    If output = "qual" Then
        th_pearson_r = qual
    ElseIf output = "ref" Then
        th_pearson_r = ref
    Else
        'Results
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "classification"
        res(1, 2) = "reference"
        res(2, 1) = qual
        res(2, 2) = ref
        
        th_pearson_r = res
    End If

End Function
