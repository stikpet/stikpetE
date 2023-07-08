Attribute VB_Name = "thumb_cohen_d"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function th_cohen_d(d, Optional qual = "sawilowsky", Optional output = "all")
Attribute th_cohen_d.VB_Description = "Rules of Thumb for Cohen d"
Attribute th_cohen_d.VB_ProcData.VB_Invoke_Func = " \n14"
    'Cohen (1988, p. 40).
    If qual = "cohen" Then
        ref = "Cohen (1988, p. 40)"
        If Abs(d) < 0.2 Then
            qual = "negligible"
        ElseIf Abs(d) < 0.5 Then
            qual = "small"
        ElseIf Abs(d) < 0.8 Then
            qual = "medium"
        Else
            qual = "large"
        End If
    
    'Lovakov and Agadullina (2021, p. 501)
    ElseIf qual = "lovakov" Then
        ref = "Lovakov and Agadullina (2021, p. 501)"
        If Abs(d) < 0.15 Then
            qual = "negligible"
            ElseIf Abs(d) < 0.35 Then
            qual = "small"
        ElseIf Abs(d) < 0.65 Then
            qual = "medium"
        Else
            qual = "large"
        End If
    
    'Rosenthal (1996, p. 45).
    ElseIf qual = "rosenthal" Then
        ref = "Rosenthal (1996, p. 45)"
        If Abs(d) < 0.2 Then
            qual = "negligible"
        ElseIf Abs(d) < 0.5 Then
            qual = "small"
        ElseIf Abs(d) < 0.8 Then
            qual = "medium"
        ElseIf Abs(d) < 1.3 Then
            qual = "large"
        Else
            qual = "very large"
        End If
    
    'Sawilowsky (2009, p. 599)
    ElseIf qual = "sawilowsky" Then
        ref = "Sawilowsky (2009, p. 599)"
        If Abs(d) < 0.1 Then
            qual = "negligible"
        ElseIf Abs(d) < 0.2 Then
            qual = "very small"
        ElseIf Abs(d) < 0.5 Then
            qual = "small"
        ElseIf Abs(d) < 0.8 Then
            qual = "medium"
        ElseIf Abs(d) < 1.2 Then
            qual = "large"
        ElseIf Abs(d) < 2 Then
            qual = "very large"
        Else
            qual = "huge"
        End If
    End If
    
    If output = "qual" Then
        th_cohen_d = qual
    ElseIf output = "ref" Then
        th_cohen_d = ref
    Else
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "classification"
        res(1, 2) = "source"
        
        res(2, 1) = qual
        res(2, 2) = ref
        th_cohen_d = res
   End If
    
End Function
   
    
