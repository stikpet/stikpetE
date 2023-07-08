Attribute VB_Name = "thumb_cohen_w"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function th_cohen_w(w, Optional qual = "cohen", Optional output = "both")
Attribute th_cohen_w.VB_Description = "Cohen's w rule-of-thumb"
Attribute th_cohen_w.VB_ProcData.VB_Invoke_Func = " \n14"

    'Qualification
    
    If qual = "cohen" Then
        'Use Cohen (1988, p. 227).
        ref = "Cohen (1988, p. 227)"
        If Abs(w) < 0.1 Then
            clas = "negligible"
        ElseIf Abs(w) < 0.3 Then
            clas = "small"
        ElseIf Abs(w) < 0.5 Then
            clas = "medium"
        Else
            clas = "large"
        End If
    End If
    
    'the output
    If output = "qual" Then
        th_cohen_w = clas
    ElseIf output = "ref" Then
        th_cohen_w = ref
    Else
        'Results
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "classification"
        res(1, 2) = "reference"
        res(2, 1) = clas
        res(2, 2) = ref
        
        th_cohen_w = res
    End If
End Function
