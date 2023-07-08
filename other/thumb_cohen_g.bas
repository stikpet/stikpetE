Attribute VB_Name = "thumb_cohen_g"
Function th_cohen_g(g, Optional qual = "cohen", Optional output = "all")
Attribute th_cohen_g.VB_Description = "Rules of Thumb for Cohen g"
Attribute th_cohen_g.VB_ProcData.VB_Invoke_Func = " \n14"

    'Cohen's rule of thumb
    If qual = "cohen" Then
    
        ref = "Cohen (1988, pp. 147-149)"
    
        'Use Cohen (1988, pp. 147-149).
        If Abs(g) < 0.05 Then
            qual = "negligible"
        ElseIf Abs(g) < 0.15 Then
            qual = "small"
        ElseIf Abs(g) < 0.25 Then
            qual = "medium"
        Else
            qual = "large"
        End If
        
    End If
    
    
    'Output
    If output = "all" Then
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "classification"
        res(1, 2) = "source"
        
        res(2, 1) = qual
        res(2, 2) = ref
        th_cohen_g = res
        
    ElseIf output = "ref" Then
        th_cohen_g = ref
    Else
        th_cohen_g = qual
    End If
    

End Function
