Attribute VB_Name = "thumb_cohen_g"
Public Sub th_cohen_g_addHelp()
Application.MacroOptions _
    Macro:="th_cohen_g", _
    Description:="Rules of Thumb for Cohen g", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the Cohen g value", _
        "optional optional the rule of thumb to be used. Currently only cohen available (also then default)", _
        "output to show, either all (default) for array result, qual for only the classification, or ref for the reference")
End Sub
Function th_cohen_g(g, Optional qual = "cohen", Optional output = "all")

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
