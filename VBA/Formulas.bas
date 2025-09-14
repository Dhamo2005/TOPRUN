Attribute VB_Name = "Module2"
Function EvalMath(rng As Range) As Variant
    Dim cell As Range
    Dim total As Double
    On Error Resume Next
    
    total = 0
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            total = total + Evaluate(cell.Value)
        End If
    Next cell
    
    ' If result is 0, return a dash instead
    If total = 0 Then
        EvalMath = "-"
    Else
        EvalMath = total
    End If
End Function

Function EvalPowercut(rng As Range) As Variant
    Dim cell As Range
    Dim total As Double
    On Error Resume Next
    
    total = 0
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            total = total + Evaluate(cell.Value)
        End If
    Next cell
    
    ' If result is 0, return a dash instead
    If total = 0 Then
        EvalPowercut = ""
    Else
        EvalPowercut = total
    End If
End Function

