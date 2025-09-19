Attribute VB_Name = "Module2"
Function EvalMath(rng As Range) As Variant
    Dim cell As Range
    Dim total As Double
    Dim matches As Object
    Dim regex As Object
    Dim val As String
    Dim i As Integer

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\d+(\.\d+)?"

    total = 0
    For Each cell In rng
        val = Trim(cell.Value)
        If val <> "" Then
            If regex.Test(val) Then
                Set matches = regex.Execute(val)
                For i = 0 To matches.Count - 1
                    total = total + CDbl(matches(i).Value)
                Next i
            End If
        End If
    Next cell

    If total = 0 Then
        EvalMath = "-"
    Else
        EvalMath = total
    End If
End Function


Function EvalPower(rng As Range) As Variant
    Dim cell As Range
    Dim total As Double
    Dim matches As Object
    Dim regex As Object
    Dim val As String
    Dim i As Integer

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\d+(\.\d+)?"

    total = 0
    For Each cell In rng
        val = Trim(cell.Value)
        If val <> "" Then
            If regex.Test(val) Then
                Set matches = regex.Execute(val)
                For i = 0 To matches.Count - 1
                    total = total + CDbl(matches(i).Value)
                Next i
            End If
        End If
    Next cell

    If total = 0 Then
        EvalPower = "0"
    Else
        EvalPower = total
    End If
End Function

