Attribute VB_Name = "Module1"
Sub UpdateMorningMeetingSlides()
    Dim pptApp As Object, pptPres As Object
    Dim ws As Worksheet
    Dim pptPath As String

    ' ---------------- Get PowerPoint Path ----------------
    Set ws = ThisWorkbook.Sheets("Publish")
    pptPath = ws.Range("C5").Value

    If Dir(pptPath) = "" Then
    MsgBox "PowerPoint file not found at:" & vbCrLf & pptPath & vbCrLf & _
           "Please verify the path in 'Publish' sheet (cell C5).", vbExclamation, "Missing File @DAMO Automation"
    ws.Activate: ws.Range("C5").Select
    Exit Sub
End If


    ' ---------------- Clear Startup Sheet Rows ----------------
    ClearStartupSheet
    ClearProcessSheet
    IgnoreNumberAsTextErrors

    ' ---------------- Launch PowerPoint ----------------
    On Error Resume Next
    Set pptApp = GetObject(Class:="PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0

    If pptApp Is Nothing Then
        MsgBox "Unable to launch PowerPoint.", vbCritical, "Launch Failed @DAMO Automation"
        Exit Sub
    End If

    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Open(pptPath)

    If pptPres Is Nothing Then
        MsgBox "Failed to open presentation.", vbCritical, "Open Failed @DAMO Automation"
        Exit Sub
    End If

    ' ---------------- Update Slides ----------------
    UpdateDateOnSlides pptPres
    UpdateSlide1 pptPres
    UpdateSlide2 pptPres
    UpdateSlide3 pptPres
    UpdateSlide4 pptPres

    ' ---------------- Final Message ----------------
    MsgBox "Morning meeting slides were published successfully.", vbInformation, "Successfully Published @DAMO Automation"
    AppActivate Application.Caption
End Sub

Sub ClearStartupSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Startup")
    
    ' Find last non-empty cell in column D starting from row 11
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < 11 Then lastRow = 11 ' Ensure minimum starting row

    ' Write serial numbers starting from 1 at row 11
    If lastRow >= 11 Then
        For i = 11 To lastRow
            ws.Cells(i, 1).Value = i - 10 ' Serial numbers: 1, 2, 3, ...
        Next i
    End If

    With ws.Range("E7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(E8:G8)"
        End If
    End With
    
    With ws.Range("E8")
        If Not .HasFormula Then
            .Formula = "=EvalPower(E11:E99)-E10"
        End If
    End With
    
    With ws.Range("F8")
        If Not .HasFormula Then
            .Formula = "=EvalPower(F11:F99)-F10"
        End If
    End With
    
    With ws.Range("G8")
        If Not .HasFormula Then
            .Formula = "=EvalPower(G11:G99)-G10"
        End If
    End With
    
    With ws.Range("E9")
        If Not .HasFormula Then
            .Formula = "=EvalPower(E10:G10)"
        End If
    End With
    
    With ws.Range("E10")
        If Not .HasFormula Then
            .Formula = "=U7"
        End If
    End With
    
    With ws.Range("F10")
        If Not .HasFormula Then
            .Formula = "=V7"
        End If
    End With
    
    With ws.Range("G10")
        If Not .HasFormula Then
            .Formula = "=W7"
        End If
    End With
    
    With ws.Range("H7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(H11:H99)"
        End If
    End With
    
    With ws.Range("I7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(I11:I99)"
        End If
    End With
    
    With ws.Range("J7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(J11:J99)"
        End If
    End With
    
    With ws.Range("K7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(K11:K99)"
        End If
    End With
    
    With ws.Range("L7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(L11:L99)"
        End If
    End With
    
    With ws.Range("M7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(M11:M99)"
        End If
    End With
    
    With ws.Range("N7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(N11:N99)"
        End If
    End With
    
    With ws.Range("O7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(O11:O99)"
        End If
    End With
    
    With ws.Range("P7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(P11:P99)"
        End If
    End With
    
    With ws.Range("Q7")
        If Not .HasFormula Then
            .Formula = "=EvalMath(Q11:Q99)"
        End If
    End With
    
    With ws.Range("R7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(R11:R99)"
        End If
    End With
    
    With ws.Range("R9")
        If Not .HasFormula Then
            .Formula = "=EvalPower(X7:Z8)"
        End If
    End With
    
    With ws.Range("U7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(U11:U99)"
        End If
    End With
    
    With ws.Range("V7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(V11:V99)"
        End If
    End With
    
    With ws.Range("W7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(W11:W99)"
        End If
    End With
    
    With ws.Range("X7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(X11:X100)"
        End If
    End With
    
    With ws.Range("Y7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(Y11:Y100)"
        End If
    End With
    
    With ws.Range("Z7")
        If Not .HasFormula Then
            .Formula = "=EvalPower(Z11:Z100)"
        End If
    End With


  ' Ensure R11 has the default formula before autofill
With ws.Range("R11")
    If Not .HasFormula Then
        .Formula = "=(EvalMath(E11:G11)*T11)"
    End If
End With

' Ensure X11 = U11*T11
With ws.Range("X11")
    If Not .HasFormula Then
        .Formula = "=U11*T11"
    End If
End With

' Ensure Y11 = V11*T11
With ws.Range("Y11")
    If Not .HasFormula Then
        .Formula = "=V11*T11"
    End If
End With

' Ensure Z11 = W11*T11
With ws.Range("Z11")
    If Not .HasFormula Then
        .Formula = "=W11*T11"
    End If
End With

' AutoFill formulas down to lastRow
If lastRow > 11 Then
    ws.Range("R11").AutoFill Destination:=ws.Range("R11:R" & lastRow)
    ws.Range("X11").AutoFill Destination:=ws.Range("X11:X" & lastRow)
    ws.Range("Y11").AutoFill Destination:=ws.Range("Y11:Y" & lastRow)
    ws.Range("Z11").AutoFill Destination:=ws.Range("Z11:Z" & lastRow)
End If
IgnoreNumberAsTextErrors

    ' Clear contents from A(lastRow+1) to Z100
    If lastRow + 1 <= 100 Then
        ws.Range("A" & lastRow + 1 & ":Z100").ClearContents
    End If
End Sub

Sub ClearProcessSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Process")
    
    ' Find last non-empty cell in column D starting from row 10
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < 10 Then lastRow = 10 ' Ensure minimum starting row

' Write serial numbers starting from 1 at row 10
If lastRow >= 10 Then
    For i = 10 To lastRow
        ws.Cells(i, 1).Value = i - 9 ' Serial numbers: 1, 2, 3, ...
    Next i
End If


    ' Ensure R10 has the default formula before autofill
    With ws.Range("R10")
        If Not .HasFormula Then
            .Formula = "=(EvalPower(E10:G10)*U10)"
        End If
    End With

With ws.Range("E7")
    If Not .HasFormula Then
        .Formula = "=EvalPower(E9:G9)"
    End If
End With

With ws.Range("E8")
    If Not .HasFormula Then
        .Formula = "=EvalPower(E9:G9)"
    End If
End With

With ws.Range("E9")
    If Not .HasFormula Then
        .Formula = "=EvalPower(E10:E53)"
    End If
End With

With ws.Range("F9")
    If Not .HasFormula Then
        .Formula = "=EvalPower(F10:F53)"
    End If
End With

With ws.Range("G9")
    If Not .HasFormula Then
        .Formula = "=EvalPower(G10:G53)"
    End If
End With

With ws.Range("H7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(H10:H53)"
    End If
End With

With ws.Range("I7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(I10:I53)"
    End If
End With

With ws.Range("J7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(J10:J53)"
    End If
End With

With ws.Range("K7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(K10:K53)"
    End If
End With

With ws.Range("L7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(L10:L53)"
    End If
End With

With ws.Range("M7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(M10:M53)"
    End If
End With

With ws.Range("N7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(N10:N53)"
    End If
End With

With ws.Range("O7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(O10:O53)"
    End If
End With

With ws.Range("P7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(P10:P53)"
    End If
End With

With ws.Range("Q7")
    If Not .HasFormula Then
        .Formula = "=EvalMath(Q10:Q53)"
    End If
End With

With ws.Range("R7")
    If Not .HasFormula Then
        .Formula = "=EvalPower(R10:R53)"
    End If
End With


    ' AutoFill formula from R10 down to R(lastRow)
    If lastRow > 10 Then
        ws.Range("R10").AutoFill Destination:=ws.Range("R10:R" & lastRow)
    End If
    
    ' Clear contents from A(lastRow+1) to Z100
    If lastRow + 1 <= 100 Then
        ws.Range("A" & lastRow + 1 & ":Z100").ClearContents
    End If
    IgnoreNumberAsTextErrors
End Sub



' ---------------- Update Date Text ----------------
Sub UpdateDateOnSlides(pptPres As Object)
    Dim ws As Worksheet
    Dim newDateText As String
    Dim pptSlide As Object, pptShape As Object
    Dim i As Long, j As Long

    Set ws = ThisWorkbook.Sheets("Publish")
    newDateText = "DATE : " & ws.Range("C4").Value

    For i = 1 To pptPres.Slides.Count
        Set pptSlide = pptPres.Slides(i)
        For j = 1 To pptSlide.Shapes.Count
            Set pptShape = pptSlide.Shapes(j)
            If pptShape.HasTextFrame Then
                If pptShape.TextFrame.HasText Then
                    If LCase(Left(Trim(pptShape.TextFrame.TextRange.Text), 4)) = "date" Then
                        pptShape.TextFrame.TextRange.Text = newDateText
                    End If
                End If
            End If
        Next j
    Next i
End Sub

' ---------------- Slide 1 ----------------
Sub UpdateSlide1(pptPres As Object)
    Dim ws As Worksheet, pptSlide As Object, pptTable As Object
    Dim rng As Range, i As Long, j As Long

    Set ws = ThisWorkbook.Sheets("Startup")
    Set pptSlide = pptPres.Slides(1)
    Set pptTable = pptSlide.Shapes(4).Table

    On Error Resume Next
    pptTable.cell(3, 5).Shape.TextFrame.TextRange.Text = ws.Range("E7").Value
    pptTable.cell(4, 5).Shape.TextFrame.TextRange.Text = ws.Range("E8").Value
    pptTable.cell(4, 6).Shape.TextFrame.TextRange.Text = ws.Range("F8").Value
    pptTable.cell(4, 7).Shape.TextFrame.TextRange.Text = ws.Range("G8").Value
    pptTable.cell(5, 5).Shape.TextFrame.TextRange.Text = ws.Range("E9").Value
    pptTable.cell(6, 5).Shape.TextFrame.TextRange.Text = ws.Range("E10").Value
    pptTable.cell(6, 6).Shape.TextFrame.TextRange.Text = ws.Range("F10").Value
    pptTable.cell(6, 7).Shape.TextFrame.TextRange.Text = ws.Range("G10").Value
    On Error GoTo 0

    For j = 8 To 19
        pptTable.cell(3, j).Shape.TextFrame.TextRange.Text = ws.Cells(7, j).Value
        pptTable.cell(5, j).Shape.TextFrame.TextRange.Text = ws.Cells(9, j).Value
    Next j

    Set rng = ws.Range("A11:S22")
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            If i + 6 <= pptTable.Rows.Count And j <= pptTable.Columns.Count Then
                pptTable.cell(i + 6, j).Shape.TextFrame.TextRange.Text = rng.Cells(i, j).Value
            End If
        Next j
    Next i
End Sub

' ---------------- Slide 2 ----------------
Sub UpdateSlide2(pptPres As Object)
    Dim ws As Worksheet, pptSlide As Object, pptTable As Object
    Dim rng As Range, i As Long, j As Long

    Set ws = ThisWorkbook.Sheets("Startup")
    Set rng = ws.Range("A23:S34")
    Set pptSlide = pptPres.Slides(2)

    For i = 1 To pptSlide.Shapes.Count
        If pptSlide.Shapes(i).HasTable Then
            Set pptTable = pptSlide.Shapes(i).Table
            Exit For
        End If
    Next i

    If pptTable Is Nothing Then Exit Sub

    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            If i + 3 <= pptTable.Rows.Count And j <= pptTable.Columns.Count Then
                pptTable.cell(i + 3, j).Shape.TextFrame.TextRange.Text = rng.Cells(i, j).Value
            End If
        Next j
    Next i
End Sub

' ---------------- Slide 3 ----------------
Sub UpdateSlide3(pptPres As Object)
    Dim ws As Worksheet, pptSlide As Object, pptTable As Object
    Dim rng As Range, lastRow As Long, i As Long, j As Long

    Set ws = ThisWorkbook.Sheets("Startup")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A35:S46")
    Set pptSlide = pptPres.Slides(3)

    For i = 1 To pptSlide.Shapes.Count
        If pptSlide.Shapes(i).HasTable Then
            Set pptTable = pptSlide.Shapes(i).Table
            Exit For
        End If
    Next i

    If pptTable Is Nothing Then Exit Sub

    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            If i + 3 <= pptTable.Rows.Count And j <= pptTable.Columns.Count Then
                pptTable.cell(i + 3, j).Shape.TextFrame.TextRange.Text = rng.Cells(i, j).Value
            End If
        Next j
    Next i
End Sub
Sub UpdateSlide4(pptPres As Object)
    Dim ws As Worksheet, pptSlide As Object, pptTable As Object
    Dim rng As Range, i As Long, j As Long

    Set ws = ThisWorkbook.Sheets("Process")
    Set pptSlide = pptPres.Slides(4)

    ' Locate the first table on the slide
    For i = 1 To pptSlide.Shapes.Count
        If pptSlide.Shapes(i).HasTable Then
            Set pptTable = pptSlide.Shapes(i).Table
            Exit For
        End If
    Next i

    If pptTable Is Nothing Then Exit Sub

    ' ---------------- Fill Main Table Range ----------------
    Set rng = ws.Range("A10:S20")
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            If i + 5 <= pptTable.Rows.Count And j <= pptTable.Columns.Count Then
                pptTable.cell(i + 5, j).Shape.TextFrame.TextRange.Text = rng.Cells(i, j).Value
            End If
        Next j
    Next i

    ' ---------------- Fill Header and Summary Cells ----------------
    On Error Resume Next
    If pptTable.Rows.Count >= 5 Then
        pptTable.cell(3, 5).Shape.TextFrame.TextRange.Text = ws.Range("E7").Value
        pptTable.cell(4, 5).Shape.TextFrame.TextRange.Text = ws.Range("E8").Value
        pptTable.cell(5, 5).Shape.TextFrame.TextRange.Text = ws.Range("E9").Value
        pptTable.cell(5, 6).Shape.TextFrame.TextRange.Text = ws.Range("F9").Value
        pptTable.cell(5, 7).Shape.TextFrame.TextRange.Text = ws.Range("G9").Value
    End If

    For j = 8 To 19
        If j <= pptTable.Columns.Count Then
            pptTable.cell(3, j).Shape.TextFrame.TextRange.Text = ws.Cells(7, j).Value
        End If
    Next j
    On Error GoTo 0
End Sub
Sub IgnoreNumberAsTextErrors()
    Dim ws As Worksheet
    Dim cell As Range

    Set ws = ActiveSheet

    On Error Resume Next
    For Each cell In ws.Range("A1:Z100")
        If cell.Errors(xlNumberAsText).Value = True Then
            cell.Errors(xlNumberAsText).Ignore = True
        End If
    Next cell
    On Error GoTo 0
End Sub




