Attribute VB_Name = "Module3"
Sub Clear_Data_Startup()
    Dim response1 As VbMsgBoxResult
    Dim response2 As VbMsgBoxResult
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Startup")
    
    ' First confirmation
    response1 = MsgBox("Are you sure you want to clear data from A11:Q100, S11:T100, and U11:X100?", vbYesNo + vbQuestion, "First Confirmation | DAMO Automation")
    If response1 <> vbYes Then Exit Sub
    
    ' Second confirmation
    response2 = MsgBox("This action cannot be undone. Do you really want to proceed?", vbYesNo + vbExclamation, "Final Confirmation | DAMO Automation")
    If response2 <> vbYes Then Exit Sub
    
    ' Clear all specified ranges
    ws.Range("A11:Q100,S7:T100,U11:X100").ClearContents
    MsgBox "Data cleared successfully.", vbInformation, "Done | DAMO Automation"
End Sub

Sub Clear_Data_Process()
    Dim response1 As VbMsgBoxResult
    Dim response2 As VbMsgBoxResult
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Process")
    
    ' First confirmation
    response1 = MsgBox("Are you sure you want to clear data from A10:Q50, S7:S50 and U10:U50?", _
                       vbYesNo + vbQuestion, "First Confirmation | DAMO Automation")
    If response1 <> vbYes Then Exit Sub
    
    ' Second confirmation
    response2 = MsgBox("This action cannot be undone. Do you really want to proceed?", _
                       vbYesNo + vbExclamation, "Final Confirmation | DAMO Automation")
    If response2 <> vbYes Then Exit Sub
    
    ' Clear specified ranges
    ws.Range("S7:S50,A10:Q50,U10:U50").ClearContents
    MsgBox "Process sheet data cleared successfully.", vbInformation, "Done | DAMO Automation"
End Sub



