Attribute VB_Name = "Module1"

Sub Combine_workbooks()
Attribute Combine_workbooks.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Combine_workbooks Macro
'
' Keyboard Shortcut: Ctrl+e
'
    'On Error GoTo ErrorHandler
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim firstWb As Workbook
    Dim newWs As Worksheet
    
    ' Disable screen updates to speed up the process
    Application.ScreenUpdating = False
    
    ' Set the first workbook as the destination
    Set firstWb = ActiveWorkbook
    
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = Sheets("Sheet1").Range("A4").Value
    Rows("3:3").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$E$13").AutoFilter Field:=3, Criteria1:="="
    Range("A15").Select
    
    ' Loop through all currently open workbooks
    For Each wb In Workbooks
        ' Skip the first workbook (the destination workbook)
        If wb.Name <> firstWb.Name And wb.Name <> "PERSONAL.XLSB" Then
            ' Loop through all sheets in the current workbook
            For Each ws In wb.Sheets
                ws.Name = ws.Range("A4").Value
                ws.Rows("3:3").AutoFilter
                ws.Range("$A$3:$E$13").AutoFilter Field:=3, Criteria1:="="
                ' Copy the sheet to the first workbook after the last sheet
                ws.Move After:=firstWb.Sheets(firstWb.Sheets.Count)
            Next ws
        End If
    Next wb
    
    ' Enable screen updates again
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "There was an error", vbOKOnly, "Information"
End Sub

