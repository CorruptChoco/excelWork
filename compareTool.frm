VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} compareTool 
   Caption         =   "Compare Tool"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   OleObjectBlob   =   "CompareTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "compareTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ClearButton_Click()

Call UserForm_Initialize

End Sub



Private Sub Label2_Click()

End Sub

Private Sub okButton_Click()
Dim prod As Workbook
Dim test As Workbook
Dim lastCell As String, ws1 As Worksheet, ws2 As Worksheet, strCell As String
Dim a As String, d As String
Dim b, c, e, f
Dim cellLetter As String, cellNumber As String

Dim f1 As String, f2 As String
f1 = file1.Value
f2 = file2.Value
Set test = Workbooks(f1 & ".xlsx")
Set prod = Workbooks(f2 & ".xlsx")


'https://analystcave.com/excel-vba-last-row-last-column-last-cell/#:~:text=To%20get%20the%20Last%20Row%20with%20data%20in,Dim%20lastRow%20as%20Range%20Set%20lastRow%20%3D%20Range%28%22A1%22%29.End%28xlDown%29
'Get Last Cell with Data in Worksheet using SpecialCells
For Y = 1 To test.Sheets.Count
    Set ws1 = test.Sheets(Y)
    Set ws2 = prod.Sheets(Y)
    
    Set lastCellws1 = ws1.Cells.SpecialCells(xlCellTypeLastCell)
    'Debug.Print "Row: " & lastCell.Row & ", Column: " & lastCell.Column
    Set lastCellws2 = ws2.Cells.SpecialCells(xlCellTypeLastCell)
    'Debug.Print "Row: " & lastCell.Row & ", Column: " & lastCell.Column
    
    a = lastCellws1.Address
    b = a

    For c = 48 To 57
        a = Replace(a, Chr(c), "")
    Next c

    For c = 1 To Len(a)
        b = Replace(b, Mid(a, c, 1), "")
    Next c
    
    d = lastCellws2.Address
    e = d

    For f = 48 To 57
        d = Replace(d, Chr(f), "")
    Next f

    For f = 1 To Len(d)
        e = Replace(e, Mid(d, f, 1), "")
    Next f
    
    If a > d Then
    cellLetter = a
    Else
    cellLetter = d
    End If
    If b > e Then
    cellNumber = b
    Else
    cellNumber = e
    End If
    
    lastCell = cellLetter & cellNumber
    
    strCell = lastCell
    strCell = Replace(strCell, "$", "")
    For Each X In test.Sheets(Y).Range("A1:" & strCell)
        If X.Value <> prod.Sheets(Y).Range(X.Address).Value Then
            prod.Sheets(Y).Range(X.Address).Interior.Color = vbYellow
            X.Interior.Color = vbYellow
        End If
    Next X
Next Y

Unload Me

End Sub

Private Sub UserForm_Initialize()
'Empty NameTextBox
file1.Value = ""

'Empty PhoneTextBox
file2.Value = ""

'Set Focus on NameTextBox
file1.SetFocus
End Sub


