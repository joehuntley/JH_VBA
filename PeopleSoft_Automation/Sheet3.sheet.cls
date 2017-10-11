Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    
    Const C_VENDOR As Long = 4
    Const C_QUOTE_ATTACHMENT As Long = 9 ' Quote attachment is fourth column
    Const HEADER_ROW As Long = 3 ' Header row is 3rd column
    
    Dim checkColVal As String
    
    Dim filePath As String
    Dim Fd As Object, fileChosen As Integer
    
    Const msoFileDialogFilePicker = 3
    
    If Target.Column = C_QUOTE_ATTACHMENT And Target.Row > HEADER_ROW Then
        checkColVal = Target.EntireRow.Columns(C_VENDOR).Value
        If checkColVal = "" Then Exit Sub ' do not do anything if check column is blank
    
        Set Fd = Application.fileDialog(msoFileDialogFilePicker)
        
        Fd.Title = "Select quote to attach"
        
        Fd.Filters.Clear
        Fd.Filters.Add "All Files", "*.*"
        
        fileChosen = Fd.Show()
        
        If fileChosen = -1 Then
            filePath = Fd.SelectedItems(1)
            Target.Value = filePath
        End If
        
        Cancel = True
    End If

End Sub
