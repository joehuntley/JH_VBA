Private Sub Worksheet_Activate()

    ThisWorkbook.ActiveSheet.PivotTables("ptPoFromQuoteQueuePoSummary").RefreshTable
    ThisWorkbook.ActiveSheet.PivotTables("ptPoQueuePoSummary").RefreshTable
    
End Sub
