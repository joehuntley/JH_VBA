Private Sub Worksheet_Activate()


    ThisWorkbook.ActiveSheet.PivotTables("ptPoQueueBudgetErrors").RefreshTable
    ThisWorkbook.ActiveSheet.PivotTables("ptPoFromQuoteQueueBudgetErrors").RefreshTable
    


End Sub
