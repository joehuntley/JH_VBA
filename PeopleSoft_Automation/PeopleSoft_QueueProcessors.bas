Attribute VB_Name = "PeopleSoft_QueueProcessors"
Option Explicit



Private Const Q_MAX_CONSECUTIVE_FAILURES = 5  ' defines


Public Sub Process_PO_Queue()


    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim C_QUEUE_ID As Integer
    Dim C_TAG As Integer
    Dim C_PO_BUSINESS_UNIT As Integer
    Dim C_PO_VENDOR_NAME_SHORT As Integer
    Dim C_PO_VENDOR_LOCATION As Integer
    Dim C_PO_BUYER_ID As Integer
    Dim C_PO_APPROVER_ID As Integer
    Dim C_PO_QUOTE As Integer
    Dim C_PO_QUOTE_ATTACHMENT As Integer
    Dim C_PO_REF As Integer
    Dim C_PO_COMMENTS As Integer
    Dim C_PO_LINE As Integer
    Dim C_PO_LINE_ITEMID As Integer
    Dim C_PO_LINE_DESC As Integer
    Dim C_PO_SCH_QTY As Integer
    Dim C_PO_SCH_PRICE As Integer
    Dim C_PO_SCH_DUE_DATE As Integer
    Dim C_PO_SCH_SHIPTO_ID As Integer
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer
    Dim C_PO_DIST_PC As Integer
    Dim C_PO_DIST_ACTIVITY_ID As Integer
    Dim C_PO_DIST_LOCATION_ID As Integer
    Dim C_PO_NUM As Integer
    Dim C_PO_AMNT_TOTAL As Integer
    Dim C_LINE_BUDGET_ERR As Integer
    Dim C_LINE_BUDGET_ERR_FUND_REQ As Integer
    Dim C_PO_ERROR As Integer
    Dim C_LINE_ERROR As Integer
    
    Dim col As Integer
    col = 0
    col = col + 1: C_QUEUE_ID = col
    col = col + 1: C_TAG = col
    col = col + 1: C_PO_BUSINESS_UNIT = col
    col = col + 1: C_PO_VENDOR_NAME_SHORT = col
    col = col + 1: C_PO_VENDOR_LOCATION = col
    col = col + 1: C_PO_BUYER_ID = col
    col = col + 1: C_PO_APPROVER_ID = col
    col = col + 1: C_PO_QUOTE = col
    col = col + 1: C_PO_QUOTE_ATTACHMENT = col
    col = col + 1: C_PO_REF = col
    col = col + 1: C_PO_COMMENTS = col
    col = col + 1: C_PO_LINE = col
    col = col + 1: C_PO_LINE_ITEMID = col
    col = col + 1: C_PO_LINE_DESC = col
    col = col + 1: C_PO_SCH_QTY = col
    col = col + 1: C_PO_SCH_PRICE = col
    col = col + 1: C_PO_SCH_DUE_DATE = col
    col = col + 1: C_PO_SCH_SHIPTO_ID = col
    col = col + 1: C_PO_DIST_BUSINESS_UNIT_PC = col
    col = col + 1: C_PO_DIST_PC = col
    col = col + 1: C_PO_DIST_ACTIVITY_ID = col
    col = col + 1: C_PO_DIST_LOCATION_ID = col
    col = col + 1: C_PO_NUM = col
    col = col + 1: C_PO_AMNT_TOTAL = col
    col = col + 1: C_LINE_BUDGET_ERR = col
    col = col + 1: C_LINE_BUDGET_ERR_FUND_REQ = col
    col = col + 1: C_PO_ERROR = col
    col = col + 1: C_LINE_ERROR = col

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_QUEUE_ID, C_QUEUE_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim POs() As PeopleSoft_PurchaseOrder
    Dim POs_DoNotProcess() As Boolean
    Dim POs_DoNotProcessCount As Integer
    
    ReDim POs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder
    ReDim POs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
           
        If idxParent > 0 Then
            idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
            
            If idxChild = 1 Then
                POs(idxParent).PO_Fields.PO_BUSINESS_UNIT = queueTableRange.Cells(curRow, C_PO_BUSINESS_UNIT).Value
                POs(idxParent).PO_Fields.VENDOR_NAME_SHORT = queueTableRange.Cells(curRow, C_PO_VENDOR_NAME_SHORT).Value
                POs(idxParent).PO_Fields.PO_HDR_VENDOR_LOCATION = queueTableRange.Cells(curRow, C_PO_VENDOR_LOCATION).Value
                POs(idxParent).PO_Fields.PO_HDR_BUYER_ID = queueTableRange.Cells(curRow, C_PO_BUYER_ID).Value
                POs(idxParent).PO_Fields.PO_HDR_APPROVER_ID = queueTableRange.Cells(curRow, C_PO_APPROVER_ID).Value
                POs(idxParent).PO_Fields.PO_HDR_QUOTE = queueTableRange.Cells(curRow, C_PO_QUOTE).Value
                POs(idxParent).PO_Fields.PO_HDR_PO_REF = queueTableRange.Cells(curRow, C_PO_REF).Value
                POs(idxParent).PO_Fields.PO_HDR_COMMENTS = queueTableRange.Cells(curRow, C_PO_COMMENTS).Value
                
                POs(idxParent).PO_Fields.Quote_Attachment_FilePath = queueTableRange.Cells(curRow, C_PO_QUOTE_ATTACHMENT).Value
            End If
           
            PeopleSoft_PurchaseOrder_AddLineSimple POs(idxParent), _
               Trim(CStr(queueTableRange.Cells(curRow, C_PO_LINE_ITEMID).Value)), _
               Trim(CStr(queueTableRange.Cells(curRow, C_PO_LINE_DESC).Value)), _
               CDec(queueTableRange.Cells(curRow, C_PO_SCH_QTY).Value), _
               CDate(queueTableRange.Cells(curRow, C_PO_SCH_DUE_DATE).Value), _
               CLng(queueTableRange.Cells(curRow, C_PO_SCH_SHIPTO_ID).Value), _
               CStr(queueTableRange.Cells(curRow, C_PO_DIST_BUSINESS_UNIT_PC).Value), _
               CStr(queueTableRange.Cells(curRow, C_PO_DIST_PC).Value), _
               CStr(queueTableRange.Cells(curRow, C_PO_DIST_ACTIVITY_ID).Value), _
               CLng(queueTableRange.Cells(curRow, C_PO_DIST_LOCATION_ID).Value), _
               CCur(queueTableRange.Cells(curRow, C_PO_SCH_PRICE).Value)



            ' if any PO Line has text in the PO NUM column, do not process entire PO
            If POs_DoNotProcess(idxParent) = False And queueTableRange.Cells(curRow, C_PO_NUM).Value <> "" Then
                POs_DoNotProcess(idxParent) = True
                POs_DoNotProcessCount = POs_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' End - Create PO objects and load from spreadsheet
    
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------


    If POs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No POs will be processed: clear any any errors and try again", vbInformation
        Exit Sub
    End If
    
    
    Dim user As String, pass As String
    
    If Prompt_UserPass(user, pass) = False Then
        MsgBox "Canceled or empty user/pass given. Quitting"
        Exit Sub
    End If
    


    Dim conseqfailCount As Integer
    
    Dim session As PeopleSoft_Session
    Dim result As Boolean
    
    
    
    session = PeopleSoft_NewSession(user, pass)
  
  

    conseqfailCount = 0
  
    For idxParent = 1 To ssMap.PARENT_COUNT
        If POs_DoNotProcess(idxParent) = False Then
        
            
            ' new in 2.11: check if quote attachment exists - if not, then error immediately.
            If POs(idxParent).PO_Fields.Quote_Attachment_FilePath <> "" Then
                ' Check if file exists
                If Dir(POs(idxParent).PO_Fields.Quote_Attachment_FilePath) = "" Then
                    result = False
                    POs(idxParent).HasError = True
                    POs(idxParent).GlobalError = "File Not Found: " & POs(idxParent).PO_Fields.Quote_Attachment_FilePath
                End If
            End If
            
            If POs(idxParent).HasError = False Then
                result = PeopleSoft_PurchaseOrder_CutPO(session, POs(idxParent))
            End If
            
            Application.ScreenUpdating = False
            
            If result Then
                conseqfailCount = 0
            
                For idxChild = 1 To POs(idxParent).PO_LineCount
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    
                    queueTableRange.Cells(curRow, C_PO_NUM).Value = POs(idxParent).PO_ID
                    
                    
                    ' Populate price
                    If POs(idxParent).PO_Lines(idxChild).Schedules(1).ScheduleFields.PRICE > 0 Then
                        If Not queueTableRange.Cells(curRow, C_PO_SCH_PRICE).Value Then queueTableRange.Cells(curRow, C_PO_SCH_PRICE).Value = POs(idxParent).PO_Lines(idxChild).Schedules(1).ScheduleFields.PRICE
                    End If
                    
                    ' Populate PO amount
                    If POs(idxParent).PO_AMNT_TOTAL > 0 Then queueTableRange.Cells(curRow, C_PO_AMNT_TOTAL).Value = POs(idxParent).PO_AMNT_TOTAL
                    
                    ' Populate budget errors
                    Dim lineBC_HasError As Boolean
                    Dim lineBC_FundReq As Currency
                    
                    lineBC_HasError = False
                    lineBC_FundReq = 0
                    
                    If POs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors Then
                        For j = 1 To POs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_LineErrorCount
                            If POs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_LineErrors(j).LINE_NBR = idxChild Then
                                lineBC_HasError = True
                                lineBC_FundReq = POs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_LineErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
                    
                    
                    
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR).Value = IIf(lineBC_HasError, "Y", "")
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR_FUND_REQ).Value = IIf(lineBC_HasError, lineBC_FundReq, "")
                    
                    
                Next idxChild
            Else
                Dim poErrString As String, lineErrString As String
                
                poErrString = ""
                
                With POs(idxParent).PO_Fields
                    If .PO_BUSINESS_UNIT_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT: " & .PO_BUSINESS_UNIT_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_APPROVER_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_APPROVER_ID: " & .PO_HDR_APPROVER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_BUYER_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_BUYER_ID: " & .PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
                    'If .PO_HDR_VENDOR_ID_Result.ValidationFailed Then _
                    '    poErrString = poErrString & "|" & "PO_HDR_VENDOR_ID: " & .PO_HDR_VENDOR_ID_Result.ValidationErrorText & vbCrLf
                    If .VENDOR_NAME_SHORT_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "VENDOR_NAME_SHORT: " & .VENDOR_NAME_SHORT_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_LOCATION: " & .PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText & vbCrLf
                        
                    If .Quote_Attachment_FilePath_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "Quote Attachment: " & .Quote_Attachment_FilePath_Result.ValidationErrorText & vbCrLf
                End With
            
            
            
                For idxChild = 1 To POs(idxParent).PO_LineCount
                    
            
                    With POs(idxParent).PO_Lines(idxChild)
                        lineErrString = ""
                        
                        If .LineFields.PO_LINE_ITEM_ID_Result.ValidationFailed Then _
                             lineErrString = lineErrString & "|" & "PO_LINE_ITEM_ID: " & .LineFields.PO_LINE_ITEM_ID_Result.ValidationErrorText & vbCrLf
                                
                        If .Schedules(1).ScheduleFields.QTY_Result.ValidationFailed Then _
                             lineErrString = lineErrString & "|" & "PO_SCH_QTY: " & .Schedules(1).ScheduleFields.QTY_Result.ValidationErrorText & vbCrLf
                        If .Schedules(1).ScheduleFields.DUE_DATE_Result.ValidationFailed Then _
                             lineErrString = lineErrString & "|" & "PO_SCH_DUE_DATE: " & .Schedules(1).ScheduleFields.DUE_DATE_Result.ValidationErrorText & vbCrLf
                        If .Schedules(1).ScheduleFields.SHIPTO_ID_Result.ValidationFailed Then _
                             lineErrString = lineErrString & "|" & "PO_SCH_SHIPTO_ID: " & .Schedules(1).ScheduleFields.SHIPTO_ID_Result.ValidationErrorText & vbCrLf
                        
                        If .Schedules(1).DistributionFields.BUSINESS_UNIT_PC_Result.ValidationFailed Then _
                            lineErrString = lineErrString & "|" & "BUSINESS_UNIT_PC: " & .Schedules(1).DistributionFields.BUSINESS_UNIT_PC_Result.ValidationErrorText & vbCrLf
                        If .Schedules(1).DistributionFields.PROJECT_CODE_Result.ValidationFailed Then _
                            lineErrString = lineErrString & "|" & "PROJECT_CODE: " & .Schedules(1).DistributionFields.PROJECT_CODE_Result.ValidationErrorText & vbCrLf
                        If .Schedules(1).DistributionFields.ACTIVITY_ID_Result.ValidationFailed Then _
                            lineErrString = lineErrString & "|" & "ACTIVITY_ID: " & .Schedules(1).DistributionFields.ACTIVITY_ID_Result.ValidationErrorText & vbCrLf
                        If .Schedules(1).DistributionFields.LOCATION_ID_Result.ValidationFailed Then _
                             lineErrString = lineErrString & "|" & "LOCATION_ID: " & .Schedules(1).DistributionFields.LOCATION_ID_Result.ValidationErrorText & vbCrLf
                    End With
            
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_PO_NUM).Value = "<ERROR>"
                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = POs(idxParent).GlobalError & IIf(Len(poErrString) > 0, vbCrLf & poErrString, "")
                    queueTableRange.Cells(curRow, C_LINE_ERROR).Value = lineErrString
                    
                    
                    queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False
                    queueTableRange.Cells(curRow, C_LINE_ERROR).WrapText = False
                    
                Next idxChild
                
                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            Application.ScreenUpdating = True
            
            Debug.Print
        
        End If
    Next idxParent

    

End Sub
Public Sub Process_PO_Queue_RetryBudgetCheck()

'WORK IN PROGRESS

    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim C_QUEUE_ID As Integer
    Dim C_TAG As Integer
    Dim C_PO_BUSINESS_UNIT As Integer
    Dim C_PO_VENDOR_NAME_SHORT As Integer
    Dim C_PO_VENDOR_LOCATION As Integer
    Dim C_PO_BUYER_ID As Integer
    Dim C_PO_APPROVER_ID As Integer
    Dim C_PO_QUOTE As Integer
    Dim C_PO_REF As Integer
    Dim C_PO_COMMENTS As Integer
    Dim C_PO_LINE As Integer
    Dim C_PO_LINE_ITEMID As Integer
    Dim C_PO_LINE_DESC As Integer
    Dim C_PO_SCH_QTY As Integer
    Dim C_PO_SCH_PRICE As Integer
    Dim C_PO_SCH_DUE_DATE As Integer
    Dim C_PO_SCH_SHIPTO_ID As Integer
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer
    Dim C_PO_DIST_PC As Integer
    Dim C_PO_DIST_ACTIVITY_ID As Integer
    Dim C_PO_DIST_LOCATION_ID As Integer
    Dim C_PO_NUM As Integer
    Dim C_PO_AMNT_TOTAL As Integer
    Dim C_LINE_BUDGET_ERR As Integer
    Dim C_LINE_BUDGET_ERR_FUND_REQ As Integer
    Dim C_PO_ERROR As Integer
    Dim C_LINE_ERROR As Integer
    
    Dim col As Integer
    col = 0
    col = col + 1: C_QUEUE_ID = col
    col = col + 1: C_TAG = col
    col = col + 1: C_PO_BUSINESS_UNIT = col
    col = col + 1: C_PO_VENDOR_NAME_SHORT = col
    col = col + 1: C_PO_VENDOR_LOCATION = col
    col = col + 1: C_PO_BUYER_ID = col
    col = col + 1: C_PO_APPROVER_ID = col
    col = col + 1: C_PO_QUOTE = col
    col = col + 1: C_PO_REF = col
    col = col + 1: C_PO_COMMENTS = col
    col = col + 1: C_PO_LINE = col
    col = col + 1: C_PO_LINE_ITEMID = col
    col = col + 1: C_PO_LINE_DESC = col
    col = col + 1: C_PO_SCH_QTY = col
    col = col + 1: C_PO_SCH_PRICE = col
    col = col + 1: C_PO_SCH_DUE_DATE = col
    col = col + 1: C_PO_SCH_SHIPTO_ID = col
    col = col + 1: C_PO_DIST_BUSINESS_UNIT_PC = col
    col = col + 1: C_PO_DIST_PC = col
    col = col + 1: C_PO_DIST_ACTIVITY_ID = col
    col = col + 1: C_PO_DIST_LOCATION_ID = col
    col = col + 1: C_PO_NUM = col
    col = col + 1: C_PO_AMNT_TOTAL = col
    col = col + 1: C_LINE_BUDGET_ERR = col
    col = col + 1: C_LINE_BUDGET_ERR_FUND_REQ = col
    col = col + 1: C_PO_ERROR = col
    col = col + 1: C_LINE_ERROR = col

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim ssMap As SPREADSHEET_MAP_2L
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_QUEUE_ID, C_QUEUE_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim PO_BCs() As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams
    Dim PO_BCs_DoNotProcess() As Boolean
    Dim PO_BCs_DoNotProcessCount As Integer
    
    ReDim PO_BCs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams
    ReDim PO_BCs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    
    ' Defauilt - do not try to save with budget check
    For idxParent = 1 To ssMap.PARENT_COUNT
        PO_BCs_DoNotProcess(idxParent) = True
    Next idxParent

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
           
        If idxParent > 0 Then
            idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
            
            If idxChild = 1 Then
                PO_BCs(idxParent).PO_BU = queueTableRange.Cells(curRow, C_PO_BUSINESS_UNIT).Value
                PO_BCs(idxParent).PO_ID = queueTableRange.Cells(curRow, C_PO_NUM).Value
            End If
           
            ' Only re-try saving with budget check if at least on PO Line has text in the PO NUM column,
            ' and one of the lines has a budget error
            If PO_BCs_DoNotProcess(idxParent) = True And queueTableRange.Cells(curRow, C_PO_NUM).Value <> "" And queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR).Value <> "" Then
                PO_BCs_DoNotProcess(idxParent) = False
                PO_BCs_DoNotProcessCount = PO_BCs_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' End - Create PO objects and load from spreadsheet
    
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------

    If PO_BCs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No POs in PO queue has budget check errors.", vbInformation
        Exit Sub
    End If
    
    
        
    Dim user As String, pass As String
    
    If Prompt_UserPass(user, pass) = False Then
        MsgBox "Canceled or empty user/pass given. Quitting"
        Exit Sub
    End If


    Dim conseqfailCount As Integer
    
    Dim session As PeopleSoft_Session
    Dim result As Boolean
    
    
    
    session = PeopleSoft_NewSession(user, pass)
  
  

    conseqfailCount = 0
  
    For idxParent = 1 To ssMap.PARENT_COUNT
        If PO_BCs_DoNotProcess(idxParent) = False Then
        
            result = PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck(session, PO_BCs(idxParent))
            
            Application.ScreenUpdating = False
            
            If result Then
                conseqfailCount = 0
                
            
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    
                                       
                    ' Populate budget errors
                    Dim lineBC_HasError As Boolean
                    Dim lineBC_FundReq As Currency
                    
                    lineBC_HasError = False
                    lineBC_FundReq = 0
                    
                    If PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors Then
                        For j = 1 To PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_LineErrorCount
                            If PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_LineErrors(j).LINE_NBR = idxChild Then
                                lineBC_HasError = True
                                lineBC_FundReq = PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_LineErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
                    
                    
                    
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR).Value = IIf(lineBC_HasError, "Y", "")
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR_FUND_REQ).Value = IIf(lineBC_HasError, lineBC_FundReq, "")
                    
                    
                Next idxChild
            Else
                'Dim poErrString As String, lineErrString As String
                
                'poErrString = ""
                
                'With POs(idxParent).PO_Fields
                '    If .PO_BUSINESS_UNIT_Result.ValidationFailed Then _
                '        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT: " & .PO_BUSINESS_UNIT_Result.ValidationErrorText & vbCrLf
                '    If .PO_HDR_APPROVER_ID_Result.ValidationFailed Then _
                '        poErrString = poErrString & "|" & "PO_HDR_APPROVER_ID: " & .PO_HDR_APPROVER_ID_Result.ValidationErrorText & vbCrLf
                '    If .PO_HDR_BUYER_ID_Result.ValidationFailed Then _
                '        poErrString = poErrString & "|" & "PO_HDR_BUYER_ID: " & .PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
                '    'If .PO_HDR_VENDOR_ID_Result.ValidationFailed Then _
                '    '    poErrString = poErrString & "|" & "PO_HDR_VENDOR_ID: " & .PO_HDR_VENDOR_ID_Result.ValidationErrorText & vbCrLf
                '    If .VENDOR_NAME_SHORT_Result.ValidationFailed Then _
                '        poErrString = poErrString & "|" & "VENDOR_NAME_SHORT: " & .VENDOR_NAME_SHORT_Result.ValidationErrorText & vbCrLf
                '    If .PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then _
                '        poErrString = poErrString & "|" & "PO_HDR_VENDOR_LOCATION: " & .PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText & vbCrLf
                'End With
            
            
            
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = "Budget Check Err: " & PO_BCs(idxParent).GlobalError
                    queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False

                Next idxChild
                
                
                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            Application.ScreenUpdating = True
            
            Debug.Print
        
        End If
    Next idxParent



    
    Debug.Print

End Sub
Public Sub Process_PO_eQuote_Q()


    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_eQuote_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim C_QUEUE_ID As Integer
    Dim C_USER_DATA As Integer
    Dim C_E_QUOTE_NBR As Integer
    Dim C_QUOTE_ATTACHMENT As Integer
    Dim C_PO_BUSINESS_UNIT As Integer
    Dim C_PO_VENDOR_ID As Integer
    Dim C_PO_VENDOR_LOCATION As Integer
    Dim C_PO_BUYER_ID As Integer
    Dim C_PO_APPROVER_ID As Integer
    Dim C_PO_REF As Integer
    Dim C_PO_COMMENTS As Integer
    
    Dim C_PO_DUE_DATE As Integer
    Dim C_PO_SHIPTO_ID As Integer
    Dim C_PO_BUSINESS_UNIT_PC As Integer
    Dim C_PO_PROJECT_CODE As Integer
    Dim C_PO_ACTIVITY_ID As Integer
    Dim C_PO_LOCATION_ID As Integer
    
    Dim C_PO_LINE As Integer
    Dim C_PO_LINE_ITEMID As Integer
    Dim C_PO_LINE_DESC As Integer
    Dim C_PO_SCH_DUE_DATE As Integer
    Dim C_PO_SCH_SHIPTO_ID As Integer
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer
    Dim C_PO_DIST_PC As Integer
    Dim C_PO_DIST_ACTIVITY_ID As Integer
    Dim C_PO_DIST_LOCATION_ID As Integer
    Dim C_PO_NUM As Integer
    Dim C_PO_AMNT_TOTAL As Integer
    Dim C_PO_BUDGET_ERR As Integer
    Dim C_PO_BUDGET_ERR_FUND_REQ As Integer
    Dim C_PO_ERROR As Integer
    Dim C_LINE_ERROR As Integer
    
    Dim col As Integer
    col = 0
    col = col + 1: C_QUEUE_ID = col
    col = col + 1: C_USER_DATA = col
    col = col + 1: C_E_QUOTE_NBR = col
    col = col + 1: C_QUOTE_ATTACHMENT = col
    col = col + 1: C_PO_BUSINESS_UNIT = col
    col = col + 1: C_PO_VENDOR_ID = col
    col = col + 1: C_PO_VENDOR_LOCATION = col
    col = col + 1: C_PO_BUYER_ID = col
    col = col + 1: C_PO_APPROVER_ID = col
    col = col + 1: C_PO_REF = col
    col = col + 1: C_PO_COMMENTS = col
    
    col = col + 1: C_PO_DUE_DATE = col
    col = col + 1: C_PO_SHIPTO_ID = col
    col = col + 1: C_PO_BUSINESS_UNIT_PC = col
    col = col + 1: C_PO_PROJECT_CODE = col
    col = col + 1: C_PO_ACTIVITY_ID = col
    col = col + 1: C_PO_LOCATION_ID = col
    
    col = col + 1: C_PO_LINE = col
    col = col + 1: C_PO_LINE_ITEMID = col
    col = col + 1: C_PO_LINE_DESC = col
    col = col + 1: C_PO_SCH_DUE_DATE = col
    col = col + 1: C_PO_SCH_SHIPTO_ID = col
    col = col + 1: C_PO_DIST_BUSINESS_UNIT_PC = col
    col = col + 1: C_PO_DIST_PC = col
    col = col + 1: C_PO_DIST_ACTIVITY_ID = col
    col = col + 1: C_PO_DIST_LOCATION_ID = col
    col = col + 1: C_PO_NUM = col
    col = col + 1: C_PO_AMNT_TOTAL = col
    col = col + 1: C_PO_BUDGET_ERR = col
    col = col + 1: C_PO_BUDGET_ERR_FUND_REQ = col
    col = col + 1: C_PO_ERROR = col
    col = col + 1: C_LINE_ERROR = col

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_QUEUE_ID, C_QUEUE_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim PO_CFQs() As PeopleSoft_PurchaseOrder_CreateFromQuoteParams
    Dim PO_CFQs_DoNotProcess() As Boolean
    Dim PO_CFQs_DoNotProcessCount As Integer
    
    ReDim PO_CFQs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder_CreateFromQuoteParams
    ReDim PO_CFQs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
           
        If idxParent > 0 Then
            idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
            
            If idxChild = 1 Then
                PO_CFQs(idxParent).E_QUOTE_NBR = Trim(queueTableRange.Cells(curRow, C_E_QUOTE_NBR).Value)
                
                
                PO_CFQs(idxParent).PO_Fields.Quote_Attachment_FilePath = queueTableRange.Cells(curRow, C_QUOTE_ATTACHMENT).Value
                
                PO_CFQs(idxParent).PO_Fields.PO_BUSINESS_UNIT = queueTableRange.Cells(curRow, C_PO_BUSINESS_UNIT).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_VENDOR_ID = queueTableRange.Cells(curRow, C_PO_VENDOR_ID).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_VENDOR_LOCATION = queueTableRange.Cells(curRow, C_PO_VENDOR_LOCATION).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_BUYER_ID = queueTableRange.Cells(curRow, C_PO_BUYER_ID).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_APPROVER_ID = queueTableRange.Cells(curRow, C_PO_APPROVER_ID).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_PO_REF = queueTableRange.Cells(curRow, C_PO_REF).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_COMMENTS = queueTableRange.Cells(curRow, C_PO_COMMENTS).Value
                
                PO_CFQs(idxParent).PO_Defaults.SCH_DUE_DATE = queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value
                PO_CFQs(idxParent).PO_Defaults.SCH_SHIPTO_ID = queueTableRange.Cells(curRow, C_PO_SHIPTO_ID).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_BUSINESS_UNIT_PC = queueTableRange.Cells(curRow, C_PO_BUSINESS_UNIT_PC).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_PROJECT_CODE = queueTableRange.Cells(curRow, C_PO_PROJECT_CODE).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_ACTIVITY_ID = queueTableRange.Cells(curRow, C_PO_ACTIVITY_ID).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_LOCATION_ID = queueTableRange.Cells(curRow, C_PO_LOCATION_ID).Value
                
                PO_CFQs(idxParent).PO_LineModCount = ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                
                ReDim PO_CFQs(idxParent).PO_LineMods(1 To PO_CFQs(idxParent).PO_LineModCount) As PeopleSoft_PurchaseOrder_CreateFromQuote_LineModification
            End If
            
            If queueTableRange.Cells(curRow, C_PO_LINE).Value <> "" Then
                                
                PO_CFQs(idxParent).PO_LineMods(idxChild).PO_Line = CInt(queueTableRange.Cells(curRow, C_PO_LINE).Value)
                
                PO_CFQs(idxParent).PO_LineMods(idxChild).PO_LINE_ITEM_ID = queueTableRange.Cells(curRow, C_PO_LINE_ITEMID).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).PO_LINE_DESC = queueTableRange.Cells(curRow, C_PO_LINE_DESC).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).SCH_DUE_DATE = queueTableRange.Cells(curRow, C_PO_SCH_DUE_DATE).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).SCH_SHIPTO_ID = queueTableRange.Cells(curRow, C_PO_SCH_SHIPTO_ID).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).DIST_BUSINESS_UNIT_PC = queueTableRange.Cells(curRow, C_PO_DIST_BUSINESS_UNIT_PC).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).DIST_PROJECT_CODE = queueTableRange.Cells(curRow, C_PO_DIST_PC).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).DIST_ACTIVITY_ID = queueTableRange.Cells(curRow, C_PO_DIST_ACTIVITY_ID).Value
                PO_CFQs(idxParent).PO_LineMods(idxChild).DIST_LOCATION_ID = queueTableRange.Cells(curRow, C_PO_DIST_LOCATION_ID).Value
            
            Else
                PO_CFQs(idxParent).PO_LineMods(idxChild).PO_Line = -9999  'will not process
            End If


            ' if there text in the PO NUM column, do not process entire PO
            If PO_CFQs_DoNotProcess(idxParent) = False And queueTableRange.Cells(curRow, C_PO_NUM).Value <> "" Then
                PO_CFQs_DoNotProcess(idxParent) = True
                PO_CFQs_DoNotProcessCount = PO_CFQs_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' End - Create PO objects and load from spreadsheet
    
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------


    If PO_CFQs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No POs will be processed: clear any any errors and try again", vbInformation
        Exit Sub
    End If
    
    
    
    
    Dim user As String, pass As String
    
    If Prompt_UserPass(user, pass) = False Then
        MsgBox "Canceled or empty user/pass given. Quitting"
        Exit Sub
    End If


    Dim conseqfailCount As Integer
    
    Dim session As PeopleSoft_Session
    Dim result As Boolean
    
    
    
    session = PeopleSoft_NewSession(user, pass)

  

    conseqfailCount = 0
  
    For idxParent = 1 To ssMap.PARENT_COUNT
        If PO_CFQs_DoNotProcess(idxParent) = False Then
        
            ' new in 2.11: check if quote attachment exists - if not, then error immediately.
            If PO_CFQs(idxParent).PO_Fields.Quote_Attachment_FilePath <> "" Then
                ' Check if file exists
                If Dir(PO_CFQs(idxParent).PO_Fields.Quote_Attachment_FilePath) = "" Then
                    result = False
                    PO_CFQs(idxParent).HasError = True
                    PO_CFQs(idxParent).GlobalError = "File Not Found: " & PO_CFQs(idxParent).PO_Fields.Quote_Attachment_FilePath
                End If
            End If
        
            If PO_CFQs(idxParent).HasError = False Then
                result = PeopleSoft_PurchaseOrder_CreateFromQuote(session, PO_CFQs(idxParent))
            End If
            
            Application.ScreenUpdating = False
            
            If result Then
                conseqfailCount = 0
            
                For idxChild = 1 To PO_CFQs(idxParent).PO_LineModCount
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    
                    queueTableRange.Cells(curRow, C_PO_NUM).Value = PO_CFQs(idxParent).PO_ID
                    
                    ' Populate PO amount
                    If PO_CFQs(idxParent).PO_AMNT_TOTAL > 0 Then queueTableRange.Cells(curRow, C_PO_AMNT_TOTAL).Value = PO_CFQs(idxParent).PO_AMNT_TOTAL
                    
                    ' Populate budget error
                    Dim BC_totalFundReq As Currency
                    
                    BC_totalFundReq = 0
                    
                    If PO_CFQs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors Then
                        For j = 1 To PO_CFQs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount
                            If PO_CFQs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j).NOT_COMMIT_AMT > 0 Then
                                BC_totalFundReq = BC_totalFundReq + PO_CFQs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
                    
                    
                    
                    queueTableRange.Cells(curRow, C_PO_BUDGET_ERR).Value = IIf(PO_CFQs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors, "Y", "")
                    queueTableRange.Cells(curRow, C_PO_BUDGET_ERR_FUND_REQ).Value = IIf(PO_CFQs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors, BC_totalFundReq, "")
                    
                    
                Next idxChild
            Else
                ' -----------------------------------
                ' Begin - Build error strings and write to spreadsheet
                ' -----------------------------------
                Dim poErrString As String, lineErrString As String
                
                poErrString = ""
                
                With PO_CFQs(idxParent).PO_Fields
                    If .PO_BUSINESS_UNIT_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT: " & .PO_BUSINESS_UNIT_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_APPROVER_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_APPROVER_ID: " & .PO_HDR_APPROVER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_BUYER_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_BUYER_ID: " & .PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_ID: " & .PO_HDR_VENDOR_ID_Result.ValidationErrorText & vbCrLf
                    'If .VENDOR_NAME_SHORT_Result.ValidationFailed Then _
                    '    poErrString = poErrString & "|" & "VENDOR_NAME_SHORT: " & .VENDOR_NAME_SHORT_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_LOCATION: " & .PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText & vbCrLf
                    If .Quote_Attachment_FilePath_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "Quote Attachment: " & .Quote_Attachment_FilePath_Result.ValidationErrorText & vbCrLf
                End With
                With PO_CFQs(idxParent).PO_Defaults
                    If .SCH_DUE_DATE_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PO_DUE_DATE: " & .SCH_DUE_DATE_Result.ValidationErrorText & vbCrLf
                    If .SCH_SHIPTO_ID_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PO_SHIPTO_ID: " & .SCH_SHIPTO_ID_Result.ValidationErrorText & vbCrLf
                    If .DIST_BUSINESS_UNIT_PC_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT_PC: " & .DIST_BUSINESS_UNIT_PC_Result.ValidationErrorText & vbCrLf
                    If .DIST_PROJECT_CODE_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_PROJECT_CODE: " & .DIST_PROJECT_CODE_Result.ValidationErrorText & vbCrLf
                    If .DIST_ACTIVITY_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_ACTIVITY_ID: " & .DIST_ACTIVITY_ID_Result.ValidationErrorText & vbCrLf
                    If .DIST_LOCATION_ID_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PO_LOCATION_ID: " & .DIST_LOCATION_ID_Result.ValidationErrorText & vbCrLf
                End With
            
            
                For idxChild = 1 To PO_CFQs(idxParent).PO_LineModCount
                    
                    lineErrString = ""
                    
                    If PO_CFQs(idxParent).PO_LineMods(idxChild).PO_Line > 0 Then
                        With PO_CFQs(idxParent).PO_LineMods(idxChild)
                            If .PO_LINE_ITEM_ID_Result.ValidationFailed Then _
                                 lineErrString = lineErrString & "|" & "PO_LINE_ITEM_ID: " & .PO_LINE_ITEM_ID_Result.ValidationErrorText & vbCrLf
                            If .SCH_DUE_DATE_Result.ValidationFailed Then _
                                 lineErrString = lineErrString & "|" & "PO_SCH_DUE_DATE: " & .SCH_DUE_DATE_Result.ValidationErrorText & vbCrLf
                            If .SCH_SHIPTO_ID_Result.ValidationFailed Then _
                                 lineErrString = lineErrString & "|" & "PO_SCH_SHIPTO_ID: " & .SCH_SHIPTO_ID_Result.ValidationErrorText & vbCrLf
                            If .DIST_BUSINESS_UNIT_PC_Result.ValidationFailed Then _
                                lineErrString = lineErrString & "|" & "PO_DIST_BUSINESS_UNIT_PC: " & .DIST_BUSINESS_UNIT_PC_Result.ValidationErrorText & vbCrLf
                            If .DIST_PROJECT_CODE_Result.ValidationFailed Then _
                                lineErrString = lineErrString & "|" & "PO_DIST_PROJECT_CODE: " & .DIST_PROJECT_CODE_Result.ValidationErrorText & vbCrLf
                            If .DIST_ACTIVITY_ID_Result.ValidationFailed Then _
                                lineErrString = lineErrString & "|" & "PO_DIST_ACTIVITY_ID: " & .DIST_ACTIVITY_ID_Result.ValidationErrorText & vbCrLf
                            If .DIST_LOCATION_ID_Result.ValidationFailed Then _
                                 lineErrString = lineErrString & "|" & "PO_DIST_LOCATION_ID: " & .DIST_LOCATION_ID_Result.ValidationErrorText & vbCrLf
                        End With
                    End If
            
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_PO_NUM).Value = "<ERROR>"
                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = PO_CFQs(idxParent).GlobalError & IIf(Len(poErrString) > 0, vbCrLf & poErrString, "")
                    queueTableRange.Cells(curRow, C_LINE_ERROR).Value = lineErrString
                    
                    
                    queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False
                    queueTableRange.Cells(curRow, C_LINE_ERROR).WrapText = False
                    
                Next idxChild
                ' -----------------------------------
                ' End - Build error strings and write to spreadsheet
                ' -----------------------------------
                
                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            Application.ScreenUpdating = True
            
            Debug.Print
        
        End If
    Next idxParent


    
End Sub


Public Sub Process_PO_BudgetCheck_Q()


    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_BudgetCheck_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim C_PO_BU As Integer
    Dim C_PO_ID As Integer
    Dim C_BC_RESULT As Integer
    Dim C_BC_FUND_REQ As Integer
    Dim C_BC_ERROR As Integer
    
    Dim col As Integer
    col = 0
    col = col + 1: C_PO_BU = col
    col = col + 1: C_PO_ID = col
    col = col + 1: C_BC_RESULT = col
    col = col + 1: C_BC_FUND_REQ = col
    col = col + 1: C_BC_ERROR = col

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim ssMap As SPREADSHEET_MAP_1L
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_1D(queueTableRange, C_PO_ID, C_PO_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim PO_BCs() As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams
    Dim PO_BCs_DoNotProcess() As Boolean
    Dim PO_BCs_DoNotProcessCount As Integer
    
    ReDim PO_BCs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams
    ReDim PO_BCs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
        
        If idxParent > 0 Then
            PO_BCs(idxParent).PO_BU = queueTableRange.Cells(curRow, C_PO_BU).Value
            PO_BCs(idxParent).PO_ID = queueTableRange.Cells(curRow, C_PO_ID).Value
               
            ' Only re-try saving with budget check if passed budget check is blank
            If PO_BCs_DoNotProcess(idxParent) = False And queueTableRange.Cells(curRow, C_BC_RESULT).Value <> "" Then
                PO_BCs_DoNotProcess(idxParent) = True
                PO_BCs_DoNotProcessCount = PO_BCs_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' End - Create PO objects and load from spreadsheet
    
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------

    If PO_BCs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No POs in budget check queue has budget has balnk result. Clear errors and try again.", vbInformation
        Exit Sub
    End If
    
    
        
    Dim user As String, pass As String
    
    If Prompt_UserPass(user, pass) = False Then
        MsgBox "Canceled or empty user/pass given. Quitting"
        Exit Sub
    End If


    Dim conseqfailCount As Integer
    
    Dim session As PeopleSoft_Session
    Dim result As Boolean
    
    
    
    session = PeopleSoft_NewSession(user, pass)
  
  

    conseqfailCount = 0
  
    For idxParent = 1 To ssMap.PARENT_COUNT
        If PO_BCs_DoNotProcess(idxParent) = False Then
        
            result = PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck(session, PO_BCs(idxParent))
            
            Application.ScreenUpdating = False
            
            
            
            If result Then
                conseqfailCount = 0
                
            
                
                Dim bcFundReq As String
                
                If PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors Then
                    bcFundReq = ""
                    
                    For j = 1 To PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount
                        With PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j)
                            bcFundReq = "," & .PROJECT_ID & ":" & .NOT_COMMIT_AMT
                        End With
                    Next j
                    
                    If Len(bcFundReq) > 0 Then bcFundReq = Mid(bcFundReq, 2)
                End If
                
                
                curRow = ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW
                
                queueTableRange.Cells(curRow, C_BC_RESULT).Value = IIf(PO_BCs(idxParent).BudgetCheck_Result.BudgetCheck_HasErrors, "FAIL", "PASS")
                queueTableRange.Cells(curRow, C_BC_FUND_REQ).Value = bcFundReq
            
                
            Else
                curRow = ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW
                
                queueTableRange.Cells(curRow, C_BC_RESULT).Value = "<ERROR>"
                
                queueTableRange.Cells(curRow, C_BC_ERROR).Value = "Budget Check Err: " & PO_BCs(idxParent).GlobalError
                queueTableRange.Cells(curRow, C_BC_ERROR).WrapText = False

                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            Application.ScreenUpdating = True
            
            Debug.Print
        
        End If
    Next idxParent



    
    Debug.Print

End Sub
Public Sub Process_PO_Receipt_Queue()



    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    ' Primary - PO_ID (Unique per PO) - defines each Rcpt
    ' Secondary - PO Line #/Schedule # - defines each Rcpt.Item
    Const HEADER_ROWS_SIZE = 0
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_Receipt_Q").Range("4:65536")
    
    ' Column numbers (in order)
    Dim C_USER_DATA As Integer
    Dim C_PO_BU As Integer
    Dim C_PO_ID As Integer
    Dim C_PO_LINE As Integer
    Dim C_PO_SCH As Integer
    Dim C_RECEIVE_QTY As Integer
    Dim C_ACCEPT_QTY As Integer
    Dim C_ITEM_ID As Integer
    Dim C_CATS_FLAG As Integer
    Dim C_TRANS_ITEM_DESC As Integer
    Dim C_RCPT_ID As Integer
    Dim C_RCPT_ERR As Integer
    Dim C_RCPT_ITEM_ERR As Integer
    
    Dim col As Integer
    col = 0
    col = col + 1: C_USER_DATA = col
    col = col + 1: C_PO_BU = col
    col = col + 1: C_PO_ID = col
    col = col + 1: C_PO_LINE = col
    col = col + 1: C_PO_SCH = col
    col = col + 1: C_RECEIVE_QTY = col
    col = col + 1: C_ACCEPT_QTY = col
    col = col + 1: C_ITEM_ID = col
    col = col + 1: C_CATS_FLAG = col
    col = col + 1: C_TRANS_ITEM_DESC = col
    col = col + 1: C_RCPT_ID = col
    col = col + 1: C_RCPT_ERR = col
    col = col + 1: C_RCPT_ITEM_ERR = col
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - Rcpt
    ' Secondary - Rcpt_Item
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    Dim i As Long
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_PO_ID, C_PO_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim Rcpts() As PeopleSoft_Receipt
    Dim Rcpts_ReceiveAllLinesFlag() As Boolean
    Dim Rcpts_DoNotProcess() As Boolean
    Dim Rcpts_DoNotProcessCount As Integer
    
    ReDim Rcpts(1 To ssMap.PARENT_COUNT) As PeopleSoft_Receipt
    ReDim Rcpts_ReceiveAllLinesFlag(1 To ssMap.PARENT_COUNT) As Boolean
    ReDim Rcpts_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    
    Dim PO_Line_Value As Variant
    

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
        
        If idxParent > 0 Then
            idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
            
            PO_Line_Value = queueTableRange.Cells(curRow, C_PO_LINE).Value
        
            If idxChild = 1 Then
                Rcpts(idxParent).PO_BU = queueTableRange.Cells(curRow, C_PO_BU).Value
                Rcpts(idxParent).PO_ID = queueTableRange.Cells(curRow, C_PO_ID).Value
                
                ' update 2.10.1: if first line # is ALL, then set recieve mode to receive all lines on the PO.
                If PO_Line_Value = "ALL" Then
                    Rcpts(idxParent).ReceiveMode = RECEIVE_ALL
                Else
                    Rcpts(idxParent).ReceiptItemCount = ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                    ReDim Rcpts(idxParent).ReceiptItems(1 To Rcpts(idxParent).ReceiptItemCount) As PeopleSoft_Receipt_Item
                End If
                
            End If
            
            ' update 2.10.1: only add PO Lines and PO_Schedule if we are not receiving all lines
            If Rcpts(idxParent).ReceiveMode <> RECEIVE_ALL Then
                If IsNumeric(queueTableRange.Cells(curRow, C_PO_LINE).Value) Then Rcpts(idxParent).ReceiptItems(idxChild).PO_Line = queueTableRange.Cells(curRow, C_PO_LINE).Value
                If IsNumeric(queueTableRange.Cells(curRow, C_PO_SCH).Value) Then Rcpts(idxParent).ReceiptItems(idxChild).PO_Schedule = queueTableRange.Cells(curRow, C_PO_SCH).Value
                If IsNumeric(queueTableRange.Cells(curRow, C_RECEIVE_QTY).Value) Then Rcpts(idxParent).ReceiptItems(idxChild).Receive_Qty = queueTableRange.Cells(curRow, C_RECEIVE_QTY).Value
            
                  
                ' any item did not properly the set PO_Line or Schedule (not numeric), do not process receipts for entire PO
                If Rcpts_DoNotProcess(idxParent) = False And (Rcpts(idxParent).ReceiptItems(idxChild).PO_Line = 0 Or Rcpts(idxParent).ReceiptItems(idxChild).PO_Schedule = 0) Then
                    Rcpts_DoNotProcess(idxParent) = True
                    Rcpts_DoNotProcessCount = Rcpts_DoNotProcessCount + 1
                End If
            End If
            
            ' if any item has something in the receipt ID, do not process receipts for entire PO
            If Rcpts_DoNotProcess(idxParent) = False And queueTableRange.Cells(curRow, C_RCPT_ID).Value <> "" Then
                Rcpts_DoNotProcess(idxParent) = True
                Rcpts_DoNotProcessCount = Rcpts_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    
    If Rcpts_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No receipts will be processed: clear any any errors and/or check PO_LINE, PO_SCH for validity. Then try again", vbInformation
        Exit Sub
    End If
    
    
   
    Dim user As String, pass As String
    
    If Prompt_UserPass(user, pass) = False Then
        MsgBox "Canceled or empty user/pass given. Quitting"
        Exit Sub
    End If
    

    Dim conseqfailCount As Integer
    
    Dim session As PeopleSoft_Session
    Dim result As Boolean
    
       
    session = PeopleSoft_NewSession(user, pass)
  
    conseqfailCount = 0



    Dim errString As String, secErrString As String
  
    ' Primary - Rcpt
    ' Secondary - Rcpt_Item
    For idxParent = 1 To ssMap.PARENT_COUNT
        If Rcpts_DoNotProcess(idxParent) = False Then
        
            result = PeopleSoft_Receipt_CreateReceipt(session, Rcpts(idxParent))
            
            
            Application.ScreenUpdating = False
            
            If result Then ' Receipt created OK
                If Rcpts(idxParent).ReceiveMode = RECEIVE_SPECIFIED Then
                    ' copy individual receipt #s each row in spreadsheet
                    For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                        curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                        
                        If Rcpts(idxParent).ReceiptItems(idxChild).HasError = False Then
                            queueTableRange.Cells(curRow, C_RCPT_ID).Value = Rcpts(idxParent).RECEIPT_ID
                        End If
                        
                        ' added in v2.10.2: include error text even if no error (it may have other useful info).
                        If Rcpts(idxParent).GlobalError <> "" Then
                            queueTableRange.Cells(curRow, C_RCPT_ERR).Value = Rcpts(idxParent).GlobalError
                        End If
                        
                    Next idxChild
                ElseIf Rcpts(idxParent).ReceiveMode = RECEIVE_ALL Then
                    ' update 2.10.1: copy all receipt lines to one row
                    Dim rcptItemLineStr As String, rcptItemQtyStr As String, rcptItemIDStr As String, rcptCatsFlagStr As String
                    Dim rcptTransItemDescStr As String
                    
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(1)
                    rcptItemLineStr = ""
                    rcptItemQtyStr = ""
                    rcptItemIDStr = ""
                    rcptCatsFlagStr = ""
                    rcptTransItemDescStr = ""
                    
                    For i = 1 To Rcpts(idxParent).ReceiptItemCount
                        rcptItemLineStr = rcptItemLineStr & Rcpts(idxParent).ReceiptItems(i).PO_Line & vbCrLf
                        rcptItemQtyStr = rcptItemQtyStr & Rcpts(idxParent).ReceiptItems(i).Receive_Qty & vbCrLf
                        rcptItemIDStr = rcptItemIDStr & Rcpts(idxParent).ReceiptItems(i).Item_ID & vbCrLf
                        rcptCatsFlagStr = rcptCatsFlagStr & Rcpts(idxParent).ReceiptItems(i).CATS_FLAG & vbCrLf
                        rcptTransItemDescStr = rcptTransItemDescStr & Rcpts(idxParent).ReceiptItems(i).TRANS_ITEM_DESC & vbCrLf
                    Next i
                    
                    ' Remove last CR-LF at end of eachs tring
                    rcptItemLineStr = Left$(rcptItemLineStr, Len(rcptItemLineStr) - Len(vbCrLf))
                    rcptItemQtyStr = Left$(rcptItemQtyStr, Len(rcptItemQtyStr) - Len(vbCrLf))
                    rcptItemIDStr = Left$(rcptItemIDStr, Len(rcptItemIDStr) - Len(vbCrLf))
                    rcptCatsFlagStr = Left$(rcptCatsFlagStr, Len(rcptCatsFlagStr) - Len(vbCrLf))
                    rcptTransItemDescStr = Left$(rcptTransItemDescStr, Len(rcptTransItemDescStr) - Len(vbCrLf))
                
                    
                    queueTableRange.Cells(curRow, C_RCPT_ID).Value = Rcpts(idxParent).RECEIPT_ID
                    queueTableRange.Cells(curRow, C_PO_LINE).Value = RTrim(rcptItemLineStr)
                    queueTableRange.Cells(curRow, C_RECEIVE_QTY).Value = RTrim(rcptItemQtyStr)
                    queueTableRange.Cells(curRow, C_ITEM_ID).Value = RTrim(rcptItemIDStr)
                    queueTableRange.Cells(curRow, C_CATS_FLAG).Value = RTrim(rcptCatsFlagStr)
                    queueTableRange.Cells(curRow, C_TRANS_ITEM_DESC).Value = RTrim(rcptTransItemDescStr)
                    
                    
                    ' added in v2.10.2: include error text even if no error (it may have other useful info).
                    If Rcpts(idxParent).GlobalError <> "" Then
                        queueTableRange.Cells(curRow, C_RCPT_ERR).Value = Rcpts(idxParent).GlobalError
                    End If
                Else
                    Err.Raise -1, , "Unknown Receive Mode (This should never happen)"
                End If
                
                conseqfailCount = 0
            Else
                
                errString = ""
                
                With Rcpts(idxParent)
                    If .PO_BU_Result.ValidationFailed Then _
                        errString = errString & "|" & "PO_BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                End With
            
                ' 2.10.1: ReceiveMode = ALL -> Maps all output to one row
                Dim childCount As Long
                childCount = ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                If Rcpts(idxParent).ReceiveMode = RECEIVE_ALL Then childCount = 1
                
                
                For idxChild = 1 To childCount
                    'secErrString = Rcpts(idxParent).ReceiptItems(idxChild).ItemError
                  
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_RCPT_ID).Value = "<ERROR>"
                    queueTableRange.Cells(curRow, C_RCPT_ERR).Value = Rcpts(idxParent).GlobalError
                    'queueTableRange.Cells(curRow, C_RCPT_ITEM_ERR).Value = secErrString
                    
                    
                    queueTableRange.Cells(curRow, C_RCPT_ID).WrapText = False
                    queueTableRange.Cells(curRow, C_RCPT_ERR).WrapText = False
                    
                Next idxChild
                
                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            ' update 2.10.1: include individual receive errorsonly if RECEIVE_SPECIFIED is
            If Rcpts(idxParent).ReceiveMode = RECEIVE_SPECIFIED Then
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT ' fixed bug: 2.
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                        
                    ' Populate individual receipt item errors
                    If Rcpts(idxParent).ReceiptItems(idxChild).HasError Then
                        queueTableRange.Cells(curRow, C_RCPT_ID).Value = "<ERROR>"
                        queueTableRange.Cells(curRow, C_RCPT_ITEM_ERR).Value = Rcpts(idxParent).ReceiptItems(idxChild).ItemError
                    End If
                    
                    'Populate receive qty if not set
                    If Rcpts(idxParent).ReceiptItems(idxChild).Receive_Qty > 0 Then
                        If Not queueTableRange.Cells(curRow, C_RECEIVE_QTY).Value Then queueTableRange.Cells(curRow, C_RECEIVE_QTY).Value = Rcpts(idxParent).ReceiptItems(idxChild).Receive_Qty
                    End If
                    
                    ' Populate accept qty, item ID, trans item desc
                    If Rcpts(idxParent).ReceiptItems(idxChild).Accept_Qty > 0 Then _
                        queueTableRange.Cells(curRow, C_ACCEPT_QTY).Value = Rcpts(idxParent).ReceiptItems(idxChild).Accept_Qty
                    If Rcpts(idxParent).ReceiptItems(idxChild).Item_ID <> "" Then _
                        queueTableRange.Cells(curRow, C_ITEM_ID).Value = Rcpts(idxParent).ReceiptItems(idxChild).Item_ID
                    If Rcpts(idxParent).ReceiptItems(idxChild).CATS_FLAG <> "" Then _
                        queueTableRange.Cells(curRow, C_CATS_FLAG).Value = Rcpts(idxParent).ReceiptItems(idxChild).CATS_FLAG
                    If Rcpts(idxParent).ReceiptItems(idxChild).TRANS_ITEM_DESC <> "" Then _
                        queueTableRange.Cells(curRow, C_TRANS_ITEM_DESC).Value = Rcpts(idxParent).ReceiptItems(idxChild).TRANS_ITEM_DESC
                    
                    
                Next idxChild
            End If
            
            
            Application.ScreenUpdating = True
            
        End If
    Next idxParent



End Sub


Sub Process_PO_ChangeOrder_Queue()

    
    Dim i As Integer, j As Integer
    

    Dim queueTableRange As Range
    Const HEADER_ROWS_SIZE = 0
    
     
    Set queueTableRange = ThisWorkbook.Worksheets("PO_ChangeOrder_Q").Range("4:65535")
    
    ' Column numbers (in order)
    Dim C_PO_BU As Integer
    Dim C_PO_ID As Integer
    Dim C_PO_DUE_DATE As Integer
    Dim C_PO_FLG_SEND_TO_VENDOR As Integer
    Dim C_CO_REASON As Integer
    Dim C_CO_STATUS As Integer
    Dim C_CO_ERROR As Integer
    
    Dim C_PO_LINE As Integer
    Dim C_PO_SCHEDULE As Integer

    
    Dim col As Integer
    col = 0
    col = col + 1: C_PO_BU = col
    col = col + 1: C_PO_ID = col
    col = col + 1: C_PO_DUE_DATE = col
    
    col = col + 1: C_PO_LINE = col
    col = col + 1: C_PO_SCHEDULE = col
    
    col = col + 1: C_PO_FLG_SEND_TO_VENDOR = col
    col = col + 1: C_CO_REASON = col
    col = col + 1: C_CO_STATUS = col
    col = col + 1: C_CO_ERROR = col
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_ID
    ' Secondary - N/A
    Dim ssMap As SPREADSHEET_MAP_2L
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_PO_ID, C_PO_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim PO_COs() As PeopleSoft_PurchaseOrder_ChangeOrder
    Dim PO_COs_DoNotProcess() As Boolean
    Dim PO_COs_DoNotProcessCount As Integer
    
    ' Begin - Create PO objects and load from spreadsheet
    ReDim PO_COs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder_ChangeOrder
    ReDim PO_COs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean

    
    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
        
        If idxParent > 0 Then
           PO_COs(idxParent).PO_BU = queueTableRange.Cells(curRow, C_PO_BU).Value
           PO_COs(idxParent).PO_ID = queueTableRange.Cells(curRow, C_PO_ID).Value
           
           If queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value <> "" Then PO_COs(idxParent).PO_DUE_DATE = queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value
           
           If queueTableRange.Cells(curRow, C_PO_FLG_SEND_TO_VENDOR).Value <> "" Then
                Select Case UCase(queueTableRange.Cells(curRow, C_PO_FLG_SEND_TO_VENDOR).Value)
                    Case "X", "Y", "YES":
                        PO_COs(idxParent).PO_HDR_FLG_SEND_TO_VENDOR = PeopleSoft_Page_CheckboxAction.SetAsChecked
                    Case "N", "NO":
                        PO_COs(idxParent).PO_HDR_FLG_SEND_TO_VENDOR = PeopleSoft_Page_CheckboxAction.SetAsUnchecked
                    Case Else
                        PO_COs(idxParent).PO_HDR_FLG_SEND_TO_VENDOR = PeopleSoft_Page_CheckboxAction.KeepExistingValue
                End Select
           End If
           
           PO_COs(idxParent).ChangeReason = queueTableRange.Cells(curRow, C_CO_REASON).Value

            ' Do not process change orders if status isn't blank
            If PO_COs_DoNotProcess(idxParent) = False And queueTableRange.Cells(curRow, C_CO_STATUS).Value <> "" Then
                PO_COs_DoNotProcess(idxParent) = True
                PO_COs_DoNotProcessCount = PO_COs_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------

    
    If PO_COs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No PO Change Orders will be processed: clear any any errors and try again", vbInformation
        Exit Sub
    End If
    
    
    Dim user As String, pass As String
    
    If Prompt_UserPass(user, pass) = False Then
        MsgBox "Canceled or empty user/pass given. Quitting"
        Exit Sub
    End If
    
    

    Dim conseqfailCount As Integer
    
    Dim session As PeopleSoft_Session
    Dim result As Boolean
    
       
    session = PeopleSoft_NewSession(user, pass)
  
    conseqfailCount = 0

    ' Primary - PO_ID
    For idxParent = 1 To ssMap.PARENT_COUNT
        If PO_COs_DoNotProcess(idxParent) = False Then
        
            result = PeopleSoft_PurchaseOrder_ProcessChangeOrder(session, PO_COs(idxParent))
            
            Application.ScreenUpdating = False
            
            curRow = ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW
            
            If result Then ' Change order OK
                
                queueTableRange.Cells(curRow, C_CO_STATUS).Value = "COMPLETE"
                
                conseqfailCount = 0
            Else
                Dim errString As String, secErrString As String
                
                errString = ""
                
                With PO_COs(idxParent)
                    'If .PO_BU_Result.ValidationFailed Then _
                    '    errString = errString & "|" & "PO_BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_DUE_DATE_Result.ValidationFailed Then _
                        errString = errString & "|" & "PO_DUE_DATE: " & .PO_DUE_DATE_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_BUYER_ID_Result.ValidationFailed Then _
                        errString = errString & "|" & "PO_BUYER_ID: " & .PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
                End With
            
                
                queueTableRange.Cells(curRow, C_CO_STATUS).Value = "<ERROR>"
                queueTableRange.Cells(curRow, C_CO_ERROR).Value = PO_COs(idxParent).GlobalError & IIf(Len(errString) > 0, vbCrLf & errString, "")
                
                queueTableRange.Cells(curRow, C_CO_ERROR).WrapText = False
                   
                
                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            Application.ScreenUpdating = True
            
        End If
    Next idxParent


End Sub

Private Function Prompt_UserPass(ByRef user As String, ByRef pass As String) As Boolean
    


    user = InputBox("Enter username (USWIN):", , Environ$("username"))
    
    If Len(user) = 0 Then Prompt_UserPass = False: Exit Function
    
    pass = InputBoxDK("Enter password:", "")
    
    If Len(pass) = 0 Then Prompt_UserPass = False: Exit Function
    
    
    Prompt_UserPass = True

End Function


