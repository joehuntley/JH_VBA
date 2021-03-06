Attribute VB_Name = "PeopleSoft_QueueProcessors"
Option Explicit

' PeopleSoft_QueueProcessors
' ------------------------------------------------
' All the routines below are meant to
'  (1) Load the spreadsheet data into PeopleSoft_Automation structures. This involves converting a flat spreadsheet into a parent-child structures with the help of SpreadsheetTableToMultiLevelMap_2D() function
'  (2) Perform any pre-process checking
'  (3) Call the respective PeopleSoft automation function for each parent unit
'  (4) Read any output and write back to the mapped row for each unit.

Private Type PS_Automation_Config_Options
    Q_MAX_CONSECUTIVE_ERRORS As Integer
End Type

Private CONFIG_OPTIONS As PS_Automation_Config_Options
Private Sub Initialize()
    ' Loads configuration options. Call before running any queues.
    
    
    Dim CFG_MAX_CONSECUTIVE_ERRORS As Integer
    Dim CFG_DEBUGOPTS_CAPTURE_EXCEPTIONS As Boolean
    Dim CFG_DEBUGOPTS_ADD_METHOD_NAMES As Boolean
    Dim CFG_DEBUGOPTS_SAVE_OUTPUT As Boolean
    Dim CFG_DEBUGOPTS_SAVE_SCREENSHOTS As Boolean
    Dim CFG_DEBUGOPTS_SAVE_PAGE_SRC As Boolean
    Dim CFG_DEBUGOPTS_QUIT_BEFORE_SAVE As Boolean
    
    CFG_MAX_CONSECUTIVE_ERRORS = ThisWorkbook.Names("CFG_MAX_CONSECUTIVE_ERRORS").RefersToRange.Value
    CFG_DEBUGOPTS_CAPTURE_EXCEPTIONS = ThisWorkbook.Names("CFG_DEBUGOPTS_CAPTURE_EXCEPTIONS").RefersToRange.Value
    CFG_DEBUGOPTS_ADD_METHOD_NAMES = ThisWorkbook.Names("CFG_DEBUGOPTS_ADD_METHOD_NAMES").RefersToRange.Value
    CFG_DEBUGOPTS_SAVE_OUTPUT = ThisWorkbook.Names("CFG_DEBUGOPTS_SAVE_OUTPUT").RefersToRange.Value
    CFG_DEBUGOPTS_SAVE_SCREENSHOTS = ThisWorkbook.Names("CFG_DEBUGOPTS_SAVE_SCREENSHOTS").RefersToRange.Value
    CFG_DEBUGOPTS_SAVE_PAGE_SRC = ThisWorkbook.Names("CFG_DEBUGOPTS_SAVE_PAGE_SRC").RefersToRange.Value
    CFG_DEBUGOPTS_QUIT_BEFORE_SAVE = ThisWorkbook.Names("CFG_DEBUGOPTS_QUIT_BEFORE_SAVE").RefersToRange.Value

    CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS = CFG_MAX_CONSECUTIVE_ERRORS
    
    PeopleSoft_SetConfigOptions captureExceptionsAsError:=CFG_DEBUGOPTS_CAPTURE_EXCEPTIONS, _
                addMethodNamesToExceptions:=CFG_DEBUGOPTS_ADD_METHOD_NAMES, _
                writeDebugOutputToFile:=CFG_DEBUGOPTS_SAVE_OUTPUT, _
                writePageSrcToFile:=CFG_DEBUGOPTS_SAVE_SCREENSHOTS, _
                takeScreenShot:=CFG_DEBUGOPTS_SAVE_PAGE_SRC, _
                quitAutomationBeforeSave:=CFG_DEBUGOPTS_QUIT_BEFORE_SAVE
    

End Sub


Public Sub Process_PO_Queue()


    Call Initialize


    Dim i As Integer, j As Integer
    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim col As Integer: col = 1
    
    Dim C_QUEUE_ID As Integer: C_QUEUE_ID = col: col = col + 1
    Dim C_USER_DATA As Integer: C_USER_DATA = col: col = col + 1
    Dim C_PO_BUSINESS_UNIT As Integer: C_PO_BUSINESS_UNIT = col: col = col + 1
    Dim C_PO_VENDOR As Integer: C_PO_VENDOR = col: col = col + 1
    Dim C_PO_VENDOR_LOCATION As Integer: C_PO_VENDOR_LOCATION = col: col = col + 1
    Dim C_PO_CONTRACT_ID As Integer: C_PO_CONTRACT_ID = col: col = col + 1
    Dim C_PO_BUYER_ID As Integer: C_PO_BUYER_ID = col: col = col + 1
    Dim C_PO_XPRESS_BID_ID As Integer: C_PO_XPRESS_BID_ID = col: col = col + 1
    Dim C_PO_QUOTE As Integer: C_PO_QUOTE = col: col = col + 1
    Dim C_QUOTE_ATTACHMENT As Integer: C_QUOTE_ATTACHMENT = col: col = col + 1
    Dim C_PO_REF As Integer: C_PO_REF = col: col = col + 1
    Dim C_PO_COMMENTS As Integer: C_PO_COMMENTS = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_CODE As Integer: C_PO_PRICE_OVERAGE_CODE = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_REASON As Integer: C_PO_PRICE_OVERAGE_REASON = col: col = col + 1
    Dim C_PO_LINE As Integer: C_PO_LINE = col: col = col + 1
    Dim C_PO_LINE_ITEMID As Integer: C_PO_LINE_ITEMID = col: col = col + 1
    Dim C_PO_LINE_DESC As Integer: C_PO_LINE_DESC = col: col = col + 1
    Dim C_PO_SCH_QTY As Integer: C_PO_SCH_QTY = col: col = col + 1
    Dim C_PO_SCH_PRICE As Integer: C_PO_SCH_PRICE = col: col = col + 1
    Dim C_PO_SCH_DUE_DATE As Integer: C_PO_SCH_DUE_DATE = col: col = col + 1
    Dim C_PO_SCH_SHIPTO_ID As Integer: C_PO_SCH_SHIPTO_ID = col: col = col + 1
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer: C_PO_DIST_BUSINESS_UNIT_PC = col: col = col + 1
    Dim C_PO_DIST_PC As Integer: C_PO_DIST_PC = col: col = col + 1
    Dim C_PO_DIST_ACTIVITY_ID As Integer: C_PO_DIST_ACTIVITY_ID = col: col = col + 1
    Dim C_PO_DIST_LOCATION_ID As Integer: C_PO_DIST_LOCATION_ID = col: col = col + 1
    Dim C_PO_NUM As Integer: C_PO_NUM = col: col = col + 1
    Dim C_PO_AMNT_TOTAL As Integer: C_PO_AMNT_TOTAL = col: col = col + 1
    Dim C_LINE_BUDGET_ERR As Integer: C_LINE_BUDGET_ERR = col: col = col + 1
    Dim C_LINE_BUDGET_ERR_FUND_REQ As Integer: C_LINE_BUDGET_ERR_FUND_REQ = col: col = col + 1
    Dim C_PO_ERROR As Integer: C_PO_ERROR = col: col = col + 1
    Dim C_LINE_ERROR As Integer: C_LINE_ERROR = col: col = col + 1
    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_QUEUE_ID, C_QUEUE_ID, HEADER_ROWS_SIZE)
    
    
    If ssMap.PARENT_COUNT = 0 Then
        MsgBox "No POs will be processed: PO count is zero. Please check that each individual PO has a QUEUE_ID assigned (1,2,3,...)", vbInformation
        Exit Sub
    End If
    
    
    Dim POs() As PeopleSoft_PurchaseOrder
    Dim POs_DoNotProcess() As Boolean
    Dim POs_DoNotProcessCount As Integer
    
    ReDim POs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder
    ReDim POs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
           
        If idxParent > 0 Then
            idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
            
            ' First spreadsheet row for this PO -> load PO field values
            If idxChild = 1 Then
                ' PO Fields
                POs(idxParent).PO_Fields.PO_BUSINESS_UNIT = queueTableRange.Cells(curRow, C_PO_BUSINESS_UNIT).Value
                POs(idxParent).PO_Fields.PO_HDR_VENDOR_LOCATION = queueTableRange.Cells(curRow, C_PO_VENDOR_LOCATION).Value
                POs(idxParent).PO_Fields.PO_HDR_BUYER_ID = queueTableRange.Cells(curRow, C_PO_BUYER_ID).Value
                POs(idxParent).PO_Fields.PO_HDR_XPRESS_BID_ID = queueTableRange.Cells(curRow, C_PO_XPRESS_BID_ID).Value
                POs(idxParent).PO_Fields.PO_HDR_QUOTE = queueTableRange.Cells(curRow, C_PO_QUOTE).Value
                POs(idxParent).PO_Fields.PO_HDR_PO_REF = queueTableRange.Cells(curRow, C_PO_REF).Value
                POs(idxParent).PO_Fields.PO_HDR_COMMENTS = queueTableRange.Cells(curRow, C_PO_COMMENTS).Value
                
                ' Default values - note: we only use the contract ID
                POs(idxParent).PO_Defaults.LINE_CONTRACT_ID = queueTableRange.Cells(curRow, C_PO_CONTRACT_ID).Value
                
                ' Budget Check Parameters
                POs(idxParent).BudgetCheck.PRICE_OVERAGE_CODE = queueTableRange.Cells(curRow, C_PO_PRICE_OVERAGE_CODE).Value
                POs(idxParent).BudgetCheck.PRICE_OVERAGE_REASON = queueTableRange.Cells(curRow, C_PO_PRICE_OVERAGE_REASON).Value
                
                ' Allow vendor to be either a number (vendor ID) or the sort vendor name
                If IsNumeric(queueTableRange.Cells(curRow, C_PO_VENDOR).Value) Then
                    POs(idxParent).PO_Fields.PO_HDR_VENDOR_ID = CLng(queueTableRange.Cells(curRow, C_PO_VENDOR).Value)
                Else
                    POs(idxParent).PO_Fields.VENDOR_NAME_SHORT = queueTableRange.Cells(curRow, C_PO_VENDOR).Value
                End If
                
                POs(idxParent).PO_Fields.Quote_Attachment_FilePath = queueTableRange.Cells(curRow, C_QUOTE_ATTACHMENT).Value
            End If
           
            
            Dim poLineNbr As Integer
            
            ' Create PO Line
            poLineNbr = PeopleSoft_PurchaseOrder_AddLine(purchaseOrder:=POs(idxParent), _
                lineItemID:=Trim(CStr(queueTableRange.Cells(curRow, C_PO_LINE_ITEMID).Value)), _
                lineItemDesc:=Trim(CStr(queueTableRange.Cells(curRow, C_PO_LINE_DESC).Value)))
                
           ' Add Schedule to PO Line
            PeopleSoft_PurchaseOrder_AddLineSchedule purchaseOrder:=POs(idxParent), poLineNbr:=poLineNbr, _
               schQty:=CCur(queueTableRange.Cells(curRow, C_PO_SCH_QTY).Value), _
               schPrice:=CCur(queueTableRange.Cells(curRow, C_PO_SCH_PRICE).Value), _
               schDueDate:=CDate(queueTableRange.Cells(curRow, C_PO_SCH_DUE_DATE).Value), _
               schShipToId:=CLng(queueTableRange.Cells(curRow, C_PO_SCH_SHIPTO_ID).Value), _
               distPcBusinessUnit:=CStr(queueTableRange.Cells(curRow, C_PO_DIST_BUSINESS_UNIT_PC).Value), _
               distProjectCode:=CStr(queueTableRange.Cells(curRow, C_PO_DIST_PC).Value), _
               distActivityID:=CStr(queueTableRange.Cells(curRow, C_PO_DIST_ACTIVITY_ID).Value), _
               distLocationID:=CLng(queueTableRange.Cells(curRow, C_PO_DIST_LOCATION_ID).Value)
            


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
        
            
            ' Set output file prefix to include the QUEUE_ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="PO_Q_" & queueTableRange.Cells(ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(1), C_QUEUE_ID).Value & "_"
            
            
            ' new in 2.11: check if quote attachment exists - if not, then error immediately.
            'If POs(idxParent).PO_Fields.Quote_Attachment_FilePath <> "" Then
            '    ' Check if file exists
            '    If Dir(POs(idxParent).PO_Fields.Quote_Attachment_FilePath) = "" Then
            '        result = False
            '        POs(idxParent).HasError = True
            '        POs(idxParent).GlobalError = "File Not Found: " & POs(idxParent).PO_Fields.Quote_Attachment_FilePath
            '    End If
            'End If
            
            result = PeopleSoft_PurchaseOrder_CutPO(session, POs(idxParent))
            
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
                    
                    If POs(idxParent).BudgetCheck.BudgetCheck_HasErrors Then
                        For j = 1 To POs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_LineErrorCount
                            If POs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_LineErrors(j).LINE_NBR = idxChild Then
                                lineBC_HasError = True
                                lineBC_FundReq = POs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_LineErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
                    
                    
                    
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR).Value = IIf(lineBC_HasError, "Y", "")
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR_FUND_REQ).Value = IIf(lineBC_HasError, lineBC_FundReq, "")
                    
                    ' Global Error may have useful info  even though it's not an error
                    If Len(POs(idxParent).GlobalError) > 0 Then queueTableRange.Cells(curRow, C_PO_ERROR).Value = POs(idxParent).GlobalError
                Next idxChild
            Else
                Dim poErrString As String, lineErrString As String
                
                poErrString = IIf(Len(POs(idxParent).GlobalError) > 0, POs(idxParent).GlobalError & vbCrLf, "")
                
                ' PO Field validation results
                With POs(idxParent).PO_Fields
                    If .PO_HDR_PO_REF_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_REF: " & .PO_HDR_PO_REF_Result.ValidationErrorText & vbCrLf
                    If .PO_BUSINESS_UNIT_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT: " & .PO_BUSINESS_UNIT_Result.ValidationErrorText & vbCrLf
                    'If .PO_HDR_APPROVER_ID_Result.ValidationFailed Then _
                    '    poErrString = poErrString & "|" & "PO_HDR_APPROVER_ID: " & .PO_HDR_APPROVER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_BUYER_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_BUYER_ID: " & .PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_ID: " & .PO_HDR_VENDOR_ID_Result.ValidationErrorText & vbCrLf
                    If .VENDOR_NAME_SHORT_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "VENDOR_NAME_SHORT: " & .VENDOR_NAME_SHORT_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_LOCATION: " & .PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_QUOTE_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_QUOTE: " & .PO_HDR_QUOTE_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_XPRESS_BID_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "XPRESS BID ID: " & .PO_HDR_XPRESS_BID_ID_Result.ValidationErrorText & vbCrLf
                        
                    If .Quote_Attachment_FilePath_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "Quote Attachment: " & .Quote_Attachment_FilePath_Result.ValidationErrorText & vbCrLf
                End With
            
                ' PO Default validation results - note: we only use the contract ID
                With POs(idxParent).PO_Defaults
                    If .LINE_CONTRACT_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "CONTRACT ID: " & .LINE_CONTRACT_ID_Result.ValidationErrorText & vbCrLf
                End With
    
            
                ' PO Budget check input validation results
                With POs(idxParent).BudgetCheck
                    If .PRICE_OVERAGE_CODE_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PRICE OVERAGE CODE: " & .PRICE_OVERAGE_CODE_Result.ValidationErrorText & vbCrLf
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
                    
                    queueTableRange.Cells(curRow, C_PO_NUM).Value = IIf(Len(POs(idxParent).PO_ID) > 0, POs(idxParent).PO_ID, "<ERROR>")
                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = poErrString
                    queueTableRange.Cells(curRow, C_LINE_ERROR).Value = lineErrString
                    
                    
                    queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False
                    queueTableRange.Cells(curRow, C_LINE_ERROR).WrapText = False
                    
                Next idxChild
                
                conseqfailCount = conseqfailCount + 1
                
            End If
            
            
            Application.ScreenUpdating = True
            
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
        
        End If
    Next idxParent

    

End Sub
Public Sub Process_PO_Queue_RetryBudgetCheck()

    Call Initialize
    

    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order) - from Process_PO_Queue()
    Dim col As Integer: col = 1

    Dim C_QUEUE_ID As Integer: C_QUEUE_ID = col: col = col + 1
    Dim C_USER_DATA As Integer: C_USER_DATA = col: col = col + 1
    Dim C_PO_BUSINESS_UNIT As Integer: C_PO_BUSINESS_UNIT = col: col = col + 1
    Dim C_PO_VENDOR As Integer: C_PO_VENDOR = col: col = col + 1
    Dim C_PO_VENDOR_LOCATION As Integer: C_PO_VENDOR_LOCATION = col: col = col + 1
    Dim C_PO_CONTRACT_ID As Integer: C_PO_CONTRACT_ID = col: col = col + 1
    Dim C_PO_BUYER_ID As Integer: C_PO_BUYER_ID = col: col = col + 1
    Dim C_PO_XPRESS_BID_ID As Integer: C_PO_XPRESS_BID_ID = col: col = col + 1
    Dim C_PO_QUOTE As Integer: C_PO_QUOTE = col: col = col + 1
    Dim C_QUOTE_ATTACHMENT As Integer: C_QUOTE_ATTACHMENT = col: col = col + 1
    Dim C_PO_REF As Integer: C_PO_REF = col: col = col + 1
    Dim C_PO_COMMENTS As Integer: C_PO_COMMENTS = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_CODE As Integer: C_PO_PRICE_OVERAGE_CODE = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_REASON As Integer: C_PO_PRICE_OVERAGE_REASON = col: col = col + 1
    Dim C_PO_LINE As Integer: C_PO_LINE = col: col = col + 1
    Dim C_PO_LINE_ITEMID As Integer: C_PO_LINE_ITEMID = col: col = col + 1
    Dim C_PO_LINE_DESC As Integer: C_PO_LINE_DESC = col: col = col + 1
    Dim C_PO_SCH_QTY As Integer: C_PO_SCH_QTY = col: col = col + 1
    Dim C_PO_SCH_PRICE As Integer: C_PO_SCH_PRICE = col: col = col + 1
    Dim C_PO_SCH_DUE_DATE As Integer: C_PO_SCH_DUE_DATE = col: col = col + 1
    Dim C_PO_SCH_SHIPTO_ID As Integer: C_PO_SCH_SHIPTO_ID = col: col = col + 1
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer: C_PO_DIST_BUSINESS_UNIT_PC = col: col = col + 1
    Dim C_PO_DIST_PC As Integer: C_PO_DIST_PC = col: col = col + 1
    Dim C_PO_DIST_ACTIVITY_ID As Integer: C_PO_DIST_ACTIVITY_ID = col: col = col + 1
    Dim C_PO_DIST_LOCATION_ID As Integer: C_PO_DIST_LOCATION_ID = col: col = col + 1
    Dim C_PO_NUM As Integer: C_PO_NUM = col: col = col + 1
    Dim C_PO_AMNT_TOTAL As Integer: C_PO_AMNT_TOTAL = col: col = col + 1
    Dim C_LINE_BUDGET_ERR As Integer: C_LINE_BUDGET_ERR = col: col = col + 1
    Dim C_LINE_BUDGET_ERR_FUND_REQ As Integer: C_LINE_BUDGET_ERR_FUND_REQ = col: col = col + 1
    Dim C_PO_ERROR As Integer: C_PO_ERROR = col: col = col + 1
    Dim C_LINE_ERROR As Integer: C_LINE_ERROR = col: col = col + 1
    
    
    
    
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
    
    ' Default state: do not try to save with budget check
    PO_BCs_DoNotProcessCount = ssMap.PARENT_COUNT
    
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
                PO_BCs_DoNotProcessCount = PO_BCs_DoNotProcessCount - 1
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
        
            ' Set output file prefix to include the PO ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="PO_Q_BC_" & PO_BCs(idxParent).PO_ID & "_"
        
        
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
                    
                    If PO_BCs(idxParent).BudgetCheck.BudgetCheck_HasErrors Then
                        For j = 1 To PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_LineErrorCount
                            If PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_LineErrors(j).LINE_NBR = idxChild Then
                                lineBC_HasError = True
                                lineBC_FundReq = PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_LineErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
                    
                    
                    
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR).Value = IIf(lineBC_HasError, "Y", "")
                    queueTableRange.Cells(curRow, C_LINE_BUDGET_ERR_FUND_REQ).Value = IIf(lineBC_HasError, lineBC_FundReq, "")
                    
                    ' Global Error is not an error but may return useful info
                    If Len(PO_BCs(idxParent).GlobalError) > 0 Then queueTableRange.Cells(curRow, C_PO_ERROR).Value = PO_BCs(idxParent).GlobalError
                Next idxChild
            Else
                Dim errStr As String
                
                With PO_BCs(idxParent)
                    errStr = "Budget Check Err: " & IIf(Len(.GlobalError) > 0, .GlobalError, "") & vbCrLf
    
                    If .PO_BU_Result.ValidationFailed Then _
                        errStr = errStr & "|" & "PO BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_ID_Result.ValidationFailed Then _
                        errStr = errStr & "|" & "PO ID: " & .PO_ID_Result.ValidationErrorText & vbCrLf
                End With
                
                ' Apply error to each row
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
        

                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = errStr
                    queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False
                Next idxChild
                
                
                conseqfailCount = conseqfailCount + 1

            End If
            
            
            Application.ScreenUpdating = True
            
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
            
        End If
    Next idxParent




End Sub

Public Sub Process_PO_eQuote_Queue()

    Call Initialize

    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_eQuote_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim col As Integer: col = 1

    Dim C_QUEUE_ID As Integer: C_QUEUE_ID = col: col = col + 1
    Dim C_USER_DATA As Integer: C_USER_DATA = col: col = col + 1
    Dim C_E_QUOTE_NBR As Integer: C_E_QUOTE_NBR = col: col = col + 1
    Dim C_QUOTE_ATTACHMENT As Integer: C_QUOTE_ATTACHMENT = col: col = col + 1
    Dim C_PO_BUSINESS_UNIT As Integer: C_PO_BUSINESS_UNIT = col: col = col + 1
    Dim C_PO_VENDOR_ID As Integer: C_PO_VENDOR_ID = col: col = col + 1
    Dim C_PO_VENDOR_LOCATION As Integer: C_PO_VENDOR_LOCATION = col: col = col + 1
    Dim C_PO_BUYER_ID As Integer: C_PO_BUYER_ID = col: col = col + 1
    Dim C_PO_XPRESS_BID_ID As Integer: C_PO_XPRESS_BID_ID = col: col = col + 1
    Dim C_PO_REF As Integer: C_PO_REF = col: col = col + 1
    Dim C_PO_COMMENTS As Integer: C_PO_COMMENTS = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_CODE As Integer: C_PO_PRICE_OVERAGE_CODE = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_REASON As Integer: C_PO_PRICE_OVERAGE_REASON = col: col = col + 1
    Dim C_PO_DUE_DATE As Integer: C_PO_DUE_DATE = col: col = col + 1
    Dim C_PO_SHIPTO_ID As Integer: C_PO_SHIPTO_ID = col: col = col + 1
    Dim C_PO_BUSINESS_UNIT_PC As Integer: C_PO_BUSINESS_UNIT_PC = col: col = col + 1
    Dim C_PO_PROJECT_CODE As Integer: C_PO_PROJECT_CODE = col: col = col + 1
    Dim C_PO_ACTIVITY_ID As Integer: C_PO_ACTIVITY_ID = col: col = col + 1
    Dim C_PO_LOCATION_ID As Integer: C_PO_LOCATION_ID = col: col = col + 1
    Dim C_PO_LINE As Integer: C_PO_LINE = col: col = col + 1
    Dim C_PO_LINE_ITEMID As Integer: C_PO_LINE_ITEMID = col: col = col + 1
    Dim C_PO_LINE_DESC As Integer: C_PO_LINE_DESC = col: col = col + 1
    Dim C_PO_SCH_DUE_DATE As Integer: C_PO_SCH_DUE_DATE = col: col = col + 1
    Dim C_PO_SCH_SHIPTO_ID As Integer: C_PO_SCH_SHIPTO_ID = col: col = col + 1
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer: C_PO_DIST_BUSINESS_UNIT_PC = col: col = col + 1
    Dim C_PO_DIST_PC As Integer: C_PO_DIST_PC = col: col = col + 1
    Dim C_PO_DIST_ACTIVITY_ID As Integer: C_PO_DIST_ACTIVITY_ID = col: col = col + 1
    Dim C_PO_DIST_LOCATION_ID As Integer: C_PO_DIST_LOCATION_ID = col: col = col + 1
    Dim C_PO_NUM As Integer: C_PO_NUM = col: col = col + 1
    Dim C_PO_AMNT_TOTAL As Integer: C_PO_AMNT_TOTAL = col: col = col + 1
    Dim C_PO_BUDGET_ERR As Integer: C_PO_BUDGET_ERR = col: col = col + 1
    Dim C_PO_BUDGET_ERR_FUND_REQ As Integer: C_PO_BUDGET_ERR_FUND_REQ = col: col = col + 1
    Dim C_PO_ERROR As Integer: C_PO_ERROR = col: col = col + 1
    Dim C_LINE_ERROR As Integer: C_LINE_ERROR = col: col = col + 1

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_QUEUE_ID, C_QUEUE_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then
        MsgBox "No POs will be processed: PO count is zero. Please check that each individual PO has a QUEUE_ID assigned (1,2,3,...)", vbInformation
        Exit Sub
    End If
    
    
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
                'PO_CFQs(idxParent).PO_Fields.PO_HDR_APPROVER_ID = queueTableRange.Cells(curRow, C_PO_APPROVER_ID).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_XPRESS_BID_ID = queueTableRange.Cells(curRow, C_PO_XPRESS_BID_ID).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_PO_REF = queueTableRange.Cells(curRow, C_PO_REF).Value
                PO_CFQs(idxParent).PO_Fields.PO_HDR_COMMENTS = queueTableRange.Cells(curRow, C_PO_COMMENTS).Value
                
                ' Set PO budget check options
                PO_CFQs(idxParent).BudgetCheck.PRICE_OVERAGE_CODE = queueTableRange.Cells(curRow, C_PO_PRICE_OVERAGE_CODE).Value
                PO_CFQs(idxParent).BudgetCheck.PRICE_OVERAGE_REASON = queueTableRange.Cells(curRow, C_PO_PRICE_OVERAGE_REASON).Value
                
                ' Set PO Defaults
                PO_CFQs(idxParent).PO_Defaults.SCH_DUE_DATE = queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value
                ' Note: We assume all distribution fields are capital -> Expense fields will require a line mod
                PO_CFQs(idxParent).PO_Defaults.DIST_CAP_SHIP_TO_ID = queueTableRange.Cells(curRow, C_PO_SHIPTO_ID).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC = queueTableRange.Cells(curRow, C_PO_BUSINESS_UNIT_PC).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_CAP_PROJECT_CODE = queueTableRange.Cells(curRow, C_PO_PROJECT_CODE).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_CAP_ACTIVITY_ID = queueTableRange.Cells(curRow, C_PO_ACTIVITY_ID).Value
                PO_CFQs(idxParent).PO_Defaults.DIST_CAP_LOCATION_ID = queueTableRange.Cells(curRow, C_PO_LOCATION_ID).Value
                
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
                PO_CFQs(idxParent).PO_LineMods(idxChild).PO_Line = -9999  'invalid line ID (negative #): will not process
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
                
            ' Set output file prefix to include the QUEUE_ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="PO_eQuote_Q_" & queueTableRange.Cells(ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(1), C_QUEUE_ID).Value & "_"
            
        
            result = PeopleSoft_PurchaseOrder_CreateFromQuote(session, PO_CFQs(idxParent))
            
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
                    
                    If PO_CFQs(idxParent).BudgetCheck.BudgetCheck_HasErrors Then
                        For j = 1 To PO_CFQs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount
                            If PO_CFQs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j).NOT_COMMIT_AMT > 0 Then
                                BC_totalFundReq = BC_totalFundReq + PO_CFQs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
                    
                    
                    
                    queueTableRange.Cells(curRow, C_PO_BUDGET_ERR).Value = IIf(PO_CFQs(idxParent).BudgetCheck.BudgetCheck_HasErrors, "Y", "")
                    queueTableRange.Cells(curRow, C_PO_BUDGET_ERR_FUND_REQ).Value = IIf(PO_CFQs(idxParent).BudgetCheck.BudgetCheck_HasErrors, BC_totalFundReq, "")
                    
                    
                Next idxChild
            Else
                ' -----------------------------------
                ' Begin - Build error strings and write to spreadsheet
                ' -----------------------------------
                Dim poErrString As String, lineErrString As String
                
                poErrString = IIf(Len(PO_CFQs(idxParent).GlobalError) > 0, PO_CFQs(idxParent).GlobalError & vbCrLf, "")
                
                With PO_CFQs(idxParent).PO_Fields
                    If .PO_HDR_PO_REF_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_REF: " & .PO_HDR_PO_REF_Result.ValidationErrorText & vbCrLf
                    If .PO_BUSINESS_UNIT_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT: " & .PO_BUSINESS_UNIT_Result.ValidationErrorText & vbCrLf
                    'If .PO_HDR_APPROVER_ID_Result.ValidationFailed Then _
                    '    poErrString = poErrString & "|" & "PO_HDR_APPROVER_ID: " & .PO_HDR_APPROVER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_XPRESS_BID_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_XPRESS_BID_ID: " & .PO_HDR_XPRESS_BID_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_BUYER_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_BUYER_ID: " & .PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_ID: " & .PO_HDR_VENDOR_ID_Result.ValidationErrorText & vbCrLf
                    'If .VENDOR_NAME_SHORT_Result.ValidationFailed Then _
                    '    poErrString = poErrString & "|" & "VENDOR_NAME_SHORT: " & .VENDOR_NAME_SHORT_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_HDR_VENDOR_LOCATION: " & .PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText & vbCrLf
                    If .PO_HDR_XPRESS_BID_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "XPRESS_BID_ID: " & .PO_HDR_XPRESS_BID_ID_Result.ValidationErrorText & vbCrLf
                    If .Quote_Attachment_FilePath_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "Quote Attachment: " & .Quote_Attachment_FilePath_Result.ValidationErrorText & vbCrLf
                End With
                With PO_CFQs(idxParent).PO_Defaults
                    If .SCH_DUE_DATE_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PO_DUE_DATE: " & .SCH_DUE_DATE_Result.ValidationErrorText & vbCrLf
                    If .DIST_CAP_SHIP_TO_ID_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PO_SHIPTO_ID: " & .DIST_CAP_SHIP_TO_ID_Result.ValidationErrorText & vbCrLf
                    If .DIST_CAP_BUSINESS_UNIT_PC_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BUSINESS_UNIT_PC: " & .DIST_CAP_BUSINESS_UNIT_PC_Result.ValidationErrorText & vbCrLf
                    If .DIST_CAP_PROJECT_CODE_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_PROJECT_CODE: " & .DIST_CAP_PROJECT_CODE_Result.ValidationErrorText & vbCrLf
                    If .DIST_CAP_ACTIVITY_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_ACTIVITY_ID: " & .DIST_CAP_ACTIVITY_ID_Result.ValidationErrorText & vbCrLf
                    If .DIST_CAP_LOCATION_ID_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PO_LOCATION_ID: " & .DIST_CAP_LOCATION_ID_Result.ValidationErrorText & vbCrLf
                End With
                With PO_CFQs(idxParent).BudgetCheck
                    If .PRICE_OVERAGE_CODE_Result.ValidationFailed Then _
                         poErrString = poErrString & "|" & "PRICE OVERAGE CODE: " & .PRICE_OVERAGE_CODE_Result.ValidationErrorText & vbCrLf
                End With
            
            
                For idxChild = 1 To PO_CFQs(idxParent).PO_LineModCount
                    
                    lineErrString = ""
                    
                    If PO_CFQs(idxParent).PO_LineMods(idxChild).PO_Line > 0 Then
                        With PO_CFQs(idxParent).PO_LineMods(idxChild)
                            lineErrString = IIf(Len(.ItemError) > 0, .ItemError & vbCrLf, "")
                        
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
                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = poErrString
                    queueTableRange.Cells(curRow, C_LINE_ERROR).Value = lineErrString
                    
                    
                    'queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False
                    'queueTableRange.Cells(curRow, C_LINE_ERROR).WrapText = False
                    
                Next idxChild
                ' -----------------------------------
                ' End - Build error strings and write to spreadsheet
                ' -----------------------------------
                
                conseqfailCount = conseqfailCount + 1
            End If
            
            
            Application.ScreenUpdating = True
            
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
            
        End If
    Next idxParent


    
End Sub
Public Sub Process_PO_eQuote_Queue_RetryBudgetCheck()

    Call Initialize
    

    Dim i As Integer, j As Integer
    Dim queueTableRange As Range
    
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_eQuote_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
 
    ' Column numbers (in order) - from Process_PO_eQuote_Queue()
    Dim col As Integer: col = 1


    Dim C_QUEUE_ID As Integer: C_QUEUE_ID = col: col = col + 1
    Dim C_USER_DATA As Integer: C_USER_DATA = col: col = col + 1
    Dim C_E_QUOTE_NBR As Integer: C_E_QUOTE_NBR = col: col = col + 1
    Dim C_QUOTE_ATTACHMENT As Integer: C_QUOTE_ATTACHMENT = col: col = col + 1
    Dim C_PO_BUSINESS_UNIT As Integer: C_PO_BUSINESS_UNIT = col: col = col + 1
    Dim C_PO_VENDOR_ID As Integer: C_PO_VENDOR_ID = col: col = col + 1
    Dim C_PO_VENDOR_LOCATION As Integer: C_PO_VENDOR_LOCATION = col: col = col + 1
    Dim C_PO_BUYER_ID As Integer: C_PO_BUYER_ID = col: col = col + 1
    Dim C_PO_XPRESS_BID_ID As Integer: C_PO_XPRESS_BID_ID = col: col = col + 1
    Dim C_PO_REF As Integer: C_PO_REF = col: col = col + 1
    Dim C_PO_COMMENTS As Integer: C_PO_COMMENTS = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_CODE As Integer: C_PO_PRICE_OVERAGE_CODE = col: col = col + 1
    Dim C_PO_PRICE_OVERAGE_REASON As Integer: C_PO_PRICE_OVERAGE_REASON = col: col = col + 1
    Dim C_PO_DUE_DATE As Integer: C_PO_DUE_DATE = col: col = col + 1
    Dim C_PO_SHIPTO_ID As Integer: C_PO_SHIPTO_ID = col: col = col + 1
    Dim C_PO_BUSINESS_UNIT_PC As Integer: C_PO_BUSINESS_UNIT_PC = col: col = col + 1
    Dim C_PO_PROJECT_CODE As Integer: C_PO_PROJECT_CODE = col: col = col + 1
    Dim C_PO_ACTIVITY_ID As Integer: C_PO_ACTIVITY_ID = col: col = col + 1
    Dim C_PO_LOCATION_ID As Integer: C_PO_LOCATION_ID = col: col = col + 1
    Dim C_PO_LINE As Integer: C_PO_LINE = col: col = col + 1
    Dim C_PO_LINE_ITEMID As Integer: C_PO_LINE_ITEMID = col: col = col + 1
    Dim C_PO_LINE_DESC As Integer: C_PO_LINE_DESC = col: col = col + 1
    Dim C_PO_SCH_DUE_DATE As Integer: C_PO_SCH_DUE_DATE = col: col = col + 1
    Dim C_PO_SCH_SHIPTO_ID As Integer: C_PO_SCH_SHIPTO_ID = col: col = col + 1
    Dim C_PO_DIST_BUSINESS_UNIT_PC As Integer: C_PO_DIST_BUSINESS_UNIT_PC = col: col = col + 1
    Dim C_PO_DIST_PC As Integer: C_PO_DIST_PC = col: col = col + 1
    Dim C_PO_DIST_ACTIVITY_ID As Integer: C_PO_DIST_ACTIVITY_ID = col: col = col + 1
    Dim C_PO_DIST_LOCATION_ID As Integer: C_PO_DIST_LOCATION_ID = col: col = col + 1
    Dim C_PO_NUM As Integer: C_PO_NUM = col: col = col + 1
    Dim C_PO_AMNT_TOTAL As Integer: C_PO_AMNT_TOTAL = col: col = col + 1
    Dim C_PO_BUDGET_ERR As Integer: C_PO_BUDGET_ERR = col: col = col + 1
    Dim C_PO_BUDGET_ERR_FUND_REQ As Integer: C_PO_BUDGET_ERR_FUND_REQ = col: col = col + 1
    Dim C_PO_ERROR As Integer: C_PO_ERROR = col: col = col + 1
    Dim C_LINE_ERROR As Integer: C_LINE_ERROR = col: col = col + 1

    
    
    
    
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
    
    ' Default state: do not try to save with budget check
    PO_BCs_DoNotProcessCount = ssMap.PARENT_COUNT
    
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
            If PO_BCs_DoNotProcess(idxParent) = True And queueTableRange.Cells(curRow, C_PO_NUM).Value <> "" And queueTableRange.Cells(curRow, C_PO_BUDGET_ERR).Value <> "" Then
                PO_BCs_DoNotProcess(idxParent) = False
                PO_BCs_DoNotProcessCount = PO_BCs_DoNotProcessCount - 1
            End If
        End If
    Next curRow
    ' End - Create PO objects and load from spreadsheet
    
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------

    If PO_BCs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No POs in PO eQuote queue has budget check errors.", vbInformation
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
        
            ' Set output file prefix to include the PO ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="PO_eQuote_Q_BC_" & PO_BCs(idxParent).PO_ID & "_"
            
            result = PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck(session, PO_BCs(idxParent))
            
            Application.ScreenUpdating = False
            
            If result Then
                conseqfailCount = 0
                
            
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    
                                       
                    ' Populate budget error
                    Dim BC_totalFundReq As Currency
                    
                    BC_totalFundReq = 0
                    
                    If PO_BCs(idxParent).BudgetCheck.BudgetCheck_HasErrors Then
                        For j = 1 To PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount
                            If PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j).NOT_COMMIT_AMT > 0 Then
                                BC_totalFundReq = BC_totalFundReq + PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j).NOT_COMMIT_AMT
                                
                                ' Actually funding required would be NOT_COMMIT_AMT - AVAIL_BUDGET_AMT
                            End If
                        Next j
                    End If
        
                    
                    queueTableRange.Cells(curRow, C_PO_BUDGET_ERR).Value = IIf(PO_BCs(idxParent).BudgetCheck.BudgetCheck_HasErrors, "Y", "")
                    queueTableRange.Cells(curRow, C_PO_BUDGET_ERR_FUND_REQ).Value = IIf(PO_BCs(idxParent).BudgetCheck.BudgetCheck_HasErrors, BC_totalFundReq, "")
                    
                    ' Global Error is not an error but may return useful info
                    If Len(PO_BCs(idxParent).GlobalError) > 0 Then queueTableRange.Cells(curRow, C_PO_ERROR).Value = PO_BCs(idxParent).GlobalError
                Next idxChild
            Else
                Dim errStr As String
                
                With PO_BCs(idxParent)
                    errStr = "Budget Check Err: " & IIf(Len(.GlobalError) > 0, .GlobalError, "") & vbCrLf
    
                    If .PO_BU_Result.ValidationFailed Then _
                        errStr = errStr & "|" & "PO BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_ID_Result.ValidationFailed Then _
                        errStr = errStr & "|" & "PO ID: " & .PO_ID_Result.ValidationErrorText & vbCrLf
                End With
            
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_PO_ERROR).Value = errStr
                    queueTableRange.Cells(curRow, C_PO_ERROR).WrapText = False

                Next idxChild
                
                
                conseqfailCount = conseqfailCount + 1
            End If
            
        
            Application.ScreenUpdating = True
            
        
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
        End If
    Next idxParent




End Sub

Public Sub Process_PO_BudgetCheck_Queue()

    Call Initialize

    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_BudgetCheck_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim col As Integer: col = 1
    
    Dim C_PO_BU As Integer: C_PO_BU = col: col = col + 1
    Dim C_PO_ID As Integer: C_PO_ID = col: col = col + 1
    Dim C_BC_RESULT As Integer: C_BC_RESULT = col: col = col + 1
    Dim C_BC_FUND_REQ As Integer: C_BC_FUND_REQ = col: col = col + 1
    Dim C_BC_ERROR As Integer: C_BC_ERROR = col: col = col + 1

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_QUEUE_ID (Unique per PO)
    ' Secondary - PO Line #
    Dim ssMap As SPREADSHEET_MAP_1L
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_1D(queueTableRange, C_PO_ID, C_PO_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then
        MsgBox "No POs will be processed: PO count is zero. Please check the PO ID field and ensure it is filled out.", vbInformation
        Exit Sub
    End If
    
    
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
        MsgBox "No POs in budget check queue has budget has blank result. Clear errors and try again.", vbInformation
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
        
            ' Set output file prefix to include the PO ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="PO_BC_Q_" & PO_BCs(idxParent).PO_ID & "_"
        
            result = PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck(session, PO_BCs(idxParent))
            
            Application.ScreenUpdating = False
            
            
            
            If result Then
                conseqfailCount = 0
                
                Dim bcFundReq As String
                
                If PO_BCs(idxParent).BudgetCheck.BudgetCheck_HasErrors Then
                    bcFundReq = ""
                    
                    For j = 1 To PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount
                        With PO_BCs(idxParent).BudgetCheck.BudgetCheck_Errors.BudgetCheck_ProjectErrors(j)
                            bcFundReq = "," & .PROJECT_ID & ":" & .NOT_COMMIT_AMT
                        End With
                    Next j
                    
                    If Len(bcFundReq) > 0 Then bcFundReq = Mid(bcFundReq, Len(",") + 1) ' Remove , at beginning
                End If
                
                
                curRow = ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW
                
                queueTableRange.Cells(curRow, C_BC_RESULT).Value = IIf(PO_BCs(idxParent).BudgetCheck.BudgetCheck_HasErrors, "FAIL", "PASS")
                queueTableRange.Cells(curRow, C_BC_FUND_REQ).Value = bcFundReq
            
                queueTableRange.Cells(curRow, C_BC_ERROR).Value = PO_BCs(idxParent).GlobalError ' Not an error but may have some info
            Else
                curRow = ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW
                
                
                Dim errStr As String
                
                With PO_BCs(idxParent)
                    errStr = "Budget Check Err: " & IIf(Len(.GlobalError) > 0, .GlobalError, "") & vbCrLf
    
                    If .PO_BU_Result.ValidationFailed Then _
                        errStr = errStr & "|" & "PO BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_ID_Result.ValidationFailed Then _
                        errStr = errStr & "|" & "PO ID: " & .PO_ID_Result.ValidationErrorText & vbCrLf
                End With
                
                queueTableRange.Cells(curRow, C_BC_RESULT).Value = "<ERROR>"

                queueTableRange.Cells(curRow, C_BC_ERROR).Value = errStr
                queueTableRange.Cells(curRow, C_BC_ERROR).WrapText = False

                conseqfailCount = conseqfailCount + 1
            End If
            
            
            Application.ScreenUpdating = True
            
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
        
        End If
    Next idxParent



    
    Debug.Print

End Sub
Public Sub Process_PO_Receipt_Queue()

    Call Initialize


    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    ' Primary - PO_ID (Unique per PO) - defines each Rcpt
    ' Secondary - PO Line #/Schedule # - defines each Rcpt.Item
    Const HEADER_ROWS_SIZE = 0
    
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_Receipt_Q").Range("4:65536")
    
    ' Column numbers (in order)
    Dim col As Integer: col = 1

    Dim C_USER_DATA As Integer: C_USER_DATA = col: col = col + 1
    Dim C_PO_BU As Integer: C_PO_BU = col: col = col + 1
    Dim C_PO_ID As Integer: C_PO_ID = col: col = col + 1
    Dim C_PO_LINE As Integer: C_PO_LINE = col: col = col + 1
    Dim C_PO_SCH As Integer: C_PO_SCH = col: col = col + 1
    Dim C_RECEIVE_QTY As Integer: C_RECEIVE_QTY = col: col = col + 1
    Dim C_ACCEPT_QTY As Integer: C_ACCEPT_QTY = col: col = col + 1
    Dim C_ITEM_ID As Integer: C_ITEM_ID = col: col = col + 1
    Dim C_CATS_FLAG As Integer: C_CATS_FLAG = col: col = col + 1
    Dim C_TRANS_ITEM_DESC As Integer: C_TRANS_ITEM_DESC = col: col = col + 1
    Dim C_RCPT_ID As Integer: C_RCPT_ID = col: col = col + 1
    Dim C_RCPT_ERR As Integer: C_RCPT_ERR = col: col = col + 1
    Dim C_RCPT_ITEM_ERR As Integer: C_RCPT_ITEM_ERR = col: col = col + 1
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - Rcpt
    ' Secondary - Rcpt_Item
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    Dim i As Long
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_PO_ID, C_PO_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then
        MsgBox "No receipts will be processed. Please check the PO ID field and ensure it is filled out.", vbInformation
        Exit Sub
    End If
    
    
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
        
            ' Set output file prefix to include the PO ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="Receipt_Q_" & Rcpts(idxParent).PO_ID & "_"
        
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
                        If Len(Rcpts(idxParent).GlobalError) > 0 Then queueTableRange.Cells(curRow, C_RCPT_ERR).Value = Rcpts(idxParent).GlobalError
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
                    If Len(Rcpts(idxParent).GlobalError) > 0 Then queueTableRange.Cells(curRow, C_RCPT_ERR).Value = Rcpts(idxParent).GlobalError
                Else
                    Err.Raise -1, , "Unknown Receive Mode (This should never happen)"
                End If
                
                conseqfailCount = 0
            Else
                
                errString = ""
                
                With Rcpts(idxParent)
                    If .PO_BU_Result.ValidationFailed Then _
                        errString = errString & "|" & "PO_BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_ID_Result.ValidationFailed Then _
                        errString = errString & "|" & "PO_ID: " & .PO_ID_Result.ValidationErrorText & vbCrLf
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
                
            End If
            
            
            ' update 2.10.1: include individual receive errorsonly if RECEIVE_SPECIFIED is set
            If Rcpts(idxParent).ReceiveMode = RECEIVE_SPECIFIED Then
                For idxChild = 1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT
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
            
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
            
        End If
    Next idxParent



End Sub
Public Sub Process_PO_ChangeOrder_Q()


    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_ChangeOrder_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim col As Integer: col = 1
    
    Dim C_PO_BU As Integer: C_PO_BU = col: col = col + 1
    Dim C_PO_ID As Integer: C_PO_ID = col: col = col + 1
    Dim C_PO_LINE As Integer: C_PO_LINE = col: col = col + 1
    Dim C_PO_SCHEDULE As Integer: C_PO_SCHEDULE = col: col = col + 1
    Dim C_PO_DUE_DATE As Integer: C_PO_DUE_DATE = col: col = col + 1
    Dim C_PO_FLG_SEND_TO_VENDOR As Integer: C_PO_FLG_SEND_TO_VENDOR = col: col = col + 1
    Dim C_CO_REASON As Integer: C_CO_REASON = col: col = col + 1
    Dim C_CO_STATUS As Integer: C_CO_STATUS = col: col = col + 1
    Dim C_CO_PO_ERROR As Integer: C_CO_PO_ERROR = col: col = col + 1
    Dim C_CO_ITEM_ERROR As Integer: C_CO_ITEM_ERROR = col: col = col + 1

    
    
    
    ' ----------------------------------------------------------------------------
    ' Begin - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------
    ' Primary - PO_ID
    ' Secondary - PO Line #
    Dim idxParent As Integer, idxChild As Integer
    Dim curRow As Integer
    
    ssMap = SpreadsheetTableToMultiLevelMap_2D(queueTableRange, C_PO_ID, C_PO_ID, HEADER_ROWS_SIZE)
    
    If ssMap.PARENT_COUNT = 0 Then Exit Sub
    
    
    Dim PO_COs() As PeopleSoft_PurchaseOrder_ChangeOrder
    Dim PO_COs_DoNotProcess() As Boolean
    Dim PO_COs_DoNotProcessCount As Integer
    
    ReDim PO_COs(1 To ssMap.PARENT_COUNT) As PeopleSoft_PurchaseOrder_ChangeOrder
    ReDim PO_COs_DoNotProcess(1 To ssMap.PARENT_COUNT) As Boolean
    

    For curRow = 1 + HEADER_ROWS_SIZE To ssMap.ROW_COUNT + HEADER_ROWS_SIZE
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
           
        If idxParent > 0 Then
            idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
            
            
            If idxChild = 1 Then
                PO_COs(idxParent).PO_BU = queueTableRange.Cells(curRow, C_PO_BU).Value
                PO_COs(idxParent).PO_ID = queueTableRange.Cells(curRow, C_PO_ID).Value
                
                
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
                
                ' Set changes to entire PO only if the PO line isn't specified
                If IsEmpty(queueTableRange.Cells(curRow, C_PO_LINE).Value) Then
                    If Not IsEmpty(queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value) Then PO_COs(idxParent).PO_DUE_DATE = queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value
                End If

                
                PO_COs(idxParent).PO_ChangeOrder_ItemCount = ssMap.PARENT_MAP(idxParent).CHILD_COUNT
                
                ReDim PO_COs(idxParent).PO_ChangeOrder_Items(1 To PO_COs(idxParent).PO_ChangeOrder_ItemCount) As PeopleSoft_PurchaseOrder_ChangeOrder_Item
            End If
            
            If queueTableRange.Cells(curRow, C_PO_LINE).Value <> "" Then
                PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).PO_Line = CInt(queueTableRange.Cells(curRow, C_PO_LINE).Value)
                PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).PO_Schedule = CInt(queueTableRange.Cells(curRow, C_PO_SCHEDULE).Value)
                
                PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).SCH_DUE_DATE = queueTableRange.Cells(curRow, C_PO_DUE_DATE).Value
            Else
                PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).PO_Line = -9999 ' No line given -> set as invalid line so it will be treated as invalid
            End If


            ' if there text in the Status column, do not process entire PO change order
            If PO_COs_DoNotProcess(idxParent) = False And queueTableRange.Cells(curRow, C_CO_STATUS).Value <> "" Then
                PO_COs_DoNotProcess(idxParent) = True
                PO_COs_DoNotProcessCount = PO_COs_DoNotProcessCount + 1
            End If
        End If
    Next curRow
    ' End - Create PO change order objects and load from spreadsheet
    
    ' ----------------------------------------------------------------------------
    ' End - Load data from spreadsheet into data structures
    ' ----------------------------------------------------------------------------


    If PO_COs_DoNotProcessCount = ssMap.PARENT_COUNT Then
        MsgBox "No PO change orders will be processed: clear any any errors and try again", vbInformation
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
        If PO_COs_DoNotProcess(idxParent) = False Then
        
        
            ' Set output file prefix to include the PO ID
            PeopleSoft_SetDebugOutputOptions filePrefix:="CO_Q_" & PO_COs(idxParent).PO_ID & "_"
        
            result = PeopleSoft_ChangeOrder_Process(session, PO_COs(idxParent))
            
            Application.ScreenUpdating = False
            
            If result Then
                conseqfailCount = 0
            
                For idxChild = 1 To PO_COs(idxParent).PO_ChangeOrder_ItemCount
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    
                    queueTableRange.Cells(curRow, C_CO_STATUS).Value = "COMPLETE"
                    
                    ' Global error may contain useful into even though it may not be an actual error
                    If Len(PO_COs(idxParent).GlobalError) > 0 Then queueTableRange.Cells(curRow, C_CO_PO_ERROR).Value = PO_COs(idxParent).GlobalError
                Next idxChild
            Else
                ' -----------------------------------
                ' Begin - Build error strings and write to spreadsheet
                ' -----------------------------------
                Dim poErrString As String, itemErrString As String
                
                poErrString = IIf(Len(PO_COs(idxParent).GlobalError) > 0, PO_COs(idxParent).GlobalError & vbCrLf, "")
                
                With PO_COs(idxParent)
                    If .PO_BU_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_BU: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_ID_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_ID: " & .PO_BU_Result.ValidationErrorText & vbCrLf
                    If .PO_DUE_DATE_Result.ValidationFailed Then _
                        poErrString = poErrString & "|" & "PO_DUE_DATE: " & .PO_DUE_DATE_Result.ValidationErrorText & vbCrLf
                End With
            
                For idxChild = 1 To PO_COs(idxParent).PO_ChangeOrder_ItemCount
                    itemErrString = ""
                    
                    If PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).PO_Line > 0 Then
                        With PO_COs(idxParent).PO_ChangeOrder_Items(idxChild)
                            itemErrString = .ItemError & vbCrLf
                        
                            If .SCH_DUE_DATE_Result.ValidationFailed Then _
                                 itemErrString = itemErrString & "|" & "SCH_DUE_DATE: " & .SCH_DUE_DATE_Result.ValidationErrorText & vbCrLf
                        End With
                    End If
            
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_CO_STATUS).Value = "<ERROR>"
                    queueTableRange.Cells(curRow, C_CO_PO_ERROR).Value = poErrString
                    queueTableRange.Cells(curRow, C_CO_ITEM_ERROR).Value = itemErrString
                    
                    
                    queueTableRange.Cells(curRow, C_CO_PO_ERROR).WrapText = False
                    queueTableRange.Cells(curRow, C_CO_ITEM_ERROR).WrapText = False
                Next idxChild
                ' -----------------------------------
                ' End - Build error strings and write to spreadsheet
                ' -----------------------------------
                
                conseqfailCount = conseqfailCount + 1
            End If
            
            
            Application.ScreenUpdating = True
            
                
            If conseqfailCount > CONFIG_OPTIONS.Q_MAX_CONSECUTIVE_ERRORS Then Exit Sub
        
        End If
    Next idxParent


    
End Sub


Public Function Prompt_UserPass(ByRef user As String, ByRef pass As String) As Boolean
    


    user = InputBox("Enter username (USWIN):", , Environ$("username"))
    
    If Len(user) = 0 Then Prompt_UserPass = False: Exit Function
    
    pass = InputBoxDK("Enter password:", "")
    
    If Len(pass) = 0 Then Prompt_UserPass = False: Exit Function
    
    
    Prompt_UserPass = True

End Function

