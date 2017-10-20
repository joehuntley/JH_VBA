Attribute VB_Name = "JH_VZW_PeopleSoft_Automation"
Option Explicit

' JH_VZW_PeopleSoft_Automation
' ------------------------------------------------------------------------------------------------------------------
' PeopleSoft Automation Module using SeleniumBasic/VBA COM library. This module is designed to be a stand-alone
' module (decoupled from the workbook). As a result, this module may be used in any other VBA project provided the
' SeleniumBasic/VBA COM library is installed.
'
' Notes for other developers: If you have made enhancements and fixes to the codebase, please record it below and send
' me the new module. I will be happy to merge any new changes or breakfixes and provide an updated module.
'
' I have purposely kept this VBA project unlocked without any password protection so that others may learn, improve,
' and re-use the code in new projects. I ask that any developer who uses this module does the same.
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------
' Changelog:
'
' 2.12.0
'   2017-10-19 (joe h)
'                         - PO Create: Added infinite timeout time after saving with budget check
'                         - PO Create: PO Defaults now behaves differently internally. The calling application may set default values
'                           via the PO_Defaults internal structure. The auto-calculated defaults only apply when the default values
'                           are not explicitly provided.
'                         - PO Defaults now has optional Contract ID field.
'                         - PO Create: Budget Check now includes inputs for overage justification and reason.
'                         - PO Create: Add Contract ID field to PO Defaults
'                         - PO Create: Add Overage Justification Fields (Reason Code, Reason) to BudgetCheck fields. Returns Reason if not populated.
'                         - PO CreateFromQuote: Added XPRESS BID ID field
'                         - PO CreateFromQuote: Fixed issue to handle cases where selecting create from quote will only sometimes trigger the onchange event.
'                         - PO CreateFromQuote: Added ItemError to line modifications
'                         - PO Change Order: Allow changing individual line/schedules
'                         - PO Change Order: Deal with case when the PO opens up starting from the budget error screen.
'                         - Retry Budget Check: Added validated field PO ID
'                         - PO Fields are now validated fields: PO_REF, QUOTE, XPRESS_BID_ID
'                         - PO Budget Check, Receipt, and ChangeOrder fields are now validated: PO BU, PO ID
'                         - Save with Budget Check now uses infinite timeouts (affects multiple features)
'                         - PeopleSoft_NavigateTo_ExistingPO now checks if landing screen is the budget check screen. If so, it returns to the PO.
' 2.11.1
'   2017-10-13 (joe h)
'                         - Added PreCheck for all automation procedures to pre-validate fields
'                         - Add debug mode option to quit before saving (stops processing before making final changes to PO/Receipt/Etc..)
'                         - Added functions PeopleSoft_PurchaseOrder_NewPO, PeopleSoft_PurchaseOrder_NewLine, PeopleSoft_PurchaseOrder_NewSchedule
'                         - Renamed PeopleSoft_PurchaseOrder_AddLineSimple -> PeopleSoft_PurchaseOrder_NewSimpleLine
'                         - PeopleSoft_Page_SuppressPopup: Removed (deprecated)
'                         - PeopleSoft_Page_WaitForProcessing: Added infinite timeout option
'                         - PeopleSoft_Page_CheckForPopup/PeopleSoft_Page_Acknowledge: Added pass through acknowledgeTimeout option
'                         - PO Create: Added infinite timeout time after saving with budget check
'                         - PO CreateFromQuote: Added ItemError to line modifications
'                         - PO Change Order: Allow changing individual line/schedules
'                         - PO Change Order: Added infinite timeout time after saving change reason
'                         - PO Change Order: Added check for popup after saving change reason
'                         - PO Change Order: Added Line Status of Open as valid before moving forward with change order.
'                         - Retry Budget Check: Added validated field PO ID
'                         - Removed PO_Defaults from PO structure (not needed)
'                         - PO Fields are now validated fields: PO_REF, QUOTE, XPRESS_BID_ID
'                         - PO Budget Check, Receipt, and ChangeOrder fields are now validated: PO BU, PO ID
'                         - PO Create now returns fields: PROJECT DESC
'                         - Added PeopleSoft_SetDebugOutputOptions to add file prefix
' 2.11.0
'   2017-10-03 (joe h)    - PO Change Order: Exits with error if PO defaults nor line item changed (modifyDefault=False and no valid line items)
'                         - PO Create from Quote: Add Retry Save with Budget Check
'                         - PO_RetryWithBudget Check: Fixed issue due to new columns
'                         - PO Create: Fixed issue where PO ID not found page.
'   2017-10-02 (joe h)    - PeopleSoft_Page_SetValidatedField overrides implicitWait (speeds up entire automation)
'                         - PO Receipt: Fixed type mismatch when Receipt Price on page is empty ($0 items)
'                         - PO Create (eQuote): Ignores spaces before/after eQuote # match checked
'                         - PeopleSoft_Page_SetValidatedField: Checks if element is disabled before setting
'                         - Converted Qty data types to Currency (fixed-point)
'                         - PO Create (CutPO): Vendor can be either set using their ID or SHORT NAME
'                         - PO Create: Removed suggested approver/approver ID from PO Fields (auto-populates)
'                         - PO Create (CutPO): Fill xPress Bid Field ID
'                         - PO Create: Expense PO issue now fixed. PO defaults calculated separately for expense chartfields.
'                         - Added Debug calls to help with debugging
'   2017-09-26 (joe h)    - PO Creation/Modify: Returns valid activity IDs when provided activity ID is invalid (adds to error message). Now fixed.
'                         - Added PeopleSoft_Page_ModalWindow_ExtractSearchTableContents: Generic utility function to read search tables in PS prompts (e.g., valid values for specific fields)
'   2017-09-22 (joe h)    - Major overhaul to Receipt_Create (formerly PurchaseOrder_ProcessReceipt). Page extraction code moved to their own functions
'                             - Receipt Lines now matched by line/schedule. (No more errors when they are out-of-order)
'                             - After receipt, has these been checked for accurancy is valid popup is ignored.
'                             - If valid receipt ID is generated, automatically acknowledges each popup regardless of message
'                             - Checks for duplicate PO Line/schedules before running (breaks the code)
'                             - Continues receiving on items even if one line is not receivable. The error is reported in the Receipt Item error
'                         - PO Creation Updates (CutPO and CreateFromQuote)
'                             - Quote attachment feature added
'                             - Excess Available is acknowledged and ignored
'                             - Create from eQuote PS issues fixed
'                             - Clicks through warning when due date is selected: Due Date selected is a weekend or a public holiday
'                             - Save with Budget Check: Increased timeout waiting period
'                             - Ignores popup if line amount is $0
'                         - Added PeopleSoft_Page_CheckForPopup. Supercedes SuppressPopup(): Checks for buttons, auto-acknowledges, and checks for expected text
'   2017-09-20 (joe h)    - PurchaseOrder_ProcessReceipt: Fixed bug: Subscript out of range error (Caused by previous receipt PO having line item count > current PO)
'
' 2.10.2
'               (joe h)    - PurchaseOrder_ProcessReceipt - If a valid receipt ID is created, automatically ackowledges each popup regardless of the message.
'
' 2.10.1
'               (joe h)    - Added chromedriver upgrade instructions to README
'               (joe h)    - PurchaseOrder_ProcessReceipts: Allows receiving of all lines/schedule using new ReceiveMode variable.
'
' 2.10.0
'               (henry c/oscar g) - Updates & Fixes due to new PS upgrade
'
' 2.9.1.3
'               (joe h)   - PS Upgrade Issue - PurchaseOrder_ProcessReceipts fixed.
'
' 2.9.1.2
'               (joe h)   - PS Upgrade Issues - Fixed issues when suppressing popup windows. This is has caused issues when creating change orders
'
' 2.9.1.1
'               (joe h)   - PS Upgrade Issue - Internal PS JS procedure hAction_win0() is no longer valid. This has caused many issues including  not allowing the
'                           creation of multi-line Pos, change orders, etc… All active hAction_win0 calls have been replaced with their submitAction_win0()
'                           equivalents. Commented out lines still need to be updated if re-used
'               (joe h)   - PS Upgrade Issue - Additional: PO ID extraction error when during saving with budget check
'
' 2.9.1
'               (joe h)   - PS Upgrade Issue - Fixed error when searching for popup containers
'
' 2.9.0
'               (joe h)   - PO_Q: When activity IDs are invalid, a list of valid activity IDs are returned in error message - Feature not working
'
' 2.8.7
'               (joe h)   - PO_Q:  Quantity no longer has to be a whole number
'                         - New Feature: PO_Receipt Q: Specify receive quantity
'
' ------------------------------------------------------------------------------------------------------------------
' Feature Requests & TODO:
'   - PO Change Order: Allow adjusting by each line
'   - PO Create: Add option to provide approval justification comments

' ------------------------------------------------
' General
' ------------------------------------------------
Type PeopleSoft_Session
    driver As SeleniumWrapper.WebDriver
    
    user As String
    pass As String
    loggedIn As Boolean
    
    LogonError As String
End Type


Type PeopleSoft_Field_ValidationResult
    ValidationFailed As Boolean
    ValidationErrorText As String
End Type


Enum PeopleSoft_Page_CheckboxAction
    KeepExistingValue = 0
    SetAsChecked = 1
    SetAsUnchecked = 2
End Enum


Type PeopleSoft_Page_PopupCheckResult
    HasPopup As Boolean
    popupText As String
    PopupElementID As String
    IsExpected As Boolean
    HasButtonOk As Boolean
    HasButtonCancel As Boolean
    HasButtonYes As Boolean
    HasButtonNo As Boolean
End Type

' ------------------------------------------------
' PO Fields
' ------------------------------------------------
Type PeopleSoft_PurchaseOrder_Fields
    ' PSoft Fields
    PO_BUSINESS_UNIT As String
    
    VENDOR_NAME_SHORT As String
    PO_HDR_VENDOR_ID As Long
    PO_HDR_VENDOR_LOCATION As String
    PO_HDR_BUYER_ID As Long
    'PO_HDR_APPROVER_ID As Long
    
    PO_HDR_PO_REF As String  ' NOTE: MAX LENGTH: 30 CHARS
    PO_HDR_COMMENTS As String
    PO_HDR_QUOTE As String
    PO_HDR_XPRESS_BID_ID As String
    
    Quote_Attachment_FilePath As String
    
    ' Field Validation Results
    PO_BUSINESS_UNIT_Result As PeopleSoft_Field_ValidationResult
    VENDOR_NAME_SHORT_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_VENDOR_ID_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_VENDOR_LOCATION_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_BUYER_ID_Result As PeopleSoft_Field_ValidationResult
    'PO_HDR_APPROVER_ID_Result As PeopleSoft_Field_ValidationResult
    
    PO_HDR_PO_REF_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_QUOTE_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_XPRESS_BID_ID_Result As PeopleSoft_Field_ValidationResult
    
    Quote_Attachment_FilePath_Result As PeopleSoft_Field_ValidationResult
    
    HasValidationError As Boolean
End Type


Type PeopleSoft_PurchaseOrder_Line_Fields
    PO_LINE_ITEM_ID As String
    PO_LINE_DESC As String
    
    PO_LINE_ITEM_ID_Result As PeopleSoft_Field_ValidationResult
End Type

Type PeopleSoft_PurchaseOrder_Schedule_Fields
    ' PSoft Fields
    DUE_DATE As Date
    SHIPTO_ID As Long
    QTY As Currency ' Use Currency as we must use a fixed point data type here
    PRICE As Currency
    
    
    ' Field Validation Results
    DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    SHIPTO_ID_Result As PeopleSoft_Field_ValidationResult
    QTY_Result As PeopleSoft_Field_ValidationResult
    PRICE_Result As PeopleSoft_Field_ValidationResult
End Type


Type PeopleSoft_PurchaseOrder_PO_Defaults
    HasGlobalError As Boolean
    GlobalError As String
    
    HasValidationError As Boolean
    
    LINE_CONTRACT_ID As String

    SCH_DUE_DATE As Date
    
    DIST_CAP_BUSINESS_UNIT_PC As String
    DIST_CAP_PROJECT_CODE As String
    DIST_CAP_ACTIVITY_ID As String
    DIST_CAP_SHIP_TO_ID As Long
    DIST_CAP_LOCATION_ID As Long
    
    DIST_EXP_BUSINESS_UNIT_PC As String
    DIST_EXP_PROJECT_CODE As String
    DIST_EXP_ACTIVITY_ID As String
    DIST_EXP_SHIP_TO_ID As Long
    DIST_EXP_LOCATION_ID As Long
    
    ' Return only fields
    DIST_CAP_PROJECT_DESC As String
    DIST_EXP_PROJECT_DESC As String
    
    ' Validation Results
    LINE_CONTRACT_ID_Result As PeopleSoft_Field_ValidationResult
    
    SCH_DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    
    DIST_CAP_BUSINESS_UNIT_PC_Result As PeopleSoft_Field_ValidationResult
    DIST_CAP_PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    DIST_CAP_ACTIVITY_ID_Result As PeopleSoft_Field_ValidationResult
    DIST_CAP_SHIP_TO_ID_Result As PeopleSoft_Field_ValidationResult
    DIST_CAP_LOCATION_ID_Result As PeopleSoft_Field_ValidationResult
    
    DIST_EXP_BUSINESS_UNIT_PC_Result As PeopleSoft_Field_ValidationResult
    DIST_EXP_PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    DIST_EXP_ACTIVITY_ID_Result As PeopleSoft_Field_ValidationResult
    DIST_EXP_SHIP_TO_ID_Result As PeopleSoft_Field_ValidationResult
    DIST_EXP_LOCATION_ID_Result As PeopleSoft_Field_ValidationResult
End Type


Type PeopleSoft_PurchaseOrder_Distribution_Fields
    ' PSoft Fields
    BUSINESS_UNIT_PC As String
    PROJECT_CODE As String
    ACTIVITY_ID As String
    LOCATION_ID As Long
    
    ' PSoft Return Only
    PROJECT_DESC As String
    
    ' Field Validation Results
    BUSINESS_UNIT_PC_Result As PeopleSoft_Field_ValidationResult
    PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    ACTIVITY_ID_Result As PeopleSoft_Field_ValidationResult
    LOCATION_ID_Result As PeopleSoft_Field_ValidationResult
End Type

Type PeopleSoft_PurchaseOrder_Schedule
    ScheduleFields As PeopleSoft_PurchaseOrder_Schedule_Fields
    DistributionFields As PeopleSoft_PurchaseOrder_Distribution_Fields
End Type

Type PeopleSoft_PurchaseOrder_Line
    LineFields As PeopleSoft_PurchaseOrder_Line_Fields
    
    Schedules() As PeopleSoft_PurchaseOrder_Schedule
    ScheduleCount As Integer
    
    HasValidationError As Boolean
End Type


Type PeopleSoft_PurchaseOrder_BudgetCheck_LineError
    LINE_NBR As Integer
    SCHED_NBR As Integer
    DISTRIB_LINE_NUM As Integer
    BUDGET_DT As String
    BUSINESS_UNIT_PC As String
    PROJECT_ID As String
    LINE_AMOUNT As Currency
    COMMIT_AMT As Currency
    NOT_COMMIT_AMT As Currency
    AVAIL_BUDGET_AMT As Currency
End Type

Type PeopleSoft_PurchaseOrder_BudgetCheck_ProjectError
    BUSINESS_UNIT_PC As String
    PROJECT_ID As String
    NOT_COMMIT_AMT As Currency
    AVAIL_BUDGET_AMT As Currency
    FUNDING_NEEDED As Currency
End Type

Type PeopleSoft_PurchaseOrder_BudgetCheckErrors
    BudgetCheck_LineErrors() As PeopleSoft_PurchaseOrder_BudgetCheck_LineError
    BudgetCheck_LineErrorCount As Integer
    
    BudgetCheck_ProjectErrors() As PeopleSoft_PurchaseOrder_BudgetCheck_ProjectError
    BudgetCheck_ProjectErrorCount As Integer
End Type

Type PeopleSoft_PurchaseOrder_BudgetCheckParams
    BudgetCheck_HasErrors As Boolean
    BudgetCheck_Errors As PeopleSoft_PurchaseOrder_BudgetCheckErrors
    
    PO_ID As String
    
    ' Inputs: Price overages
    PRICE_OVERAGE_CODE As String
    PRICE_OVERAGE_REASON As String
    
    PRICE_OVERAGE_CODE_Result As PeopleSoft_Field_ValidationResult
    
    HasGlobalError As Boolean
    GlobalError As String
End Type

Type PeopleSoft_PurchaseOrder
    PO_ID As String
     
    PO_AMNT_FTM_TOTAL As Currency
    PO_AMNT_TOTAL As Currency
    PO_AMNT_MERCH_TOTAL As Currency
    
    PO_Fields As PeopleSoft_PurchaseOrder_Fields
    PO_Lines() As PeopleSoft_PurchaseOrder_Line
    PO_LineCount As Integer
    
    
    PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults
    
    HasError As Boolean
    GlobalError As String
    
    BudgetCheck As PeopleSoft_PurchaseOrder_BudgetCheckParams
    
End Type


' ------------------------------------------------
' PO Create From Quote Fields
' ------------------------------------------------
Type PeopleSoft_PurchaseOrder_CreateFromQuote_LineModification
    PO_Line As Integer
    
    
    PO_LINE_ITEM_ID As String
    PO_LINE_DESC As String
    
    SCH_DUE_DATE As Date
    SCH_SHIPTO_ID As Long
    
    
    DIST_BUSINESS_UNIT_PC As String
    DIST_PROJECT_CODE As String
    DIST_ACTIVITY_ID As String
    DIST_LOCATION_ID As Long
    
    PO_LINE_ITEM_ID_Result As PeopleSoft_Field_ValidationResult
    SCH_DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    SCH_SHIPTO_ID_Result As PeopleSoft_Field_ValidationResult

    DIST_BUSINESS_UNIT_PC_Result As PeopleSoft_Field_ValidationResult
    DIST_PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    DIST_ACTIVITY_ID_Result As PeopleSoft_Field_ValidationResult
    DIST_LOCATION_ID_Result As PeopleSoft_Field_ValidationResult
    
    HasValidationError As Boolean
    ItemError As String
End Type

Type PeopleSoft_PurchaseOrder_CreateFromQuoteParams
    PO_ID As String

    E_QUOTE_NBR As String
    E_QUOTE_NBR_Result As PeopleSoft_Field_ValidationResult

     
    PO_AMNT_FTM_TOTAL As Currency
    PO_AMNT_TOTAL As Currency
    PO_AMNT_MERCH_TOTAL As Currency
    
    PO_Fields As PeopleSoft_PurchaseOrder_Fields
    
    PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults
    
    PO_LineMods() As PeopleSoft_PurchaseOrder_CreateFromQuote_LineModification
    PO_LineModCount As Integer
    
    
    
    HasError As Boolean
    GlobalError As String
    
    BudgetCheck As PeopleSoft_PurchaseOrder_BudgetCheckParams

End Type

' ------------------------------------------------
' PO Change Order Types
' ------------------------------------------------
Type PeopleSoft_PurchaseOrder_ChangeOrder_Item
    PO_Line As Integer
    PO_Schedule As Integer
    PO_ItemID As String
    
    SCH_DUE_DATE As Date
    SCH_DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    
    HasError As Boolean
    ItemError As String
End Type


Type PeopleSoft_PurchaseOrder_ChangeOrder
    PO_BU As String
    PO_ID As String
    
    PO_BU_Result As PeopleSoft_Field_ValidationResult
    PO_ID_Result As PeopleSoft_Field_ValidationResult
    
    ' PO Defaults
    PO_DUE_DATE As Date
    PO_PROJECT_CODE As String
    
    PO_DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    PO_PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    
    ' PO Fields
    'PO_HDR_BUYER_ID As Long
    'PO_HDR_BUYER_ID_Result As PeopleSoft_Field_ValidationResult
    
    'PO_HDR_PO_REF As String
    
    PO_HDR_FLG_SEND_TO_VENDOR As PeopleSoft_Page_CheckboxAction
    
    PO_ChangeOrder_Items() As PeopleSoft_PurchaseOrder_ChangeOrder_Item
    PO_ChangeOrder_ItemCount As Integer
    
    
    ChangeReason As String
    
    
    BudgetCheck As PeopleSoft_PurchaseOrder_BudgetCheckParams
    
    HasError As Boolean
    GlobalError As String
End Type


' ------------------------------------------------
' PO Receipt Types
' ------------------------------------------------
Enum PeopleSoft_Receive_Mode
    RECEIVE_SPECIFIED = 0
    RECEIVE_ALL = 1
End Enum

Type PeopleSoft_Receipt_Item
    PO_Line As Integer
    PO_Schedule As Integer
    
    
    CATS_FLAG As String
    
    Item_ID As String
    TRANS_ITEM_DESC As String
    
    Receive_Qty As Currency ' Currenty type: Need to use fixed point type for accuracy.
    Accept_Qty As Currency
    
    IsNotReceivable As Boolean ' Returns True if not receivable (receive checkbox is greyed out)
    HasError As Boolean
    ItemError As String
End Type

Type PeopleSoft_Receipt
    PO_BU As String
    PO_ID As String
    
    PO_BU_Result As PeopleSoft_Field_ValidationResult
    PO_ID_Result As PeopleSoft_Field_ValidationResult
    
    RECEIPT_ID As String
    
    ReceiveMode As PeopleSoft_Receive_Mode
    
    ReceiptItems() As PeopleSoft_Receipt_Item
    ReceiptItemCount As Long
    
    HasGlobalError As Boolean
    GlobalError As String
End Type


' Internal type for storing unreceived items extracted from a PS page
Private Type PeopleSoft_ReceiptPage_UnreceivedItem
    PO_ID As String
    PO_Line As Long
    PO_Schedule As Long
    PO_Qty As Currency ' Fixed-point data type
    PO_Item_ID As String
    CATS_FLAG As String
    
    PO_TRANS_ITEM_DESC As String
       
    IsReceivable As Boolean
    
    PageTableRowIndex As Integer
End Type

' Internal type for storing ReceiptLine extracted from the receipt PS page
Private Type PeopleSoft_ReceiptPage_ReceiptLine
    Receipt_Line As Integer
    Item_ID As String
    Description As String
    
    
    Receipt_Qty As Currency ' Fixed-point data type (high accuracy)
    Accept_Qty As Currency ' Fixed-point data type (high accuracy)
    Receipt_Price As Currency ' Fixed-point data type (high accuracy)
    
    Status As String
    
    PO_Line As Long
    PO_Schedule As Long
    
    PageTableRowIndex As Integer
End Type

' ------------------------------------------------
' PO - Retry Save With Budget Check Types
' ------------------------------------------------
Type PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams
    PO_BU As String
    PO_ID As String
    
    PO_BU_Result As PeopleSoft_Field_ValidationResult
    PO_ID_Result As PeopleSoft_Field_ValidationResult
    
    BudgetCheck As PeopleSoft_PurchaseOrder_BudgetCheckParams
    
    HasGlobalError As Boolean
    GlobalError As String
End Type
' ------------------------------------------------
' Constants
' ------------------------------------------------
Private Const URI_BASE As String = "https://erpprd-fnprd.erp.vzwcorp.com/"
Private Const PS_URI_LOGIN As String = "https://websso.vzwcorp.com/siteminderagent/forms/vzwsso/sso_login.asp?TARGET=https://websso.vzwcorp.com/profileChk/chkProfile.asp?Orig_Trgt=HTTPS%3a%2f%2ferpprd-fnprd%2eerp%2evzwcorp%2ecom%2fpsp%2fps%2fEMPLOYEE%2fERP%2fh%2f%3ftab%3dDEFAULT"
Private Const PS_URI_PO_EXPRESS As String = "https://erpprd-fnprd.erp.vzwcorp.com/psc/ps/EMPLOYEE/ERP/c/MANAGE_PURCHASE_ORDERS.PURCHASE_ORDER_EXP.GBL"
Private Const PS_URI_RECEIPT_ADD As String = "https://erpprd-fnprd.erp.vzwcorp.com/psc/ps/EMPLOYEE/ERP/c/MANAGE_SHIPMENTS.RECV_PO.GBL"

Private Const TIMEOUT_INFINITE = -1 ' Use for Unlimited
Private Const TIMEOUT_LONG = 60 * 10 ' 10min (seconds)
Private Const TIMEOUT_IMPLICIT_WAIT = 3000 ' 3 seconds (milliseconds)

' ------------------------------------------------
' Debug Stuff
' ------------------------------------------------
Private Type PeopleSoft_Debug_Options
    InitFlag As Boolean

    CaptureExceptions As Boolean                     ' Raised errors/exceptions are captured and returned as errors (keeps flow going)
    AddMethodNamePrefixToExceptions As Boolean       ' Prefix local function name to errors
    
    QuitBeforeSave As Boolean                        ' Quit PS automation before making final changes (useful for testing)
    
    SaveDebugInfo_WriteDebugOutputToFile As Boolean  ' Save Debug_Print() output to file
    SaveDebugInfo_WriteSrcToFile As Boolean          ' Save WebDriver page HTML source to file
    SaveDebugInfo_TakeScreenShot As Boolean          ' Save WebDriver screenshot to file
    
    SaveDebugInfo_FilePrefix As String               ' Prefix of file to be used
End Type

Private DEBUG_OPTIONS As PeopleSoft_Debug_Options


Public Function GetSeleniumVersion() As String

    Dim assy As New SeleniumWrapper.Assembly
    
    GetSeleniumVersion = assy.GetVersion
    

End Function
Public Sub PeopleSoft_SetConfigOptions(Optional captureExceptionsAsError As Boolean = False, Optional addMethodNamesToExceptions As Boolean = False, _
    Optional writeDebugOutputToFile As Boolean = False, Optional writePageSrcToFile As Boolean = False, Optional takeScreenShot As Boolean = False, Optional quitAutomationBeforeSave As Boolean = False)
    
   
    DEBUG_OPTIONS.InitFlag = True

    DEBUG_OPTIONS.CaptureExceptions = captureExceptionsAsError
    DEBUG_OPTIONS.AddMethodNamePrefixToExceptions = addMethodNamesToExceptions
    
    DEBUG_OPTIONS.SaveDebugInfo_WriteDebugOutputToFile = writeDebugOutputToFile
    DEBUG_OPTIONS.SaveDebugInfo_WriteSrcToFile = writePageSrcToFile
    DEBUG_OPTIONS.SaveDebugInfo_TakeScreenShot = takeScreenShot
    
    DEBUG_OPTIONS.QuitBeforeSave = quitAutomationBeforeSave

End Sub
Public Sub PeopleSoft_SetDebugOutputOptions(Optional filePrefix As String)
    
    If IsMissing(filePrefix) = False Then
        DEBUG_OPTIONS.SaveDebugInfo_FilePrefix = filePrefix
    End If
End Sub
Private Sub PeopleSoft_SaveDebugInfo(driver As SeleniumWrapper.WebDriver)

    Dim dirPath As String
    Dim fileNamePrefix As String, fileHandle As Long
    
    dirPath = ThisWorkbook.Path & "\"
    fileNamePrefix = dirPath & _
                        IIf(Len(DEBUG_OPTIONS.SaveDebugInfo_FilePrefix) > 0, DEBUG_OPTIONS.SaveDebugInfo_FilePrefix, "") & _
                        "PS_" & Format$(Now(), "YYYYmmddHhmmSs")
    
    Debug_Print "PeopleSoft_SaveDebugInfo: Generating debug info files with prefix: " & fileNamePrefix
    
    If DEBUG_OPTIONS.SaveDebugInfo_WriteSrcToFile Then
        fileHandle = FreeFile
        
        Open fileNamePrefix & "_src.html" For Output As #fileHandle
            Write #fileHandle, driver.PageSource
        Close #fileHandle
    End If
    
    If DEBUG_OPTIONS.SaveDebugInfo_WriteDebugOutputToFile Then Debug_ToFile (fileNamePrefix & "_debug.txt")
    
    If DEBUG_OPTIONS.SaveDebugInfo_TakeScreenShot Then driver.captureEntirePageScreenshot fileNamePrefix & "_SS.png"

End Sub
' Utility function to check if a field value has multiple lines and sets validation result accordingly
Private Function PeopleSoft_ValidateField_IsSingleLine(fieldName As String, fieldVal As String, ByRef validationResult As PeopleSoft_Field_ValidationResult) As Boolean

    
    If Len(fieldVal) > 0 Then
        If InStr(1, fieldVal, vbCr) > 0 Or InStr(1, fieldVal, vbLf) > 0 Then
            validationResult.ValidationErrorText = fieldName & " does not allow multiple lines as input"
            validationResult.ValidationFailed = True
            
            PeopleSoft_ValidateField_IsSingleLine = False
            Exit Function
        End If
    End If
    
    PeopleSoft_ValidateField_IsSingleLine = True

End Function

' -----------------------------------------------------------------------------------------------
Public Function PeopleSoft_NewSession(user As String, pass As String) As PeopleSoft_Session


    
    Debug_Init reset:=True
    
    Debug_Print "PeopleSoft_NewSession called"
    
    Dim session As PeopleSoft_Session
    Dim driver As New SeleniumWrapper.WebDriver
    
    
   
    If DEBUG_OPTIONS.InitFlag = False Then PeopleSoft_SetConfigOptions
    
    'If True Then
    '    ' Setup global debug options: DEBUG MODE
    '    DEBUG_OPTIONS.CaptureExceptions = False
    '    DEBUG_OPTIONS.AddMethodNamePrefixToExceptions = False
    '
    '    DEBUG_OPTIONS.SaveDebugInfo_WriteDebugOutputToFile = True
    '    DEBUG_OPTIONS.SaveDebugInfo_WriteSrcToFile = True
    '    DEBUG_OPTIONS.SaveDebugInfo_TakeScreenShot = True
    'Else
    '    ' Setup global debug options: PRODUCTION MODE
    '    DEBUG_OPTIONS.CaptureExceptions = True
    '    DEBUG_OPTIONS.AddMethodNamePrefixToExceptions = True
    '
    '    DEBUG_OPTIONS.SaveDebugInfo_WriteDebugOutputToFile = True
    '    DEBUG_OPTIONS.SaveDebugInfo_WriteSrcToFile = True
    '    DEBUG_OPTIONS.SaveDebugInfo_TakeScreenShot = True
    'End If
    
    
    Set session.driver = driver
    
    
    session.user = user
    session.pass = pass
    session.loggedIn = False
    
    PeopleSoft_NewSession = session

End Function


Public Function PeopleSoft_Login(ByRef session As PeopleSoft_Session) As Boolean
    
    Debug_Print "PeopleSoft_Login: called"
    
    
    On Error GoTo ExceptionThrown
    
    Dim driver As SeleniumWrapper.WebDriver
    
    Set driver = session.driver
    
    
    session.LogonError = ""
    
    If Not session.loggedIn Then
        driver.Start "chrome", URI_BASE
        driver.setImplicitWait TIMEOUT_IMPLICIT_WAIT
        
        
        driver.get PS_URI_LOGIN
        
          
        driver.findElementByName("USER").Clear
        driver.findElementByName("USER").SendKeys session.user
        driver.findElementByName("password").Clear
        driver.findElementByName("password").SendKeys session.pass
        driver.findElementByName("btn_logon").Click
        
        driver.waitForPageToLoad 5000 ' wait up to 5s
        
        
        
        Dim By As New SeleniumWrapper.By, weErrorBoxMsg As SeleniumWrapper.WebElement
        Dim errMsg As String, errBoxExists As Boolean
        
        errBoxExists = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[starts-with(@id,'ErrorBox')]//p/b"))
        
        If errBoxExists Then
            errMsg = driver.findElementByXPath(".//*[starts-with(@id,'ErrorBox')]//p/b").text
                    
            
            Debug_Print "PeopleSoft_Login: failed: " & errMsg
            session.LogonError = "PeopleSoft Login Failed: " & errMsg
            PeopleSoft_Login = False
            Exit Function
        End If
    
        
        session.loggedIn = True
    End If
    
    PeopleSoft_Login = session.loggedIn
    Exit Function
  
ExceptionThrown:
    session.LogonError = "PeopleSoft_Login Exception: " & Err.Description
    
    PeopleSoft_Login = False

End Function

Public Function PeopleSoft_NavigateTo_AddPO(ByRef session As PeopleSoft_Session, PO_BU As String, ByRef PO_BU_Result As PeopleSoft_Field_ValidationResult) As Boolean

    Debug_Print "PeopleSoft_NavigateTo_AddPO called"
    
    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown


    Dim driver As SeleniumWrapper.WebDriver

    Set driver = session.driver

    driver.get PS_URI_PO_EXPRESS
    
    Dim PO_BU_default As String
    
    ' Check if auto-filled PO BU is correct. If not,enter the correct PO BU
    If PO_BU <> "" Then
        PO_BU_default = driver.findElementById("PO_ADD_SRCH_BUSINESS_UNIT").getAttribute("value")
    
        If PO_BU_default <> PO_BU Then
            PeopleSoft_Page_SetValidatedField driver, ("PO_ADD_SRCH_BUSINESS_UNIT"), PO_BU, PO_BU_Result
                
            If PO_BU_Result.ValidationFailed Then GoTo ValidationFail
        End If
    End If
    
    driver.findElementById("#ICSearch").Click
    'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSearch');" ' work-around - can't click 'Add'

    PeopleSoft_Page_WaitForProcessing driver
    
    
    PeopleSoft_NavigateTo_AddPO = True
    Exit Function

ValidationFail:
    PeopleSoft_NavigateTo_AddPO = False
    Exit Function
    
ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_NavigateTo_AddPO Exception: " & Err.Description, Err.Helpfile, Err.HelpContext

End Function
Public Function PeopleSoft_NavigateTo_ExistingPO(ByRef session As PeopleSoft_Session, PO_BU As String, PO_ID As String) As Boolean
    
    
    Debug_Print "PeopleSoft_NavigateTo_ExistingPO called (" & Debug_VarListString("PO BU", PO_BU, "PO ID", PO_ID) & ")"
    
    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As New SeleniumWrapper.WebDriver
    

    
    Set driver = session.driver
    
    
    driver.get PS_URI_PO_EXPRESS
    
    
    ' TODO: How to deal with the below error message
    ' <p class="psloginmessagelarge">We are not able to process your request at this time. Please close your web browser and try your request again. If this problem continues, please contact the ITSC
    '      and provide them with the details of what you were attempting to do when this problem occurred, along with any other details needed to
    '      reproduce this issue.</p>
    
    
    'driver.waitForElementPresent "css=#RECV_PO_ADD_BUSINESS_UNIT"
    
    
    ' Switch from Add to Find
    driver.runScript "javascript:submitAction_win0(document.win0,'#ICSwitchMode');"
    
    PeopleSoft_Page_WaitForProcessing driver
    
    
    Dim PO_BU_default As String
    
    ' Check if auto-filled PO BU is correct. If not,enter the correct PO BU
    If PO_BU <> "" Then
        PO_BU_default = driver.findElementById("PO_SRCH_BUSINESS_UNIT").getAttribute("value")
    
        If PO_BU_default <> PO_BU Then
            driver.findElementById("PO_SRCH_BUSINESS_UNIT").Clear
            driver.findElementById("PO_SRCH_BUSINESS_UNIT").SendKeys PO_BU
        End If
    End If
    
    
    
    
    driver.findElementById("PO_SRCH_PO_ID").SendKeys PO_ID
    driver.findElementById("PO_SRCH_OPRID_ENTERED_BY").Clear
    
    
    
    driver.findElementById("#ICSearch").Click
    'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSearch');"
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
    
    
    ' Check if we are in the budget error page, if so go back to the PO page
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[text()='PO Budget Check Errors']")) Then
        driver.findElementById("#ICCancel").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICCancel');"
        PeopleSoft_Page_WaitForProcessing driver
    End If

    
    If PeopleSoft_Page_ElementExists(driver, By.id("PO_PNLS_PB_PAGE_TITLE_PO")) = False Then
        Debug_Print "PeopleSoft_NavigateTo_ExistingPO: PO page not found"
        PeopleSoft_NavigateTo_ExistingPO = False
        Exit Function
    End If
    
    
    Debug_Print "PeopleSoft_NavigateTo_ExistingPO: PO page found!"
    

    PeopleSoft_NavigateTo_ExistingPO = True
    Exit Function
    
    
ExceptionThrown:
    PeopleSoft_NavigateTo_ExistingPO = False
    Err.Raise Err.Number, Err.Source, "PeopleSoft_NavigateTo_ExistingPO Exception: " & Err.Description, Err.Helpfile, Err.HelpContext
    
    
End Function
Public Function PeopleSoft_PurchaseOrder_CutPO(ByRef session As PeopleSoft_Session, ByRef purchaseOrder As PeopleSoft_PurchaseOrder) As Boolean


    Debug_Print "PeopleSoft_PurchaseOrder_CutPO called (" & Debug_VarListString("PO Ref", purchaseOrder.PO_Fields.PO_HDR_PO_REF) & ")"
    

    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown


    ' Pre-Check
    If PeopleSoft_PurchaseOrder_CutPO_PreCheck(purchaseOrder) = False Then
        PeopleSoft_PurchaseOrder_CutPO = False
        Exit Function
    End If

    Dim driver As SeleniumWrapper.WebDriver

    
    If Not session.loggedIn Then
        session.loggedIn = PeopleSoft_Login(session)
        
        If Not session.loggedIn Then
            purchaseOrder.GlobalError = "Logon Error: " & session.LogonError
            purchaseOrder.HasError = True
            
            PeopleSoft_PurchaseOrder_CutPO = False
            Exit Function
        End If
    End If
    
    


    Set driver = session.driver


    
    Call PeopleSoft_NavigateTo_AddPO(session, purchaseOrder.PO_Fields.PO_BUSINESS_UNIT, purchaseOrder.PO_Fields.PO_BUSINESS_UNIT_Result)
    If purchaseOrder.PO_Fields.PO_BUSINESS_UNIT_Result.ValidationFailed Then GoTo ValidationFail
    
    
    If purchaseOrder.PO_Fields.PO_HDR_VENDOR_ID > 0 Then
        ' If vendor ID provided: use vendor ID.
        Dim weID As String
        weID = driver.findElementByXPath(".//input[starts-with(@id,'PO_HDR_VENDOR_ID')]").getAttribute("id")
        PeopleSoft_Page_SetValidatedField driver, (weID), Format(purchaseOrder.PO_Fields.PO_HDR_VENDOR_ID, "0000000000"), purchaseOrder.PO_Fields.PO_HDR_VENDOR_ID_Result
        If purchaseOrder.PO_Fields.PO_HDR_VENDOR_ID_Result.ValidationFailed Then GoTo ValidationFail
    Else
        ' Otherwise: use vendor name short
        PeopleSoft_Page_SetValidatedField driver, ("VENDOR_VENDOR_NAME_SHORT"), purchaseOrder.PO_Fields.VENDOR_NAME_SHORT, purchaseOrder.PO_Fields.VENDOR_NAME_SHORT_Result
        If purchaseOrder.PO_Fields.VENDOR_NAME_SHORT_Result.ValidationFailed Then GoTo ValidationFail
    End If
    
    
    ' Vendor location
    If Len(purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("Z_VNDR_PNLS_WRK_VNDR_LOC"), _
            purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION, purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result
        If purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then GoTo ValidationFail
    Else
        ' If vendor location not provided: check if it has a valid value. If not, then we cannot continue
        Dim vendorLocationText As String
        vendorLocationText = driver.findElementById("Z_VNDR_PNLS_WRK_VNDR_LOC").getAttribute("value")
        
        If vendorLocationText = "" Then
            purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result.ValidationFailed = True
            purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText = "Vendor location is required for this vendor"
            GoTo ValidationFail
        End If
    End If
    

    
    ' Buyer ID
    PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_BUYER_ID"), _
        CStr(purchaseOrder.PO_Fields.PO_HDR_BUYER_ID), purchaseOrder.PO_Fields.PO_HDR_BUYER_ID_Result
        
    If purchaseOrder.PO_Fields.PO_HDR_BUYER_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    ' PO Reference
    If Len(purchaseOrder.PO_Fields.PO_HDR_PO_REF) > 0 Then
        driver.findElementById("PO_HDR_PO_REF").Clear
        driver.findElementById("PO_HDR_PO_REF").SendKeys purchaseOrder.PO_Fields.PO_HDR_PO_REF
    End If
    
    
    If Len(purchaseOrder.PO_Fields.PO_HDR_XPRESS_BID_ID) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_Z_XPRESS_BID_ID"), CStr(purchaseOrder.PO_Fields.PO_HDR_XPRESS_BID_ID), purchaseOrder.PO_Fields.PO_HDR_XPRESS_BID_ID_Result
        ' Note: Here we ignore validation result since the field is not actually validated field by PS
    End If
    
    
    
    ' -------------------------------------------------------------------
    ' Begin - Header Section
    ' -------------------------------------------------------------------
    If Len(purchaseOrder.PO_Fields.PO_HDR_QUOTE) > 0 Then
        Debug_Print "PeopleSoft_PurchaseOrder_CutPO: Navigating to PO header page"
        
        ' Only if quote field provided
    
        driver.findElementById("PO_PNLS_WRK_GOTO_HDR_DTL").Click
         PeopleSoft_Page_WaitForProcessing driver
        
        driver.waitForElementPresent "css=#PO_HDR_Z_QUOTE_NBR"
        
        driver.findElementById("PO_HDR_Z_QUOTE_NBR").Clear
        driver.findElementById("PO_HDR_Z_QUOTE_NBR").SendKeys purchaseOrder.PO_Fields.PO_HDR_QUOTE
    
        
        driver.findElementById("#ICSave").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
        PeopleSoft_Page_WaitForProcessing driver
    End If
    ' -------------------------------------------------------------------
    ' End - Header Section
    ' -------------------------------------------------------------------
    
    ' Fill PO Comments & Attach Quote
    If PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page(driver, purchaseOrder.PO_Fields) = False Then
        ' TODO: Add .HasValidationError calculation to get GlobalError
        purchaseOrder.GlobalError = "Error occurred when filling in comments page, or attaching quote"
        GoTo ValidationFail
    End If
    

    
    Dim PO_Line As Integer
    Dim PO_LineCount As Integer
    Dim PO_pageLineIndex As Integer
    Dim PO_pageScheduleIndex As Integer
    Dim PO_Line_Schedule As Integer
    
    Dim isExpenseLine As Boolean, lineItemID As String
    
    ' -------------------------------------------------------------------
    ' Begin - Calculate and fill defaults
    ' -------------------------------------------------------------------
    PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc purchaseOrder
    PeopleSoft_PurchaseOrder_PO_Defaults_Fill driver, purchaseOrder.PO_Defaults
    
    ' Begin - Transfer information from defaults to each corresponding line/schedule, where applicable
    For PO_Line = 1 To purchaseOrder.PO_LineCount
        ' Determine if expense line
        isExpenseLine = False
        lineItemID = UCase$(purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_ITEM_ID)
        If lineItemID Like "EXP-*" Or lineItemID Like "CR-EXP-*" Then isExpenseLine = True
    
        For PO_Line_Schedule = 1 To purchaseOrder.PO_Lines(PO_Line).ScheduleCount
        
            ' If there was a validation error -> Transfer Validation Results/Extracted fields from defaults to each line/schedule
            ' provided that the field value is the same as the default
            If purchaseOrder.PO_Defaults.HasValidationError Then
                With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule)
                
                    ' Expense/Capital agnostic fields
                    If .ScheduleFields.DUE_DATE > 0 And .ScheduleFields.DUE_DATE = purchaseOrder.PO_Defaults.SCH_DUE_DATE Then
                        .ScheduleFields.DUE_DATE_Result = purchaseOrder.PO_Defaults.SCH_DUE_DATE_Result
                        purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    End If
                    
                    ' Expense/Capital specific fields
                    If isExpenseLine Then
                        If .ScheduleFields.SHIPTO_ID > 0 And .ScheduleFields.SHIPTO_ID = purchaseOrder.PO_Defaults.DIST_EXP_SHIP_TO_ID Then
                            .ScheduleFields.SHIPTO_ID_Result = purchaseOrder.PO_Defaults.DIST_EXP_SHIP_TO_ID_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If Len(.DistributionFields.BUSINESS_UNIT_PC) > 0 And .DistributionFields.BUSINESS_UNIT_PC = purchaseOrder.PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC Then
                            .DistributionFields.BUSINESS_UNIT_PC_Result = purchaseOrder.PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If Len(.DistributionFields.PROJECT_CODE) > 0 And .DistributionFields.PROJECT_CODE = purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_CODE Then
                            .DistributionFields.PROJECT_CODE_Result = purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_CODE_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If Len(.DistributionFields.ACTIVITY_ID) > 0 And .DistributionFields.ACTIVITY_ID = purchaseOrder.PO_Defaults.DIST_EXP_ACTIVITY_ID Then
                            .DistributionFields.ACTIVITY_ID_Result = purchaseOrder.PO_Defaults.DIST_EXP_ACTIVITY_ID_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If .DistributionFields.LOCATION_ID > 0 And .DistributionFields.LOCATION_ID = purchaseOrder.PO_Defaults.DIST_EXP_LOCATION_ID Then
                            .DistributionFields.LOCATION_ID_Result = purchaseOrder.PO_Defaults.DIST_EXP_LOCATION_ID_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                    Else
                        If .ScheduleFields.SHIPTO_ID > 0 And .ScheduleFields.SHIPTO_ID = purchaseOrder.PO_Defaults.DIST_CAP_SHIP_TO_ID Then
                            .ScheduleFields.SHIPTO_ID_Result = purchaseOrder.PO_Defaults.DIST_CAP_SHIP_TO_ID_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If Len(.DistributionFields.BUSINESS_UNIT_PC) > 0 And .DistributionFields.BUSINESS_UNIT_PC = purchaseOrder.PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC Then
                            .DistributionFields.BUSINESS_UNIT_PC_Result = purchaseOrder.PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If Len(.DistributionFields.PROJECT_CODE) > 0 And .DistributionFields.PROJECT_CODE = purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_CODE Then
                            .DistributionFields.PROJECT_CODE_Result = purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_CODE_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If Len(.DistributionFields.ACTIVITY_ID) > 0 And .DistributionFields.ACTIVITY_ID = purchaseOrder.PO_Defaults.DIST_CAP_ACTIVITY_ID Then
                            .DistributionFields.ACTIVITY_ID_Result = purchaseOrder.PO_Defaults.DIST_CAP_ACTIVITY_ID_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                        If .DistributionFields.LOCATION_ID > 0 And .DistributionFields.LOCATION_ID = purchaseOrder.PO_Defaults.DIST_CAP_LOCATION_ID Then
                            .DistributionFields.LOCATION_ID_Result = purchaseOrder.PO_Defaults.DIST_CAP_LOCATION_ID_Result
                            purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                        End If
                    End If
                    
                End With
            
            End If
            
            ' Transfer extracted fields where applicable
            With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule)
                If isExpenseLine Then
                    ' Copy project description if project codes are the same
                    If Len(purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_DESC) > 0 And .DistributionFields.PROJECT_CODE = purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_CODE Then _
                        .DistributionFields.PROJECT_DESC = purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_DESC
                Else
                    ' Copy project description if project codes are the same
                    If Len(purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_DESC) > 0 And .DistributionFields.PROJECT_CODE = purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_CODE Then _
                        .DistributionFields.PROJECT_DESC = purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_DESC
                End If
            End With
            
            
        Next PO_Line_Schedule
        
        
        If purchaseOrder.PO_Lines(PO_Line).HasValidationError Then purchaseOrder.PO_Defaults.HasValidationError = True
    Next PO_Line
        
    If purchaseOrder.PO_Defaults.HasValidationError Then
        purchaseOrder.GlobalError = "One or more line fields failed PS validation"
        GoTo ValidationFail
    End If
    ' End - Transfer information from defaults to each corresponding line/schedule, where applicable
    
    ' -------------------------------------------------------------------
    ' End - Calculate and fill defaults
    ' -------------------------------------------------------------------
    
    ' -------------------------------------------------------------------
    ' Begin - Add individual lines to PO
    ' -------------------------------------------------------------------
    Debug_Print "PeopleSoft_PurchaseOrder_CutPO: Begin PO Line/Schedule Adds (" & Debug_VarListString("Line Count", purchaseOrder.PO_LineCount) & ")"
    
    
    PO_Line = 1
    PO_pageLineIndex = 0
    PO_pageScheduleIndex = 0
    PO_LineCount = purchaseOrder.PO_LineCount 'UBound(purchaseOrder.PO_Lines)
    
    
    ' Add X items
    If PO_LineCount > 1 Then
        driver.runScript "javascript:document.win0.ICAddCount.value = " & (PO_LineCount - 1) & "; submitAction_win0(document.win0,'PO_LINE_SCROLL$newm$0$$0'); " ' work-around
        PeopleSoft_Page_WaitForProcessing driver
    End If
        
    ' Expand All
    driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
    PeopleSoft_Page_WaitForProcessing driver
    
    
    ' Begin - Add multiple schedules
    Dim PO_Line_ScheduleIndex As Integer
    Dim PO_AnyLineHasMultiSchedules As Boolean
    
    PO_Line_ScheduleIndex = 0
    PO_AnyLineHasMultiSchedules = True
    
    For PO_Line = 1 To PO_LineCount
        Dim PO_Line_ScheduleCount As Integer
        
        PO_Line_ScheduleCount = UBound(purchaseOrder.PO_Lines(PO_Line).Schedules)
        
        If PO_Line_ScheduleCount > 1 Then
            PO_AnyLineHasMultiSchedules = True
            
            driver.runScript "javascript:document.win0.ICAddCount.value = " & (PO_Line_ScheduleCount - 1) & "; javascript:submitAction_win0(document.win0,'PO_LINE_SHIP_SCROL$newm$" & PO_Line_ScheduleIndex & "$$" & (PO_Line - 1) & "'); " ' work-around
            PeopleSoft_Page_WaitForProcessing driver
        End If
        
        
        PO_Line_ScheduleIndex = PO_Line_ScheduleIndex + PO_Line_ScheduleCount
    Next PO_Line
    
    If PO_AnyLineHasMultiSchedules Then
        driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
        PeopleSoft_Page_WaitForProcessing driver
    End If
    ' End - Add multiple schedules
    
    
    
    
    For PO_Line = 1 To PO_LineCount
        Debug_Print "PeopleSoft_PurchaseOrder_CutPO: Processing Line #" & PO_Line & " (" & _
                Debug_VarListString("Item ID", purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_ITEM_ID) & ")"
   
        
        PeopleSoft_PurchaseOrder_Fill_PO_Line driver, purchaseOrder, PO_Line, PO_pageScheduleIndex
        If purchaseOrder.HasError Then GoTo ValidationFail
        
        If purchaseOrder.PO_Lines(PO_Line).HasValidationError Then GoTo ValidationFail
        
        PO_pageScheduleIndex = PO_pageScheduleIndex + purchaseOrder.PO_Lines(PO_Line).ScheduleCount
    Next PO_Line
    
    
    
    
    driver.runScript "javascript:submitAction_win0(document.win0,'CALCULATE_TAXES');" ' Fix for 2.9.1.1  due to PS upgrade
    'driver.findElementById("CALCULATE_TAXES").Click
    PeopleSoft_Page_WaitForProcessing driver

    
    Dim amntStr As String
    
    ' Total
    amntStr = driver.findElementById("PO_PNLS_WRK_PO_AMT_TTL").text
    purchaseOrder.PO_AMNT_TOTAL = CurrencyFromString(amntStr)
    
    ' Total w/o Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_MERCH_AMT_TTL").text
    purchaseOrder.PO_AMNT_MERCH_TOTAL = CurrencyFromString(amntStr)
    
    ' Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_ADJ_AMT_TTL_LBL").text
    purchaseOrder.PO_AMNT_FTM_TOTAL = CurrencyFromString(amntStr)
    
    
    If DEBUG_OPTIONS.QuitBeforeSave Then
        purchaseOrder.GlobalError = "Debug option QuitBeforeSave enabled. Processing halted before saving. To prevent this from occurring, disable QuitBeforeSave option."
        purchaseOrder.HasError = True
        PeopleSoft_PurchaseOrder_CutPO = False
        Exit Function
    End If
    
    Dim result As Boolean
    
    result = PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver, purchaseOrder.BudgetCheck)
    
    purchaseOrder.PO_ID = purchaseOrder.BudgetCheck.PO_ID
    
    
    
    If result = False Then
        purchaseOrder.GlobalError = purchaseOrder.BudgetCheck.GlobalError
        purchaseOrder.HasError = purchaseOrder.BudgetCheck.HasGlobalError
        
        PeopleSoft_PurchaseOrder_CutPO = False
        Exit Function
    End If
    
    
    Debug_Print "PeopleSoft_PurchaseOrder_CutPO: complete (" & Debug_VarListString("PO ID", purchaseOrder.PO_ID) & ")"
    
    PeopleSoft_PurchaseOrder_CutPO = True
    Exit Function
    
    

CreatePoFailed:
ValidationFail:
    PeopleSoft_SaveDebugInfo driver
    PeopleSoft_PurchaseOrder_CutPO = False
    Exit Function
    
ExceptionThrown:
    PeopleSoft_SaveDebugInfo driver
    purchaseOrder.HasError = True
    purchaseOrder.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_CutPO = False


End Function
Private Function PeopleSoft_PurchaseOrder_PO_Fields_PreCheck(poFields As PeopleSoft_PurchaseOrder_Fields) As Boolean

    
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Fields_PreCheck: PO (" & Debug_VarListString( _
        "PO REF", poFields.PO_HDR_PO_REF, _
        "BUYER ID", poFields.PO_HDR_BUYER_ID, _
        "VENDOR ID", poFields.PO_HDR_VENDOR_ID, _
        "VENDOR NAME SHORT", poFields.VENDOR_NAME_SHORT, _
        "PO QUOTE", poFields.PO_HDR_QUOTE, _
        "XPRESS BID ID", poFields.PO_HDR_XPRESS_BID_ID, _
        "Quote FilePath", poFields.Quote_Attachment_FilePath) _
        & ")"
    
    
    
    ' Begin - Check required fields
    If Len(poFields.PO_BUSINESS_UNIT) = 0 Then
        poFields.PO_BUSINESS_UNIT_Result.ValidationErrorText = "PO BUSINESS UNIT is required."
        poFields.PO_BUSINESS_UNIT_Result.ValidationFailed = True
        poFields.HasValidationError = True
    End If
    If poFields.PO_HDR_BUYER_ID = 0 Then
        poFields.PO_HDR_BUYER_ID_Result.ValidationErrorText = "BUYER ID is required."
        poFields.PO_HDR_BUYER_ID_Result.ValidationFailed = True
        poFields.HasValidationError = True
    End If
    If poFields.PO_HDR_VENDOR_ID = 0 And Len(poFields.VENDOR_NAME_SHORT_Result) = 0 Then
        poFields.PO_HDR_VENDOR_ID_Result.ValidationErrorText = "PO VENDOR is required."
        poFields.PO_HDR_VENDOR_ID_Result.ValidationFailed = True
        poFields.VENDOR_NAME_SHORT_Result.ValidationErrorText = "PO VENDOR is required."
        poFields.VENDOR_NAME_SHORT_Result.ValidationFailed = True
        poFields.HasValidationError = True
    End If
    If Len(poFields.PO_HDR_PO_REF) = 0 Then
        poFields.PO_HDR_PO_REF_Result.ValidationErrorText = "PO REF is required."
        poFields.PO_HDR_PO_REF_Result.ValidationFailed = True
        poFields.HasValidationError = True
    End If
    
    If poFields.HasValidationError Then GoTo PreCheckFailed
    ' End - Check required fields
    
    ' Begin - validate fields: ensure single-line fields are actually a single line
    If PeopleSoft_ValidateField_IsSingleLine("PO BUSINESS UNIT", poFields.PO_BUSINESS_UNIT, poFields.PO_BUSINESS_UNIT_Result) = False Then poFields.HasValidationError = True
    If PeopleSoft_ValidateField_IsSingleLine("PO VENDOR NAME SHORT", poFields.VENDOR_NAME_SHORT, poFields.VENDOR_NAME_SHORT_Result) = False Then poFields.HasValidationError = True
    If PeopleSoft_ValidateField_IsSingleLine("PO REF", poFields.PO_HDR_PO_REF, poFields.PO_HDR_PO_REF_Result) = False Then poFields.HasValidationError = True
    If PeopleSoft_ValidateField_IsSingleLine("PO QUOTE", poFields.PO_HDR_QUOTE, poFields.PO_HDR_QUOTE_Result) = False Then poFields.HasValidationError = True
    If PeopleSoft_ValidateField_IsSingleLine("XPRESS BID ID", poFields.PO_HDR_XPRESS_BID_ID, poFields.PO_HDR_XPRESS_BID_ID_Result) = False Then poFields.HasValidationError = True
    
    If PeopleSoft_ValidateField_IsSingleLine("Quote Attachment FilePath", poFields.Quote_Attachment_FilePath, poFields.Quote_Attachment_FilePath_Result) = False Then poFields.HasValidationError = True

    If poFields.HasValidationError Then GoTo PreCheckFailed
    ' End - validate fields: ensure single-line fields are actually a single line
    
    
    ' Check if Quote_Attachment_FilePath exists
    If Len(poFields.Quote_Attachment_FilePath) > 0 Then
        ' Check if file exists
        If Dir(poFields.Quote_Attachment_FilePath) = "" Then
            poFields.HasValidationError = True
            poFields.Quote_Attachment_FilePath_Result.ValidationFailed = True
            poFields.Quote_Attachment_FilePath_Result.ValidationErrorText = "File Not Found: " & poFields.Quote_Attachment_FilePath
            
            GoTo PreCheckFailed
        End If
    End If
    
    
    PeopleSoft_PurchaseOrder_PO_Fields_PreCheck = True
    Exit Function
    
PreCheckFailed:
    poFields.HasValidationError = True
    PeopleSoft_PurchaseOrder_PO_Fields_PreCheck = False
    
    
    Dim globalErrorStr As String: globalErrorStr = ""
    
    If poFields.PO_BUSINESS_UNIT_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.PO_BUSINESS_UNIT_Result.ValidationErrorText & vbCrLf
    If poFields.PO_HDR_BUYER_ID_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.PO_HDR_BUYER_ID_Result.ValidationErrorText & vbCrLf
    If poFields.PO_HDR_VENDOR_ID_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.PO_HDR_VENDOR_ID_Result.ValidationErrorText & vbCrLf
    If poFields.VENDOR_NAME_SHORT_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.VENDOR_NAME_SHORT_Result.ValidationErrorText & vbCrLf
    If poFields.PO_HDR_PO_REF_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.PO_HDR_PO_REF_Result.ValidationErrorText & vbCrLf
    If poFields.PO_HDR_QUOTE_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.PO_HDR_QUOTE_Result.ValidationErrorText & vbCrLf
    If poFields.PO_HDR_XPRESS_BID_ID_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.PO_HDR_XPRESS_BID_ID_Result.ValidationErrorText & vbCrLf
    If poFields.Quote_Attachment_FilePath_Result.ValidationFailed Then globalErrorStr = globalErrorStr & poFields.Quote_Attachment_FilePath_Result.ValidationErrorText & vbCrLf

    
    ' Remove CR-LF at end of error string if it exists
    If Len(globalErrorStr) > 0 Then
        If Right$(globalErrorStr, Len(vbCrLf)) = vbCrLf Then globalErrorStr = Left$(globalErrorStr, Len(globalErrorStr) - Len(vbCrLf))
    End If
    
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Fields_PreCheck: failed! " & globalErrorStr
    
End Function

Private Function PeopleSoft_PurchaseOrder_CutPO_PreCheck(purchaseOrder As PeopleSoft_PurchaseOrder) As Boolean

    
    Debug_Print "PeopleSoft_PurchaseOrder_CutPO_PreCheck: PO (" & Debug_VarListString( _
        "PO REF", purchaseOrder.PO_Fields.PO_HDR_PO_REF) _
        & ")"
    
       
    If PeopleSoft_PurchaseOrder_PO_Fields_PreCheck(purchaseOrder.PO_Fields) = False Then
        purchaseOrder.GlobalError = "One or more PO fields failed validation"
        GoTo PreCheckFailed
    End If
    
    
    '---------------------------------------------------------------
    ' Begin - Perform PO Line/Schedule Pre-Checks
    '---------------------------------------------------------------
    Dim PO_Line As Integer, PO_Schedule As Integer
    Dim hasAnyValidationError As Boolean
    
    If purchaseOrder.PO_LineCount < 1 Then
        purchaseOrder.GlobalError = "PO must have at least one PO line and schedule"
        GoTo PreCheckFailed
    End If
    
    ' Check if each line has at least one schedule
    For PO_Line = 1 To purchaseOrder.PO_LineCount
        If purchaseOrder.PO_Lines(PO_Line).ScheduleCount < 1 Then
            purchaseOrder.GlobalError = "PO Line " & PO_Line & " must have one or more schedules"
            purchaseOrder.HasError = True
        End If
    Next PO_Line
    
    If purchaseOrder.HasError Then GoTo PreCheckFailed
    
    
    ' Begin - Check required fields
    hasAnyValidationError = False
    
    For PO_Line = 1 To purchaseOrder.PO_LineCount
        ' PO Line required fields
        With purchaseOrder.PO_Lines(PO_Line)
            If Len(.LineFields.PO_LINE_ITEM_ID) = 0 Then
                .LineFields.PO_LINE_ITEM_ID_Result.ValidationErrorText = "PO_LINE_ITEM_ID is required."
                .LineFields.PO_LINE_ITEM_ID_Result.ValidationFailed = True
                .HasValidationError = True
                hasAnyValidationError = True
            End If
        End With
    
        For PO_Schedule = 1 To purchaseOrder.PO_Lines(PO_Line).ScheduleCount
            With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Schedule)
                If .ScheduleFields.DUE_DATE = 0 And purchaseOrder.PO_Defaults.SCH_DUE_DATE = 0 Then
                    .ScheduleFields.DUE_DATE_Result.ValidationErrorText = "SCH DUE DATE is required."
                    .ScheduleFields.DUE_DATE_Result.ValidationFailed = True
                    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    hasAnyValidationError = True
                End If
                If .ScheduleFields.QTY = 0 Then
                    .ScheduleFields.QTY_Result.ValidationErrorText = "SCH QTY is required."
                    .ScheduleFields.QTY_Result.ValidationFailed = True
                    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    hasAnyValidationError = True
                End If
                If .ScheduleFields.SHIPTO_ID = 0 And purchaseOrder.PO_Defaults.DIST_CAP_SHIP_TO_ID = 0 Then
                    .ScheduleFields.SHIPTO_ID_Result.ValidationErrorText = "SCH SHIP TO ID is required."
                    .ScheduleFields.SHIPTO_ID_Result.ValidationFailed = True
                    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    hasAnyValidationError = True
                End If
                
                ' NOTE: Distribution fields may or may not be required depending on the item type (cap/expense). We will forego any pre-check and let PeopleSoft handle it.
            End With
        Next PO_Schedule
    Next PO_Line
    
    If hasAnyValidationError Then
        purchaseOrder.GlobalError = "One or more required PO Line/Schedule fields are missing data"
        GoTo PreCheckFailed
    End If
    ' End - Check required fields
    
    
    ' Begin - validate fields: ensure single-line fields are actually a single line
    hasAnyValidationError = False
    
    For PO_Line = 1 To purchaseOrder.PO_LineCount
        ' PO Line single-line fields
        With purchaseOrder.PO_Lines(PO_Line)
            If PeopleSoft_ValidateField_IsSingleLine("PO_LINE_ITEM_ID", .LineFields.PO_LINE_ITEM_ID, .LineFields.PO_LINE_ITEM_ID_Result) = False Then
                .HasValidationError = True
                hasAnyValidationError = True
            End If
        End With
    
        For PO_Schedule = 1 To purchaseOrder.PO_Lines(PO_Line).ScheduleCount
            With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Schedule)
                If PeopleSoft_ValidateField_IsSingleLine("BUSINESS_UNIT_PC", .DistributionFields.BUSINESS_UNIT_PC, .DistributionFields.BUSINESS_UNIT_PC_Result) = False Then
                    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    hasAnyValidationError = True
                End If
                If PeopleSoft_ValidateField_IsSingleLine("PROJECT_CODE", .DistributionFields.PROJECT_CODE, .DistributionFields.PROJECT_CODE_Result) = False Then
                    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    hasAnyValidationError = True
                End If
                If PeopleSoft_ValidateField_IsSingleLine("ACTIVITY_ID", .DistributionFields.ACTIVITY_ID, .DistributionFields.ACTIVITY_ID_Result) = False Then
                    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True
                    hasAnyValidationError = True
                End If
            End With
        Next PO_Schedule
    Next PO_Line
    
    If hasAnyValidationError Then
        purchaseOrder.GlobalError = "One or more required PO Line/Schedule field failed validation"
        GoTo PreCheckFailed
    End If
    ' End - validate fields: ensure single-line fields are actually a single line
    
    
    '---------------------------------------------------------------
    ' End - Perform PO Line/Schedule Pre-Checks
    '---------------------------------------------------------------

    
    PeopleSoft_PurchaseOrder_CutPO_PreCheck = True
    Exit Function
    
PreCheckFailed:
    purchaseOrder.HasError = True
    PeopleSoft_PurchaseOrder_CutPO_PreCheck = False
    
    Debug_Print "PeopleSoft_PurchaseOrder_CutPO_PreCheck: failed! " & purchaseOrder.GlobalError
    
End Function
Public Function PeopleSoft_PurchaseOrder_CreateFromQuote(ByRef session As PeopleSoft_Session, ByRef poCFQ As PeopleSoft_PurchaseOrder_CreateFromQuoteParams) As Boolean

    
    
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote called (" & Debug_VarListString("E_QUOTE_NBR", poCFQ.E_QUOTE_NBR, "PO Ref", poCFQ.PO_Fields.PO_HDR_PO_REF) & ")"
    
    
    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown
    
    ' Perform Pre-Check
    If PeopleSoft_PurchaseOrder_CreateFromQuote_PreCheck(poCFQ) = False Then
        poCFQ.HasError = True
        PeopleSoft_PurchaseOrder_CreateFromQuote = False
        Exit Function
    End If


    Dim driver As SeleniumWrapper.WebDriver, By As New SeleniumWrapper.By

    
    If Not session.loggedIn Then
        session.loggedIn = PeopleSoft_Login(session)
        
        If Not session.loggedIn Then
            poCFQ.GlobalError = "Logon Error: " & session.LogonError
            poCFQ.HasError = True
            
            PeopleSoft_PurchaseOrder_CreateFromQuote = False
            Exit Function
        End If
    End If
    
    


    Set driver = session.driver


    
    Call PeopleSoft_NavigateTo_AddPO(session, poCFQ.PO_Fields.PO_BUSINESS_UNIT, poCFQ.PO_Fields.PO_BUSINESS_UNIT_Result)
    If poCFQ.PO_Fields.PO_BUSINESS_UNIT_Result.ValidationFailed Then GoTo ValidationFail
    
    
    
    PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_BUYER_ID"), CStr(poCFQ.PO_Fields.PO_HDR_BUYER_ID), poCFQ.PO_Fields.PO_HDR_BUYER_ID_Result
    If poCFQ.PO_Fields.PO_HDR_BUYER_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    
    If Len(poCFQ.PO_Fields.PO_HDR_PO_REF) > 0 Then
        driver.findElementById("PO_HDR_PO_REF").Clear
        driver.findElementById("PO_HDR_PO_REF").SendKeys poCFQ.PO_Fields.PO_HDR_PO_REF
    End If
    
    
    If Len(poCFQ.PO_Fields.PO_HDR_XPRESS_BID_ID) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_Z_XPRESS_BID_ID"), CStr(poCFQ.PO_Fields.PO_HDR_XPRESS_BID_ID), poCFQ.PO_Fields.PO_HDR_XPRESS_BID_ID_Result
        ' Note: Here we ignore validation result since the field is not actually validated field by PS
    End If
    
    
    
    'Dim elemSelect As SeleniumWrapper.Select
    Dim elemSelect As SeleniumWrapper.WebElement
    

    ' Select CopyFrom eQuote and force load next page
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: select copy from eQuote"
    
    Set elemSelect = driver.findElementById("PO_COPY_TMPLT_W_COPY_PO_FROM")
    elemSelect.AsSelect.selectByText "eQuote"
    PeopleSoft_Page_WaitForProcessing driver
    driver.Wait 1000
    
    ' Create from Quote
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: waiting for 'Create from Quote' text"
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//h1[@class='PSSRCHTITLE' and text()='Create from Quote']")) = False Then
        Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: 'Create from Quote' not found. Forcing change event."
        driver.runScript "javascript: var elem = document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM'); addchg_win0(elem); submitAction_win0(elem.form,elem.name);"
        PeopleSoft_Page_WaitForProcessing driver
    End If
    
    driver.waitForElementPresent "xpath=.//h1[text()='Create from Quote']"
    

    ' Type Vendor ID
    PeopleSoft_Page_SetValidatedField driver, ("Z_E_QT_WRK_VENDOR_ID"), Format(poCFQ.PO_Fields.PO_HDR_VENDOR_ID, "0000000000"), poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result
    If poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    ' Type Quote Number
    PeopleSoft_Page_SetValidatedField driver, ("Z_E_QT_WRK_Z_QUOTE_NBR"), poCFQ.E_QUOTE_NBR, poCFQ.E_QUOTE_NBR_Result
    If poCFQ.E_QUOTE_NBR_Result.ValidationFailed Then GoTo ValidationFail
    
    ' Click Search
    driver.findElementById("Z_E_QT_WRK_REFRESH").Click
    PeopleSoft_Page_WaitForProcessing driver
    
    
    Dim pageQuoteNbr As String
    pageQuoteNbr = driver.findElementById("Z_E_QT_CPPO_VW_Z_QUOTE_NBR$0").text
    
    
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: pageQuoteNbr is '" & pageQuoteNbr & "'"

    ' Sanity check: check if loaded quote is the same as the provided E_QUOTE_NBR
    If pageQuoteNbr <> poCFQ.E_QUOTE_NBR Then
        poCFQ.HasError = True
        poCFQ.GlobalError = "Sanity check failed: quote # mismatch. Quote # on page '" & pageQuoteNbr & "' does not match provided quote # '" & poCFQ.E_QUOTE_NBR & "'"
        GoTo ValidationFail
    End If
    
    
    ' Click Apply
    driver.findElementById("Z_E_QT_WRK_APPLY").Click
    'driver.runScript "javascript:hAction_win0(document.win0,'Z_E_QT_WRK_APPLY', 0, 0, 'Apply', false, true)"
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG

   
    ' Fill in PO Defaults:
    PeopleSoft_PurchaseOrder_PO_Defaults_Fill driver, poCFQ.PO_Defaults
    If poCFQ.PO_Defaults.HasValidationError Then GoTo ValidationFail
    
    ' Fill PO Comments & Attach Quote
    Dim fillResult As Boolean
    fillResult = PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page(driver, poCFQ.PO_Fields)
    If Not fillResult Then GoTo ValidationFail ' TODO: Add .HasValidationError calculation
    

    

    
    ' -------------------------------------------------------------------
    ' Begin - Modify existing lines as specified
    ' -------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim tmpIdx As Integer
    
    Dim PO_LineMod As Integer
    Dim PO_LineModCount As Integer
    Dim PO_pageLineIndex As Integer
    Dim PO_pageScheduleIndex As Integer
    Dim PO_Line_Schedule As Integer
    
    Dim validLineModCount As Integer
    
    ' will not process line modifications which have a line # of 0 or less
    validLineModCount = 0
    For i = 1 To poCFQ.PO_LineModCount
        If poCFQ.PO_LineMods(i).PO_Line > 0 Then validLineModCount = validLineModCount + 1
    Next i
    
    
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: Processing line modifications (" & Debug_VarListString("LineModCount", poCFQ.PO_LineModCount, "validLineModCount", validLineModCount) & ")"
    
    If validLineModCount > 0 Then
       
        ' Expand All
        driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
        PeopleSoft_Page_WaitForProcessing driver
        
        For PO_LineMod = 1 To poCFQ.PO_LineModCount
            'PO_LineMod = PO_LineMod_SortedIdx(i)
            
            
            If poCFQ.PO_LineMods(PO_LineMod).PO_Line > 0 Then
                Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: Processing Line Mod #" & PO_LineMod & " (" & _
                    Debug_VarListString("PO Line", poCFQ.PO_LineMods(PO_LineMod).PO_Line, "Item ID", poCFQ.PO_LineMods(PO_LineMod).PO_LINE_ITEM_ID) & ")"
   
            
                ' Note: We ASSUME each line has a single schedule here
                PO_pageLineIndex = poCFQ.PO_LineMods(PO_LineMod).PO_Line - 1
                PO_pageScheduleIndex = poCFQ.PO_LineMods(PO_LineMod).PO_Line - 1
                    
                PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_INV_ITEM_ID$" & PO_pageLineIndex), _
                    poCFQ.PO_LineMods(PO_LineMod).PO_LINE_ITEM_ID, poCFQ.PO_LineMods(PO_LineMod).PO_LINE_ITEM_ID_Result
                
                If poCFQ.PO_LineMods(PO_LineMod).PO_LINE_ITEM_ID_Result.ValidationFailed Then GoTo ValidationFail
                
                
                If poCFQ.PO_LineMods(PO_LineMod).SCH_DUE_DATE > 0 Then
                    PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_DUE_DT$" & PO_pageScheduleIndex), _
                        Format(poCFQ.PO_LineMods(PO_LineMod).SCH_DUE_DATE, "mm/dd/yyyy"), poCFQ.PO_LineMods(PO_LineMod).SCH_DUE_DATE_Result
                    
                    If poCFQ.PO_LineMods(PO_LineMod).SCH_DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
                End If
                
                If poCFQ.PO_LineMods(PO_LineMod).SCH_SHIPTO_ID > 0 Then
                    PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_SHIPTO_ID$" & PO_pageScheduleIndex), _
                        CStr(poCFQ.PO_LineMods(PO_LineMod).SCH_SHIPTO_ID), poCFQ.PO_LineMods(PO_LineMod).SCH_SHIPTO_ID_Result
                    
                    If poCFQ.PO_LineMods(PO_LineMod).SCH_SHIPTO_ID_Result.ValidationFailed Then GoTo ValidationFail
                End If
            
                ' - Begin - Distribution fields (may be needed for expense items)
                If Len(poCFQ.PO_LineMods(PO_LineMod).DIST_BUSINESS_UNIT_PC) > 0 Then
                    PeopleSoft_Page_SetValidatedField driver, ("BUSINESS_UNIT_PC$" & PO_pageScheduleIndex), _
                        poCFQ.PO_LineMods(PO_LineMod).DIST_BUSINESS_UNIT_PC, poCFQ.PO_LineMods(PO_LineMod).DIST_BUSINESS_UNIT_PC_Result
                    
                    If poCFQ.PO_LineMods(PO_LineMod).DIST_BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
                End If
            
                If Len(poCFQ.PO_LineMods(PO_LineMod).DIST_PROJECT_CODE) > 0 Then
                    PeopleSoft_Page_SetValidatedField driver, ("PROJECT_ID$" & PO_pageScheduleIndex), _
                        poCFQ.PO_LineMods(PO_LineMod).DIST_PROJECT_CODE, poCFQ.PO_LineMods(PO_LineMod).DIST_PROJECT_CODE_Result
                    
                    If poCFQ.PO_LineMods(PO_LineMod).DIST_PROJECT_CODE_Result.ValidationFailed Then GoTo ValidationFail
                End If
            
                If Len(poCFQ.PO_LineMods(PO_LineMod).DIST_ACTIVITY_ID) > 0 Then
                    PeopleSoft_Page_SetValidatedField driver, ("ACTIVITY_ID$" & PO_pageScheduleIndex), _
                        poCFQ.PO_LineMods(PO_LineMod).DIST_ACTIVITY_ID, poCFQ.PO_LineMods(PO_LineMod).DIST_ACTIVITY_ID_Result
                    
                    If poCFQ.PO_LineMods(PO_LineMod).DIST_ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
                End If
            
                If poCFQ.PO_LineMods(PO_LineMod).DIST_LOCATION_ID > 0 Then
                    PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_DISTRIB_LOCATION$" & PO_pageScheduleIndex), _
                        CStr(poCFQ.PO_LineMods(PO_LineMod).DIST_LOCATION_ID), poCFQ.PO_LineMods(PO_LineMod).DIST_LOCATION_ID_Result
                    
                    If poCFQ.PO_LineMods(PO_LineMod).DIST_LOCATION_ID_Result.ValidationFailed Then GoTo ValidationFail
                End If
                ' - End - Distribution fields (may be needed for expense items)
            
            End If
            
        Next PO_LineMod
    End If
    ' -------------------------------------------------------------------
    ' End - Modify existing lines as specified
    ' -------------------------------------------------------------------
    
    'If anyLineHasValidationError Then GoTo ValidationFail
    
       
    driver.runScript "javascript:submitAction_win0(document.win0,'CALCULATE_TAXES');" ' Fix for 2.9.1.1  due to PS upgrade
    'driver.findElementById("CALCULATE_TAXES").Click
    PeopleSoft_Page_WaitForProcessing driver

    
    Dim amntStr As String
    
    ' Total
    amntStr = driver.findElementById("PO_PNLS_WRK_PO_AMT_TTL").text
    poCFQ.PO_AMNT_TOTAL = CurrencyFromString(amntStr)
    
    ' Total w/o Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_MERCH_AMT_TTL").text
    poCFQ.PO_AMNT_MERCH_TOTAL = CurrencyFromString(amntStr)
    
    ' Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_ADJ_AMT_TTL_LBL").text
    poCFQ.PO_AMNT_FTM_TOTAL = CurrencyFromString(amntStr)
    
    
    
    If DEBUG_OPTIONS.QuitBeforeSave Then
        poCFQ.GlobalError = "Debug option QuitBeforeSave enabled. Processing halted before saving. To prevent this from occurring, disable QuitBeforeSave option."
        poCFQ.HasError = True
        PeopleSoft_PurchaseOrder_CreateFromQuote = False
        Exit Function
    End If
    
    
    Dim result As Boolean
    
    result = PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver, poCFQ.BudgetCheck)
    
    If result = False Then
        poCFQ.GlobalError = poCFQ.BudgetCheck.GlobalError
        poCFQ.HasError = poCFQ.BudgetCheck.HasGlobalError
        
        PeopleSoft_PurchaseOrder_CreateFromQuote = False
        Exit Function
    End If
    
    poCFQ.PO_ID = poCFQ.BudgetCheck.PO_ID
    
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote: complete (" & Debug_VarListString("PO ID", poCFQ.PO_ID) & ")"
    
    PeopleSoft_PurchaseOrder_CreateFromQuote = True
    Exit Function
    
    
ValidationFail:
    PeopleSoft_SaveDebugInfo driver
    PeopleSoft_PurchaseOrder_CreateFromQuote = False
    Exit Function
    
ExceptionThrown:
    PeopleSoft_SaveDebugInfo driver
    poCFQ.HasError = True
    poCFQ.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_CreateFromQuote = False


End Function
Private Function PeopleSoft_PurchaseOrder_CreateFromQuote_PreCheck(poCFQ As PeopleSoft_PurchaseOrder_CreateFromQuoteParams) As Boolean

    
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote_PreCheck: PO (" & Debug_VarListString( _
        "PO REF", poCFQ.PO_Fields.PO_HDR_PO_REF, _
        "EQUOTE", poCFQ.E_QUOTE_NBR) _
        & ")"
    
    
    ' PeopleSoft_PurchaseOrder_PO_Fields_PreCheck() checks if either VENDOR ID *or* VENDOR SHORT NAME is given.
    ' When creating from quote, only creating using vendor ID is allowed. So we must explicitly check for it here.
    If poCFQ.PO_Fields.PO_HDR_VENDOR_ID = 0 Then
        poCFQ.PO_Fields.HasValidationError = True
        poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result.ValidationFailed = True
        poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result.ValidationErrorText = "VENDOR ID is required"
        poCFQ.GlobalError = "One or more PO fields failed validation"
        GoTo PreCheckFailed
    End If
    
       
     ' Check standard fields
    If PeopleSoft_PurchaseOrder_PO_Fields_PreCheck(poCFQ.PO_Fields) = False Then
        poCFQ.GlobalError = "One or more PO fields failed validation"
        GoTo PreCheckFailed
    End If
    
    ' Begin - Check required fields
    If Len(poCFQ.E_QUOTE_NBR) = 0 Then
        poCFQ.E_QUOTE_NBR_Result.ValidationFailed = True
        poCFQ.E_QUOTE_NBR_Result.ValidationErrorText = "E QUOTE NBR is required"
        poCFQ.GlobalError = "One or more PO fields failed validation"
        GoTo PreCheckFailed
    End If
    ' End - Check required fields
    
    ' Begin - validate fields: ensure single-line fields are actually a single line
    If PeopleSoft_ValidateField_IsSingleLine("E QUOTE NBR", poCFQ.E_QUOTE_NBR, poCFQ.E_QUOTE_NBR_Result) = False Then
        poCFQ.GlobalError = "One or more PO fields failed validation"
        GoTo PreCheckFailed
    End If
    ' Begin - validate fields: ensure single-line fields are actually a single line
    
    
    '---------------------------------------------------------------
    ' Begin - Perform PO Default Pre-Checks
    '---------------------------------------------------------------
    ' Note: We are checking capital chartfields only. It is assumed expense
    ' chartfields will be a line mod (since at least the EXP-* item code) will need
    ' to be specified
    
    ' Begin - Check required fields
    With poCFQ.PO_Defaults
        If .SCH_DUE_DATE = 0 Then
            .SCH_DUE_DATE_Result.ValidationErrorText = "SCH DUE DATE is required."
            .SCH_DUE_DATE_Result.ValidationFailed = True
            .HasValidationError = True
        End If
        If .DIST_CAP_SHIP_TO_ID = 0 Then
            .DIST_CAP_SHIP_TO_ID_Result.ValidationErrorText = "SHIP TO ID is required."
            .DIST_CAP_SHIP_TO_ID_Result.ValidationFailed = True
            .HasValidationError = True
        End If
        If Len(.DIST_CAP_BUSINESS_UNIT_PC) = 0 Then
            .DIST_CAP_BUSINESS_UNIT_PC_Result.ValidationErrorText = "BUSINESS_UNIT_PC is required."
            .DIST_CAP_BUSINESS_UNIT_PC_Result.ValidationFailed = True
            .HasValidationError = True
        End If
        If Len(.DIST_CAP_PROJECT_CODE) = 0 Then
            .DIST_CAP_PROJECT_CODE_Result.ValidationErrorText = "PROJECT_CODE is required."
            .DIST_CAP_PROJECT_CODE_Result.ValidationFailed = True
            .HasValidationError = True
        End If
     
        ' Note: we assume Location ID and Activity ID is not required. They may or may not be required
        ' by PeopleSoft depending on project type. If there is an error, it will be discovered during
        ' the browswer automation
        
    End With
        
    If poCFQ.PO_Defaults.HasValidationError Then
        poCFQ.PO_Defaults.GlobalError = "One or more PO fields failed validation"
        poCFQ.PO_Defaults.HasGlobalError = True
        poCFQ.GlobalError = poCFQ.PO_Defaults.GlobalError
        GoTo PreCheckFailed
    End If
    ' End - Check required fields
    
    
    ' Begin - validate fields: ensure single-line fields are actually a single line
    With poCFQ.PO_Defaults
        If PeopleSoft_ValidateField_IsSingleLine("BUSINESS_UNIT_PC", .DIST_CAP_BUSINESS_UNIT_PC, .DIST_CAP_BUSINESS_UNIT_PC_Result) = False Then .HasValidationError = True
        If PeopleSoft_ValidateField_IsSingleLine("PROJECT_CODE", .DIST_CAP_PROJECT_CODE, .DIST_CAP_PROJECT_CODE_Result) = False Then .HasValidationError = True
        If PeopleSoft_ValidateField_IsSingleLine("ACTIVITY_ID", .DIST_CAP_ACTIVITY_ID, .DIST_CAP_ACTIVITY_ID_Result) = False Then .HasValidationError = True
    End With
    
    If poCFQ.PO_Defaults.HasValidationError Then
        poCFQ.PO_Defaults.GlobalError = "One or more PO fields failed validation"
        poCFQ.PO_Defaults.HasGlobalError = True
        poCFQ.GlobalError = poCFQ.PO_Defaults.GlobalError
        GoTo PreCheckFailed
    End If
    ' End - validate fields: ensure single-line fields are actually a single line
    '---------------------------------------------------------------
    ' End - Perform PO Default Pre-Checks
    '---------------------------------------------------------------
    
    
    '---------------------------------------------------------------
    ' Begin - Perform PO Line Modification Pre-Checks
    '---------------------------------------------------------------
    Dim anyItemHasErrors As Boolean
    Dim i As Integer, j As Integer
    
    ' Begin - Check for duplicate PO lines
    anyItemHasErrors = False
    
    For i = 1 To poCFQ.PO_LineModCount
        For j = i + 1 To poCFQ.PO_LineModCount
            If poCFQ.PO_LineMods(i).PO_Line = poCFQ.PO_LineMods(j).PO_Line Then
                poCFQ.PO_LineMods(i).HasValidationError = True
                poCFQ.PO_LineMods(j).HasValidationError = True
                poCFQ.PO_LineMods(i).ItemError = "Duplicate line"
                poCFQ.PO_LineMods(j).ItemError = "Duplicate line"
                anyItemHasErrors = True
            End If
        Next j
    Next i
    
    If anyItemHasErrors Then
        poCFQ.GlobalError = "Line modifications has at least one duplicate line"
        GoTo PreCheckFailed
    End If
    ' End - Check for duplicate PO lines
    
    
    ' Begin - Validate Line modifcations
    anyItemHasErrors = False

    For i = 1 To poCFQ.PO_LineModCount
        With poCFQ.PO_LineMods(i)
            ' Valid line modification?
            If .PO_Line > 0 Then
                ' Begin - validate fields: ensure single-line fields are actually a single line
                If PeopleSoft_ValidateField_IsSingleLine("PO_LINE_ITEM_ID", .PO_LINE_ITEM_ID, .PO_LINE_ITEM_ID_Result) = False Then .HasValidationError = True
                If PeopleSoft_ValidateField_IsSingleLine("BUSINESS_UNIT_PC", .DIST_BUSINESS_UNIT_PC, .DIST_BUSINESS_UNIT_PC_Result) = False Then .HasValidationError = True
                If PeopleSoft_ValidateField_IsSingleLine("PROJECT_CODE", .DIST_PROJECT_CODE, .DIST_PROJECT_CODE_Result) = False Then .HasValidationError = True
                If PeopleSoft_ValidateField_IsSingleLine("ACTIVITY_ID", .DIST_ACTIVITY_ID, .DIST_ACTIVITY_ID_Result) = False Then .HasValidationError = True
                
                If .HasValidationError Then
                    .ItemError = "Line modification has pre-check errors"
                    anyItemHasErrors = True
                End If
                ' End - validate fields: ensure single-line fields are actually a single line
            End If
        End With
    Next i
    ' End - Validate Line modifcations
    
    If anyItemHasErrors Then
        poCFQ.GlobalError = "One or more line modification items has pre-check errors"
        GoTo PreCheckFailed
    End If
    
    
    
    '---------------------------------------------------------------
    ' Begin - Perform PO Line Modification Pre-Checks
    '---------------------------------------------------------------

    
    PeopleSoft_PurchaseOrder_CreateFromQuote_PreCheck = True
    Exit Function
    
PreCheckFailed:
    poCFQ.HasError = True
    PeopleSoft_PurchaseOrder_CreateFromQuote_PreCheck = False
    
    Debug_Print "PeopleSoft_PurchaseOrder_CreateFromQuote_PreCheck: failed! " & poCFQ.GlobalError
    
End Function
Public Function PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc(purchaseOrder As PeopleSoft_PurchaseOrder) As Boolean

    ' Auto calculates PO defaults. A field has a default value when:
    ' (1) A default value is not already specified in the calculatedDefaults structure
    ' (2) AND all PO lines/schedules/distributions have the same value
    
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc called"

    Dim calculatedDefaults As PeopleSoft_PurchaseOrder_PO_Defaults
    Dim PO_Line As Integer, PO_Line_Schedule As Integer
    
    '-------------------------------------------------------------------------------------------------------------------
    ' Begin - Calculate the PO Defaults from line/schedule values
    '-------------------------------------------------------------------------------------------------------------------
    calculatedDefaults.SCH_DUE_DATE = 0
    
    calculatedDefaults.DIST_CAP_BUSINESS_UNIT_PC = ""
    calculatedDefaults.DIST_CAP_PROJECT_CODE = ""
    calculatedDefaults.DIST_CAP_ACTIVITY_ID = ""
    calculatedDefaults.DIST_CAP_SHIP_TO_ID = 0
    calculatedDefaults.DIST_CAP_LOCATION_ID = 0
    
    calculatedDefaults.DIST_EXP_BUSINESS_UNIT_PC = ""
    calculatedDefaults.DIST_EXP_PROJECT_CODE = ""
    calculatedDefaults.DIST_EXP_ACTIVITY_ID = ""
    calculatedDefaults.DIST_EXP_SHIP_TO_ID = 0
    calculatedDefaults.DIST_EXP_LOCATION_ID = 0
    
    Dim isExpenseLine As Boolean
    Dim alreadyProcessedFirstCapitalLine As Boolean, alreadyProcessedFirstExpenseLine As Boolean
    Dim lineItemID As String
    
    alreadyProcessedFirstCapitalLine = False
    alreadyProcessedFirstExpenseLine = False
   
    For PO_Line = 1 To purchaseOrder.PO_LineCount
        For PO_Line_Schedule = 1 To purchaseOrder.PO_Lines(PO_Line).ScheduleCount
            'Check if expense line
            isExpenseLine = False
            lineItemID = UCase$(purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_ITEM_ID)
            
            If Len(lineItemID) >= Len("EXP-") Then
                If Left$(lineItemID, Len("EXP-")) = "EXP-" Then isExpenseLine = True
                
                If isExpenseLine = False And Len(lineItemID) >= Len("CR-EXP-") Then
                    If Left$(lineItemID, Len("CR-EXP-")) = "CR-EXP-" Then isExpenseLine = True
                End If
            End If
            
            
            With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule)
                ' Process fields which are capital/expense agnostic
                If PO_Line = 1 And PO_Line_Schedule = 1 Then
                    ' first line: use to set default values
                    calculatedDefaults.SCH_DUE_DATE = .ScheduleFields.DUE_DATE
                Else
                    ' different field values -> then set default to null-equivalent value
                    If calculatedDefaults.SCH_DUE_DATE <> .ScheduleFields.DUE_DATE Then calculatedDefaults.SCH_DUE_DATE = 0
                End If
                
                If isExpenseLine Then
                    If alreadyProcessedFirstExpenseLine = False Then
                        ' first capital project line: set as default value
                        alreadyProcessedFirstExpenseLine = True
                        
                        calculatedDefaults.DIST_EXP_BUSINESS_UNIT_PC = .DistributionFields.BUSINESS_UNIT_PC
                        calculatedDefaults.DIST_EXP_PROJECT_CODE = .DistributionFields.PROJECT_CODE
                        calculatedDefaults.DIST_EXP_ACTIVITY_ID = .DistributionFields.ACTIVITY_ID
                        calculatedDefaults.DIST_EXP_SHIP_TO_ID = .ScheduleFields.SHIPTO_ID
                        calculatedDefaults.DIST_EXP_LOCATION_ID = .DistributionFields.LOCATION_ID
                    Else
                        If calculatedDefaults.DIST_EXP_BUSINESS_UNIT_PC <> .DistributionFields.BUSINESS_UNIT_PC Then calculatedDefaults.DIST_EXP_BUSINESS_UNIT_PC = ""
                        If calculatedDefaults.DIST_EXP_PROJECT_CODE <> .DistributionFields.PROJECT_CODE Then calculatedDefaults.DIST_EXP_PROJECT_CODE = ""
                        If calculatedDefaults.DIST_EXP_ACTIVITY_ID <> .DistributionFields.ACTIVITY_ID Then calculatedDefaults.DIST_EXP_ACTIVITY_ID = ""
                        If calculatedDefaults.DIST_EXP_SHIP_TO_ID <> .ScheduleFields.SHIPTO_ID Then calculatedDefaults.DIST_EXP_SHIP_TO_ID = 0
                        If calculatedDefaults.DIST_EXP_LOCATION_ID <> .DistributionFields.LOCATION_ID Then calculatedDefaults.DIST_EXP_LOCATION_ID = 0
                    End If
                Else
                    If alreadyProcessedFirstCapitalLine = False Then
                        ' first capital project line: set as default value
                        alreadyProcessedFirstCapitalLine = True
                        
                        calculatedDefaults.DIST_CAP_BUSINESS_UNIT_PC = .DistributionFields.BUSINESS_UNIT_PC
                        calculatedDefaults.DIST_CAP_PROJECT_CODE = .DistributionFields.PROJECT_CODE
                        calculatedDefaults.DIST_CAP_ACTIVITY_ID = .DistributionFields.ACTIVITY_ID
                        calculatedDefaults.DIST_CAP_SHIP_TO_ID = .ScheduleFields.SHIPTO_ID
                        calculatedDefaults.DIST_CAP_LOCATION_ID = .DistributionFields.LOCATION_ID
                    Else
                        If calculatedDefaults.DIST_CAP_BUSINESS_UNIT_PC <> .DistributionFields.BUSINESS_UNIT_PC Then calculatedDefaults.DIST_CAP_BUSINESS_UNIT_PC = ""
                        If calculatedDefaults.DIST_CAP_PROJECT_CODE <> .DistributionFields.PROJECT_CODE Then calculatedDefaults.DIST_CAP_PROJECT_CODE = ""
                        If calculatedDefaults.DIST_CAP_ACTIVITY_ID <> .DistributionFields.ACTIVITY_ID Then calculatedDefaults.DIST_CAP_ACTIVITY_ID = ""
                        If calculatedDefaults.DIST_CAP_SHIP_TO_ID <> .ScheduleFields.SHIPTO_ID Then calculatedDefaults.DIST_CAP_SHIP_TO_ID = 0
                        If calculatedDefaults.DIST_CAP_LOCATION_ID <> .DistributionFields.LOCATION_ID Then calculatedDefaults.DIST_CAP_LOCATION_ID = 0
                    End If
                End If
            End With
        

            
        Next PO_Line_Schedule
    Next PO_Line
    
    
    If calculatedDefaults.DIST_CAP_PROJECT_CODE = "" Then ' Activity & Location default requires as project code default
        calculatedDefaults.DIST_CAP_ACTIVITY_ID = ""
        calculatedDefaults.DIST_CAP_LOCATION_ID = 0
    End If
    '-------------------------------------------------------------------------------------------------------------------
    ' End - Calculate the PO Defaults from line/schedule values
    '-------------------------------------------------------------------------------------------------------------------
        
    ' Set defaults to calculated defaults, but only if the default value hasn't already been specified
    If purchaseOrder.PO_Defaults.SCH_DUE_DATE = 0 Then purchaseOrder.PO_Defaults.SCH_DUE_DATE = calculatedDefaults.SCH_DUE_DATE
    
    If Len(purchaseOrder.PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC) = 0 Then purchaseOrder.PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC = calculatedDefaults.DIST_EXP_BUSINESS_UNIT_PC
    If Len(purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_CODE) = 0 Then purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_CODE = calculatedDefaults.DIST_EXP_PROJECT_CODE
    If Len(purchaseOrder.PO_Defaults.DIST_EXP_ACTIVITY_ID) = 0 Then purchaseOrder.PO_Defaults.DIST_EXP_ACTIVITY_ID = calculatedDefaults.DIST_EXP_ACTIVITY_ID
    If purchaseOrder.PO_Defaults.DIST_EXP_SHIP_TO_ID = 0 Then purchaseOrder.PO_Defaults.DIST_EXP_SHIP_TO_ID = calculatedDefaults.DIST_EXP_SHIP_TO_ID
    If purchaseOrder.PO_Defaults.DIST_EXP_LOCATION_ID = 0 Then purchaseOrder.PO_Defaults.DIST_EXP_LOCATION_ID = calculatedDefaults.DIST_EXP_LOCATION_ID
    
    If Len(purchaseOrder.PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC) = 0 Then purchaseOrder.PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC = calculatedDefaults.DIST_CAP_BUSINESS_UNIT_PC
    If Len(purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_CODE) = 0 Then purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_CODE = calculatedDefaults.DIST_CAP_PROJECT_CODE
    If Len(purchaseOrder.PO_Defaults.DIST_CAP_ACTIVITY_ID) = 0 Then purchaseOrder.PO_Defaults.DIST_CAP_ACTIVITY_ID = calculatedDefaults.DIST_CAP_ACTIVITY_ID
    If purchaseOrder.PO_Defaults.DIST_CAP_SHIP_TO_ID = 0 Then purchaseOrder.PO_Defaults.DIST_CAP_SHIP_TO_ID = calculatedDefaults.DIST_CAP_SHIP_TO_ID
    If purchaseOrder.PO_Defaults.DIST_CAP_LOCATION_ID = 0 Then purchaseOrder.PO_Defaults.DIST_CAP_LOCATION_ID = calculatedDefaults.DIST_CAP_LOCATION_ID
    
    
    
    
    Debug.Print "PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc"
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc: PO Defaults (Common): " & Debug_VarListString("SCH_DUE_DATE", purchaseOrder.PO_Defaults.SCH_DUE_DATE)
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc: PO Defaults (CAPITAL): " & Debug_VarListString( _
        "BUSINESS_UNIT_PC", purchaseOrder.PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC, _
        "PROJECT_CODE", purchaseOrder.PO_Defaults.DIST_CAP_PROJECT_CODE, _
        "ACTIVITY_ID", purchaseOrder.PO_Defaults.DIST_CAP_ACTIVITY_ID, _
        "SHIP_TO_ID", purchaseOrder.PO_Defaults.DIST_CAP_SHIP_TO_ID, _
        "LOCATION_ID", purchaseOrder.PO_Defaults.DIST_CAP_LOCATION_ID)
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc: PO Defaults (EXPENSE): " & Debug_VarListString( _
        "BUSINESS_UNIT_PC", purchaseOrder.PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC, _
        "PROJECT_CODE", purchaseOrder.PO_Defaults.DIST_EXP_PROJECT_CODE, _
        "ACTIVITY_ID", purchaseOrder.PO_Defaults.DIST_EXP_ACTIVITY_ID, _
        "SHIP_TO_ID", purchaseOrder.PO_Defaults.DIST_EXP_SHIP_TO_ID, _
        "LOCATION_ID", purchaseOrder.PO_Defaults.DIST_EXP_LOCATION_ID)
    
    
    PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc = True

End Function
Private Function PeopleSoft_PurchaseOrder_PO_Defaults_Fill(driver As SeleniumWrapper.WebDriver, ByRef PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults) As Boolean

     
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Defaults_Fill called"


    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown

    Dim isAnyDefaultSpecified As Boolean
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    
    
    isAnyDefaultSpecified = False
    
    
    If Len(PO_Defaults.LINE_CONTRACT_ID) > 0 Then isAnyDefaultSpecified = True
    
    If PO_Defaults.SCH_DUE_DATE > 0 Then isAnyDefaultSpecified = True
    
    If Len(PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_CAP_PROJECT_CODE) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_CAP_ACTIVITY_ID) > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.DIST_CAP_SHIP_TO_ID > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.DIST_CAP_LOCATION_ID > 0 Then isAnyDefaultSpecified = True
    
    If Len(PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_EXP_PROJECT_CODE) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_EXP_ACTIVITY_ID) > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.DIST_EXP_SHIP_TO_ID > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.DIST_EXP_LOCATION_ID > 0 Then isAnyDefaultSpecified = True

    
    If isAnyDefaultSpecified = False Then
        Debug_Print "PeopleSoft_PurchaseOrder_PO_Defaults_Fill: no default specified"
        PeopleSoft_PurchaseOrder_PO_Defaults_Fill = True
        Exit Function
    End If
        
    
     ' Click PO Defaults
    driver.findElementById("PO_PNLS_WRK_GOTO_DEFAULTS").Click
    'driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_GOTO_DEFAULTS');"
    PeopleSoft_Page_WaitForProcessing driver
     
     
    popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False)
     
    If popupResult.HasPopup Then
        PO_Defaults.GlobalError = "Unexpected Popup:" & popupResult.popupText
        PO_Defaults.HasGlobalError = True
    
        PeopleSoft_PurchaseOrder_PO_Defaults_Fill = False
        Exit Function
    End If
    
    'driver.waitForElementPresent "css=#PO_HDR_Z_QUOTE_NBR"
        
    ' Line Defaults
    If Len(PO_Defaults.LINE_CONTRACT_ID) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PO_DFLT_TBL_CNTRCT_ID"), PO_Defaults.LINE_CONTRACT_ID, PO_Defaults.LINE_CONTRACT_ID_Result
        If PO_Defaults.LINE_CONTRACT_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If
        
    ' Schedule
    If PO_Defaults.SCH_DUE_DATE > 0 Then
        PeopleSoft_Page_SetValidatedField driver:=driver, fieldElementID:=("PO_DFLT_TBL_DUE_DT"), fieldValue:=Format(PO_Defaults.SCH_DUE_DATE, "mm/dd/yyyy"), _
            validationResult:=PO_Defaults.SCH_DUE_DATE_Result, _
            expectedPopupContents:="*Due Date selected is a weekend or a public holiday*"
        
        If PO_Defaults.SCH_DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
    End If

    ' Expand Distributions
    driver.runScript "javascript:submitAction_win0(document.win0,'PO_DFLT_DISTRIB$htab$0');"
    PeopleSoft_Page_WaitForProcessing driver
    

    ' Fill capital distributions
    If Len(PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("BUSINESS_UNIT_PC$0"), PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC, PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC_Result
        If PO_Defaults.DIST_CAP_BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
    End If
    If Len(PO_Defaults.DIST_CAP_PROJECT_CODE) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PROJECT_ID$0"), PO_Defaults.DIST_CAP_PROJECT_CODE, PO_Defaults.DIST_CAP_PROJECT_CODE_Result
        If PO_Defaults.DIST_CAP_PROJECT_CODE_Result.ValidationFailed Then GoTo ValidationFail
    End If
    If Len(PO_Defaults.DIST_CAP_ACTIVITY_ID) > 0 Then
        PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID driver, _
            "ACTIVITY_ID$0", PO_Defaults.DIST_CAP_ACTIVITY_ID, PO_Defaults.DIST_CAP_ACTIVITY_ID_Result, "ACTIVITY_ID$prompt$0"
        If PO_Defaults.DIST_CAP_ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If
    If PO_Defaults.DIST_CAP_SHIP_TO_ID > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PO_DFLT_DISTRIB_SHIPTO_ID$0"), CStr(PO_Defaults.DIST_CAP_SHIP_TO_ID), PO_Defaults.DIST_CAP_SHIP_TO_ID_Result
        If PO_Defaults.DIST_CAP_SHIP_TO_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If
    
    If PO_Defaults.DIST_CAP_LOCATION_ID > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("LOCATION$0"), CStr(PO_Defaults.DIST_CAP_LOCATION_ID), PO_Defaults.DIST_CAP_LOCATION_ID_Result
        If PO_Defaults.DIST_CAP_LOCATION_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If
    
    ' Fill expense distributions
    If Len(PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("Z_EXP_PC_BU$0"), PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC, PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC_Result
        If PO_Defaults.DIST_EXP_BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
    End If
    If Len(PO_Defaults.DIST_EXP_PROJECT_CODE) > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PROJECT_ID_2$0"), PO_Defaults.DIST_EXP_PROJECT_CODE, PO_Defaults.DIST_EXP_PROJECT_CODE_Result
        If PO_Defaults.DIST_EXP_PROJECT_CODE_Result.ValidationFailed Then GoTo ValidationFail
    End If
    If Len(PO_Defaults.DIST_EXP_ACTIVITY_ID) > 0 Then
        PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID driver, _
            "ACTIVITY_ID_2$0", PO_Defaults.DIST_EXP_ACTIVITY_ID, PO_Defaults.DIST_EXP_ACTIVITY_ID_Result, "ACTIVITY_ID_2$prompt$0"
        If PO_Defaults.DIST_EXP_ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If
    If PO_Defaults.DIST_EXP_SHIP_TO_ID > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PO_DFLT_DISTRIB_SHIPTO_ID_DEFAULT$0"), CStr(PO_Defaults.DIST_EXP_SHIP_TO_ID), PO_Defaults.DIST_EXP_SHIP_TO_ID_Result
        If PO_Defaults.DIST_EXP_SHIP_TO_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If
    
    If PO_Defaults.DIST_EXP_LOCATION_ID > 0 Then
        PeopleSoft_Page_SetValidatedField driver, ("PO_DFLT_DISTRIB_Z_EXP_CF1$0"), CStr(PO_Defaults.DIST_EXP_LOCATION_ID), PO_Defaults.DIST_EXP_LOCATION_ID_Result
        If PO_Defaults.DIST_EXP_LOCATION_ID_Result.ValidationFailed Then GoTo ValidationFail
    End If


    ' TODO: Need to implement expense chartfields:
    ' Exp Cost Center: PO_DFLT_DISTRIB_Z_EXP_DEPTID$0
    ' Exp Product Code: PO_DFLT_DISTRIB_Z_EXP_PROD$0
    
    
    ' Extract some fields from page
    If Len(PO_Defaults.DIST_CAP_PROJECT_CODE) > 0 Then _
        PO_Defaults.DIST_CAP_PROJECT_DESC = driver.findElementById("Z_PROJ_CAP_VW_DESCR$0").text
    If Len(PO_Defaults.DIST_EXP_PROJECT_CODE) > 0 Then _
        PO_Defaults.DIST_EXP_PROJECT_DESC = driver.findElementById("Z_PROJ_EXP_VW1_DESCR$0").text
    
    
    ' Click save
    driver.findElementById("#ICSave").Click
    'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
        
    
    
    PeopleSoft_PurchaseOrder_PO_Defaults_Fill = True
    Exit Function
  
  
ValidationFail:
    PO_Defaults.HasValidationError = True
    PeopleSoft_PurchaseOrder_PO_Defaults_Fill = False
    Exit Function
    

ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_PurchaseOrder_PO_Defaults_Fill Exception: " & Err.Description, Err.Helpfile, Err.HelpContext
    

End Function
Private Function PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page(driver As SeleniumWrapper.WebDriver, poFields As PeopleSoft_PurchaseOrder_Fields) As Boolean
    
    Debug_Print "PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page called"
    
    
    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown
    
    Dim we As WebElement
    Dim weCollection As WebElementCollection
    
    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult
    
    ' -------------------------------------------------------------------
    ' Begin - Comments Section
    ' -------------------------------------------------------------------
    If Len(poFields.PO_HDR_COMMENTS) > 0 Or Len(poFields.Quote_Attachment_FilePath) > 0 Then
        driver.findElementById("COMM_WRK1_COMMENTS_PB").Click
        PeopleSoft_Page_WaitForProcessing driver
         
         
         ' Fill in PO Approver -> Now disabled.
        'If False Then
        '    driver.waitForElementPresent "css=#PO_HDR_Z_SUG_APPRVR"
        '
        '    PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_Z_SUG_APPRVR"), CStr(poFields.PO_HDR_APPROVER_ID), poFields.PO_HDR_APPROVER_ID_Result
        '    If poFields.PO_HDR_APPROVER_ID_Result.ValidationFailed Then GoTo ValidationFail
        'End If
        
        If Len(poFields.PO_HDR_COMMENTS) > 0 Then
            Debug_Print "PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page: Entering comments"
            
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").Clear
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").SendKeys poFields.PO_HDR_COMMENTS
            PeopleSoft_Page_WaitForProcessing driver
        End If
        
        
        
        ' If quote file provided -> attach quote
        If Len(poFields.Quote_Attachment_FilePath) > 0 Then
            Debug_Print "PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page: attaching quote: " & poFields.Quote_Attachment_FilePath
            
            driver.findElementById("PV_ATTACH_WRK_SCM_UPLOAD$0").Click
            'driver.runScript "javascript:submitAction_win2(document.win2, 'PV_ATTACH_WRK_SCM_UPLOAD$0');"
            PeopleSoft_Page_WaitForProcessing driver
            
            
            'driver.Wait 1000
            
            Dim modalPopupIndex As Integer
            
            
            modalPopupIndex = PeopleSoft_Page_CheckForModal(driver)
                
            driver.switchToFrame "ptModFrame_" & modalPopupIndex
            
            driver.findElementByName("#ICOrigFileName").SendKeys poFields.Quote_Attachment_FilePath
            PeopleSoft_Page_WaitForProcessing driver
                        
            ' CLick upload button and wait for processing
            ' <input type="button" class="PSPUSHBUTTON" value="Upload" onclick="doModalMFormSubmit_win0(this.form,'#ICOK');" psaccesskey="\">
            'driver.findElementByXPath(".//form[@name='win2']/descendant::input[@value='Upload']").Click
            driver.runScript "javascript: var elems = document.getElementsByName('#ICOrigFileName'); doModalMFormSubmit_win0(elems[0].form,'#ICOK');"
            driver.selectFrame "relative=top"
            PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG ' May need some time to upload file here ?
            
            ' Check for file upload popup: Attachment failed to upload
            popupCheckResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False)
            
            If popupCheckResult.HasPopup Then
                poFields.Quote_Attachment_FilePath_Result.ValidationFailed = True
                poFields.Quote_Attachment_FilePath_Result.ValidationErrorText = "Attach quote failed: Unexpected popup: " & popupCheckResult.popupText
                GoTo ValidationFail
            End If
                        
            
            ' Check if file was successfully uploaded
            Dim uploadedFilename As String
            Set we = driver.findElementById("PV_ATTACH_WRK_ATTACHUSERFILE$0")
            uploadedFilename = we.text
            
            If Len(Trim(UCase(uploadedFilename))) = 0 Then
                ' Need a better method here than raising an exception
                poFields.Quote_Attachment_FilePath_Result.ValidationFailed = True
                poFields.Quote_Attachment_FilePath_Result.ValidationErrorText = "Attach quote failed: could not verify if quote was sucessfully uploaded."
                GoTo ValidationFail
            End If
        End If
        
        
        
    
        driver.findElementById("#ICSave").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
        PeopleSoft_Page_WaitForProcessing driver
    
    End If
    ' -------------------------------------------------------------------
    ' End - Comments Section
    ' -------------------------------------------------------------------
    
    PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page = True
    Exit Function

  
ValidationFail:
    poFields.HasValidationError = True
    PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page = False
    Exit Function

ExceptionThrown:
    PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page = False
    Err.Raise Err.Number, Err.Source, "PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page Exception: " & Err.Description, Err.Helpfile, Err.HelpContext

End Function

Private Function PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID(driver As SeleniumWrapper.WebDriver, fieldElementID As String, fieldValue As String, ByRef validationResult As PeopleSoft_Field_ValidationResult, promptElementID As String) As String

    'On Error GoTo ErrOccurred
    
    
    
    Dim activityListString As String: activityListString = ""

    PeopleSoft_Page_SetValidatedField driver, fieldElementID, fieldValue, validationResult


    
    If validationResult.ValidationFailed Then
        If validationResult.ValidationErrorText Like "*Invalid value*" Then
            Dim tmpFVR As PeopleSoft_Field_ValidationResult
            
            'Clear field
            PeopleSoft_Page_SetValidatedField driver, fieldElementID, "", tmpFVR, False
        
        
            
            ' Simulates clicking on the spyglass. Extracts the activity IDs from the popup.
            driver.findElementById(promptElementID).Click
            PeopleSoft_Page_WaitForProcessing driver
            
            Dim activities() As Variant
            activities = PeopleSoft_Page_ModalWindow_ExtractSearchTableContents(driver, "Activity")
            
            
            validationResult.ValidationErrorText = "Invalid activity ID. Valid values for this project: " & Join(activities, ",")
   
    
        End If
    End If
    
    PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID = activityListString
    
    Exit Function
    
ErrOccurred:

    validationResult.ValidationErrorText = validationResult.ValidationErrorText & vbCrLf & vbCrLf & "PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID Exception: " & Err.Description
    
    

End Function
' PeopleSoft_Page_ModalWindow_ExtractSearchTableContents: Utility function to extract contents of a PS search table from a modal window
Private Function PeopleSoft_Page_ModalWindow_ExtractSearchTableContents(driver As SeleniumWrapper.WebDriver, Optional returnColumns As Variant) As Variant()

    
    If IsEmpty(returnColumns) Then
        returnColumns = Array()
    Else
        If Not IsArray(returnColumns) Then returnColumns = Array(returnColumns)
    End If
    
    Dim returnColNames() As String
    Dim returnColNums() As Long ' Column #s for returnColumns
    Dim returnColumnCount As Long
    
    returnColumnCount = UBound(returnColumns) - LBound(returnColumns) + 1

    
    Dim modalIndex As Integer
    Dim By As New SeleniumWrapper.By
    
    modalIndex = PeopleSoft_Page_CheckForModal(driver)

    If modalIndex < 0 Then Exit Function ' Modal window not found
    
    driver.switchToFrame "ptModFrame_" & modalIndex
    
    If Not PeopleSoft_Page_ElementExists(driver, By.id("PTSRCHRESULTS")) Then Exit Function ' No search table
    

    Dim i As Long, j As Long
    Dim we As WebElement, weCollection As WebElementCollection
    Dim pageColName As String
    Dim columnCount As Long, rowCount As Long
    
    ' Get Columns
    Set weCollection = driver.findElementsByXPath(".//table[@id='PTSRCHRESULTS']/descendant::th[@class='PSSRCHRESULTSHDR']")
    columnCount = weCollection.Count - 1
    
    ' Begin - Populate returnColNums() based on column names
    If returnColumnCount > 0 Then
        ReDim returnColNums(1 To returnColumnCount) As Long
        
        For i = 1 To columnCount
            pageColName = weCollection.Item(i).text
            
            ' See if this column is in the list of return columns
            For j = 1 To returnColumnCount
                If StrComp(returnColumns(j - LBound(returnColumns) - 1), pageColName, vbTextCompare) = 0 Then
                    ' found it
                    returnColNums(j) = i
                End If
            Next j
        Next i
        
        ' Check to see if one or more return columns could not be found
        Dim missingColNamesStr As String
        
        For i = 1 To returnColumnCount
            If returnColNums(i) = 0 Then missingColNamesStr = missingColNamesStr & returnColumns(i - LBound(returnColumns) + 1) & ","
        Next i
        
        If Len(missingColNamesStr) > 0 Then
            missingColNamesStr = Left$(missingColNamesStr, Len(missingColNamesStr) - 1) ' remove extra ,
            Err.Raise -1, , "Missing columns in modal window search table: " & missingColNamesStr
        End If
    Else
        ' return ALL columns
        returnColumnCount = columnCount
        
        ReDim returnColNames(1 To columnCount) As String
        ReDim returnColNums(1 To columnCount) As Long
        
        For i = 1 To columnCount
            returnColNames(i) = weCollection.Item(i).text
            returnColNums(i) = i
        Next i
        
        ' Set ByRef reference parameter to colunm names (returned to calling function)
        returnColumns = returnColNames
    End If
    

    rowCount = driver.getXpathCount(".//table[@id='PTSRCHRESULTS']/descendant::tr") - 1
    
    Dim ret() As Variant
    
    If returnColumnCount = 1 Then
        ' Return 1D array
        ReDim ret(1 To rowCount) As Variant
    Else
        ' Return 2D array
        ReDim ret(1 To rowCount, 1 To returnColumnCount) As Variant
    End If
    
    For i = 0 To rowCount - 1
        For j = 1 To returnColumnCount
            Set we = driver.findElementById("RESULT" & (returnColNums(j) + 2) & "$" & i)
            
            If returnColumnCount = 1 Then
                ret(i + 1) = we.text
            Else
                ret(i + 1, j) = we.text
            End If
        Next j
    Next i
    
    
    driver.selectFrame "relative=top"
    driver.runScript "javascript:doCloseModal(" & modalIndex & ");"
    
    PeopleSoft_Page_ModalWindow_ExtractSearchTableContents = ret


End Function
Public Function PeopleSoft_PurchaseOrder_Fill_PO_Line(driver As SeleniumWrapper.WebDriver, ByRef purchaseOrder As PeopleSoft_PurchaseOrder, PO_Line As Integer, ByVal PO_pageScheduleIndex As Integer) As Boolean

    Debug.Assert PO_Line > 0 And PO_Line <= purchaseOrder.PO_LineCount
    

    
    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown
    
    
    Dim PO_Line_Schedule As Integer, PO_Line_ScheduleCount As Integer
    
    
    ' Begin - Enter Line Fields
    With purchaseOrder.PO_Lines(PO_Line)
        PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_INV_ITEM_ID$" & (PO_Line - 1)), _
            .LineFields.PO_LINE_ITEM_ID, .LineFields.PO_LINE_ITEM_ID_Result
        
        If .LineFields.PO_LINE_ITEM_ID_Result.ValidationFailed Then GoTo ValidationFail
        
        
        Dim tmpValResult As PeopleSoft_Field_ValidationResult
        
        PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_DESCR254_MIXED$" & (PO_Line - 1)), _
                .LineFields.PO_LINE_DESC, tmpValResult
            
        If tmpValResult.ValidationFailed Then GoTo ValidationFail
    
    End With
    ' End - Enter Line Fields
    
    
    PO_Line_ScheduleCount = purchaseOrder.PO_Lines(PO_Line).ScheduleCount
    
    For PO_Line_Schedule = 1 To PO_Line_ScheduleCount
        ' Begin - Enter Schedule Fields
        
        With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule)
            ' PENDING DELETION: ' Due date set or PO default due date is not set
            ' PENDING DELETION: If .ScheduleFields.DUE_DATE > 0 Or purchaseOrder.PO_Defaults.SCH_DUE_DATE = 0 Then
            
            ' Due date set or PO default due date is not set
            If .ScheduleFields.DUE_DATE > 0 Then
                PeopleSoft_Page_SetValidatedField driver:=driver, fieldElementID:=("PO_LINE_SHIP_DUE_DT$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                    fieldValue:=Format(.ScheduleFields.DUE_DATE, "mm/dd/yyyy"), _
                    validationResult:=.ScheduleFields.DUE_DATE_Result, _
                    expectedPopupContents:="*Due Date selected is a weekend or a public holiday*"
                    
    
                If .ScheduleFields.DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
            End If
            
        
            'Debug.Print
            PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_SHIPTO_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                CStr(.ScheduleFields.SHIPTO_ID), .ScheduleFields.SHIPTO_ID_Result
            
            If .ScheduleFields.SHIPTO_ID_Result.ValidationFailed Then GoTo ValidationFail
            
            
            If .ScheduleFields.QTY > 0 Then
                PeopleSoft_Page_SetValidatedField driver:=driver, fieldElementID:=("PO_LINE_SHIP_QTY_PO$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                    fieldValue:=CStr(.ScheduleFields.QTY), validationResult:=.ScheduleFields.QTY_Result, _
                    expectedPopupContents:="*Vendor item price is not available*"
                
                ' ^- Ignore: Vendor item price is not available - use item standard price.....The vendor item price was not setup, or the correspondin....
                
                If .ScheduleFields.QTY_Result.ValidationFailed Then GoTo ValidationFail
            End If
            
            
            ' Retrieve price Dim priceStr As String
            Dim priceStr As String, priceVal As Currency
            
            priceStr = driver.findElementById("PO_LINE_SHIP_PRICE_PO$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)).getAttribute("value")
            priceVal = CurrencyFromString(priceStr)
            
            ' Price given? Change price if PO default price is different from what is given. Otherwise, retrieve the price from page
            If .ScheduleFields.PRICE > 0 Or True Then
                If .ScheduleFields.PRICE <> priceVal Then
                
                    PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_PRICE_PO$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                        CStr(.ScheduleFields.PRICE), .ScheduleFields.PRICE_Result
                    
                    If .ScheduleFields.PRICE_Result.ValidationFailed Then GoTo ValidationFail
                    
    
                End If
            Else
                 .ScheduleFields.PRICE = priceVal
            End If
            ' End - Enter Schedule Fields
            
            ' Begin - Enter Distribution Fields
            
            PeopleSoft_Page_SetValidatedField driver, _
                ("BUSINESS_UNIT_PC$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                .DistributionFields.BUSINESS_UNIT_PC, .DistributionFields.BUSINESS_UNIT_PC_Result
            
            If .DistributionFields.BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
            
            
            PeopleSoft_Page_SetValidatedField driver, _
                ("PROJECT_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                .DistributionFields.PROJECT_CODE, .DistributionFields.PROJECT_CODE_Result
            
            If .DistributionFields.PROJECT_CODE_Result.ValidationFailed Then GoTo ValidationFail
            
            
            PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID driver, _
                    ("ACTIVITY_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                    .DistributionFields.ACTIVITY_ID, .DistributionFields.ACTIVITY_ID_Result, _
                    ("ACTIVITY_ID$prompt$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1))
            
            If .DistributionFields.ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
            
            
            If .DistributionFields.LOCATION_ID > 0 Then
                PeopleSoft_Page_SetValidatedField driver, _
                    ("PO_LINE_DISTRIB_LOCATION$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                    CStr(.DistributionFields.LOCATION_ID), .DistributionFields.LOCATION_ID_Result
                
                If .DistributionFields.LOCATION_ID_Result.ValidationFailed Then GoTo ValidationFail
            End If
                
            ' TODO: Additional Chartfields for expenses
            '   Cost Center: DEPTID$X
            '   Product Code: PRODUCT$X
            
            ' Extract fields from page
            If Len(.DistributionFields.PROJECT_CODE) > 0 Then _
                .DistributionFields.PROJECT_DESC = driver.findElementById("Z_PROJ_BUGL_FVW_DESCR$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)).text
         
        End With
        
        PO_pageScheduleIndex = PO_pageScheduleIndex + 1
    Next PO_Line_Schedule
    
    
            
    PeopleSoft_PurchaseOrder_Fill_PO_Line = True
    Exit Function
  
  
ValidationFail:
    purchaseOrder.PO_Lines(PO_Line).HasValidationError = True

    PeopleSoft_PurchaseOrder_Fill_PO_Line = False
    
    Exit Function
    
    
ExceptionThrown:
    purchaseOrder.HasError = True
    purchaseOrder.GlobalError = "PeopleSoft_PurchaseOrder_Fill_PO_Line Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_Fill_PO_Line = False
    
    

End Function


Public Function PeopleSoft_ChangeOrder_Process(ByRef session As PeopleSoft_Session, ByRef poChangeOrder As PeopleSoft_PurchaseOrder_ChangeOrder) As Boolean
    
    Debug_Print "PeopleSoft_ChangeOrder_Process called (" & Debug_VarListString("PO ID", poChangeOrder.PO_ID) & ")"
    
    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As New SeleniumWrapper.WebDriver
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    Dim i As Long
    
    
    If PeopleSoft_ChangeOrder_Process_PreCheck(poChangeOrder) = False Then
        ' GlobalError and HasError set by PreCheck
        GoTo ChangeOrderFailed
    End If
    
    
    
    PeopleSoft_Login session
    
    If Not session.loggedIn Then
        poChangeOrder.GlobalError = "Logon Error: " & session.LogonError
        poChangeOrder.HasError = True
        
        PeopleSoft_ChangeOrder_Process = False
        Exit Function
    End If

    
    Set driver = session.driver
    
    
    If PeopleSoft_NavigateTo_ExistingPO(session, poChangeOrder.PO_BU, poChangeOrder.PO_ID) = False Then
        poChangeOrder.GlobalError = "PO navigation failed"
        poChangeOrder.HasError = True
        
        GoTo ChangeOrderFailed
    End If
    
    
    
    
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[contains(text(),'Purchase order being processed by batch programs')]")) Then
        Debug_Print "PeopleSoft_ChangeOrder_Process: Check for batch processing: In-Use (fail)"
        poChangeOrder.GlobalError = "PO is currently being processed by other programs. Try again later."
        poChangeOrder.HasError = True
        
        GoTo ChangeOrderFailed
    End If
    
    Debug_Print "PeopleSoft_ChangeOrder_Process: Check for batch processing: OK"
    
    ' -------------------------------------------------------------------
    ' Begin - Comments Section
    ' -------------------------------------------------------------------
    Debug_Print "PeopleSoft_ChangeOrder_Process: Comments section"
    
    If poChangeOrder.PO_HDR_FLG_SEND_TO_VENDOR <> PeopleSoft_Page_CheckboxAction.KeepExistingValue Then
        Dim weCmtsLink As SeleniumWrapper.WebElement
        
        If PeopleSoft_Page_ElementExists(driver, By.id("COMM_WRK1_COMMENTS1_PB")) Then
            ' Edit Comments
            'driver.executeScript "javascript:submitAction_win0(document.win0,'COMM_WRK1_COMMENTS1_PB');"
            Set weCmtsLink = driver.findElementById("COMM_WRK1_COMMENTS1_PB")
        Else
            ' Add Comments
            'driver.executeScript "javascript:submitAction_win0(document.win0,'COMM_WRK1_COMMENTS_PB');"
            Set weCmtsLink = driver.findElementById("COMM_WRK1_COMMENTS_PB")
        End If
        
        weCmtsLink.Click
        PeopleSoft_Page_WaitForProcessing driver
        
        
        

        driver.waitForElementPresent "css=#PO_HDR_Z_SUG_APPRVR"
        
        
        
        Dim chkElem As SeleniumWrapper.WebElement
        
        Set chkElem = driver.findElementById("PO_COMMENTS_PUBLIC_FLG$0")
        
        If chkElem.Selected = True And poChangeOrder.PO_HDR_FLG_SEND_TO_VENDOR = SetAsUnchecked Then
            ' checked but should be unchecked
            chkElem.Click
        ElseIf chkElem.Selected = False And poChangeOrder.PO_HDR_FLG_SEND_TO_VENDOR = SetAsChecked Then
            ' unchecked but should be checked
            chkElem.Click
        End If
    
    
        driver.findElementById("#ICSave").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
        PeopleSoft_Page_WaitForProcessing driver
        
        
        ' Check if approver changed. Hit OK if so
        'If PeopleSoft_Page_ElementExists(driver, By.ID("PSTEXT")) Then
        '    Dim msgText As String
        '
        '    msgText = driver.findElementById("PSTEXT").Text
        '
        '    If InStr(1, msgText, "has assigned delegation") > 0 Then ' Warning -- The user Last1,First1 (1234567)  has assigned delegation to Last2,First2 (7654321) . (23200,238) This will result in Suggested approver being updated accordingly
        '        driver.findElementById("#ICOK").Click
        '
        '        PeopleSoft_Page_WaitForProcessing driver
        '    End If
        'End If
        
    End If
    ' -------------------------------------------------------------------
    ' End - Comments Section
    ' -------------------------------------------------------------------
  
    
    ' -------------------------------------------------------------------
    ' Begin - PO Defaults Section
    ' -------------------------------------------------------------------
    Dim result As Boolean
    Dim modifyDefaults As Boolean
    
    modifyDefaults = poChangeOrder.PO_DUE_DATE > 0
    
    Debug_Print "PeopleSoft_ChangeOrder_Process: PO Default Section (" & Debug_VarListString("modifyDefaults", modifyDefaults) & ")"
    
    If modifyDefaults Then
        
        ' Re-use code for filling PO defaults, except only use the due date field
        Dim PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults
        
        PO_Defaults.SCH_DUE_DATE = poChangeOrder.PO_DUE_DATE
    
        result = PeopleSoft_PurchaseOrder_PO_Defaults_Fill(driver, PO_Defaults)
        
        Debug_Print "PeopleSoft_ChangeOrder_Process: PO Default Fill Result (" & Debug_VarListString("Result", result) & ")"
        
        poChangeOrder.PO_DUE_DATE_Result = PO_Defaults.SCH_DUE_DATE_Result
        
        If result = False Then
            poChangeOrder.HasError = True
            If PO_Defaults.HasGlobalError Then poChangeOrder.GlobalError = PO_Defaults.GlobalError
            
            GoTo ChangeOrderFailed
        End If
        
        ' Suppress expected popups - may be multiple
        Do
            popupResult = PeopleSoft_Page_CheckForPopup(driver, acknowledgePopup:=True, raiseErrorIfUnexpected:=False, _
                            expectedContent:=Array( _
                                "*Default values will be applied only to PO lines that are not received or invoiced*", _
                                "*This action will create a change order*", _
                                "*This PO has been dispatched, add/delete/change a line or schedule will create a change order*" _
                                ))
            If popupResult.HasPopup = False Then Exit Do
            
            If popupResult.IsExpected = False Then
                poChangeOrder.HasError = True
                poChangeOrder.GlobalError = "Unexpected Popup: " & popupResult.popupText
        
                GoTo ChangeOrderFailed
            End If
            
            PeopleSoft_Page_WaitForProcessing driver
        Loop While popupResult.HasPopup
        
    End If
    ' -------------------------------------------------------------------
    ' End - PO Defaults Section
    ' -------------------------------------------------------------------
    
    ' Calculate number of valid change order items
    Dim validChangeOrderItemCount As Long: validChangeOrderItemCount = 0
    
    For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
        If poChangeOrder.PO_ChangeOrder_Items(i).PO_Line > 0 And poChangeOrder.PO_ChangeOrder_Items(i).PO_Schedule > 0 Then validChangeOrderItemCount = validChangeOrderItemCount + 1
    Next i
    
    
    ' Santiy check: either default value was modified OR there is at least one change item. If not, then
    ' do not continue as there is no action being taken
    If modifyDefaults = False And validChangeOrderItemCount = 0 Then
        poChangeOrder.HasError = True
        poChangeOrder.GlobalError = "Invalid change order request: No PO defaults modified or valid change order items."
        GoTo ChangeOrderFailed
    End If
    

    If validChangeOrderItemCount > 0 Then
        result = PeopleSoft_ChangeOrder_ProcessLineScheduleItems(driver, poChangeOrder)
        If poChangeOrder.HasError Then GoTo ChangeOrderFailed
    
        ' Check for individual item failures:
        For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
            If poChangeOrder.PO_ChangeOrder_Items(i).HasError Then GoTo ChangeOrderFailed
            If poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
        Next i
    End If
    
    
    ' TODO: Check if change made (e.g., due date was actually changed). Exit if not
    
    
        
    If DEBUG_OPTIONS.QuitBeforeSave Then
        poChangeOrder.GlobalError = "Debug option QuitBeforeSave enabled. Processing halted before saving. To prevent this from occurring, disable QuitBeforeSave option."
        poChangeOrder.HasError = True
        PeopleSoft_ChangeOrder_Process = False
        Exit Function
    End If
    
    
    
    
    driver.findElementById("PO_KK_WRK_PB_BUDGET_CHECK").Click
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_INFINITE
    
    
    If PeopleSoft_Page_ElementExists(driver, By.id("PO_CHNG_REASON_COMMENTS$0")) Then
        
        driver.findElementById("PO_CHNG_REASON_COMMENTS$0").Clear
        driver.findElementById("PO_CHNG_REASON_COMMENTS$0").SendKeys poChangeOrder.ChangeReason
        
        
        
        driver.findElementById("#ICSave").Click
        PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_INFINITE
        
        popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True)
        
        If popupResult.HasPopup Then
            poChangeOrder.HasError = True
            poChangeOrder.GlobalError = "Unexpected popup: " & popupResult.popupText
            GoTo ChangeOrderFailed
        End If
    End If
    
    PeopleSoft_ChangeOrder_Process = True
    Exit Function
    
ValidationFail:
ChangeOrderFailed:
    poChangeOrder.HasError = True
    PeopleSoft_ChangeOrder_Process = False
    PeopleSoft_SaveDebugInfo driver
    Exit Function
    
ExceptionThrown:
    poChangeOrder.HasError = True
    poChangeOrder.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_ChangeOrder_Process = False
    PeopleSoft_SaveDebugInfo driver

End Function
Private Function PeopleSoft_ChangeOrder_Process_PreCheck(poChangeOrder As PeopleSoft_PurchaseOrder_ChangeOrder) As Boolean

    
    Debug_Print "PeopleSoft_ChangeOrder_Process_PreCheck: Change Order (" & Debug_VarListString("PO ID", poChangeOrder.PO_ID, "PO DUE DATE", poChangeOrder.PO_DUE_DATE, "SendToVendor", poChangeOrder.PO_HDR_FLG_SEND_TO_VENDOR) & ")"
    
    Dim i As Long, j As Long
    
    For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
        With poChangeOrder.PO_ChangeOrder_Items(i)
            Debug_Print "PeopleSoft_ChangeOrder_Process_PreCheck: Change Order Item: (" & Debug_VarListString("PO Line", .PO_Line, "PO Schedule", .PO_Schedule, "SCH_DUE_DATE", .SCH_DUE_DATE) & ")"
        End With
    Next i
    
    
    ' Begin - Check required fields
    If Len(poChangeOrder.PO_BU) = 0 Then
        poChangeOrder.PO_BU_Result.ValidationErrorText = "PO BU is required."
        poChangeOrder.PO_BU_Result.ValidationFailed = True
        poChangeOrder.HasError = True
    End If
    If Len(poChangeOrder.PO_ID) = 0 Then
        poChangeOrder.PO_ID_Result.ValidationErrorText = "PO BU is required."
        poChangeOrder.PO_ID_Result.ValidationFailed = True
        poChangeOrder.HasError = True
    End If
    
    If poChangeOrder.HasError Then
        poChangeOrder.GlobalError = "One or more required fields are missing"
        GoTo PreCheckFailed
    End If
    
    ' Change reason is not validated return in global error
    If Len(poChangeOrder.ChangeReason) = 0 Then
        poChangeOrder.GlobalError = "Change Reason is required"
        GoTo PreCheckFailed
    End If
    ' End - Check required fields
    

    ' Begin - validate fields: ensure single-line fields are actually a single line
    If PeopleSoft_ValidateField_IsSingleLine("PO BU", poChangeOrder.PO_BU, poChangeOrder.PO_BU_Result) = False Then poChangeOrder.HasError = True
    If PeopleSoft_ValidateField_IsSingleLine("PO ID", poChangeOrder.PO_ID, poChangeOrder.PO_ID_Result) = False Then poChangeOrder.HasError = True
    
    If poChangeOrder.HasError Then
        poChangeOrder.GlobalError = "One or more fields failed pre-check"
        GoTo PreCheckFailed
    End If
    ' Begin - validate fields: ensure single-line fields are actually a single line
    

    Dim hasPoDefaultFieldChange As Boolean
    Dim hasValidChangeOrderItem As Boolean
    Dim anyItemHasErrors As Boolean
    
    hasPoDefaultFieldChange = False
    hasValidChangeOrderItem = False
    anyItemHasErrors = True = False
    
    If poChangeOrder.PO_DUE_DATE > 0 Then hasPoDefaultFieldChange = True


    ' Begin - Check for duplicate PO lines/schedules in ChangeOrderItems
    anyItemHasErrors = False
    
    For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
        For j = i + 1 To poChangeOrder.PO_ChangeOrder_ItemCount
            If poChangeOrder.PO_ChangeOrder_Items(i).PO_Line = poChangeOrder.PO_ChangeOrder_Items(j).PO_Line _
              And poChangeOrder.PO_ChangeOrder_Items(i).PO_Schedule = poChangeOrder.PO_ChangeOrder_Items(j).PO_Schedule Then
                poChangeOrder.PO_ChangeOrder_Items(i).HasError = True
                poChangeOrder.PO_ChangeOrder_Items(j).HasError = True
                poChangeOrder.PO_ChangeOrder_Items(i).ItemError = "Duplicate line/schedule"
                poChangeOrder.PO_ChangeOrder_Items(j).ItemError = "Duplicate line/schedule"
                anyItemHasErrors = True
            End If
        Next j
    Next i
    
    If anyItemHasErrors Then
        poChangeOrder.HasError = True
        poChangeOrder.GlobalError = "Duplicate line/schedule in one more change order items"
        
        GoTo PreCheckFailed
    End If
    ' End - Check for duplicate PO lines/schedules in ChangeOrderItems
    
    
    ' Begin - Validate ChangeOrderItems
    anyItemHasErrors = False

    For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
        With poChangeOrder.PO_ChangeOrder_Items(i)
            ' Valid change order line?
            If .PO_Line > 0 Then
                hasValidChangeOrderItem = True
                
                If .PO_Schedule < 1 Then
                    .HasError = True
                    .ItemError = .ItemError & "PO Schedule must be zero or greater" & vbCrLf
                    anyItemHasErrors = True
                End If
                ' Does the line item have any changes?
                If .SCH_DUE_DATE = 0 Then
                    .HasError = True
                    .ItemError = .ItemError & "Change Order Item does not specify any changes." & vbCrLf
                    anyItemHasErrors = True
                End If
            End If
        End With
    Next i
    
    If anyItemHasErrors Then
        poChangeOrder.HasError = True
        poChangeOrder.GlobalError = "One or more change items have pre-check errors"
        
        GoTo PreCheckFailed
    End If
    ' End - Validate ChangeOrderItems
    
    
    If hasPoDefaultFieldChange = False And hasValidChangeOrderItem = False Then
        poChangeOrder.HasError = True
        poChangeOrder.GlobalError = "Change Order does not specify any changes. Set one or more change order fields."
        GoTo PreCheckFailed
    End If
    
    
    PeopleSoft_ChangeOrder_Process_PreCheck = True
    Exit Function

PreCheckFailed:
    poChangeOrder.HasError = True
    Debug_Print "PeopleSoft_ChangeOrder_Process_PreCheck: Failed! " & poChangeOrder.GlobalError
    PeopleSoft_ChangeOrder_Process_PreCheck = False

End Function
Private Function PeopleSoft_ChangeOrder_ProcessLineScheduleItems(driver As SeleniumWrapper.WebDriver, poChangeOrder As PeopleSoft_PurchaseOrder_ChangeOrder) As Boolean
    ' Assume we are starting from the PO page
    
    If poChangeOrder.PO_ChangeOrder_ItemCount <= 0 Then
        ' No items to process
        PeopleSoft_ChangeOrder_ProcessLineScheduleItems = True
        Exit Function
    End If
    
    Dim By As New SeleniumWrapper.By
    
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    
    Dim i As Long
    Dim paginationText As String, posTo As Integer, posOf As Integer
    Dim pageLineFrom As Integer, pageLineTo As Integer, pageLineTotal As Integer
    Dim anyChangeOrderItemsOnPage As Boolean
    
    Dim isSinglePagePO As Boolean
    Dim numProcessed As Integer
    
    Dim we As WebElement, weCollection As WebElementCollection
    Dim weID As String
    
    Dim PO_ChangeOrder_ItemProcessedFlag() As Boolean
    ReDim PO_ChangeOrder_ItemProcessedFlag(1 To poChangeOrder.PO_ChangeOrder_ItemCount) As Boolean
   
    isSinglePagePO = True
    numProcessed = 0
        
    ' At a high level, the general approach taken is to loop through each PO page which consists of all the possible lines items
    ' As we visit each page, we check to see if any of the change order items are on the current page
    Do
        anyChangeOrderItemsOnPage = False
        
        ' Find element with item count. Example: 1 to 75 of 77
        If PeopleSoft_Page_ElementExists(driver, By.id("PO_SCR_NAV_WRK_SRCH_RSLT_MSG")) Then
            isSinglePagePO = False
        
            ' Extract starting line # and last line #s on page
            paginationText = driver.findElementById("PO_SCR_NAV_WRK_SRCH_RSLT_MSG").text
        
            posTo = InStr(1, paginationText, " to ")
            posOf = InStr(1, paginationText, " of ")
            
            Debug.Assert posTo > 0
            Debug.Assert posOf > 0
            Debug.Assert posOf > posTo
            
            pageLineFrom = Mid(paginationText, 1, posTo - 1)
            pageLineTo = Mid(paginationText, posTo + Len(" to "), posOf - posTo - Len(" to "))
            pageLineTotal = Mid(paginationText, posOf + Len(" of "))
            
            ' Check to see if any of the change order lines items are on the page
            anyChangeOrderItemsOnPage = False
            
            For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
                If pageLineFrom <= poChangeOrder.PO_ChangeOrder_Items(i).PO_Line And poChangeOrder.PO_ChangeOrder_Items(i).PO_Line <= pageLineTo Then
                    anyChangeOrderItemsOnPage = True
                    Exit For
                End If
            Next i
        Else
            pageLineFrom = 1
            pageLineTo = 9999
        End If
        
        ' ------------------------------
        ' Begin - Multi-page Workaround
        ' ------------------------------
        ' For some reason, if the PO spans multiple pages, moving from the first page to the second does not work (the browser hangs).
        ' Therefore, we can only process a change order for items on the first page. If any change order item
        ' exists outside of the first page, an error will be thrown and the change order canceled.
        ' This entire section can be removed after the issue is fixed.
        If False Then
            Dim anyLineEditsOutsideOfPage As Boolean: anyLineEditsOutsideOfPage = False
            
            For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
                If pageLineFrom > poChangeOrder.PO_ChangeOrder_Items(i).PO_Line Or poChangeOrder.PO_ChangeOrder_Items(i).PO_Line > pageLineTo Then
                    anyLineEditsOutsideOfPage = True
                    poChangeOrder.PO_ChangeOrder_Items(i).HasError = True
                    poChangeOrder.PO_ChangeOrder_Items(i).ItemError = "Cannot process change order for item: line exists outside of first PO page"
                    Exit For
                End If
            Next i
            
            If anyLineEditsOutsideOfPage Then
                poChangeOrder.HasError = True
                poChangeOrder.GlobalError = "Change order needs to be performed manually: one or more lines exists outside of first PO page"
            
                GoTo ChangeOrderFailed
            End If
        End If
        ' ------------------------------
        '- End - Multi-page Workaround
        ' ------------------------------
        
        
        If anyChangeOrderItemsOnPage Or isSinglePagePO Then
            Dim pageLineIndex As Integer, pageScheduleIndex As Integer
            'Dim lineIndex As Integer
            
            pageLineIndex = -1 ' -1 = invalid value
            pageScheduleIndex = -1   '-1 = invalid value
            
            
            
            ' Expand All Lines/Schedules
            driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
            PeopleSoft_Page_WaitForProcessing driver
            
            For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
                If pageLineFrom <= poChangeOrder.PO_ChangeOrder_Items(i).PO_Line And poChangeOrder.PO_ChangeOrder_Items(i).PO_Line <= pageLineTo Then
                    pageLineIndex = poChangeOrder.PO_ChangeOrder_Items(i).PO_Line - pageLineFrom
                    
                    
                    If pageLineIndex < 0 Then
                        poChangeOrder.PO_ChangeOrder_Items(i).HasError = True
                        poChangeOrder.PO_ChangeOrder_Items(i).ItemError = "Cannot process change order for item: calculated line index is invalid"
                        
                        GoTo ChangeOrderFailed
                    End If
                    
                    
                    Dim poLineStatus As String
                    poLineStatus = driver.findElementById("PO_LINE_CANCEL_STATUS$" & pageLineIndex).text
                    
                    If poLineStatus <> "Active" And poLineStatus <> "Open" Then
                        poChangeOrder.PO_ChangeOrder_Items(i).HasError = True
                        poChangeOrder.PO_ChangeOrder_Items(i).ItemError = "Cannot process change order for item: line status is not Active or Open"
                        
                        GoTo ChangeOrderFailed
                    End If
                    
                    ' Begin - Determine the schedule index in the page by looking at the index for the schedule captions
                    ' Get elements which have the schedule ID
                    Set weCollection = driver.findElementsByXPath(".//*[@id='ACE_PO_LINE_SHIP_SCROL$" & pageLineIndex & "']/descendant::span[starts-with(@id,'PO_LINE_SHIP_SCHED_NBR')]")
                     
                    pageScheduleIndex = -1
                    
                    For Each we In weCollection
                        If Not IsNumeric(we.text) Then
                            poChangeOrder.PO_ChangeOrder_Items(i).HasError = True
                            poChangeOrder.PO_ChangeOrder_Items(i).ItemError = "Unexpected error in page: Schedule element is not numeric. Value is: " & we.text
                        End If
                        
                        If CInt(we.text) = poChangeOrder.PO_ChangeOrder_Items(i).PO_Schedule Then
                            ' Extract schedule index from span ID (Example: PO_LINE_SHIP_SCHED_NBR$10) <--- extract the 10 at the end. This is our page schedule index
                            weID = we.getAttribute("id")
                            pageScheduleIndex = CInt(Mid$(weID, InStr(1, weID, "$") + 1))
                            Exit For
                        End If
                    Next we
                    
                    
                    If pageScheduleIndex < 0 Then
                        poChangeOrder.PO_ChangeOrder_Items(i).HasError = True
                        poChangeOrder.PO_ChangeOrder_Items(i).ItemError = "Cannot process change order for item: line schedule does not exist or not displayed on page"
                        
                        GoTo ChangeOrderFailed
                    End If
                    
                    
                    ' Expand Schedule
                    'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_PB$" & lineIndex & "', 0, 0, 'Expand Schedule Section', false, true);"
                    'PeopleSoft_Page_WaitForProcessing driver
                    
                    ' Expand Distribution
                    ' Click PO_PNLS_PB_EXPAND_PB$232$$0
                     'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_PB$232$$0', 0, 0, 'Expand Distribution Section', false, true);"
                    'PeopleSoft_Page_WaitForProcessing driver
                    
                    
                    
                    'driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
    
                    
                    ' Note since 2.9.1.1,
                    '<a id="PO_PNLS_WRK_Z_CHANGE_DIST$0" href="javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_Z_CHANGE_DIST$0', false, true);" tabindex="893" name="PO_PNLS_WRK_Z_CHANGE_DIST$0">
                    'a id="PO_PNLS_WRK_GOTO_SCHED_DTLS$0" href="javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_GOTO_SCHED_DTLS$0');" tabindex="584" name="PO_PNLS_WRK_GOTO_SCHED_DTLS$0">
                    '<a id="PO_PNLS_WRK_GOTO_LINE_DTLS$2" href="javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_GOTO_LINE_DTLS$2');" tabindex="557" name="PO_PNLS_WRK_GOTO_LINE_DTLS$2">
                    
                   ' <a href="javascript:hAction_win0(document.win0,'PO_PNLS_WRK_CHANGE_LINE', 0, 0, 'Create Line Change', false, true);" tabindex="16" id="PO_PNLS_WRK_CHANGE_LINE" name="PO_PNLS_WRK_CHANGE_LINE"><img border="0" title="Create Line Change" alt="Create Line Change" name="PO_PNLS_WRK_CHANGE_LINE$IMG" src="/cs/ps/cache/PS_DELTA_ICN_1.gif"></a>
                    
                    'Dim tmp As String
                    'tmp = driver.findElementById("PO_LINE_SHIP_DUE_DT$" & (pageScheduleIndex)).getAttribute("disabled")
                    
                    
                    ' Click on change order triangle for this schedule
                    driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_CHANGE_SHIP$" & pageScheduleIndex & "');" ' Fix for 2.9.1.1  due to PS upgrade
                    PeopleSoft_Page_WaitForProcessing driver
                    
                    If poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE > 0 Then
                        PeopleSoft_Page_SetValidatedField driver:=driver, fieldElementID:=("PO_LINE_SHIP_DUE_DT$" & pageScheduleIndex), _
                            fieldValue:=Format(poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE, "mm/dd/yyyy"), validationResult:=poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE_Result, _
                            expectedPopupContents:="*Due Date selected is a weekend or a public holiday*"
                            
                            
                        If poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE_Result.ValidationFailed Then GoTo ChangeOrderFailed ' ValidationFail
                    End If
                    
                    PO_ChangeOrder_ItemProcessedFlag(i) = True
                    numProcessed = numProcessed + 1
                End If
            Next i
            
            
        End If
        
        
        If pageLineTo < pageLineTotal And numProcessed < poChangeOrder.PO_ChangeOrder_ItemCount And Not isSinglePagePO Then
            ' Next page
            driver.findElementById("PO_SCR_NAV_WRK_NEXT_ITEM_BUTTON").Click
            PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
            
            popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True)
            PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
            ' Do we need to check the status of any of these popups?
        End If
    Loop Until pageLineTo = pageLineTotal Or isSinglePagePO
        
    
    


    PeopleSoft_ChangeOrder_ProcessLineScheduleItems = True
    Exit Function
    
ChangeOrderFailed:
    PeopleSoft_ChangeOrder_ProcessLineScheduleItems = False

End Function
Public Function PeopleSoft_Receipt_CreateReceipt(ByRef session As PeopleSoft_Session, ByRef rcpt As PeopleSoft_Receipt) As Boolean

    
    Debug_Print "PeopleSoft_Receipt_CreateReceipt called (" & Debug_VarListString("PO ID", rcpt.PO_ID, "Mode", IIf(rcpt.ReceiveMode = RECEIVE_SPECIFIED, "<SPECIFIED>", "<ALL>"), "ReceiptItemsCount", rcpt.ReceiptItemCount) & ")"


    If PeopleSoft_Receipt_CreateReceipt_PreCheck(rcpt) = False Then
        ' GlobalError and HasError set by PreCheck
        GoTo ReceiptFailed
    End If


    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown

    'Dim session As PeopleSoft_Session
    Dim driver As New SeleniumWrapper.WebDriver
    Dim elem As WebElement
    
    
    Set driver = session.driver
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim i As Integer, j As Integer
    
    
    PeopleSoft_Login session
    
    
    If Not session.loggedIn Then
        rcpt.GlobalError = "Logon Error: " & session.LogonError
        rcpt.HasGlobalError = True
        
        GoTo ReceiptFailed
    End If
    
    
    driver.get PS_URI_RECEIPT_ADD
    
    
    driver.waitForElementPresent "css=#RECV_PO_ADD_BUSINESS_UNIT"
    
    
    Dim PO_BU_default As String
    
    ' Check if auto-filled PO BU is correct. If not,enter the correct PO BU
    If rcpt.PO_BU <> "" Then
        
        Set elem = driver.findElementById("RECV_PO_ADD_BUSINESS_UNIT")
        
        PO_BU_default = elem.getAttribute("value")
    
        If PO_BU_default <> rcpt.PO_BU Then
            PeopleSoft_Page_SetValidatedField driver, ("RECV_PO_ADD_BUSINESS_UNIT"), rcpt.PO_BU, rcpt.PO_BU_Result
            If rcpt.PO_BU_Result.ValidationFailed Then GoTo ValidationFailed
        End If
    End If
    
    
    driver.findElementById("#ICSearch").Click
    
    PeopleSoft_Page_WaitForProcessing driver
    
    ' Enter PO ID
    driver.findElementById("PO_PICK_ORD_WRK_ORDER_ID").Clear
    driver.findElementById("PO_PICK_ORD_WRK_ORDER_ID").SendKeys rcpt.PO_ID
    
    
    driver.findElementById("PO_PICK_ORD_WRK_PB_FETCH_PO").Click
    PeopleSoft_Page_WaitForProcessing driver
    
    
    
    ' ------------------------------------------------------
    ' Begin - Map unreceived items on page to receipt items
    ' ------------------------------------------------------
    Dim unreceivedItems() As PeopleSoft_ReceiptPage_UnreceivedItem
    Dim unreceivedItemCount As Long
    
    
    unreceivedItemCount = PeopleSoft_Receipt_ExtractUnreceivedItemsFromPage(driver, unreceivedItems)
    
    If unreceivedItemCount = 0 Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "No receivable items on this PO: all items are already received"
        GoTo ReceiptFailed
    End If
    
    ' Bug fix: Exit here if none of the items on the page can be received
    Dim anyItemIsReceivable As Boolean: anyItemIsReceivable = False
    
    For i = 1 To unreceivedItemCount
        anyItemIsReceivable = anyItemIsReceivable Or unreceivedItems(i).IsReceivable
    Next i
    
    If anyItemIsReceivable = False Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "No receivable items on this PO: remaining items cannot be received in PeopleSoft"
        GoTo ReceiptFailed
    End If
    
    ' Holds the mapping between the ReceiptItems() and the UnreceivedItems() on the page
    '    mapReceiptItemsToPageUnreceivedItems(<index of rcpt.ReceiptItems() item>) = <index of unreceivedItems() item>
    Dim mapReceiptItemsToPageUnreceivedItems() As Long
    
    Dim idx As Long
    
    If rcpt.ReceiveMode = RECEIVE_SPECIFIED Then
        ' RECEIVE_SPECIFIED: pre-allocate index map array, and map items between
        mapReceiptItemsToPageUnreceivedItems = PeopleSoft_Receipt_MapReceiptItemsToPageUnreceivedItems(rcpt.ReceiptItems, rcpt.ReceiptItemCount, unreceivedItems, unreceivedItemCount)
    
         ' receive specified: map each row to the corresponding specific line/schedule in ReceiptItems()
        For i = 1 To rcpt.ReceiptItemCount
            ' If found valid mapping, copy item data
            If mapReceiptItemsToPageUnreceivedItems(i) > 0 Then
                idx = mapReceiptItemsToPageUnreceivedItems(i)
                
                If rcpt.ReceiptItems(i).Item_ID = "" Then rcpt.ReceiptItems(i).Item_ID = unreceivedItems(idx).PO_Item_ID
                rcpt.ReceiptItems(i).CATS_FLAG = unreceivedItems(idx).CATS_FLAG
                rcpt.ReceiptItems(i).TRANS_ITEM_DESC = unreceivedItems(idx).PO_TRANS_ITEM_DESC
                rcpt.ReceiptItems(i).IsNotReceivable = Not unreceivedItems(idx).IsReceivable
                
                ' Not receivable => Qty received = 0
                If unreceivedItems(idx).IsReceivable = False Then
                    rcpt.ReceiptItems(i).Receive_Qty = 0
                    rcpt.ReceiptItems(i).Accept_Qty = 0
                End If
            End If
        Next i
    Else ' rcpt.ReceiveMode = RECEIVE_ALL
        ' update 2.10.1:
        ' RECEIVE_ALL: If we are receiving all lines, then we will return all line item information, rather than match the specific line items
        '   As a result, The mapping will result in the same index for both arrays
        
        ' Receiving all lines means we need to return the ReceiptItems() as output
        rcpt.ReceiptItemCount = unreceivedItemCount
        ReDim rcpt.ReceiptItems(1 To rcpt.ReceiptItemCount) As PeopleSoft_Receipt_Item
        ReDim mapReceiptItemsToPageUnreceivedItems(1 To rcpt.ReceiptItemCount) As Long
        
        For i = 1 To unreceivedItemCount
            mapReceiptItemsToPageUnreceivedItems(i) = i ' Since we are copying. We can perform in order:
            
            rcpt.ReceiptItems(i).PO_Line = unreceivedItems(i).PO_Line
            rcpt.ReceiptItems(i).PO_Schedule = unreceivedItems(i).PO_Schedule
            rcpt.ReceiptItems(i).Item_ID = unreceivedItems(i).PO_Item_ID
            rcpt.ReceiptItems(i).TRANS_ITEM_DESC = unreceivedItems(i).PO_TRANS_ITEM_DESC
            rcpt.ReceiptItems(i).CATS_FLAG = unreceivedItems(i).CATS_FLAG
            rcpt.ReceiptItems(i).IsNotReceivable = Not unreceivedItems(i).IsReceivable
            
            rcpt.ReceiptItems(i).Receive_Qty = IIf(unreceivedItems(i).IsReceivable, unreceivedItems(i).PO_Qty, 0)
            rcpt.ReceiptItems(i).Accept_Qty = IIf(unreceivedItems(i).IsReceivable, unreceivedItems(i).PO_Qty, 0)
        Next i
    End If
    ' ------------------------------------------------------
    ' End - Map receivable  items on page to receipt items
    ' ------------------------------------------------------


    'Debug.Print
    
    Dim numUnmatchedItems As Integer: numUnmatchedItems = 0
    Dim numReceivableItems As Integer: numReceivableItems = 0
    Dim rowIndex As Integer
    
    ' Go through mapping/receive items. Click checkbox to receive.
    ' Check if any of the receipt items have not been mapped. If so,
    ' it has already been received or it is not receivable by the user
    For i = 1 To rcpt.ReceiptItemCount
        If mapReceiptItemsToPageUnreceivedItems(i) > 0 Then
            rowIndex = unreceivedItems(mapReceiptItemsToPageUnreceivedItems(i)).PageTableRowIndex
                                                                           
            If rcpt.ReceiptItems(i).IsNotReceivable = False Then
                ' Check the checkbox
                driver.findElementById("RECV_PO_SCHEDULE$" & rowIndex).Click
                
                numReceivableItems = numReceivableItems + 1
            End If
            
            rcpt.ReceiptItems(i).HasError = False
        Else
            ' The requested receipt item could not be mapped to an unreceived item. (This can only occur when
            ' receive mode = specified.)
            
            Debug.Assert rcpt.ReceiveMode = RECEIVE_SPECIFIED
            
            rcpt.ReceiptItems(i).HasError = True
            rcpt.ReceiptItems(i).ItemError = "Cannot receive on this item: not receivable or already received."
            
            numUnmatchedItems = numUnmatchedItems + 1
        End If
    Next i
    
    If numUnmatchedItems = rcpt.ReceiptItemCount Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "None of the specified PO lines/schedules can be received on PO."
        
        GoTo ReceiptFailed
    End If


    'Debug.Print

    
    ' Navigate to receipt page
    'driver.findElementById("#ICSave").Click
    driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');"
    PeopleSoft_Page_WaitForProcessing driver
    
    

    
    Dim pageReceiptLines() As PeopleSoft_ReceiptPage_ReceiptLine
    Dim pageReceiptLineCount As Long

    pageReceiptLineCount = PeopleSoft_Receipt_ExtractReceiptLinesFromPage(driver, pageReceiptLines)


    ' Sanity check: The number of receipt lines should match the number of checkboxes we clicked. Really, this should never happen.
    If pageReceiptLineCount <> numReceivableItems Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Number of receipt lines does not match items checked on previous page"
        
        GoTo ReceiptFailed
    End If
    
    
    
    ' Holds the mapping between the ReceiptItems() and the ReceiptLines() on the page
    '    mapReceiptItemsToPageReceiptLines(<index of rcpt.ReceiptItems() item>) = <index of pageReceiptLines() item>
    Dim mapReceiptItemsToPageReceiptLines() As Long
    
    mapReceiptItemsToPageReceiptLines = PeopleSoft_Receipt_MapReceiptItemsToPageReceiptLines(rcpt.ReceiptItems, rcpt.ReceiptItemCount, pageReceiptLines, pageReceiptLineCount)
    

    ' -----------------------------------------------------------
    ' Begin - Sanity Checks: if receipt lines match the input ReceiptLines. Adjust ReceiptQty and return AcceptQty as needed.
    ' -----------------------------------------------------------
    Dim pageRcptLineIdx As Long
    Dim anyItemHasErrors As Boolean
    
    ' Sanity check to ensure all receivable items are mapped to a receipt line before we start
    anyItemHasErrors = False
    
    For i = 1 To rcpt.ReceiptItemCount
        pageRcptLineIdx = mapReceiptItemsToPageReceiptLines(i)
        
        If rcpt.ReceiptItems(i).HasError = False And rcpt.ReceiptItems(i).IsNotReceivable = False And mapReceiptItemsToPageReceiptLines(i) < 1 Then
            rcpt.ReceiptItems(i).HasError = True
            rcpt.ReceiptItems(i).ItemError = "Failed check: receivable item not mapped to receive line"
            anyItemHasErrors = True
        End If
    Next i
    
    If anyItemHasErrors Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "One or more receivable items were not mapped to a receive line"
        
        GoTo ReceiptFailed
    End If

    ' Receive on specific quantities
    If rcpt.ReceiveMode = RECEIVE_SPECIFIED Then
        Debug_Print "PeopleSoft_Receipt_CreateReceipt: Receive specified mode -> Receive on specific quantities"
    
        anyItemHasErrors = False
        
        For i = 1 To rcpt.ReceiptItemCount
            pageRcptLineIdx = mapReceiptItemsToPageReceiptLines(i)
            
            
            Debug_Print "PeopleSoft_Receipt_CreateReceipt: Receipt Line #" & i & " (" & Debug_VarListString("HasError", rcpt.ReceiptItems(i).HasError, "Receivable", Not rcpt.ReceiptItems(i).IsNotReceivable, "pageRcptLineIdx", pageRcptLineIdx) & ")"
            
            Debug_Print "PeopleSoft_Receipt_CreateReceipt: PageReceiptLine Index #" & pageRcptLineIdx & " (" & Debug_VarListString("Accept Qty", pageReceiptLines(pageRcptLineIdx).Accept_Qty) & ")"
            
            
            If rcpt.ReceiptItems(i).HasError = False And rcpt.ReceiptItems(i).IsNotReceivable = False Then
                rcpt.ReceiptItems(i).Accept_Qty = pageReceiptLines(pageRcptLineIdx).Accept_Qty
                        
                ' Check: receive quantity is less than accept qty
                If rcpt.ReceiptItems(i).Receive_Qty > 0 Then ' Receipt qty specified
                    If rcpt.ReceiptItems(i).Receive_Qty > rcpt.ReceiptItems(i).Accept_Qty Then
                        Debug_Print "PeopleSoft_Receipt_CreateReceipt: Receipt Line #" & i & ": Receive qty is greater than accepted qty (Accept Qty: " & rcpt.ReceiptItems(i).Accept_Qty & ")"
                        rcpt.ReceiptItems(i).HasError = True
                        rcpt.ReceiptItems(i).ItemError = "Receive qty is greater than accepted qty (Accept Qty: " & rcpt.ReceiptItems(i).Accept_Qty & ")"
                        anyItemHasErrors = True
                    End If
                End If
                        
                If rcpt.ReceiptItems(i).HasError = False Then
                     ' If Receipt qty specified -> receive on the specified amount, otherwise receive all
                    If rcpt.ReceiptItems(i).Receive_Qty > 0 Then
                        Dim tmpValidationResult As PeopleSoft_Field_ValidationResult
                        
                        ' Fill in Receive Qtr
                        PeopleSoft_Page_SetValidatedField driver, ("RECV_LN_SHIP_QTY_SH_RECVD$" & pageReceiptLines(pageRcptLineIdx).PageTableRowIndex), _
                            CStr(rcpt.ReceiptItems(i).Receive_Qty), tmpValidationResult
                        
                        If tmpValidationResult.ValidationFailed Then
                            Debug_Print "PeopleSoft_Receipt_CreateReceipt: Receipt Line #" & i & ": RECEIVE QTY ERROR: " & tmpValidationResult.ValidationErrorText
                            rcpt.ReceiptItems(i).HasError = True
                            rcpt.ReceiptItems(i).ItemError = "RECEIVE QTY ERROR: " & tmpValidationResult.ValidationErrorText
                            anyItemHasErrors = True
                        End If
                    Else
                        ' No receipt qty given. Receive on all and return the qty.
                        rcpt.ReceiptItems(i).Receive_Qty = pageReceiptLines(pageRcptLineIdx).Receipt_Qty
                    End If
                End If

            End If
        Next i
        
        
        If anyItemHasErrors Then
            rcpt.HasGlobalError = True
            rcpt.GlobalError = "Error occurred when modifying receipt lines. See each line item for details."
            GoTo ReceiptFailed
        End If
        
    End If
        
        
        
    If DEBUG_OPTIONS.QuitBeforeSave Then
        rcpt.GlobalError = "Debug option QuitBeforeSave enabled. Processing halted before saving. To prevent this from occurring, disable QuitBeforeSave option."
        rcpt.GlobalError = True
        PeopleSoft_Receipt_CreateReceipt = False
        Exit Function
    End If
    
            
    ' Save Receipt
    'driver.findElementById("#ICSave").Click
    driver.runScript "javascript:setSaveText_win0('Saving...');submitAction_win0(document.win0, '#ICSave');"
    
    
    ' Wait for "Saving..." to stop.
    driver.waitForElementPresent "css=#SAVED_win0"
    'driver.findElementById("processing").waitForCssValue "visibility", "visible"
    driver.findElementById("SAVED_win0").waitForCssValue "visibility", "hidden"
    
    
     
    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult
    
    
    popupCheckResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False, _
        raiseErrorIfUnexpected:=False, expectedContent:="*Have these receipt quantities been checked for accuracy*")
    
    
    If popupCheckResult.HasPopup = False Or (popupCheckResult.HasPopup And popupCheckResult.IsExpected = False) Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Did not receive expected popup: Have these receipt quantities been checked for accuracy?" _
                            & IIf(popupCheckResult.HasPopup, vbCrLf & "Popup received: " & popupCheckResult.popupText, "")
        
        GoTo ReceiptFailed
    End If
    
    
    ' We received correct popup -> acknowledge
    PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbYes
    PeopleSoft_Page_WaitForProcessing driver
    
            
            
    


    
    
    ' Check for receipt ID.
    rcpt.RECEIPT_ID = driver.findElementById("RECV_HDR_RECEIVER_ID").text
    rcpt.RECEIPT_ID = Trim(rcpt.RECEIPT_ID)
    Debug.Print "Receipt ID: " & rcpt.RECEIPT_ID
    
    
    
    If Not IsNumeric(rcpt.RECEIPT_ID) Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Non-numeric receipt ID not found on page: " & rcpt.RECEIPT_ID
    
        GoTo CancelReceiptAndExit
    End If
    
    
    ' Receipt ID provided -> at this point it doesnt matter what shows up, just acknowledge it
    Dim popupCountCheck As Integer: popupCountCheck = 0
    
    Do
        popupCheckResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True, raiseErrorIfUnexpected:=False, _
            expectedContent:=Array("*This means the receipt is being updated by the receipt integration process*"))
        If popupCheckResult.HasPopup = False Then Exit Do
        
        If popupCheckResult.IsExpected = False Then
            popupCountCheck = popupCountCheck + 1
            Debug_Print "Popup received after Receipt " & popupCountCheck & ": " & popupCheckResult.popupText
            rcpt.GlobalError = rcpt.GlobalError & "Popup received after Receipt " & popupCountCheck & ": " & popupCheckResult.popupText & vbCrLf
        End If
        
        PeopleSoft_Page_WaitForProcessing driver
    Loop
        
        
    
    
    
    PeopleSoft_Receipt_CreateReceipt = True
    Exit Function
    
    
CancelReceiptAndExit:
    ' Begin - Cancel Receipt
    'driver.findElementById("RECV_HDR_WK_PB_CANCEL_RECPT").Click
    'PeopleSoft_Page_WaitForProcessing driver
    
    
    'PopupText = PeopleSoft_Page_SuppressPopup(driver, vbYes)
    'popUpIsExpected = InStr(1, popUpText, "Canceling Receipt cannot be reversed.") > 0
    
    'If popUpIsExpected = False Then
    '    rcpt.HasGlobalError = True
    '    rcpt.GlobalError = "Unexpected popup: " & popUpText
    '
    '    GoTo ReceiptFailed
    'End If
    
    'PeopleSoft_Page_WaitForProcessing driver
    ' End - Cancel Receipt



ValidationFailed:
ReceiptFailed:
    PeopleSoft_Receipt_CreateReceipt = False
    
    Debug_Print "PeopleSoft_Receipt_CreateReceipt: ERROR: " & rcpt.GlobalError
    PeopleSoft_SaveDebugInfo driver
    Exit Function
       
ExceptionThrown:
    PeopleSoft_Receipt_CreateReceipt = False
    
    rcpt.HasGlobalError = True
    rcpt.GlobalError = "Exception: " & Err.Description
    
    Debug_Print "PeopleSoft_Receipt_CreateReceipt: ERROR: " & rcpt.GlobalError
    PeopleSoft_SaveDebugInfo driver



End Function
Private Function PeopleSoft_Receipt_CreateReceipt_PreCheck(rcpt As PeopleSoft_Receipt) As Boolean

    
    Debug_Print "PeopleSoft_Receipt_CreateReceipt_PreCheck: Receipt (" & Debug_VarListString("PO BU", rcpt.PO_BU, "PO ID", rcpt.PO_ID) & ")"
    
    Dim i As Long, j As Long
    
    For i = 1 To rcpt.ReceiptItemCount
        With rcpt.ReceiptItems(i)
            Debug_Print "PeopleSoft_Receipt_CreateReceipt_PreCheck: Receipt Item: (" & Debug_VarListString("PO Line", .PO_Line, "PO Schedule", .PO_Schedule, "Receive Qty", .Receive_Qty) & ")"
        End With
    Next i
    
    
    ' Begin - Check required fields
    If Len(rcpt.PO_BU) = 0 Then
        rcpt.PO_BU_Result.ValidationErrorText = "PO BU is required."
        rcpt.PO_BU_Result.ValidationFailed = True
        rcpt.HasGlobalError = True
    End If
    If Len(rcpt.PO_ID) = 0 Then
        rcpt.PO_ID_Result.ValidationErrorText = "PO BU is required."
        rcpt.PO_ID_Result.ValidationFailed = True
        rcpt.HasGlobalError = True
    End If
    
    If rcpt.HasGlobalError Then
        rcpt.GlobalError = "One or more required fields are missing"
        GoTo PreCheckFailed
    End If
    ' End - Check required fields
    

    
    ' Begin - validate fields: ensure single-line fields are actually a single line
    If PeopleSoft_ValidateField_IsSingleLine("PO BU", rcpt.PO_BU, rcpt.PO_BU_Result) = False Then rcpt.HasGlobalError = True
    If PeopleSoft_ValidateField_IsSingleLine("PO ID", rcpt.PO_ID, rcpt.PO_ID_Result) = False Then rcpt.HasGlobalError = True
    
    If rcpt.HasGlobalError Then
        rcpt.GlobalError = "One or more fields failed pre-check"
        GoTo PreCheckFailed
    End If
    ' Begin - validate fields: ensure single-line fields are actually a single line


    Dim anyItemHasErrors As Boolean


    ' Begin - Check for duplicate PO lines/schedules in Receipt Items
    anyItemHasErrors = False
    
    For i = 1 To rcpt.ReceiptItemCount
        For j = i + 1 To rcpt.ReceiptItemCount
            If rcpt.ReceiptItems(i).PO_Line = rcpt.ReceiptItems(j).PO_Line _
              And rcpt.ReceiptItems(i).PO_Schedule = rcpt.ReceiptItems(j).PO_Schedule Then
                rcpt.ReceiptItems(i).HasError = True
                rcpt.ReceiptItems(j).HasError = True
                rcpt.ReceiptItems(i).ItemError = "Duplicate line/schedule"
                rcpt.ReceiptItems(j).ItemError = "Duplicate line/schedule"
                anyItemHasErrors = True
            End If
        Next j
    Next i
    
    If anyItemHasErrors Then
        rcpt.GlobalError = "Duplicate line/schedule in one more receipt items"
        GoTo PreCheckFailed
    End If
    ' End - Check for duplicate PO lines/schedules in  Receipt Items
    
    
    ' Begin - Validate Receipt Items
    anyItemHasErrors = False

    For i = 1 To rcpt.ReceiptItemCount
        With rcpt.ReceiptItems(i)
            If .PO_Line < 1 Then
                .HasError = True
                .ItemError = .ItemError & "PO Line must be zero or greater" & vbCrLf
                anyItemHasErrors = True
            End If
            If .PO_Schedule < 1 Then
                .HasError = True
                .ItemError = .ItemError & "PO Schedule must be zero or greater" & vbCrLf
                anyItemHasErrors = True
            End If
            If .Receive_Qty < 0 Then
                .HasError = True
                .ItemError = .ItemError & "Receive qty must be zero (receive entire line) or greater (receive specified amount)" & vbCrLf
                anyItemHasErrors = True
            End If
        End With
    Next i
    
    If anyItemHasErrors Then
        rcpt.GlobalError = "One or more change items have pre-check errors"
        GoTo PreCheckFailed
    End If
    ' End - Validate Receipt Items
    
    

    
    
    PeopleSoft_Receipt_CreateReceipt_PreCheck = True
    Exit Function

PreCheckFailed:
    rcpt.HasGlobalError = True
    Debug_Print "PeopleSoft_Receipt_CreateReceipt_PreCheck: Failed! " & rcpt.GlobalError
    PeopleSoft_Receipt_CreateReceipt_PreCheck = False

End Function
' PeopleSoft_Receipt_ExtractUnreceivedItems: Extracts all unreceived items from PS Receipt page. Assumes we already navigated to page. Returns count but populated the paremeter unreceivedItems
Private Function PeopleSoft_Receipt_ExtractUnreceivedItemsFromPage(driver As SeleniumWrapper.WebDriver, ByRef unreceivedItems() As PeopleSoft_ReceiptPage_UnreceivedItem) As Long

    Dim By As New SeleniumWrapper.By

    Dim unreceivedItemCount As Long
    
    unreceivedItemCount = 0
    PeopleSoft_Receipt_ExtractUnreceivedItemsFromPage = 0
    
    If Not PeopleSoft_Page_ElementExists(driver, By.id("win0divPO_PICK_ORD_WS$0")) Then ' fix 2.9.1.3
        ' No receivable items on this PO
        Exit Function
    End If
    
    
    If Not PeopleSoft_Page_ElementExists(driver, By.id("PO_PICK_ORD_WRK_Z_IN_CATS_FLAG$0")) Then
        ' No receivable items on this PO
        Exit Function
    End If
    
    
    ' The following script has to be executed because selenium can only operate on visible elements. The retreived
    ' rows on the page by default is limited to a height of 400 or so pixels and forces the user to use scrollbars to
    ' see the rest of the items. This script modifies the height to include ALL items, regardless of how lengthy the
    ' page becomes.
    driver.runScript "javascript: document.getElementById('divgblPO_PICK_ORD_WS$0').style.height ='auto'; " & _
                                  "document.getElementById('divgbrPO_PICK_ORD_WS$0').style.height ='auto'; "
                                  
    ' in some cases, the Save, Cancel and Refresh buttons cover the checkbox. Move them to the upper part of the page
    driver.runScript "javascript: var elem = document.getElementById('#ICSave'); elem.style.position = 'absolute'; elem.style.top = 0;"
    driver.runScript "javascript: var elem = document.getElementById('#ICCancel'); elem.style.position = 'absolute'; elem.style.top = 0;"
    driver.runScript "javascript: var elem = document.getElementById('#ICRefresh'); elem.style.position = 'absolute'; elem.style.top = 0;"
    
    Dim numReturnedRows As Long, rowIndex As Long
    
    numReturnedRows = driver.getXpathCount(".//*[contains(@id,'ftrPO_PICK_ORD_WS$0_row')]")
    
    ' if one entry, check PO ID. If blank, then there aren't any receivable lines.
    If numReturnedRows = 1 Then
        Dim tmpPO_ID As String
        
        tmpPO_ID = driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_PO_ID$0').textContent;")
        tmpPO_ID = Trim(Replace(tmpPO_ID, Chr$(160), Chr$(32))) ' Remove spaces, and non-breaking spaces
        
        If Len(tmpPO_ID) = 0 Then
            ' No receivable items on this PO
            Exit Function
        End If
    End If
    
    
    unreceivedItemCount = numReturnedRows
    ReDim unreceivedItems(1 To unreceivedItemCount) As PeopleSoft_ReceiptPage_UnreceivedItem
    
    
    PeopleSoft_Receipt_ExtractUnreceivedItemsFromPage = unreceivedItemCount
    
    Debug_Print "PeopleSoft_Receipt_ExtractUnreceivedItemsFromPage: Unreceived Items"
    Debug_CM_Start 7
    Debug_CM_PrintRow "Line", "Schedule", "Qty", "Item ID", "CATS Flag", "Description", "Chk Disabled"
    
    
    For rowIndex = 0 To numReturnedRows - 1
        Dim Row_CheckDisabled As String
      
        unreceivedItems(rowIndex + 1).PageTableRowIndex = rowIndex
        
        ' Print Debug Row
        Debug_CM_Print 1, driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_LINE_NBR$" & rowIndex & "').textContent;")
        Debug_CM_Print 2, driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_SCHED_NBR$" & rowIndex & "').textContent;")
        Debug_CM_Print 3, driver.executeScript("return document.getElementById('QTY_PO$" & rowIndex & "').textContent;")
        Debug_CM_Print 4, driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndex & "').textContent;")
        Debug_CM_Print 5, driver.executeScript("return document.getElementById('PO_PICK_ORD_WRK_Z_IN_CATS_FLAG$" & rowIndex & "').textContent;")
        Debug_CM_Print 6, driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_DESCR254_MIXED$" & rowIndex & "').textContent;")
        Debug_CM_Print 7, driver.executeScript("return document.getElementById('RECV_PO_SCHEDULE$" & rowIndex & "').disabled;")
        
      
        ' workaround because driver.findElementById(X).Text doesn't always return a value and is very slow
        unreceivedItems(rowIndex + 1).PO_ID = driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_PO_ID$" & rowIndex & "').textContent;")              'driver.findElementById("PO_PICK_ORD_WS_PO_ID$" & rowIndex).Text
        unreceivedItems(rowIndex + 1).PO_Line = CLng(driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_LINE_NBR$" & rowIndex & "').textContent;"))     'CInt(driver.findElementById("PO_PICK_ORD_WS_LINE_NBR$" & rowIndex).Text)
        unreceivedItems(rowIndex + 1).PO_Schedule = CLng(driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_SCHED_NBR$" & rowIndex & "').textContent;"))     'CInt(driver.findElementById("PO_PICK_ORD_WS_SCHED_NBR$" & rowIndex).Text)
        unreceivedItems(rowIndex + 1).PO_Qty = CLng(driver.executeScript("return document.getElementById('QTY_PO$" & rowIndex & "').textContent;"))                       'CInt(driver.findElementById("QTY_PO$" & rowIndex).Text)
        unreceivedItems(rowIndex + 1).PO_Item_ID = driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndex & "').textContent;")        'driver.findElementById("PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndex).Text
        unreceivedItems(rowIndex + 1).CATS_FLAG = driver.executeScript("return document.getElementById('PO_PICK_ORD_WRK_Z_IN_CATS_FLAG$" & rowIndex & "').textContent;")                'driver.findElementById("PO_PICK_ORD_WS_PO_ID$" & rowIndex).Text
       
        unreceivedItems(rowIndex + 1).PO_TRANS_ITEM_DESC = driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_DESCR254_MIXED$" & rowIndex & "').textContent;")
        'unreceivedItems(rowIndex + 1).PO_TRANS_ITEM_DESC = driver.findElementById("PO_PICK_ORD_WS_DESCR254_MIXED$" & rowIndex).Text
       
        Row_CheckDisabled = driver.executeScript("return document.getElementById('RECV_PO_SCHEDULE$" & rowIndex & "').disabled;")
        
        unreceivedItems(rowIndex + 1).IsReceivable = Not CBool(Row_CheckDisabled)
        
    Next rowIndex
    
    Debug_CM_EndFlush colSep:="|"
    

End Function
' PeopleSoft_Receipt_ExtractReceiptItemsFromPage: Extracts all receipt items from PS Receipt page. Assumes we already navigated to page. Returns count but populated the paremeter pageReceiptItems
Private Function PeopleSoft_Receipt_ExtractReceiptLinesFromPage(driver As SeleniumWrapper.WebDriver, ByRef pageReceiptLines() As PeopleSoft_ReceiptPage_ReceiptLine) As Long

    
    Debug_Print "PeopleSoft_Receipt_ExtractReceiptLinesFromPage called"

    Dim pageReceiptLineCount As Long
    
    pageReceiptLineCount = 0
    PeopleSoft_Receipt_ExtractReceiptLinesFromPage = 0
    
    ' Simulate "View All"
    driver.runScript "javascript:submitAction_win0(document.win0,'RECV_LN_SHIP$hviewall$0');"
    PeopleSoft_Page_WaitForProcessing driver
    
    
    pageReceiptLineCount = driver.getXpathCount(".//*[contains(@id,'ftrRECV_LN_SHIP$0_row')]")
    Debug_Print "PeopleSoft_Receipt_ExtractReceiptLinesFromPage: Receipt Line Count: " & pageReceiptLineCount
    If pageReceiptLineCount = 0 Then Exit Function
    
    
    ' Simulate "Show All Columns"
    driver.runScript "javascript:submitAction_win0(document.win0,'RECV_LN_SHIP$htab$0');"
    PeopleSoft_Page_WaitForProcessing driver
    

    ReDim pageReceiptLines(1 To pageReceiptLineCount) As PeopleSoft_ReceiptPage_ReceiptLine
    
    
    Dim rowIndex As Integer
    Dim pageItemContent As Variant
    
    ' Print start of table
    Debug_Print "PeopleSoft_Receipt_ExtractReceiptLinesFromPage: Receipt Lines"
    Debug_CM_Start 8
    Debug_CM_PrintRow "Rcpt Line", "Item ID", "PO Line", "PO Schedule", "Rcpt Qty", "Accept Qty", "Status", "Description"
    
    ' Note: We need to use javascript to get the element content/values since nothing is returned when items are not visible on the page
    For rowIndex = 0 To pageReceiptLineCount - 1
        pageReceiptLines(rowIndex + 1).PageTableRowIndex = rowIndex
        
        ' Print debug line
        Debug_CM_Print 1, driver.executeScript("javascript: return document.getElementById('RECV_LN_NBR$" & rowIndex & "').textContent;")
        Debug_CM_Print 2, driver.executeScript("javascript: return document.getElementById('INV_ITEM_ID$" & rowIndex & "').textContent;")
        Debug_CM_Print 3, driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_LINE_NBR$" & rowIndex & "').textContent;")
        Debug_CM_Print 4, driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_SCHED_NBR$" & rowIndex & "').textContent;")
        Debug_CM_Print 5, driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_QTY_SH_RECVD$" & rowIndex & "').value;")
        Debug_CM_Print 6, driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_QTY_SH_ACCPT$" & rowIndex & "').textContent;")
        Debug_CM_Print 7, driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_RECV_SHIP_STATUS$" & rowIndex & "').textContent;")
        Debug_CM_Print 8, driver.executeScript("javascript: return document.getElementById('DESCR$" & rowIndex & "').textContent;")
        
        
        
        pageReceiptLines(rowIndex + 1).Receipt_Line = CInt(driver.executeScript("javascript: return document.getElementById('RECV_LN_NBR$" & rowIndex & "').textContent;"))
        
        pageReceiptLines(rowIndex + 1).Item_ID = driver.executeScript("javascript: return document.getElementById('INV_ITEM_ID$" & rowIndex & "').textContent;")
        pageReceiptLines(rowIndex + 1).Description = driver.executeScript("javascript: return document.getElementById('DESCR$" & rowIndex & "').textContent;")
        pageReceiptLines(rowIndex + 1).Status = driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_RECV_SHIP_STATUS$" & rowIndex & "').textContent;")
        
        ' Fix: sometimes receipt price is blank: only set if it's available
        pageItemContent = driver.executeScript("javascript: return document.getElementById('PRICE_RECV$" & rowIndex & "').textContent;")
        If IsNumeric(pageItemContent) Then pageReceiptLines(rowIndex + 1).Receipt_Price = CCur(pageItemContent)
        
        
        pageReceiptLines(rowIndex + 1).Receipt_Qty = CCur(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_QTY_SH_RECVD$" & rowIndex & "').value;")) ' <-- note JS .value
        pageReceiptLines(rowIndex + 1).Accept_Qty = CCur(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_QTY_SH_ACCPT$" & rowIndex & "').textContent;"))
        
        pageReceiptLines(rowIndex + 1).PO_Line = CLng(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_LINE_NBR$" & rowIndex & "').textContent;"))
        pageReceiptLines(rowIndex + 1).PO_Schedule = CLng(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_SCHED_NBR$" & rowIndex & "').textContent;"))
        
    Next rowIndex
    
    Debug_CM_EndFlush
    
    PeopleSoft_Receipt_ExtractReceiptLinesFromPage = pageReceiptLineCount

End Function
' PeopleSoft_Receipt_MapReceiptLineToPageReceiptLines: Utility function to map ReceiptLines() to receipt lines on the peoplesoft page unreceived lines
Private Function PeopleSoft_Receipt_MapReceiptItemsToPageUnreceivedItems(ReceiptItems() As PeopleSoft_Receipt_Item, ReceiptItemCount As Long, pageUnreceivedLines() As PeopleSoft_ReceiptPage_UnreceivedItem, pageUnreceivedLineCount As Long) As Long()


    Dim mapFromCount As Long, mapToCount As Long
    
    mapFromCount = ReceiptItemCount
    mapToCount = pageUnreceivedLineCount
    
    Dim map() As Long
    ReDim map(1 To mapFromCount) As Long
    
    Dim i As Long, j As Long
    
    ' Initialize mapping where -1 means unresolved index
    For i = 1 To mapFromCount
        map(i) = -1 ' -1 means undefined
    Next i
    
     ' receive specified: map each row to the corresponding specific line/schedule in ReceiptItems()
    For i = 1 To mapFromCount
        For j = 1 To mapToCount
            If ReceiptItems(i).PO_Line = pageUnreceivedLines(j).PO_Line And ReceiptItems(i).PO_Schedule = pageUnreceivedLines(j).PO_Schedule Then
                ' If ITEM ID is specified, check to make sure the ITEM ID matches as well
                If ReceiptItems(i).Item_ID = "" Or ReceiptItems(i).Item_ID = pageUnreceivedLines(j).PO_Item_ID Then
                    map(i) = j
                End If
            End If
        Next j
    Next i
       
    PeopleSoft_Receipt_MapReceiptItemsToPageUnreceivedItems = map

End Function
'PeopleSoft_Receipt_MapReceiptLineToPageReceiptLines: Utility function to map ReceiptLines() to receipt lines on the peoplesoft page receipt lines
Private Function PeopleSoft_Receipt_MapReceiptItemsToPageReceiptLines(ReceiptItems() As PeopleSoft_Receipt_Item, ReceiptItemCount As Long, pageReceiptLines() As PeopleSoft_ReceiptPage_ReceiptLine, pageReceiptLineCount As Long) As Long()


    Dim mapFromCount As Long, mapToCount As Long
    
    mapFromCount = ReceiptItemCount
    mapToCount = pageReceiptLineCount
    
    Dim map() As Long
    ReDim map(1 To mapFromCount) As Long
    
    Dim i As Long, j As Long
    
    ' Initialize mapping where -1 means unresolved index
    For i = 1 To mapFromCount
        map(i) = -1 ' -1 means undefined
    Next i
    
     ' receive specified: map each row to the corresponding specific line/schedule in ReceiptItems()
    For i = 1 To mapFromCount
        For j = 1 To mapToCount
            If ReceiptItems(i).PO_Line = pageReceiptLines(j).PO_Line And ReceiptItems(i).PO_Schedule = pageReceiptLines(j).PO_Schedule Then
                ' If ITEM ID is specified, check to make sure the ITEM ID matches as well
                If ReceiptItems(i).Item_ID = "" Or ReceiptItems(i).Item_ID = pageReceiptLines(j).Item_ID Then
                    map(i) = j
                End If
            End If
        Next j
    Next i
       
    PeopleSoft_Receipt_MapReceiptItemsToPageReceiptLines = map

End Function
Public Function PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck(ByRef session As PeopleSoft_Session, ByRef poRetryBC As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams) As Boolean
    
    
    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown
    
    ' Perform pre-check
    If PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck_PreCheck(poRetryBC) = False Then
        GoTo RetryBCFailed
    End If
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As SeleniumWrapper.WebDriver
    
    
    PeopleSoft_Login session
    
    If Not session.loggedIn Then
        poRetryBC.GlobalError = "Logon Error: " & session.LogonError
        GoTo RetryBCFailed
    End If

    
    Set driver = session.driver
    
    
    If PeopleSoft_NavigateTo_ExistingPO(session, poRetryBC.PO_BU, poRetryBC.PO_ID) = False Then
        poRetryBC.GlobalError = "Failed to navigate to PO"
        GoTo RetryBCFailed
    End If
    
    
    
    
    ' Skip if PO is Dispatched or Approved.
    Dim poStatusText As String
    poStatusText = driver.findElementById("PSXLATITEM_XLATSHORTNAME").text
    
    If poStatusText = "Approved" Or poStatusText = "Dispatched" Then
        poRetryBC.GlobalError = "Skipped. PO Status is " & poStatusText  ' Not really an error
        PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = True
        Exit Function
    End If
    
        
    If DEBUG_OPTIONS.QuitBeforeSave Then
        poRetryBC.GlobalError = "Debug option QuitBeforeSave enabled. Processing halted before saving. To prevent this from occurring, disable QuitBeforeSave option."
        poRetryBC.HasGlobalError = True
        PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = False
        Exit Function
    End If
    
    
    Dim result As Boolean
    
    result = PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver, poRetryBC.BudgetCheck)
    
    If result = False Then
        poRetryBC.GlobalError = poRetryBC.BudgetCheck.GlobalError
        poRetryBC.HasGlobalError = poRetryBC.BudgetCheck.HasGlobalError
        
        GoTo RetryBCFailed
    End If
    
    
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = True
    Exit Function
    
    
ValidationFailed:
RetryBCFailed:
    poRetryBC.HasGlobalError = True
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = False
    Exit Function
       
ExceptionThrown:
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = False
    
    poRetryBC.HasGlobalError = True
    poRetryBC.GlobalError = "PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck Exception: " & Err.Description
    


End Function
Private Function PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck_PreCheck(ByRef poRetryBC As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams) As Boolean

    
    Debug_Print "PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck: PO (" & Debug_VarListString( _
        "PO BU", poRetryBC.PO_BU, _
        "PO ID", poRetryBC.PO_ID) _
        & ")"
    
    
    
    ' Begin - Check required fields
    If Len(poRetryBC.PO_BU) = 0 Then
        poRetryBC.PO_BU_Result.ValidationErrorText = "PO BU is required."
        poRetryBC.PO_BU_Result.ValidationFailed = True
        poRetryBC.HasGlobalError = True
    End If
    If Len(poRetryBC.PO_ID) = 0 Then
        poRetryBC.PO_ID_Result.ValidationErrorText = "PO ID is required."
        poRetryBC.PO_ID_Result.ValidationFailed = True
        poRetryBC.HasGlobalError = True
    End If
    
    If poRetryBC.HasGlobalError Then
        poRetryBC.GlobalError = "One or required fields are missing"
        GoTo PreCheckFailed
    End If
    ' End - Check required fields
    
    ' Begin - validate fields: ensure single-line fields are actually a single line
    If PeopleSoft_ValidateField_IsSingleLine("PO BU", poRetryBC.PO_BU, poRetryBC.PO_BU_Result) = False Then poRetryBC.HasGlobalError = True
    If PeopleSoft_ValidateField_IsSingleLine("PO ID", poRetryBC.PO_ID, poRetryBC.PO_ID_Result) = False Then poRetryBC.HasGlobalError = True
    

    If poRetryBC.HasGlobalError Then
        poRetryBC.GlobalError = "One or more fields failed pre-check"
        GoTo PreCheckFailed
    End If
    ' End - validate fields: ensure single-line fields are actually a single line
    
    
    
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck_PreCheck = True
    Exit Function
    
PreCheckFailed:
    poRetryBC.HasGlobalError = True
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck_PreCheck = False

    
    Debug_Print "PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck_PreCheck: failed! " & poRetryBC.GlobalError
    
End Function

Public Function PeopleSoft_Page_SetValidatedField(ByRef driver As SeleniumWrapper.WebDriver, ByVal fieldElementID As String, ByVal fieldValue As String, ByRef validationResult As PeopleSoft_Field_ValidationResult, Optional ignoreEmptyValues As Boolean = True, Optional expectedPopupContents As Variant) As Boolean

    
    Debug_Print "PeopleSoft_Page_SetValidatedField called (" & Debug_VarListString("fieldElementID", fieldElementID, "fieldValue", fieldValue) & ")"

    validationResult.ValidationFailed = False
    validationResult.ValidationErrorText = ""
    
    ' Dont bother if value is empty string and option to ignore empty values is true
    If Len(fieldValue) = 0 And ignoreEmptyValues Then
        PeopleSoft_Page_SetValidatedField = True
        Exit Function
    End If
    
    Dim elID As String, elVal As String, elDisabled As String
    
    elID = Replace$(fieldElementID, "'", "\'")
    
    elDisabled = driver.executeScript("return document.getElementById('" & elID & "').disabled;")
    
    If elDisabled <> "" Then
        If CBool(elDisabled) = True Then
            validationResult.ValidationFailed = True
            validationResult.ValidationErrorText = "Element is disabled"
            PeopleSoft_Page_SetValidatedField = True
            Exit Function
        End If
    End If
    
    elVal = driver.executeScript("return document.getElementById('" & elID & "').value;")
    

    If fieldValue <> elVal Then
        
        'tryNo = 1
        
    
        ' sanitize fieldValue
        fieldValue = Replace(fieldValue, "'", "\'") ' escape quuotes
        fieldValue = Replace(fieldValue, vbCrLf, "\n") ' replace new lines with newline character
        fieldValue = Replace(fieldValue, vbCr, "\n") ' replace new lines with newline character
        fieldValue = Replace(fieldValue, vbLf, "\n") ' replace new lines with newline character
        
        
        
        'fieldElement.Click
        'fieldElement.Clear
        driver.Wait 100
        'fieldElement.SendKeys fieldValue
        driver.runScript "javascript:document.getElementById('" & elID & "').value = '" & fieldValue & "';"
    
        
              
  
        ' Force field check
        driver.runScript "javascript:oChange_win0=document.getElementById('" & elID & "');addchg_win0(oChange_win0);submitAction_win0(oChange_win0.form,oChange_win0.name);"
        PeopleSoft_Page_WaitForProcessing driver
        
        
        
        Dim popupResult As PeopleSoft_Page_PopupCheckResult
        
        driver.setImplicitWait 100 ' new in 2.11: override implicit wait to 100ms (speeds up field entering)
        
        ' Wait for pop up: if popup not expected -> validation failed
        popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True, raiseErrorIfUnexpected:=False, expectedContent:=expectedPopupContents)
        PeopleSoft_Page_WaitForProcessing driver
        
        driver.setImplicitWait TIMEOUT_IMPLICIT_WAIT ' new in 2.11: restore implicit wait
        
        If popupResult.HasPopup And popupResult.IsExpected = False Then
            validationResult.ValidationErrorText = popupResult.popupText
            validationResult.ValidationFailed = True
            
            PeopleSoft_Page_SetValidatedField = False
            Exit Function
        End If
        
    
    End If


   
    PeopleSoft_Page_SetValidatedField = True
    Exit Function
    
    'PeopleSoft_Page_SetValidatedField = Not fieldValResult.ValidationFailed

End Function
' Utility function to create the PO data structure in one line. Must use PeopleSoft_PurchaseOrder_AddLine or AddLineSimple() to add lines
Public Function PeopleSoft_PurchaseOrder_NewPO(poBU As String, buyerID As Long, vendor As Variant, poReference As String) As PeopleSoft_PurchaseOrder

    Dim po As PeopleSoft_PurchaseOrder
    
    po.PO_Fields.PO_BUSINESS_UNIT = poBU
    po.PO_Fields.PO_HDR_BUYER_ID = buyerID
    po.PO_Fields.PO_HDR_PO_REF = poReference
    
    If IsNumeric(vendor) Then
        po.PO_Fields.PO_HDR_VENDOR_ID = CLng(vendor)
    Else
        po.PO_Fields.VENDOR_NAME_SHORT = vendor
    End If
    
    PeopleSoft_PurchaseOrder_NewPO = po

End Function
' Utility function to add a line to a PO structure. Returns line number. Must be followed by PeopleSoft_PurchaseOrder_NewSchedule() or schedule added manually.
Public Function PeopleSoft_PurchaseOrder_AddLine(ByRef purchaseOrder As PeopleSoft_PurchaseOrder, lineItemID As String, Optional lineItemDesc As String) As Integer
    
    purchaseOrder.PO_LineCount = purchaseOrder.PO_LineCount + 1

    ReDim Preserve purchaseOrder.PO_Lines(1 To purchaseOrder.PO_LineCount) As PeopleSoft_PurchaseOrder_Line
    
    With purchaseOrder.PO_Lines(purchaseOrder.PO_LineCount)
        .ScheduleCount = 0
    
        .LineFields.PO_LINE_ITEM_ID = lineItemID
        If IsMissing(lineItemDesc) = False Then .LineFields.PO_LINE_DESC = lineItemDesc
    End With
    
    PeopleSoft_PurchaseOrder_AddLine = purchaseOrder.PO_LineCount
    
End Function
' Utility function to add a schedule to a PO line structure. Returns schedule number
Public Function PeopleSoft_PurchaseOrder_AddLineSchedule(ByRef purchaseOrder As PeopleSoft_PurchaseOrder, poLineNbr As Integer, schQty As Currency, schDueDate As Date, schShipToId As Long, distPcBusinessUnit As String, distProjectCode As String, distActivityID As String, Optional distLocationID As Long = 0, Optional schPrice As Currency = 0) As Integer

    
    Debug.Assert poLineNbr > 0
    Debug.Assert poLineNbr <= purchaseOrder.PO_LineCount

    
    purchaseOrder.PO_Lines(poLineNbr).ScheduleCount = purchaseOrder.PO_Lines(poLineNbr).ScheduleCount + 1
    
    ReDim Preserve purchaseOrder.PO_Lines(poLineNbr).Schedules(1 To purchaseOrder.PO_Lines(poLineNbr).ScheduleCount) As PeopleSoft_PurchaseOrder_Schedule
    
    With purchaseOrder.PO_Lines(poLineNbr).Schedules(purchaseOrder.PO_Lines(poLineNbr).ScheduleCount)
        .ScheduleFields.DUE_DATE = schDueDate
        .ScheduleFields.SHIPTO_ID = schShipToId
        .ScheduleFields.QTY = schQty
        .ScheduleFields.PRICE = schPrice
        .DistributionFields.BUSINESS_UNIT_PC = distPcBusinessUnit
        .DistributionFields.PROJECT_CODE = distProjectCode
        .DistributionFields.ACTIVITY_ID = distActivityID
        .DistributionFields.LOCATION_ID = distLocationID
    End With
    
    PeopleSoft_PurchaseOrder_AddLineSchedule = purchaseOrder.PO_Lines(poLineNbr).ScheduleCount
    
End Function
' Utility function to add a line to a PO structure with a single schedule. Returns line number
Public Function PeopleSoft_PurchaseOrder_AddLineSimple(ByRef purchaseOrder As PeopleSoft_PurchaseOrder, lineItemID As String, lineItemDesc As String, schQty As Currency, schDueDate As Date, schShipToId As Long, distPcBusinessUnit As String, distProjectCode As String, distActivityID As String, Optional distLocationID As Long = 0, Optional schPrice As Currency = 0) As Integer

    
    Dim PO_LineCount As Integer
    
    
    PO_LineCount = purchaseOrder.PO_LineCount + 1

    ReDim Preserve purchaseOrder.PO_Lines(1 To PO_LineCount) As PeopleSoft_PurchaseOrder_Line
    
    ReDim purchaseOrder.PO_Lines(PO_LineCount).Schedules(1 To 1) As PeopleSoft_PurchaseOrder_Schedule
    
    purchaseOrder.PO_Lines(PO_LineCount).ScheduleCount = 1
    
    purchaseOrder.PO_Lines(PO_LineCount).LineFields.PO_LINE_ITEM_ID = lineItemID
    purchaseOrder.PO_Lines(PO_LineCount).LineFields.PO_LINE_DESC = lineItemDesc
    
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.DUE_DATE = schDueDate
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.SHIPTO_ID = schShipToId
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.QTY = schQty
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.PRICE = schPrice
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.BUSINESS_UNIT_PC = distPcBusinessUnit
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.PROJECT_CODE = distProjectCode
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.ACTIVITY_ID = distActivityID
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.LOCATION_ID = distLocationID
    
    purchaseOrder.PO_LineCount = PO_LineCount
    
    PeopleSoft_PurchaseOrder_AddLineSimple = PO_LineCount
    
End Function
Public Function PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver As SeleniumWrapper.WebDriver, ByRef budgetCheckParams As PeopleSoft_PurchaseOrder_BudgetCheckParams) As Boolean

    
    Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck called"

    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown


    ' ---------------------------------------------------------------------
    ' Begin - Save w/ Budget Check
    ' ---------------------------------------------------------------------
    
    Dim By As New SeleniumWrapper.By
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    
    
    
    Dim swByPOId As SeleniumWrapper.By
    Dim wePOId As SeleniumWrapper.WebElement
    
    Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: Clicking 'Save with Budget Check'"
    driver.findElementById("PO_KK_WRK_PB_BUDGET_CHECK").Click
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_INFINITE
    
    ' Acknowledge/Take action with popups
    Do
        popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False)
        If popupResult.HasPopup = False Then Exit Do
        
        If popupResult.popupText Like "*below PO line schedules exist with $0.00 or blank pricing*" Then
            ' Acknowledge Popup with message: The below PO line schedules exist with $0.00 or blank pricing.
            PeopleSoft_Page_AcknowledgePopup driver, popupResult, vbOK
            PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_INFINITE
        ElseIf popupResult.popupText Like "*Vendor * requires a Valid Contract*" Then
            ' Acknowlede popup with: Warning -- Vendor XXX requires a Valid Contract. Note: we will cancel the PO at this time.
            PeopleSoft_Page_AcknowledgePopup driver, popupResult, vbOK
            PeopleSoft_Page_WaitForProcessing driver
            
            budgetCheckParams.GlobalError = "Unexpected Popup: " & popupResult.popupText
                
            '  Acknowledge the popup and return the PO ID but with errors
            ' If PO ID provided, then grab PO ID
            If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[starts-with(@id,'PO_HDR_PO_ID')]")) Then
                Set wePOId = driver.findElementByXPath(".//*[starts-with(@id,'PO_HDR_PO_ID')]")
                Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: Found PO_HDR_PO_ID field: " & Debug_VarListString(wePOId.getAttribute("id"), wePOId.text)
                
                If Len(wePOId.text) > 0 And wePOId.text <> "NEXT" Then
                    budgetCheckParams.PO_ID = wePOId.text
                    budgetCheckParams.GlobalError = "PO ID '" & wePOId.text & "' generated with unexpected popup: " & popupResult.popupText
                End If
            End If
            
            GoTo BudgetCheckFatalError
        Else
            ' Unexpected popup
            budgetCheckParams.GlobalError = "Unexpected Popup: " & popupResult.popupText
            GoTo BudgetCheckFatalError
        End If
    Loop
    
    If budgetCheckParams.HasGlobalError Then GoTo BudgetCheckFatalError
    
    
    ' Begin - Deal with the new screen which asks about quantities in available excess
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//title[contains(text(),'Excess Available')]")) Then
        Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: Excess Available popup page found"
        
        driver.findElementById("Z_CAT_AVL_WRK_IGNORE_PB").Click
        driver.runScript "javascript: submitAction_win0(document.win0,'Z_CAT_AVL_WRK_IGNORE_PB');"
        PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_INFINITE
        
        
        popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False)
        
        If popupResult.HasPopup Then ' Error while saving
            budgetCheckParams.GlobalError = popupResult.popupText
            GoTo BudgetCheckFatalError
        End If
    End If
    ' End - Deal with the new screen which asks about quantities in available excess
    
    ' Begin - Deal with screen where there are price overages in PO
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//h1[@class='PSSRCHTITLE' and text()='PO Item Price Overages']")) Then
        Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: PO Item Price Overages page found"
        
        If PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage(driver, budgetCheckParams) = False Then
            ' Note: PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage will populate budgetCheckParams.GlobalError
            GoTo BudgetCheckFatalError
        End If
    End If
    ' End - Deal with screen where there are price overages in PO

    
    
    ' The PO ID will show up in one of two elements:
    '     Budget Check Failed -> Z_KK_ERR_WRK_PO_ID
    '     Budget Check Pass -> PO_HDR_PO_ID*
    '
    ' We will check for both. In some cases, neither is available immediately
    ' so we need try a few times or error out.
    Dim elementExists_PO_ID_budgetCheckFailed As Boolean  ' Case: budget check failed
    Dim elementExists_PO_ID_budgetCheckPassed As Boolean  ' Case: budget check passed
    Dim tryNo As Integer
    
    tryNo = 0
    
    Do
        tryNo = tryNo + 1
        
        elementExists_PO_ID_budgetCheckFailed = False
        elementExists_PO_ID_budgetCheckPassed = False
    
        elementExists_PO_ID_budgetCheckFailed = PeopleSoft_Page_ElementExists(driver, By.id("Z_KK_ERR_WRK_PO_ID"))
        If elementExists_PO_ID_budgetCheckFailed = False Then elementExists_PO_ID_budgetCheckPassed = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[starts-with(@id,'PO_HDR_PO_ID')]"))
        
        
        Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: PO ID element exists (" & Debug_VarListString("tryNo", tryNo, "elementExists_PO_ID_budgetCheckPassed", elementExists_PO_ID_budgetCheckPassed, "budgetCheckFailed", elementExists_PO_ID_budgetCheckFailed) & ")"
    Loop Until elementExists_PO_ID_budgetCheckFailed Or elementExists_PO_ID_budgetCheckPassed Or tryNo = 3
    
    If elementExists_PO_ID_budgetCheckFailed = False And elementExists_PO_ID_budgetCheckPassed = False Then
        budgetCheckParams.GlobalError = "Could not find PO ID on page: manual check required"
        budgetCheckParams.HasGlobalError = True
        
        GoTo BudgetCheckFatalError
    End If
    
    If elementExists_PO_ID_budgetCheckPassed Then
        ' Budget check passed
        Set wePOId = driver.findElementByXPath(".//*[starts-with(@id,'PO_HDR_PO_ID')]") 'driver.findElementByid("PO_HDR_PO_ID$14$")
        Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: Found header ID element (" & Debug_VarListString("elementID", wePOId.getAttribute("id"), "PO ID", wePOId.text) & ")"
        
        If wePOId.text = "NEXT" Then ' Error while saving
            budgetCheckParams.GlobalError = "Unknown error - Invalid PO ID Generated: " & wePOId.text
            budgetCheckParams.HasGlobalError = True
            
            GoTo BudgetCheckFatalError
        Else
            budgetCheckParams.PO_ID = wePOId.text
        End If
    Else
        ' Budget check failed
        Set wePOId = driver.findElementById("Z_KK_ERR_WRK_PO_ID")
        Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck: Found header ID element (" & Debug_VarListString("elementID", wePOId.getAttribute("id"), "PO ID", wePOId.text) & ")"
        
        budgetCheckParams.PO_ID = wePOId.text
        
        PeopleSoft_PurchaseOrder_BudgetCheck_ExtractFromPage driver, budgetCheckParams
    End If


    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = True
    Exit Function
    ' ---------------------------------------------------------------------
    ' End - Save w/ Budget Check
    ' ---------------------------------------------------------------------
    
    
BudgetCheckFatalError:
    budgetCheckParams.HasGlobalError = True
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
    Exit Function
    
ExceptionThrown:
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
    Err.Raise Err.Number, Err.Source, "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck Exception: " & Err.Description, Err.Helpfile, Err.HelpContext
    

End Function
Public Function PeopleSoft_PurchaseOrder_BudgetCheck_ExtractFromPage(driver As SeleniumWrapper.WebDriver, ByRef budgetCheckParams As PeopleSoft_PurchaseOrder_BudgetCheckParams) As Boolean

    Debug_Print "PeopleSoft_PurchaseOrder_BudgetCheck_ExtractFromPage called"
    Dim By As New SeleniumWrapper.By
    
    ' Click View All - by Line
    If PeopleSoft_Page_ElementExists(driver, By.id("Z_KK_PO_ERR_VW$hviewall$0")) Then
        'driver.findElementById("Z_KK_PO_ERR_VW$hviewall$0").Click
        driver.runScript "javascript:submitAction_win0(document.win0,'Z_KK_PO_ERR_VW$hviewall$0');"
        PeopleSoft_Page_WaitForProcessing driver
    End If
    
    ' Click View All - by Project
    If PeopleSoft_Page_ElementExists(driver, By.id("Z_KK_PRJ_ERR_VW$hviewall$0")) Then
        'driver.findElementById("Z_KK_PRJ_ERR_VW$hviewall$0").Click
        driver.runScript "javascript:submitAction_win0(document.win0,'Z_KK_PRJ_ERR_VW$hviewall$0');"
        PeopleSoft_Page_WaitForProcessing driver
    End If

  
    Dim PO_ErrorCount As Integer
    Dim PO_ErrorIndex As Integer
    
    Dim i As Integer
    
    
    budgetCheckParams.BudgetCheck_HasErrors = True
    
    ' Begin - Line Errors
    PO_ErrorCount = CInt(driver.getXpathCount(".//*[contains(@id,'trZ_KK_PO_ERR_VW$0_row')]"))
    
    budgetCheckParams.BudgetCheck_Errors.BudgetCheck_LineErrorCount = PO_ErrorCount
    
    
    ReDim budgetCheckParams.BudgetCheck_Errors.BudgetCheck_LineErrors(1 To PO_ErrorCount) As PeopleSoft_PurchaseOrder_BudgetCheck_LineError

    
    For i = 1 To PO_ErrorCount
        PO_ErrorIndex = i - 1
        
        With budgetCheckParams.BudgetCheck_Errors.BudgetCheck_LineErrors(i)
            .LINE_NBR = CInt(driver.findElementById("Z_KK_PO_ERR_VW_LINE_NBR$" & PO_ErrorIndex).text)
            .SCHED_NBR = CInt(driver.findElementById("Z_KK_PO_ERR_VW_SCHED_NBR$" & PO_ErrorIndex).text)
            .DISTRIB_LINE_NUM = CInt(driver.findElementById("Z_KK_PO_ERR_VW_DISTRIB_LINE_NUM$" & PO_ErrorIndex).text)
            .BUDGET_DT = driver.findElementById("Z_KK_PO_ERR_VW_BUDGET_DT$" & PO_ErrorIndex).text
            .BUSINESS_UNIT_PC = driver.findElementById("Z_KK_PO_ERR_VW_BUSINESS_UNIT_PC$" & PO_ErrorIndex).text
            .PROJECT_ID = driver.findElementById("Z_KK_PO_ERR_VW_PROJECT_ID$" & PO_ErrorIndex).text
            .LINE_AMOUNT = CurrencyFromString(driver.findElementById("Z_KK_PO_ERR_VW_MONETARY_AMOUNT$" & PO_ErrorIndex).text)
            .COMMIT_AMT = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_COMMIT_AMT$" & PO_ErrorIndex).text)
            .NOT_COMMIT_AMT = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_NOT_COMMIT_AMT$" & PO_ErrorIndex).text)
            .AVAIL_BUDGET_AMT = CurrencyFromString(driver.findElementById("Z_KK_PO_ERR_VW_Z_BUDGET_AMT$" & PO_ErrorIndex).text)
        End With
    Next i
    ' End - Line Errors
    
    ' Begin - Project Errors
    PO_ErrorCount = CInt(driver.getXpathCount(".//*[contains(@id,'trZ_KK_PRJ_ERR_VW$0_row')]"))
    
    budgetCheckParams.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount = PO_ErrorCount
    
    
    ReDim budgetCheckParams.BudgetCheck_Errors.BudgetCheck_ProjectErrors(1 To PO_ErrorCount) As PeopleSoft_PurchaseOrder_BudgetCheck_ProjectError

    
    ' Extract Project Budget Check Errors from field
    For i = 1 To PO_ErrorCount
        PO_ErrorIndex = i - 1
        
        With budgetCheckParams.BudgetCheck_Errors.BudgetCheck_ProjectErrors(i)
            .BUSINESS_UNIT_PC = driver.findElementById("Z_KK_PRJ_ERR_VW_BUSINESS_UNIT_PC$" & PO_ErrorIndex).text
            .PROJECT_ID = driver.findElementById("Z_KK_PRJ_ERR_VW_PROJECT_ID$" & PO_ErrorIndex).text
            .NOT_COMMIT_AMT = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_NOT_COMMIT_AMT2$" & PO_ErrorIndex).text)
            .AVAIL_BUDGET_AMT = CurrencyFromString(driver.findElementById("Z_KK_PRJ_ERR_VW_Z_BUDGET_AMT$" & PO_ErrorIndex).text)
            .FUNDING_NEEDED = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_KK_BAL_AMT$" & PO_ErrorIndex).text)
        End With
    Next i
    ' End - Project Errors
    
    PeopleSoft_PurchaseOrder_BudgetCheck_ExtractFromPage = True
    Exit Function


End Function
Public Function PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage(driver As SeleniumWrapper.WebDriver, budgetCheckParams As PeopleSoft_PurchaseOrder_BudgetCheckParams) As Boolean

    Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage called (" & Debug_VarListString("PRICE_OVERAGE_CODE", budgetCheckParams.PRICE_OVERAGE_CODE, "PRICE_OVERAGE_REASON", budgetCheckParams.PRICE_OVERAGE_REASON) & ")"

    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown
    
    
    Dim By As New SeleniumWrapper.By

    Dim rowCount As Integer, rowPageIndex As Integer
    Dim PO_Line As Integer, PO_Schedule As Integer
    
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    
    'If PeopleSoft_Page_ElementExists(driver, By.XPath(".//h1[@class='PSSRCHTITLE' and text()='PO Item Price Overages']")) = False Then
    If PeopleSoft_Page_ElementExists(driver, By.id("gbZ_PO_RSN_CD$0")) = False Then
        ' No table found -> nothing to do. Or should we error here?
        PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage = True
        Exit Function
    End If
    
    ' Check to see if a price overage code & reason are provided. If not, exit right away
    If Len(budgetCheckParams.PRICE_OVERAGE_CODE) = 0 Then 'Or Len(budgetCheckParams.PRICE_OVERAGE_REASON) = 0 Then
        budgetCheckParams.GlobalError = "The PO has price item overages from the contract price in PeopleSoft. A valid Price Overage Code and Reason must be provided for the PO."
        GoTo Failed
    End If
    
    
    rowCount = driver.getXpathCount(".//table[@id='gbZ_PO_RSN_CD$0']//tr[starts-with(@id,'trZ_PO_RSN_CD$0_row')]")
    

    ' apply price overage and reason to each item in the PO
    For rowPageIndex = 0 To rowCount - 1
        PO_Line = CInt(driver.findElementById("Z_PO_RSN_CD_LINE_NBR2$" & rowPageIndex).text)
        PO_Schedule = CInt(driver.findElementById("Z_PO_RSN_CD_SCHED_NBR2$" & rowPageIndex).text)
    
        PeopleSoft_Page_SetValidatedField driver, "Z_PO_RSN_CD_REASON_CD$" & rowPageIndex, budgetCheckParams.PRICE_OVERAGE_CODE, budgetCheckParams.PRICE_OVERAGE_CODE_Result
        
        If budgetCheckParams.PRICE_OVERAGE_CODE_Result.ValidationFailed Then
            budgetCheckParams.GlobalError = "PRICE OVERAGE CODE Validation Error: " & budgetCheckParams.PRICE_OVERAGE_CODE_Result.ValidationErrorText
            GoTo Failed
        End If
        
        ' If reason is not explictly given, then extract the reason description from the page and use it for the reason:
        If Len(budgetCheckParams.PRICE_OVERAGE_REASON) = 0 Then budgetCheckParams.PRICE_OVERAGE_REASON = driver.findElementById("PO_PNLS_WRK_DESCR50_MIXED$" & rowPageIndex).text
        
        driver.findElementById("Z_PO_RSN_CD_COMMENTS_254$" & rowPageIndex).SendKeys budgetCheckParams.PRICE_OVERAGE_REASON
    Next rowPageIndex
    
    'driver.findElementById("PO_PNLS_WRK_PB_OK$0").Click
    driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_PB_OK$0');"
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_INFINITE
    
    
    ' Check for popup
    ' Typical (all result in error):
    ' 1) Reason Code and Comments must be entered for all lines
    ' 2) PO item/kit price entered exceeds vendor contracted price. An attachment must be added to proceed. (23200,512)
    popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False)
    
    If popupResult.HasPopup Then
        budgetCheckParams.GlobalError = "PRICE OVERAGE Page Popup: " & popupResult.popupText
        GoTo Failed
    End If
    

    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage = True
    Exit Function
    
Failed:
    budgetCheckParams.HasGlobalError = True
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage = False

    Debug_Print "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage failed! " & budgetCheckParams.GlobalError
    Exit Function
    
ExceptionThrown:
    budgetCheckParams.HasGlobalError = True
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage = False
    Err.Raise Err.Number, Err.Source, "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck_FillPriceOveragePage Exception: " & Err.Description, Err.Helpfile, Err.HelpContext
    

End Function

Public Function PeopleSoft_Page_ElementExists(driver As SeleniumWrapper.WebDriver, weBy As SeleniumWrapper.By, Optional timeoutms As Long) As Boolean

    On Error GoTo ElementNotFoundOrError:

    Dim we As SeleniumWrapper.WebElement
    
    
    Set we = driver.findElement(weBy, timeoutms)
    
    If Not we Is Nothing Then
        PeopleSoft_Page_ElementExists = True
        Exit Function
    End If
    
    
ElementNotFoundOrError:

    PeopleSoft_Page_ElementExists = False
    

End Function
Private Function PeopleSoft_Page_GetElementText(driver As SeleniumWrapper.WebDriver, ByVal elementID As String, Optional ByVal default As Variant) As Variant

    Dim text As Variant
    
    elementID = Replace(elementID, "'", "\'")

    text = driver.executeScript("return document.getElementById('" & elementID & "').textContent;")
    
    PeopleSoft_Page_GetElementText = text
    
    If Not IsMissing(default) Then
        If Not text <> "" Then
            PeopleSoft_Page_GetElementText = default
        End If
    End If
    

End Function
Private Function PeopleSoft_Page_GetInputElementValue(driver As SeleniumWrapper.WebDriver, ByVal elementID As String, Optional ByVal default As Variant) As Variant

    Dim text As Variant
    
    elementID = Replace(elementID, "'", "\'")

    text = driver.executeScript("return document.getElementById('" & elementID & "').value;")
    
    PeopleSoft_Page_GetInputElementValue = text
    
    If Not IsMissing(default) Then
        If Not text <> "" Then
            PeopleSoft_Page_GetInputElementValue = default
        End If
    End If

End Function

Public Sub PeopleSoft_Page_WaitForProcessing(driver As SeleniumWrapper.WebDriver, Optional ByVal timeout_s As Long = 60, Optional waitForLoader As Boolean = False)

    
    Const POLL_INTERVAL_MS As Double = 500 ' Poll every 0.5 s
    
    Dim iter As Integer
    Dim loader_inProcess As Boolean, proc_visibility As Variant
    Dim loader_exists As Boolean
    Dim infiniteTimeout As Boolean
    
    infiniteTimeout = False
    
    If timeout_s = TIMEOUT_INFINITE Then
        infiniteTimeout = True ' No timeout
    ElseIf timeout_s = 0 Then
        timeout_s = 60  'Default to 60s if timeout not given
    End If
    
    Debug.Assert timeout_s >= -1
    
    
    Dim MAX_ITER As Long
    
    If infiniteTimeout = False Then MAX_ITER = timeout_s * 1000 / POLL_INTERVAL_MS
    iter = 0
    
    
    ' If waitForLoader is set -> wait for page loader to exist before the next step
    If waitForLoader Then
        Do
            loader_exists = driver.executeScript("return (loader != null);")
            If loader_exists Then Exit Do
            
            driver.Wait POLL_INTERVAL_MS
        
            DoEvents
            iter = iter + 1
        Loop Until (iter > MAX_ITER And infiniteTimeout = False) Or loader_exists
        
    
        If infiniteTimeout = False And iter > MAX_ITER Then
            Err.Raise 513, , "PeopleSoft_Page_WaitForProcessing: Timeout during waitForLoader"
        End If
    End If
    
    ' Processing is over when two actions happen (for good measure, both must occur):
    '   (1) The processing icon is no longer visible
    '   (2) When the PeopleSoft internal loader is no longer active and processing
    '
    Do
    
        loader_inProcess = driver.executeScript("return (loader != null && loader.GetInProcess());")
        proc_visibility = driver.executeScript("return document.getElementById('WAIT_win0').style.visibility;")
         
        driver.Wait POLL_INTERVAL_MS
        
        DoEvents
    
        iter = iter + 1
    Loop Until (iter > MAX_ITER And infiniteTimeout = False) Or (proc_visibility <> "visible" And loader_inProcess = False)
    

    If infiniteTimeout = False And iter > MAX_ITER Then
        Err.Raise 513, , "PeopleSoft_Page_WaitForProcessing: Timeout"
    End If
    

End Sub
Public Function PeopleSoft_Page_CheckForModal(driver As SeleniumWrapper.WebDriver) As Integer
    ' Returns index # of modal if found (starts at 0). Returns -1 if not found, -2 if error
    
    
    On Error GoTo NotFoundOrErr
    
    Dim wePopupModals As WebElementCollection
    
    
    Set wePopupModals = driver.findElementsByXPath(".//*[starts-with(@id,'ptMod_')]", 100)
    
    If wePopupModals.Count = 0 Then
        Debug_Print "PeopleSoft_Page_CheckForModal: modal window not found"
        PeopleSoft_Page_CheckForModal = -1
        Exit Function
    End If
    
    
    Dim elemID As String, modalIndexStr As String
    
    elemID = wePopupModals(0).getAttribute("id")
    modalIndexStr = Right$(elemID, Len(elemID) - Len("ptMod_"))
    
    If Not IsNumeric(modalIndexStr) Then
        Debug_Print "PeopleSoft_Page_CheckForModal: error occurred when parsing modal element ID '" & elemID & "'"
        PeopleSoft_Page_CheckForModal = -2
        Exit Function
    End If
    
    PeopleSoft_Page_CheckForModal = CLng(modalIndexStr)
    
    Debug_Print "PeopleSoft_Page_CheckForModal: modal window found (index=" & modalIndexStr & ")"
        
    Exit Function
    
NotFoundOrErr:
    Debug_Print "PeopleSoft_Page_CheckForModal: error occurred"
    PeopleSoft_Page_CheckForModal = -2
    

End Function
Public Function PeopleSoft_Page_CheckForPopup(driver As SeleniumWrapper.WebDriver, Optional acknowledgePopup As Boolean = False, Optional raiseErrorIfUnexpected As Boolean = True, Optional expectedContent As Variant) As PeopleSoft_Page_PopupCheckResult

    
    Debug_Print "PeopleSoft_Page_CheckForPopup called (" & Debug_VarListString("acknowledgePopup", acknowledgePopup, "raiseErrorIfUnexpected", raiseErrorIfUnexpected, "expectedContent (is provided)", Not IsMissing(expectedContent)) & ")"


    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo PopupNotFoundOrErr
    
    
    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult
    
    
    popupCheckResult.HasPopup = False


    Dim we As SeleniumWrapper.WebElement, By As New SeleniumWrapper.By
    Dim wePopupModals As WebElementCollection
    
    
    Set wePopupModals = driver.findElementsByXPath(".//*[contains(@id,'ptModContent_')]", 100)
    
    'no popup modals found?
    If wePopupModals.Count = 0 Then
        Debug.Print "PeopleSoft_Page_CheckForPopup: No popup found"
        
        PeopleSoft_Page_CheckForPopup = popupCheckResult
        Exit Function
    End If
    
    popupCheckResult.HasPopup = True
    
    popupCheckResult.PopupElementID = wePopupModals(0).getAttribute("id")
    
    
    
    
    ' get buttons visible on alert: slow method (if item doesn't exist, then it hangs)
    'popupCheckResult.HasButtonOk = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICOK']"), 10)
    'popupCheckResult.HasButtonCancel = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICCancel']"), 10)
    'popupCheckResult.HasButtonYes = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICYes']"), 10)
    'popupCheckResult.HasButtonNo = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICNo']"), 10)
    
    ' get buttons visible on alert: fast method
    popupCheckResult.HasButtonOk = driver.executeScript("javascript: return document.getElementById('#ICOK') != null;")
    popupCheckResult.HasButtonCancel = driver.executeScript("javascript: return document.getElementById('#ICCancel') != null;")
    popupCheckResult.HasButtonYes = driver.executeScript("javascript: return document.getElementById('#ICYes') != null;")
    popupCheckResult.HasButtonNo = driver.executeScript("javascript: return document.getElementById('#ICNo') != null;")
    
    
    ' Get alert text
    Set we = driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='alertmsg']/span")
    popupCheckResult.popupText = we.text
    
    
    
    ' Check to see if the popup text matches any the patterns in expectContent - allow for either array or string
    Dim expectedPatterns() As Variant, expectedPattern As Variant
    Dim expectedDebugStr As String, i As Long
    
    expectedDebugStr = "NULL"
     
    ' Compare popup text with any of the strings listed in expectedContent() to determine if popup is expected
    If Not IsMissing(expectedContent) Then
        ' Ensure expectedContents is an array
        If IsArray(expectedContent) Then
            expectedPatterns = expectedContent          ' Array with N items
        Else
            expectedPatterns = Array(expectedContent)   ' Array with 1 item
        End If
        
        expectedDebugStr = "'" & Join(expectedPatterns, "','" & "'")
        
        For Each expectedPattern In expectedPatterns
            If popupCheckResult.popupText Like CStr(expectedPattern) Then
                popupCheckResult.IsExpected = True
                Exit For
            End If
        Next
    End If
    
    
    Dim sanitizedText As String, buttonText As String
    
    ' Replace new lines with "\n"
    sanitizedText = popupCheckResult.popupText
    sanitizedText = Replace(sanitizedText, vbCrLf, "\n")
    sanitizedText = Replace(sanitizedText, vbCr, "\n")
    sanitizedText = Replace(sanitizedText, vbLf, "\n")
    
    ' Create ButtonText
    buttonText = Join(Array(IIf(popupCheckResult.HasButtonYes, "Yes", ""), IIf(popupCheckResult.HasButtonNo, "No", ""), IIf(popupCheckResult.HasButtonOk, "OK", ""), IIf(popupCheckResult.HasButtonCancel, "Cancel", "")), "|")
        
    
    Debug_Print "PeopleSoft_Page_CheckForPopup: ID='" & popupCheckResult.PopupElementID & "', Expected=" & popupCheckResult.IsExpected & ", " _
                & "Buttons=(" & buttonText & "), " _
                & "Text='" & sanitizedText & "', ExpectedContents=" & expectedDebugStr & ""
    
    If raiseErrorIfUnexpected And Not IsMissing(expectedContent) And popupCheckResult.IsExpected = False Then
        On Error GoTo 0
        Err.Raise -1, , "Unexpected Popup: " & popupCheckResult.popupText
        On Error GoTo PopupNotFoundOrErr
    End If

    ' Acknowledge if requested
    If acknowledgePopup Then
        If popupCheckResult.HasButtonOk Then
            PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbOK
        ElseIf popupCheckResult.HasButtonYes Then
            PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbYes
        ElseIf popupCheckResult.HasButtonCancel Then
            PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbCancel
        Else
            PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbNo
        End If
    End If
                
    
    PeopleSoft_Page_CheckForPopup = popupCheckResult
    
    
    Exit Function
    
PopupNotFoundOrErr:
    popupCheckResult.HasPopup = False
    popupCheckResult.PopupElementID = ""
    popupCheckResult.popupText = ""
    
    PeopleSoft_Page_CheckForPopup = popupCheckResult
    
    Debug_Print "PeopleSoft_Page_CheckForPopup: No popup found or error: " & Err.Description

End Function
Public Sub PeopleSoft_Page_AcknowledgePopup(driver As SeleniumWrapper.WebDriver, ByRef popupCheckResult As PeopleSoft_Page_PopupCheckResult, clickButton As VbMsgBoxResult)
    
    
    Debug_Print "PeopleSoft_Page_AcknowledgePopup called"
    
    
    If DEBUG_OPTIONS.AddMethodNamePrefixToExceptions Then On Error GoTo ExceptionThrown
    
    If clickButton = vbOK Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICOK']").Click
    ElseIf clickButton = vbCancel Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICCancel']").Click
    ElseIf clickButton = vbYes Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICYes']").Click
    ElseIf clickButton = vbNo Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICNo']").Click
    End If
    
    'PeopleSoft_Page_WaitForProcessing driver, timeout_s
    
    Exit Sub
    
ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_Page_AcknowledgePopup: " & Err.Description, Err.Helpfile, Err.HelpContext

End Sub

Private Function CurrencyFromString(strCur As String) As Currency

    strCur = Replace(strCur, ",", "")
    
    If IsNumeric(strCur) Then
        CurrencyFromString = CCur(strCur)
    Else
        CurrencyFromString = 0
    End If

End Function



